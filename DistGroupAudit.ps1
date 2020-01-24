# Convert secure string back to plain text
function Get-PlainText()
{
	[CmdletBinding()]
	param
	(
		[parameter(Mandatory = $true)]
		[System.Security.SecureString]$SecureString
	)
	BEGIN { }
	PROCESS
	{
		$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString);

		try
		{
			return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr);
		}
		finally
		{
			[Runtime.InteropServices.Marshal]::FreeBSTR($bstr);
		}
	}
	END { }
}

# Used to retrieve an access token
function Get-AccessToken{
    param(
        [Parameter(ParameterSetName='PlainText',Mandatory=$true)]
        [string]$PlainTextSecret,
        [Parameter(ParameterSetName='SecureString',Mandatory=$true)]
        [string]$Secret,
        [Parameter(Mandatory=$true)]
        [string]$ClientID,
        [Parameter(Mandatory=$true)]
        [string]$TenantName
    )
    <#
    .DESCRIPTION
    Used to retrieve a Token from a registered MSGraphAPI App, to be used to make calls to the MSGraphAPI
    Use $Secret to create an secure string to be used with Get-AccessToken.
    $Secret = 'PlainTextSecretGoesHere'| ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    This string is bound to the System and User that creates it.  It's not perfect but it works pretty well.
    It can be stored in a text file for later use/programatic access under a service account.
    Pass $Secret in for the value of the -Secret, if you intend to paste in the plain text value use -PlainTextSecret

    .EXAMPLE
    C:\PS>$Global:Token = Get-AccessToken -Secret 'encrypted_secret_goes_here' -ClientID 'graph_api_client_id_goes_here' -TenantName 'tenant_id_goes_here'

    Returns a token after providing an application secret using the SecureString method

    .EXAMPLE
    C:\PS>$Global:Token = Get-AccessToken -PlainTextSecret 'AppSecret_PlainText' -ClientID '00000000-1111-2222-3333-444444444444' -TenantName '00000000-1111-2222-3333-444444444444'

    Returns a token after providing an application secret in plain text.

    #>
    switch ($PSCmdlet.ParameterSetName) {
        'PlainText' {
            $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $clientID
            Client_Secret = $PlainTextSecret
            } 
            $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
            Return $TokenResponse.access_token
        }
        'SecureString'{
            $clientSecret = Get-PlainText -SecureString ($secret | ConvertTo-SecureString)
            $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $clientID
            Client_Secret = $clientSecret
            } 
            $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
            Return $TokenResponse.access_token
        }
    }
}

function Get-MSGraphEmail{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Token,
        [Parameter(Mandatory=$true)]
        [string]$UserObjectID,
        [Parameter(Mandatory=$true)]
        [string]$RecipientAddress,
        [Parameter(Mandatory=$true)]
        [string]$StartDate
    )
    <#
    .DESCRIPTION
    By default this will return 10 items matching the RecipientAddress value from the specified mailbox (UserObjectID). It is possible to recurse by using the .'@odata.nextLink' to get next page of results, this isn't currently implemented for speed reasons.

    .EXAMPLE
    c:\PS>$MatchingEmails = Get-MSGraphEmail -Token 'Token' -UserObjectID 'ObjectID' -StartDate '2020-01-22' -RecipientAddress example@example.com
    #>
    
    

    $Filters = '$search=' +'"'+$RecipientAddress+'"'
        
    $Request = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$UserObjectID/messages?$filters"
        Headers = @{Authorization = "Bearer $Token"}
        Method = 'Get'
        ErrorAction = 'Stop'
        ErrorVariable = 'ErrGetEmail'
    }

    $Data = Invoke-RestMethod @Request
    $Emails = ($Data | Select-Object Value).Value
    # Filter out return values since I can't sort out how to search for emails actually sent to the specified address.
    # Only return messages after the specified cutoff date.
    Return $Emails | Where-Object{$_.toRecipients.emailAddress.address -eq $RecipientAddress -and $_.receivedDateTime -ge $StartDate}
}

# Actual processing done below this line, make sure to choose your method of getting tokens, see the comments in the function on creating a string to pass in for the -secret option

# Connect to exchange online shell and get DistributionGroups and their members
# The exchange online shell module is required.

Connect-ExchangeOnlineShell

# Create an empty list to add groups to
$DistroGroups  = New-Object 'System.Collections.Generic.List[System.Object]'

# Filter out security enabled groups from Hybrid environment, you cannot convert them anyway.
Get-DistributionGroup -ResultSize Unlimited | Where-Object{!($_.GroupType -like "*Security*")} | ForEach-Object{$DistroGroups.Add($_)}

# Get an access token
# Pick a method and uncomment it

# $Global:Token = Get-AccessToken -Secret 'encrypted_secret_goes_here' -ClientID 'graph_api_client_id_goes_here' -TenantName 'tenant_id_goes_here'
# Alternate method using a plaintext graphapi key, do not recommend storing this value in a script, only use this method ad-hoc, copy/paste into a powershell session.
# $Global:Token = Get-AccessToken -PlainTextSecret 'plain_text_secret' -ClientID 'graph_api_client_id_goes_here' -TenantName 'tenant_id_goes_here'

$TokenTimer = New-Object system.diagnostics.stopwatch
$TokenTimer.start()
$ListActivity = New-Object 'System.Collections.Generic.List[System.Object]'

$i = 0
$TotalGroups = ($DistroGroups | Measure-Object).count
Foreach($DistGroup in $DistroGroups){
    $i++
    Write-host  "On Group $i of $TotalGroups." -ForegroundColor White
    $GroupMembers = New-Object 'System.Collections.Generic.List[System.Object]'
    $Temp = Get-DistributionGroupMember $DistGroup.Name
    $Temp | Where-Object{$_.RecipientType -eq "UserMailbox"} | ForEach-Object{
        $TempUserObject = [pscustomobject]@{
            InList = $DistGroup.PrimarySmtpAddress
            MemberID = $_.Identity
            ObjectID = $_.ExternalDirectoryObjectId
        }
        $GroupMembers.Add($TempuserObject)
    }
    $SubGroups = $Temp | Where-Object{$_.RecipientType -like "*Group*"}

    if($SubGroups){
        Write-Host "$(($Subgroups | Measure-Object).count) nested groups found in $($Distgroup.name)" -ForegroundColor Yellow
        $TempSubGroups = New-Object 'System.Collections.Generic.List[System.Object]'
        do{
            Foreach($SubGroup in $SubGroups){
                Write-Host "Getting members of $($SubGroup.Name)" -ForegroundColor Green
                $Temp = Get-DistributionGroupMember $SubGroup.Name
                $Temp | Where-Object{$_.RecipientType -eq "UserMailbox"} | ForEach-Object{
                    $TempUserObject = [pscustomobject]@{
                        InList = $DistGroup.PrimarySmtpAddress
                        MemberID = $_.Identity
                        ObjectID = $_.ExternalDirectoryObjectId
                    }
                    $GroupMembers.Add($TempuserObject)                    
                }
                $Temp | Where-Object{$_.RecipientType -like "*Group*" -and !($_.RecipientType -like "*Security*")}|ForEach-Object{$TempSubGroups.Add($_)}
            }
            $SubGroups = $null
            $SubGroups = $TempSubGroups
        }
        While($SubGroups)
    }
    # Check 10% of the recipients in the list for email to reduce processing time.
    $MemberCount = ($GroupMembers | Measure-Object).Count
    $EndIndex = [math]::round($MemberCount * .10)

    foreach($Member in $GroupMembers[0..$EndIndex]){
        $MatchingEmails = 0
        # Check that the token does not need to be renewed
        if($TokenTimer.Elapsed.Minutes -ge 50 -or $TokenTimer.Elapsed.Hours -ge 1){
            # Pick a method and uncomment it

            # $Global:Token = Get-AccessToken -Secret 'encrypted_secret_goes_here' -ClientID 'graph_api_client_id_goes_here' -TenantName 'tenant_id_goes_here'
            # Alternate method using a plaintext graphapi key, do not recommend storing this value in a script, only use this method ad-hoc, copy/paste into a powershell session.
            # $Global:Token = Get-AccessToken -PlainTextSecret 'plain_text_secret' -ClientID 'graph_api_client_id_goes_here' -TenantName 'tenant_id_goes_here'
        }

        # Set StartDate to pull emails from that date forward, limited by the returned results of Get-MSGraphEmail as currently written
        $MatchingEmails += (Get-MSGraphEmail -Token $Token -UserObjectID $Member.ObjectID -StartDate '2020-01-22' -RecipientAddress $Member.InList | Measure-Object).count
    }
    $TempResultsObject = [pscustomobject]@{
        List = $DistGroup.PrimarySmtpAddress
        MessagesReceived = $MatchingEmails        
    }
    $ListActivity.Add($TempResultsObject)
    Clear-Host
}

$ListActivity | Export-csv c:\temp\ListActivity.csv -NoTypeInformation

