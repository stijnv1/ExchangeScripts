#
# RemoveMailForwardO365Mailbox.ps1
#
param
(
	[Parameter(Mandatory=$true)]
    [string]$LogPath,

	[Parameter(Mandatory=$true)]
    [string]$ADGroupToBeProcessed
)

Function WriteToLog
{
	param
	(
		[string]$LogPath,
		[string]$TextValue,
		[bool]$WriteError
	)

	Try
	{
		#create log file name
		$thisDate = (Get-Date -DisplayHint Date).ToLongDateString()
		$LogFileName = "DisableForwardSMTP_$thisDate.log"

		#write content to log file
		if ($WriteError)
		{
			Add-Content -Value "[ERROR $(Get-Date -DisplayHint Time)] $TextValue" -Path "$LogPath\$LogFileName"
		}
		else
		{
			Add-Content -Value "[INFO $(Get-Date -DisplayHint Time)] $TextValue" -Path "$LogPath\$LogFileName"
		}
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		Write-Host "Error occured in WriteToLog function: $ErrorMessage" -ForegroundColor Red
	}

}

Try
{
	#create credentials for O365 remote powershell connection
	$UserCredential = Get-Credential

	#connect to Office 365 Exchange via remote powershell
	$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $O365Session

	#get all users for which the forward smtp address must be removed
	$adusers = Get-ADGroupMember $adgroup | Get-ADUser

	foreach ($aduser in $adusers)
	{
		Try
		{
			#set forward smtp address empty
			Set-Mailbox $aduser.UserPrincipalName -ForwardingSmtpAddress ""
		}
		Catch
		{
			$ErrorMessage = $_.Exception.Message
            WriteToLog -LogPath $LogPath -TextValue "Error occured during processing of user $($aduser.Name). The exact error message = $ErrorMessage" -WriteError $true
            $ErrorActionPreference = "stop"
            Remove-PSSession $O365Session 
		}
	}
}

Catch
{
	$ErrorMessage = $_.Exception.Message
    WriteToLog -LogPath $LogPath -TextValue "Error occured during execution of the script. Error message = $ErrorMessage" -WriteError $true
    Remove-PSSession $O365Session
}


