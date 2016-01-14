#
# CreateO365MailboxesWithForwardingMail.ps1
#

param
(
    [string]$LogPath,
    [string]$RemoteRoutingMailDomain,
    [string]$AADServerName,
    [string]$emailAddressPolicyName,
	[string]$ExchangeServerName,
	[string]$GroupsOUDistinguishedName,
	[string]$O365LicenseName
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
		$LogFileName = "CreateADUser_$thisDate.log"

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
    #create connection to O365
    $UserCredential = Get-Credential

    #create remote powershell session to onpremise Exchange 2013
    $OnPremExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServerName/PowerShell/ -Authentication Kerberos

    #start remote powershell session to onpremise Exchange 2013 server
    Import-PSSession $OnPremExchangeSession

    #connect to Azure AD
    Connect-MsolService -Credential $UserCredential

    #$ADGroups = Get-ADGroup -SearchBase $GroupsOUDistinguishedName
    $ADGroups = Get-ADGroup "GG-S-MailForward-odice"

    foreach ($adgroup in $ADGroups)
    {
        WriteToLog -LogPath $LogPath -TextValue "Start Processing AD Group $($adgroup.Name)" -WriteError $false
        $adusers = Get-ADGroupMember $adgroup | Get-ADUser -Properties streetAddress
        
        foreach ($aduser in $adusers)
        {
            Try
            {
                $mailAlias = ($aduser.UserPrincipalName).SubString(0,$aduser.UserPrincipalName.IndexOf("@"))

                #check whether remote mailbox already exists for this user. If exists, no further processing is needed for this user
                if (Get-RemoteMailbox $mailAlias -ErrorAction silentlycontinue)
                {
                    Write-Host "[SKIP]: User $mailAlias already has an Office 365 mailbox, user is skipped ..." -ForegroundColor Yellow
                    WriteToLog -LogPath $LogPath -TextValue "User $mailAlias already has a mailbox in Office 365, user is skipped" -WriteError $false
                }
                else
                {
                    #enable remote mailbox
					Write-Host "[CREATE]: Create Office 365 mailbox for user $($aduser.Name) with mailalias $mailAlias ..." -ForegroundColor Yellow 
                    WriteToLog -LogPath $LogPath -TextValue "Create Office 365 mailbox for user $($aduser.Name) with mailalias $mailAlias"
                    Enable-RemoteMailbox $aduser.UserPrincipalName -RemoteRoutingAddress "$mailAlias@$RemoteRoutingMailDomain" | Out-Null

                    #assign O365 license
                    Set-MsolUser -UserPrincipalName $aduser.UserPrincipalName -UsageLocation BE
                    Set-MsolUserLicense -UserPrincipalName $aduser.UserPrincipalName -AddLicenses $O365LicenseName 

                    #run the e-mail address policy update
                    Get-EmailAddressPolicy | ? Name -Like "*$emailAddressPolicyName*" | Update-EmailAddressPolicy
                }

                #activate code below if something went wrong during license provisioning
                #if ((Get-MsolUser -UserPrincipalName $aduser.UserPrincipalName).isLicensed)
                #{
                #    Write-Host "User already licensed"
                #}
                #else
                #{
                #    Write-Host "Set license"
                #    Set-MsolUserLicense -UserPrincipalName $aduser.UserPrincipalName -AddLicenses $O365LicenseName
                #}
            }
            Catch
            {
                $ErrorMessage = $_.Exception.Message
                WriteToLog -LogPath $LogPath -TextValue "Error occured during processing of user $($aduser.Name). The exact error message = $ErrorMessage" -WriteError $true
                $ErrorActionPreference = "stop"
                Remove-PSSession $OnPremExchangeSession   
            }
        }

        #start Azure AD sync
        WriteToLog -LogPath $LogPath -TextValue "Start Azure AD sync on AAD server $AADServerName ..." -WriteError $false
        Invoke-Command -ComputerName $AADServerName -ScriptBlock {& cmd.exe /c "D:\Program Files\Microsoft Azure AD Sync\Bin\DirectorySyncClientCmd.exe"}

        #close onpremise exchange powershell session
        WriteToLog -LogPath $LogPath -TextValue "Closing on-premise Exchange remote powershell session..." -WriteError $false
        Remove-PSSession $OnPremExchangeSession

        #start O365 remote powershell session
        $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        Import-PSSession $O365Session

        foreach ($aduser in $adusers)
        {
            Try
            {
                $mailAlias = ($aduser.UserPrincipalName).SubString(0,$aduser.UserPrincipalName.IndexOf("@"))

                #script has to wait for the mailbox to be created in Office 365
                $mailboxCreated = $false
                do
                {
                    if (Get-Mailbox $mailAlias -ErrorAction SilentlyContinue)
                    {
                        Write-Host "Mailbox of user $mailAlias is created in Office 365, continue script ..." -ForegroundColor Yellow
                        $mailboxCreated = $true                     
                    }
                    else
                    {
                        Write-Host "Mailbox of user $mailAlias is not created yet in Office 365, start sleep for 1 minute ..." -ForegroundColor Yellow
                        Start-Sleep -Seconds 60
                    }
                }
                while (!$mailboxCreated)


                $prefixAddress = $aduser.streetAddress

                #check whether forwarding address is already configured, if configured, skip user
                if ((Get-Mailbox $mailAlias).ForwardingSmtpAddress)
                {
                    Write-Host "[SKIP]: User $mailAlias already has a forwarding SMTP address configured, skip user ..." -ForegroundColor Green
                    WriteToLog -LogPath $LogPath -TextValue "User $mailAlias already has a forwarding SMTP address configured, skip user"
                }
                else
                {
                    Write-Host "[CREATE]: User $($aduser.Name) has no forwarding SMTP address configured yet, start creating forwarding SMT Address ..." -ForegroundColor Green
                    WriteToLog -LogPath $LogPath -TextValue "Set forwarding mail address for user $($aduser.Name) to $mailAlias@$prefixAddress" -WriteError $false
                    Set-Mailbox $aduser.UserPrincipalName -ForwardingSmtpAddress "$mailAlias@$prefixAddress"
                }
            }

            catch
            {
                $ErrorMessage = $_.Exception.Message
                WriteToLog -LogPath $LogPath -TextValue "Error occured during processing of user $($aduser.Name). The exact error message = $ErrorMessage" -WriteError $true
                $ErrorActionPreference = "stop"
                Remove-PSSession $O365Session 
            }

            
        }

        WriteToLog -LogPath $LogPath -TextValue "Closing Office 365 remote powershell session..." -WriteError $false
        Remove-PSSession $O365Session
    }

}
Catch
{
    #Remove-PSSession $Session
    $ErrorMessage = $_.Exception.Message
    WriteToLog -LogPath $LogPath -TextValue "Error occured during execution of the script. Error message = $ErrorMessage" -WriteError $true
    Remove-PSSession $O365Session
    Remove-PSSession $OnPremExchangeSession
}
Finally
{
    #Write-Host "Closing remote powershell session to Office 365"
    #Remove-PSSession $Session
}