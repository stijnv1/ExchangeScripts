$mailboxes = Get-Mailbox
$exportUserArray = @()
$Member = @{
		MemberType = "NoteProperty"
		Force = $true
	}

#foreach ($mailbox in $mailboxes)
#{
#    if ($mailbox.RecipientTypeDetails -ne "UserMailbox")
#    {
#        $adUser = Get-ADUser -Filter {UserPrincipalName -eq $mailbox.UserPrincipalName}

#        if ($adUser.GivenName -eq $null)
#        {
#            Write-Host "Updating user $($adUser.UserPrincipalName)" -ForegroundColor Yellow
#            $aduser | Set-ADUser -GivenName $mailbox.Alias -Surname $mailbox.RecipientTypeDetails
#        }
#    }
#}

foreach ($mailbox in $mailboxes)
{
    $adUser = Get-ADUser -Filter {UserPrincipalName -eq $mailbox.UserPrincipalName}

    $userObject = New-Object psobject
    
    $userObject | Add-Member @Member -Name "GivenName" -Value $adUser.GivenName
    $userObject | Add-Member @Member -Name "Surname" -Value $adUser.Surname
    $userObject | Add-Member @Member -Name "MailboxType" -Value $mailbox.RecipientTypeDetails
    $userObject | Add-Member @Member -Name "Alias" -Value $mailbox.Alias
    $userObject | Add-Member @Member -Name "PrimaryAddress" -Value $mailbox.PrimarySmtpAddress

    #export the alias e-mail addresses
    $EmailAddressesStringArray = [string[]]$mailbox.emailaddresses
    [int]$i = 0
    foreach ($emailAddress in $EmailAddressesStringArray)
    {
        if (($emailAddress | Select-String -Pattern "SMTP:") -or ($emailAddress | Select-String -Pattern "smtp:"))
        {
            if (!($emailAddress | Select-String -Pattern "adam.local") -and !($emailAddress.Replace("SMTP:","") -eq $mailbox.PrimarySmtpAddress))
            {
                $emailAddress = $emailAddress.Replace("smtp:","")
                $emailAddress = $emailAddress.Replace("SMTP:","")
                $userObject | Add-Member @Member -Name "EmailAddress_$i" -Value $emailAddress
                $i++
            }
        }
    }

    $exportUserArray += $userObject


    #Write-Host "AD User= $($adUser.SurName) $($adUser.GivenName)"
    #Write-Host "Mailbox = $($mailbox.RecipientTypeDetails)" -ForegroundColor Green
    #if ($mailbox.RecipientTypeDetails -ne "UserMailbox")
    #{
    #    Write-Host "Alias = $($mailbox.Alias) - number of characters =" $($mailbox.Alias).length -ForegroundColor Yellow
    #}
    #Write-Host "-------------------------------------"
}

$exportUserArray | Export-Csv -Path C:\Sources\userExport.csv -NoTypeInformation