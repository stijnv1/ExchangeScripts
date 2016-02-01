#
# AddEmailAliasses.ps1
#
param
(
    [Parameter(Mandatory=$true)]
    [string]$CSVFilePath,

    [Parameter(Mandatory=$true)]
    [string]$LogDirPath,

    [Parameter(Mandatory=$true)]
    [string]$ExchangeServerName,

	[Parameter(Mandatory=$true)]
	[stirng]$AADConnectServerName
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
		$LogFileName = "UpdateAlias_$thisDate.log"

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

Function AddEmailAddressRemoteMailbox
{
    param
    (
        [string]$userUPN,
        [Array]$Aliasses,
        [string]$ExchangeServerName,
        [string]$LogDirPath
    )

    Try
    {
        #add mail addresses
        foreach ($mailaddress in $Aliasses)
        {
            if (Set-RemoteMailbox -Identity $userUPN -EmailAddresses @{add="$mailaddress"} -ErrorAction stop)
            {
                WriteToLog -LogPath $LogDirPath -TextValue "Added e-mail address $mailaddress successfully to Office 365 mailbox" -WriteError $false
            }
            else
            {
                #WriteToLog -LogPath $LogDirPath -TextValue "Error occured while adding e-mail address $mailaddress : $error[0]" -WriteError $true
            }
        }
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
	    WriteToLog -LogPath $LogDirPath -TextValue "Error occured in AddEmailAddressRemoteMailbox function while processing user $($userUPN) : $ErrorMessage" -WriteError $true
	    Write-Host "Error occured in AddEmailAddressRemoteMailbox function while processing user $($userUPN) : $ErrorMessage" -ForegroundColor Red
    }

}

Function AddEmailAddressOnPremMailbox
{
    param
    (
        [string]$userUPN,
        [Array]$Aliasses,
        [string]$ExchangeServerName,
        [string]$LogDirPath
    )

    Try
    {
        #add mail addresses
        foreach ($mailaddress in $Aliasses)
        {
            if (Set-Mailbox -Identity $userUPN -EmailAddresses @{add="$mailaddress"} -ErrorAction stop)
            {
                WriteToLog -LogPath $LogDirPath -TextValue "Added e-mail address $mailaddress successfully to onpremise mailbox" -WriteError $false
            }
            else
            {
                #WriteToLog -LogPath $LogDirPath -TextValue "Error occured while adding e-mail address $mailaddress : $error[0]" -WriteError $true
            }
        }
    }

    Catch
    {
        $ErrorMessage = $_.Exception.Message
	    WriteToLog -LogPath $LogDirPath -TextValue "Error occured in AddEmailAddressOnPremMailbox function while processing user $($userUPN) : $ErrorMessage" -WriteError $true
	    Write-Host "Error occured in AddEmailAddressOnPremMailbox function while processing user $($userUPN) : $ErrorMessage" -ForegroundColor Red
    }
}

Try
{
    #import CSV data
	$UserData = Import-Csv $CSVFilePath -Delimiter ","
    
    #get number of alias columns by counting noteproperty minus 1. UPN column is the column which must be ignored. Script assumes that all
    #address columns start with value "Alias" followed by a sequence number
    #for example: Alias1, Alias2, ... 
    #this count is the maximum number of aliasses that are defined. it is possible that some users do not have this maximum of defined addresses
    $numberOfAddresses = (($UserData | Get-Member -MemberType NoteProperty).Count) - 1

    #create remote powershell session to onpremise exchange
    $OnPremExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServerName/PowerShell/ -Authentication Kerberos
    Import-PSSession $OnPremExchangeSession
    
	#start looping through user data and create objects
	foreach ($userObject in $UserData)
	{
        #create array of defined a-mail addresses
        $addressArray = @()
        $arrayCounter = 0
        $emptyAliasDefined = $false

        Write-Host "Processing user with UPN $($userObject.UPN) ..." -ForegroundColor Yellow
        WriteToLog -LogPath $LogDirPath -TextValue "Processing user with UPN $($userObject.UPN) ..." -WriteError $false
        do
        {
            $columnName = (($userObject | Get-Member -MemberType NoteProperty))[$arrayCounter].Name

            if ($userObject.$columnName)
            {
                Write-Host "Adding e-mail address $($userObject.$columnName) to address array of user ..." -ForegroundColor Green
                WriteToLog -LogPath $LogDirPath -TextValue "Adding e-mail address $($userObject.$columnName) to address array of user" -WriteError $false
                $addressArray += $userObject.$columnName
                $arrayCounter++
            }
            else
            {
                $emptyAliasDefined = $true
            }
        }
        while ((!$emptyAliasDefined) -and ($arrayCounter -lt $numberOfAddresses))

        Write-Host "Starting function to add e-mail address(es) to mailbox of user $($userObject.UPN) ..." -ForegroundColor Gray
        WriteToLog -LogPath $LogDirPath -TextValue "Starting function to add e-mail address(es) to mailbox of user $($userObject.UPN) ...`n" -WriteError $false

        if (Get-RemoteMailbox $userObject.UPN -ErrorAction SilentlyContinue)
        {
            Write-Host "User mailbox is an Office 365 mailbox ...`n`n" -ForegroundColor Green
            AddEmailAddressRemoteMailbox -userUPN $userObject.UPN -Aliasses $addressArray -ExchangeServerName $ExchangeServerName -LogDirPath $LogDirPath
        }
        elseif (Get-Mailbox $userObject.UPN -ErrorAction SilentlyContinue)
        {
            Write-Host "User mailbox is an onpremise mailbox ...`n`n" -ForegroundColor Green
            AddEmailAddressOnPremMailbox -userUPN $userObject.UPN -Aliasses $addressArray -ExchangeServerName $ExchangeServerName -LogDirPath $LogDirPath
        }
    }
    Remove-PSSession $OnPremExchangeSession

    #start Azure AD sync
    Write-Host "Starting Azure AD Sync ..." -ForegroundColor DarkYellow
    Invoke-Command -ComputerName $AADConnectServerName -ScriptBlock {& cmd.exe /c "D:\Program Files\Microsoft Azure AD Sync\Bin\DirectorySyncClientCmd.exe"}
}

Catch
{
    $ErrorMessage = $_.Exception.Message
    Remove-PSSession $OnPremExchangeSession
	WriteToLog -LogPath $LogDirPath -TextValue "Error occured: $ErrorMessage" -WriteError $true
	Write-Host "Error occured: $error" -ForegroundColor Red
}