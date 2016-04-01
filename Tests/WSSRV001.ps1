#requires -Modules ExchangeAnalyzer

#This function checks to see if the Exchange Install Drive has greater than 30% free space
Function Run-WSSRV001()
{

   [CmdletBinding()]
    param()

    $TestID = "WSSRV001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        $ServerName = $server.name
        if ($server.AdminDisplayVersion -match "Version 14") {$ver = "V14"}
        if ($server.AdminDisplayVersion -match "Version 15") {$ver = "V15"}

        Write-Verbose "Checking $ServerName."

        # Null out variables for clean results
        $installpath = $null
        $reg = $null

        # Get the Exchange Install Drive
        try
        {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ServerName)
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $installpath = "Unable to connect to registry"
        }

        if (!($installpath))
        {
            $installpath = $reg.OpenSubKey("SOFTWARE\Microsoft\ExchangeServer\$ver\Setup").GetValue("MsiInstallPath")
        }
        $ExchangeInstallDrive = $InstallPath.substring(0,2)

        # Get free Space for all drives on the Exchange Server
        $LogicalDisks = Get-WMIObject win32_logicaldisk -computername $ServerName
    
        # Go through each drive on the server
        foreach ($LogicalDisk in $LogicalDisks)
        {

            # Get the current drive in the loop
            $drive = $LogicalDisk.DeviceID

            # Look just for the Exchange Install Drive for percent free space
            if ($drive -eq $ExchangeInstallDrive) {
                $free = $LogicalDisk.freespace
                $size = $LogicalDisk.size

                # Calculate percent free space
                $percent2 = ($free/$size)*100

                # Round the percentage
                $percent = [math]::Round($percent2,2)

                # Check to see if the Exchange Install Drive has 30% or greater free space
                if ($percent -lt "15") {
                    write-verbose "The Exchange Install Drive has less than 30% free space."
                    $FailedList += $($ServerName)
                } 
                if (($percent -gt "15") -and ($percent -lt "30")) {
                   write-verbose "The Exchange Install Drive has greater than 15% and less than 30% free space."
                    $WarningList += $($ServerName)
                }
                if ($percent -gt "30") {
                    write-verbose "The Exchange Install Drive has greater than 30% free space."
                    $PassedList += $($ServerName)
                }
            }
        }
    }

    #Roll the object to be returned to the results
    $ReportObj = Get-TestResultObject -ExchangeAnalyzerTests $ExchangeAnalyzerTests `
                                      -TestId $TestID `
                                      -PassedList $PassedList `
                                      -FailedList $FailedList `
                                      -WarningList $WarningList `
                                      -InfoList $InfoList `
                                      -ErrorList $ErrorList `
                                      -Verbose:($PSBoundParameters['Verbose'] -eq $true)

    return $ReportObj
}

Run-WSSRV001