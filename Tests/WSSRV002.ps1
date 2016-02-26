#requires -Modules ExchangeAnalyzer

#This function checks to see if the Exchange Install Drive is greater than 130 GB
Function Run-WSSRV002()
{

   [CmdletBinding()]
    param()

    $TestID = "WSSRV002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        $name = $server.name
        if ($server.AdminDisplayVersion -match "Version 14") {$ver = "V14"}
        if ($server.AdminDisplayVersion -match "Version 15") {$ver = "V15"}

        Write-Verbose "Checking $name."

        # Null out variables for clean results
        $installpath = $null
        $reg = $null

        # Get the Exchange Install Drive
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$name)
        } catch {
            Write-Warning $_.Exception.Message
            $installpath = "Unable to connect to registry"
        }
        if (!($installpath)) {
            $installpath = $reg.OpenSubKey("SOFTWARE\Microsoft\ExchangeServer\$ver\Setup").GetValue("MsiInstallPath")
        }
        $ExchangeDrive = $InstallPath.substring(0,2)

        # Get free Space for all drives on the Exchange Server
        $space = Get-WMIObject win32_logicaldisk -computername $name
    
        # Go through each drive on the server
        foreach ($line in $space) {

            # Pull just the drive letter
            $drive = $line.DeviceID

            # Look just for the Exchange Install Drive to find how large the volume is
            if ($drive -eq $ExchangeDrive) {
                $size = $line.size
                $SizeGB = $size/1073741824
                write-verbose "The Exchange Install  Drive is $sizeGB GB in size."

                # Check to see if the Exchange Install Drive is larger than 130 GB
                if ($SizeGB -gt 130) {
                    write-verbose "The Exchange Install Drive is over 130 GB in size."
                    $PassedList += $($name)
                } 
                if (($SizeGB -gt 100) -and ($sizeGB -lt 130)) {
                    write-verbose "The Exchange Install Drive could be larger than it is at $sizeGB GB in size."
                    $WarningList += $($name)
                }
                if ($SizeGB -lt 100) {                
                    write-verbose "The Exchange Install Drive is too small at $sizeGB GB in size."
                    $FailedList += $($name)
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

Run-WSSRV002