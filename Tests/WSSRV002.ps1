#requires -Modules ExchangeAnalyzer

#This function checks to see if the C: Drive is greater than 130 GB
Function Run-WSSRV002()
{

   [CmdletBinding()]
    param()

    $TestID = "WSSRV002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        $name = $server.name

        # Get free Space for all drives on the Exchange Server
        $space = Get-WMIObject win32_logicaldisk -computername $name
    
        # Go through each drive on the server
        foreach ($line in $space) {

            # Pull just the drive letter
            $drive = $line.DeviceID

            # Look just for the C: drive to find how large the volume is
            if ($drive -eq "C:") {
                $size = $line.size

                # Check to see if the C: drive is larger than 130 GB
                if ($size -gt 139586437120) {
                    $PassedList += $($name)
                } else {
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
                                      -ErrorList $ErrorList `
                                      -Verbose:($PSBoundParameters['Verbose'] -eq $true)

    return $ReportObj
}

Run-WSSRV002