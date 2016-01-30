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
        $space = Get-WMIObject win32_logicaldisk -computername $name
    
        foreach ($line in $space) {
            $drive = $line.DeviceID
            if ($drive -eq "C:") {
                $size = $line.size
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