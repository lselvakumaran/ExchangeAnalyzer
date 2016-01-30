#requires -Modules ExchangeAnalyzer

#This function checks to see if the C: Drive has greater than 30% free space
Function Run-WSSRV001()
{

   [CmdletBinding()]
    param()

    $TestID = "WSSRV001"
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
                $free = $line.freespace
                $size = $line.size

                if (($free -ne $null) -or ($size -ne $null)) {
                    $percent2 = ($free/$size)*100
                    $percent = [math]::Round($percent2,2)
                    if ($percent -lt "30") {
                        $FailedList += $($name)
                    } else {
                        $PassedList += $($name)
                    }
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

Run-WSSRV001