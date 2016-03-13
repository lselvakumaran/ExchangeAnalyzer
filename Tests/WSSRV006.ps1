#requires -Modules ExchangeAnalyzer

#This function checks the pagefile to see if it is over 32778 MB in size
Function Run-WSSRV006()
{
   [CmdletBinding()]
    param()

    $TestID = "WSSRV006"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    $PageFileDataList | foreach-object {
        # Prepare all variables
        $up=$_.up;$name = $_.name; $managed=$_.page_managed;$page_min=$_.page_min;$page_max=$_.page_max;$currentpagefile=$_.page_current;$RAMinMB=$_.RAMinMB
        
        if ($up -eq $false) {
            $FailedList += $($name)
            write-verbose "The server $name is unavailable for testing."
        } else {

            # Check if the pagefile is managed
            if ($managed -ne $true) {

                # Check if the maximum pagefile is greater than 32 GB + 10 MB
                if ($page_max -gt 32778) {
                    $FailedList += $($name)
                    write-verbose "The $name server has a PageFile size over 32778 MB in size."
                    write-verbose "The PageFile set to $CurrentPageFile MB"
                } else {         
                    $PassedList += $($name)
                    write-verbose "The $name server has a PageFile size less than 32778 MB in size."
                }
            } else {
                if ($CurrentPageFile -gt 32778) {
                    $FailedList += $($name)
                    write-verbose "The $name server has a PageFile size over 32778 MB in size."
                    write-verbose "The PageFile set to $CurrentPageFile MB"
                } else {
                    $PassedList += $($name)
                    write-verbose "The $name server has a PageFile size less than 32778 MB in size."
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

Run-WSSRV006