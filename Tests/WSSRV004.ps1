#requires -Modules ExchangeAnalyzer

#This function checks the pagefile min and max sizes
Function Run-WSSRV004()
{
   [CmdletBinding()]
    param()

    $TestID = "WSSRV004"
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
                   
                # Check if Initial and Maximum Pagefile are the same
                    if ($page_min -eq $page_max) {
                    $PassedList += $($name)
                    write-verbose "The $name server has the initial and maximum PageFile set to the same value."
                    write-verbose "The initial PageFile is set to $page_min."
                    write-verbose "The maximum PageFile is set to $page_max."
                } else {
                    $FailedList += $($name)
                    write-verbose "The $name server does not have the initial and maximum PageFile set to the same value."
                    write-verbose "The initial PageFile is set to $page_min."
                    write-verbose "The maximum PageFile is set to $page_max."
                }
            } else {
                $FailedList += $($name)
                write-verbose "The PageFile is System Managed and fails this test."
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
    Write-Verbose "----- Ending test $TestID" 
}

Run-WSSRV004