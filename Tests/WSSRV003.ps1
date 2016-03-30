#requires -Modules ExchangeAnalyzer

#This function checks the pagefile is set to RAM + 10MB
Function Run-WSSRV003()
{
   [CmdletBinding()]
    param()

    $TestID = "WSSRV003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    $PageFileDataList | foreach-object {
        
        # Prepare all variables
        $up = $_.up
        $name = $_.name
        $managed=$_.page_managed
        $page_min=$_.page_min
        $page_max=$_.page_max
        $currentpagefile=$_.page_current
        $RAMinMB=$_.RAMinMB
        $RAMIdeal = $RAMinMB + 10

        if ($up -eq $false) {
            $FailedList += $($name)
            write-verbose "The server $name is unavailable for testing."
        } else {

            # Check if the pagefile is managed
            if ($managed -ne $true) {
                
                # Check to make sure that the PageFile is the same size as RAM + 10Mb
                if ($page_max -like $RAMIdeal) {
                    $PassedList += $($name)
                    write-verbose "The $name server has the correct PageFile size of $RAMIdeal MB."
                } else {
                    $FailedList += $($name)
                    write-verbose "The $name server does not have the correct PageFile size of $RAMIdeal MB."
                    write-verbose "The PageFile set to $CurrentPageFile MB."
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
}

Run-WSSRV003