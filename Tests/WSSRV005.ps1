#requires -Modules ExchangeAnalyzer

#This function checks the PageFile to see if it is System Managed
Function Run-WSSRV005()
{
   [CmdletBinding()]
    param()

    $TestID = "WSSRV005"
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
        
            # Check to see if PageFile is System Managed
            if ($Managed -ne $true) {
                $PassedList += $($name)
                write-verbose "The PageFile on server $name is not System Managed."
            } else {
                $FailedList += $($name)
                write-verbose "The PageFile on server $name is System Managed." 
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

Run-WSSRV005