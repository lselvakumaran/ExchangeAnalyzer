#requires -Modules ExchangeAnalyzer

#This function checks to see if all Domain Controllers are Global Catalogs
Function Run-AD003()
{
    [CmdletBinding()]
    param()

    $TestID = "AD003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()
    $TestSuccess = $null


    foreach ($DomainController in $ADDomainControllers)
    {
        $ServerName = $DomainController.name
        Write-Verbose "Checking the AD Domain Controller $ServerName to see if it is a Global Catalog Server"
        $GCValue = (Get-ADDomainController -server $ServerName).isglobalcatalog

        #If GCValue is true then the test is a pass and if it is false then the test fails with a warning
        if ($GCValue -eq $true) {

            write-verbose "The $ServerName is a Global Catalog server"
            $PassedList += $($ServerName)
        } else {

            write-verbose "The $ServerName is not a Global Catalog server"
            $InfoList += "$($ServerName) - is not a Global Catalog server"
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

Run-AD003