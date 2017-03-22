#requires -Modules ExchangeAnalyzer

#This function checks to see if the Exchange Server power management is set to High Performance
Function Run-HW002()
{

   [CmdletBinding()]
    param()

    $TestID = "HW002"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        # Set Variables for the current loop
        $CurrPlan = $null
        $name = $server.name
        $ConnectionError = $false
        $tryWMI = $false

        try {
            Get-CIMInstance -Class win32_powerplan -ComputerName $name  -Namespace root\cimv2\power -Filter "isactive='true'" -erroraction STOP
        } catch {
            $tryWMI = $true
            Write-Verbose "$($TestID): Was not able to acquire information for $name via CIM"    
        }
        if ($tryWMI) {
            try {
                Get-WmiObject -Class win32_powerplan -ComputerName $name  -Namespace root\cimv2\power -Filter "isactive='true'" -erroraction STOP
            } catch {
                Write-Verbose "$($TestID): Was not able to acquire information for $name via WMI"    
                $ConnectionError = $true
            }
        }

        If ($ConnectionError -eq $false) {
            try {    
                $CurrPlan = (Get-CIMInstance -Class win32_powerplan -ComputerName $name  -Namespace root\cimv2\power -Filter "isactive='true'" -erroraction STOP).elementname
            } catch {
                Write-Verbose "$($TestID): Was not able to acquire information for $name via CIM"
                $tryWMI = $true
            }
            if ($tryWMI) {
                try {
                    $CurrPlan = (Get-WMIObject -Class win32_powerplan -ComputerName $name  -Namespace root\cimv2\power -Filter "isactive='true'" -erroraction STOP).elementname
                } catch {
                    Write-Verbose "$($TestID): Was not able to acquire information for $name via WMI"
                }
            }
            
            # Validate if the server is set to use the High Performance power plan
            if ($CurrPlan -eq "High performance") {
                write-verbose "The power plan on $name is set to High Performance."
                $PassedList += $($name)
            } else {
                write-verbose "The power plan on $name is not set to High Performance and is currently set to $currplan."
                $WarningList += $($name)
            }
        } else {
            write-verbose "There was an issue connecting to the $name server. "
            $WarningList += $($name)
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

Run-HW002