#requires -Modules ExchangeAnalyzer

#This function checks the pagefile size and if it is managed
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

    foreach ($server in $exchangeservers) {
        $name = $server.name

        # Get RAM of remote server
        $RAMinMB = (Get-WmiObject -Class win32_physicalmemory -computer $name | measure-object -property capacity -sum).sum/1048576
        write-verbose "The $name server has $RAMinMB MB of RAM installed"

        # Get IP Address of remote server
        $ipV4 = Test-Connection -ComputerName $name -Count 1  | Select IPV4Address
        $ip = $ipV4.IPV4Address

        #Pagefile check
        $managed = Get-WmiObject -ComputerName $server -Class Win32_ComputerSystem | % {$_.AutomaticManagedPagefile}

        # Check to see if PageFile is System Managed
        if ($Managed -ne $true) {
            $PassedList += $($name)
            write-verbose "The PageFile on server $name is not System Managed."
        } else {
            $FailedList += $($name)
            write-verbose "The PageFile on server $name is System Managed."
                
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