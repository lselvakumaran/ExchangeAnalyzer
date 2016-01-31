#requires -Modules ExchangeAnalyzer

#This function checks the pagefile size and if it is managed
Function Run-WSSRV003()
{

   [CmdletBinding()]
    param()

    $TestID = "WSSRV003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        $name = $server.name

        # Get RAM of remote server
        $RAMinMB = (Get-WmiObject -Class win32_physicalmemory -computer $server | measure-object -property capacity -sum).sum/1048576
        write-verbose "The $name server has $RAMinMB MB of RAM installed"

        # Get IP Address of remote server
        $ipV4 = Test-Connection -ComputerName $name -Count 1  | Select IPV4Address
        $ip = $ipV4.IPV4Address

        #Pagefile check
        $managed = Get-WmiObject -ComputerName $server -Class Win32_ComputerSystem | % {$_.AutomaticManagedPagefile}
        $RAMIdeal = $RAMinMB+10
        $CurrentPageFile = (Get-WmiObject -computer $server Win32_PageFileUsage).allocatedbasesize

        # Get Initial and Maximum PageFileSize
        $is = wmic /node:"$ip" pagefileset get initialsize
        $ms = wmic /node:"$ip" pagefileset get maximumsize

        # Split pagefile value to get number only
        $PFInitSize = $is.Split([char]0x000D)
        $PFMaxSize = $is.Split([char]0x000D)

        # Select the numerical value from the PageFile variables
        $pfis = $PFInitSize[2]
        $pfms = $PFMAXSize[2]

        # Remove extra spaces
        $pfis = $pfis -replace '\s',''
        $pfms = $pfms -replace '\s',''

        # Check to see if PageFile is System Managed
        if ($Managed -ne $true) {
            write-verbose "The $name server does not have a system managed PageFile."
            # Check if Initial and Maximum Pagefile are the same
            if ($pfis -eq $pfms) {

                # Check to make sure that the PageFile is the same size as RAM + 10Mb
                if ($pfms -like $RAMIdeal) {
                    $PassedList += $($name)
                    write-verbose "The $name server has the correct PageFile size of $RAMIdeal MB."
                } else {
                    $FailedList += $($name)
                    write-verbose "The $name server does not have the correct PageFile size of $RAMIdeal MB."
                    write-host "The PageFile set to $CurrentPageFile MB."
                }
            } else {
                $FailedList += $($name)
                write-verbose "The $name server does not have the initial and maximum PageFile set to the same value."
                write-verbose "The initial PageFile is set to $pfis."
                write-verbose "The maximum PageFile is set to $pfmx."
            }
        } else {
            $FailedList += $($name)
            write-verbose "The $name server has a system managed PageFile."
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

Run-WSSRV003