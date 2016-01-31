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

        # Get free Space for all drives on the Exchange Server
        $space = Get-WMIObject win32_logicaldisk -computername $name
    
        # Go through each drive on the server
        foreach ($line in $space) {

            # Pull just the drive letter
            $drive = $line.DeviceID

            # Look just for the C: drive for percent free space
            if ($drive -eq "C:") {
                $free = $line.freespace
                $size = $line.size

                # Calculate percent free space
                $percent2 = ($free/$size)*100

                # Round the percentage
                $percent = [math]::Round($percent2,2)

                # Check to se if C: drive has 30% or greater free space
                if ($percent -lt "30") {
                    $FailedList += $($name)
                } else {
                    $PassedList += $($name)
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