#requires -Modules ExchangeAnalyzer

#This function checks to see if the Exchange Server NIC power management is set to High Performance
Function Run-NET001()
{

   [CmdletBinding()]
    param()

    $TestID = "NET001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    foreach ($server in $exchangeservers) {
        $name = $server.name

        # Get all NICs that are valid on the server for network traffic
        $NICs = Get-WmiObject -computer $name -Class Win32_NetworkAdapter|Where-Object{$_.PNPDeviceID -notlike "ROOT\*" -and $_.Manufacturer -ne "Microsoft" -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22} 
        Foreach($NIC in $NICs) {

            # Get the Device ID of the NIC to help search the registry for the power management setting
            $NICName = $NIC.Name
            $DeviceID = $NIC.DeviceID
            If([Int32]$DeviceID -lt 10) {
                $DeviceNumber = "000"+$DeviceID 
            } Else {
                $DeviceNumber = "00"+$DeviceID
            }

            #Get value for the pnpcapabilities value - NIC Power Management state
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $name )
            $pnp = $reg.OpenSubKey("SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\$DeviceNumber").GetValue("PnPCapabilities")
  
            # Chekc $pnp to see if it equals 24 (Power Management off)
            if ($PnP -ne 24) {
                write-verbose "This NIC does not have Power Management enabled."
                $WarningList += $($name)
            } else {
                write-verbose "This NIC has Power Management enabled." 
                $PassedList += $($name)
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

Run-NET001