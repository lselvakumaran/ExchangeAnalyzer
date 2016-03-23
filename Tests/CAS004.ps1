#requires -Modules ExchangeAnalyzer

#Checks to see if the RC4 cipher is enabled (should be disabled)
Function Run-CAS004()
{
    [CmdletBinding()]
    param()

    $TestID = "CAS004"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    foreach ($server in $ExchangeServers) {
        # Set the inital value
        $name = $server.name
        $up = $true
        $success = $true
        try {
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$name)
        } catch {
            $up = $false
        }

        if ($up -eq $true) {
    
            # Check the registry path - Cipher Stack
            $check1 = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\")
    
            if ($check1 -ne $null) {
           
                # Registry check - RC4 128/128\
                $check2a = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\")
            
                if ($check2a -ne $null) {
                    $check2b = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\").GetValue("Enabled")
                
                    if ($check2b -ne 0) {
                        $success = $false
                    }
                } else {
                    $success = $false
                }

                # Registry check - RC4 40/128\
                $check3a = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\")
            
                if ($check3a -ne $null) {
                    $check3b = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\").GetValue("Enabled")
                
                    if ($check3b -ne 0) {
                        $success = $false
                    }
                } else {
                    $success = $false
                }

                # Registry check 4 - RC4 56/128\
                $check4a = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\")
            
            if ($check4a -ne $null) {
                $check4b = $registry.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\").GetValue("Enabled")
                
                if ($check4b -ne 0) {
                    $success = $false
                }
            } else {
                $success = $false
            }
        } else {
            $succes = $false
            write-verbose "The server $name is down or unreachable."
        }
        
        # Decide if the test has failed based off of the values missing or present and what value is there if present
        If ($Success -eq $true) {
            $PassedList += $name
            write-verbose "RC4 is disabled on the server $name!!"
        } else {
            $FailedList += $name
                write-verbose "RC4 is not disabed on server $name!!"
            }
        } else {
            $FailedList += $name
            write-verbose "The server $name is down/unreachable and RC4 settings cannot be verified."
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

Run-CAS004