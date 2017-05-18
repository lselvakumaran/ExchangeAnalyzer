#requires -Modules ExchangeAnalyzer

#This function tests each Exchange server to verify a supported version of the .NET framework istalled
Function Run-EXSRV003()
{
    [CmdletBinding()]
    param()

    $TestID = "EXSRV003"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $WarningList = @()
    $InfoList = @()
    $ErrorList = @()

    #Import the CSV file containing Net Framework support information
    Write-Verbose "Importing CSV file for .NET support matrix"
    $NetFXSupportMatrix = Import-CSV "$($MyDir)\Data\NETFXSupportMatrix.csv"

    foreach ($server in $exchangeservers) {
        # Set the inital value
        
        $ServerName = $server.name
        $NetFxRelease = $null
        $NetFXSupportStatus = $null

        Write-Verbose "Checking $ServerName"

        $NetFxRelease = Get-ExARegistryValue -Host $ServerName -Hive LocalMachine -Key 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Value 'Release'

        Write-Verbose "Release: $NetFXRelease"

        If ($NetFxRelease)
        {
            # Results of the check
            switch ($NETFxRelease)
            {
                378389 {$NetFXVersion = ".NET Framework 4.5"}
                378675 {$NetFXVersion = ".NET Framework 4.5.1"}
                379893 {$NetFXVersion = ".NET Framework 4.5.2"}
		393297 {$NetFXVersion = ".NET Framework 4.6"}
                394271 {$NetFXVersion = ".NET Framework 4.6.1"}
                394802 {$NetFXVersion = ".NET Framework 4.6.2"}
		394806 {$NetFXVersion = ".NET Framework 4.6.2"}
                460805 {$NetFXVersion = ".NET Framework 4.7"}
                default {$NetFxVersion = "Unknown"}
            }
            
            Write-Verbose ".NET FX Version is: $NetFXVersion"
            Set-ExAServerProperty -Server $Server -Property '.NET Framework' -Value $NetFXVersion
            $ServerVersion = Get-ExAServerProperty -Server $Server -Property 'BuildDescription'

            Write-Verbose "Server version is: $ServerVersion"
            
            $NetFXSupportStatus = ($NetFXSupportMatrix | Where {$_.ExchangeDescription -eq $ServerVersion}).$NetFXVersion

            if (-not($NetFXSupportStatus)) {
                $NetFXSupportStatus = "Unknown"
            }
            
            Write-Verbose "Support status: $NetFXSupportStatus"

            if ($NetFXSupportStatus -eq "Supported") {
                $PassedList += $ServerName
            }
            elseif ($NetFXSupportStatus -eq "Unsupported") {
                $FailedList += $ServerName
            }
            elseif ($NetFxSupportStatus -eq "Unknown") {
                $WarningList += $ServerName
            }
            else {
                $ErrorList += $ServerName
            }

        }
        else
        {
            $WarningList += "$ServerName - unable to connect to registry"
            Write-Verbose "Unable to connect to registry of $ServerName"
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

Run-EXSRV003
