#requires -Modules ExchangeAnalyzer

#This function tests each Exchange site to determine whether more than one CAS URL/namespace
#exists for each HTTPS service.
Function Run-CAS001()
{
    [CmdletBinding()]
    param()

    $TestID = "CAS001"
    Write-Verbose "----- Starting test $TestID"

    $PassedList = @()
    $FailedList = @()
    $ErrorList = @()

    $sites = @($ClientAccessServers | Group-Object -Property:Site | Select-Object -Property Name)

    # Get the URLs for each site, and if more than one URL exists for a HTTPS service in the same site it is
    # considered a fail.
    foreach ($site in $sites)
    {
        $SiteName = ($Site.Name).Split("/")[-1]

        Write-Verbose "Processing $SiteName"

        $SiteOWAInternalUrls = @()
        $SiteOWAExternalUrls = @()

        $SiteECPInternalUrls = @()
        $SiteECPExternalUrls = @()

        $SiteOABInternalUrls = @()
        $SiteOABExternalUrls = @()
    
        $SiteRPCInternalUrls = @()
        $SiteRPCExternalUrls = @()

        $SiteEWSInternalUrls = @()
        $SiteEWSExternalUrls = @()

        $SIteMAPIInternalUrls = @()
        $SiteMAPIExternalUrls = @()

        $SiteActiveSyncInternalUrls = @()
        $SiteActiveSyncExternalUrls = @()

        $SiteAutodiscoverUrls = @()

        $CASinSite = @($ClientAccessServers | Where-Object -FilterScript {$_.Site -eq $site.Name})

        Write-Verbose "Getting OWA Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            Write-Verbose "Server: $($CAS.Name)"
            $CASOWAUrls = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property OWAInternal,OWAExternal)
            foreach ($CASOWAUrl in $CASOWAUrls)
            {
                if (!($SiteOWAInternalUrls -Contains $CASOWAUrl.OWAInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASOWAUrl.OWAInternal.ToLower()))
                {
                    $SiteOWAInternalUrls += $CASOWAUrl.OWAInternal.ToLower()
                }
                if (!($SiteOWAExternalUrls -Contains $CASOWAUrl.OWAExternal.ToLower()) -and  ![String]::IsNullOrEmpty($CASOWAUrl.OWAExternal.ToLower()))
                {
                    $SiteOWAExternalUrls += $CASOWAUrl.OWAExternal.ToLower()
                }
            }
        }

        if ($SiteOWAInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOWAExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ECP Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASECPUrls = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property ECPInternal,ECPExternal)
            foreach ($CASECPUrl in $CASECPUrls)
            {
                if (!($SiteECPInternalUrls -Contains $CASECPUrl.ECPInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASECPUrl.ECPInternal.ToLower()))
                {
                    $SiteECPInternalUrls += $CASECPUrl.ECPInternal.ToLower()
                }
                if (!($SiteECPExternalUrls -Contains $CASECPUrl.ECPInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASECPUrl.ECPInternal.ToLower()))
                {
                    $SiteECPExternalUrls += $CASECPUrl.ECPExternal.ToLower()
                }
            }
        }

        if ($SiteECPInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteECPExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting OAB Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASOABUrls = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property OABInternal,OABExternal)
            foreach ($CASOABUrl in $CASOABUrls)
            {
                if (!($SiteOABInternalUrls -Contains $CASOABUrl.OABInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASOABUrl.OABInternal.ToLower()))
                {
                    $SiteOABInternalUrls += $CASOABUrl.OABInternal.ToLower()
                }
                if (!($SiteOABExternalUrls -Contains $CASOABUrl.OABExternal.ToLower()) -and  ![String]::IsNullOrEmpty($CASOABUrl.OABExternal.ToLower()))
                {
                    $SiteOABExternalUrls += $CASOABUrl.OABExternal.ToLower()
                }
            }
        }

        if ($SiteOABInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteOABExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting RPC Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $OA = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property OAInternal,OAExternal)
            [string]$OAInternalHostName = $OA.OAInternal
            [string]$OAExternalHostName = $OA.OAExternal

            [string]$OAInternalUrl = "https://$($OAInternalHostName.ToLower())/rpc"
            [string]$OAExternalUrl = "https://$($OAExternalHostName.ToLower())/rpc"

            if (!($SiteRPCInternalUrls -Contains $OAInternalUrl) -and ![String]::IsNullOrEmpty($OAInternalHostName))
            {
                $SiteRPCInternalUrls += $OAInternalUrl
            }
            if (!($SiteRPCExternalUrls -Contains $OAExternalUrl) -and ![String]::IsNullOrEmpty($OAExternalHostName))
            {
                $SiteRPCExternalUrls += $OAExternalUrl
            }
        }

        if ($SiteRPCInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteRPCExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting EWS Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASEWSUrls = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property EWSInternal,EWSExternal)
            foreach ($CASEWSUrl in $CASEWSUrls)
            {
                if (!($SiteEWSInternalUrls -Contains $CASEWSUrl.EWSInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASEWSUrl.EWSInternal.ToLower()))
                {
                    $SiteEWSInternalUrls += $CASEWSUrl.EWSInternal.ToLower()
                }
                if (!($SiteEWSExternalUrls -Contains $CASEWSUrl.EWSExternal.ToLower()) -and ![String]::IsNullOrEmpty($CASEWSUrl.EWSExternal.ToLower()))
                {
                    $SiteEWSExternalUrls += $CASEWSUrl.EWSExternal.ToLower()
                }
            }
        }

        if ($SiteEWSInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteEWSExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting MAPI Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASMAPIUrls = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property-Object -Property MAPIInternal,MAPIExternal)
            foreach ($CASMAPIUrl in $CASMAPIUrls)
            {
                if (!($SiteMAPIInternalUrls -Contains $CASMAPIUrl.MAPIInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASMAPIUrl.MAPIInternal.ToLower()))
                {
                    $SiteMAPIInternalUrls += $CASMAPIUrl.MAPIInternal.ToLower()
                }
                if (!($SiteMAPIExternalUrls -Contains $CASMAPIUrl.MAPIExternal.ToLower()) -and ![String]::IsNullOrEmpty($CASMAPIUrl.MAPIExternal.ToLower()))
                {
                    $SiteMAPIExternalUrls += $CASMAPIUrl.MAPIExternal.ToLower()
                }
            }
        }

        if ($SiteMAPIInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteMAPIExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting ActiveSync Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            $CASActiveSyncUrls = @($CASURls | Where-Object -FilterScript {$_.Name -eq $CAS.Name} | Select-Object -Property EASInternal,EASExternal)
            foreach ($CASActiveSyncUrl in $CASActiveSyncUrls)
            {
                if (!($SiteActiveSyncInternalUrls -Contains $CASActiveSyncUrl.EASInternal.ToLower()) -and ![String]::IsNullOrEmpty($CASActiveSyncUrl.EASInternal.ToLower()))
                {
                    $SiteActiveSyncInternalUrls += $CASActiveSyncUrl.EASInternal.ToLower()
                }
                if (!($SiteActiveSyncExternalUrls -Contains $CASActiveSyncUrl.EASExternal.ToLower()) -and  ![String]::IsNullOrEmpty($CASActiveSyncUrl.EASExternal.ToLower()))
                {
                    $SiteActiveSyncExternalUrls += $CASActiveSyncUrl.EASExternal.ToLower()
                }
            }
        }

        if ($SiteActiveSyncInternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }
        if ($SiteActiveSyncExternalUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        Write-Verbose "Getting AutoDiscover Urls for site $SiteName"
        foreach ($CAS in $CASinSite)
        {
            #$CASServer = Get-ClientAccessServer $CAS.Name
            $AutoDUrl = @($CASURLs | Where-Object -FilterScript {$_.Name -ieq $CAS.Name} | Select-Object -Property AutoD)
            [string]$AutodiscoverSCP = $AutoDUrl
            $CASAutodiscoverUrl = $AutodiscoverSCP.Replace("/Autodiscover.xml","")
            if (!($SiteAutodiscoverUrls -Contains $CASAutodiscoverUrl.ToLower())) {$SiteAutodiscoverUrls += $CASAutodiscoverUrl.ToLower()}
        }

        if ($SiteAutodiscoverUrls.Count -gt 1) { if ($FailedList -notcontains $SiteName) { $FailedList += $SiteName} }

        #If the site is not in FailedList by now, add it to $PassedList
        if ($FailedList -notcontains $SiteName) { $PassedList += $SiteName }
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

Run-CAS001