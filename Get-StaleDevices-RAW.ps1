<#
.SYNOPSIS
    Extract workstation account details from AzureAD, Intune and on-prem AD, and analyse for dormancy
.DESCRIPTION
    This script retrieves device details from AzureAD, Intune and on-prem ad for analysis
.PARAMETER TenantID
    Specify the Azure AD tenant ID in the code below.
.PARAMETER ClientID
    Specify the service principal, also known as app registration, Client ID (also known as Application ID) in the code below.
.PARAMETER State
    Nothing to see here... Yet
.EXAMPLE
    .\Get-StaleDevices.ps1 -Verbose
.NOTES
    FileName:    Get-StaleDevices.ps1
    Author:      Christopher Baxter
    Contact:     https://github.com/christopherbaxter
    Created:     2020-12-04
    Updated:     2022-01-24

    I wish to thank Anders Ahl and Nickolaj Andersen as well as the authors of the functions I have used below, without these guys, this script would not exist.

    You will notice that I have hashed out a number of lines in the script. I have done this for a very specific reason. I have had numberous issues with reliable data extraction
    from the various locations. The script will export the data in each section. In my environment, some of the extractions take more than an hour. The lines that have been hashed will allow for faster testing\troubleshooting
    
    You will also notice some strange code for retrying and for token renewal. By default, your token is valid for an hour, depending on the size of your estate,
    this could be WAY too short. Basically, the way the script has been hacked together, you will not have an issue with this, in most cases. I am able to keep the script running for 8 hours (the amount of time my elevated rights remain valid),
    without any trouble.

    Depending on the size of your environment, the RAM util gets quite high, the script also runs a little slowly as it moves along. I have added a number of 'cleanups' into the script, this helps significantly.
    
    In order to configure your Service Principal, follow this guide. https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/

#>
#Requires -Modules "MSAL.PS,ActiveDirectory,AzureAD,ImportExcel,JoinModule,PSReadline"
[CmdletBinding(SupportsShouldProcess = $TRUE)]
param(
    #PLEASE make sure you have specified your details below, else edit this and use the switches\variables in command line.
    [parameter(Mandatory = $False, HelpMessage = "Specify the Azure AD tenant ID.")]
    [ValidateNotNullOrEmpty()]
    #[string]$TenantID = "", # Populate this with your TenantID, this will then allow the script to run without asking for the details
    [string]$TenantID,

    [parameter(Mandatory = $False, HelpMessage = "Specify the service principal, also known as app registration, Client ID (also known as Application ID).")]
    [ValidateNotNullOrEmpty()]
    #[string]$ClientID = "" # Populate this with your ClientID\ApplicationID of your Service Principal, this will then allow the script to run without asking for the details
    [string]$ClientID
)
Begin {}
Process {
    # Functions
    function Invoke-MSGraphOperation {
        <#
        .SYNOPSIS
            Perform a specific call to Graph API, either as GET, POST, PATCH or DELETE methods.
            
        .DESCRIPTION
            Perform a specific call to Graph API, either as GET, POST, PATCH or DELETE methods.
            This function handles nextLink objects including throttling based on retry-after value from Graph response.
            
        .PARAMETER Get
            Switch parameter used to specify the method operation as 'GET'.
            
        .PARAMETER Post
            Switch parameter used to specify the method operation as 'POST'.
            
        .PARAMETER Patch
            Switch parameter used to specify the method operation as 'PATCH'.
            
        .PARAMETER Put
            Switch parameter used to specify the method operation as 'PUT'.
            
        .PARAMETER Delete
            Switch parameter used to specify the method operation as 'DELETE'.
            
        .PARAMETER Resource
            Specify the full resource path, e.g. deviceManagement/auditEvents.
            
        .PARAMETER Headers
            Specify a hash-table as the header containing minimum the authentication token.
            
        .PARAMETER Body
            Specify the body construct.
            
        .PARAMETER APIVersion
            Specify to use either 'Beta' or 'v1.0' API version.
            
        .PARAMETER ContentType
            Specify the content type for the graph request.
            
        .NOTES
            Author:      Nickolaj Andersen & Jan Ketil Skanke & (very little) Christopher Baxter
            Contact:     @JankeSkanke @NickolajA
            Created:     2020-10-11
            Updated:     2020-11-11
    
            Version history:
            1.0.0 - (2020-10-11) Function created
            1.0.1 - (2020-11-11) Tested in larger environments with 100K+ resources, made small changes to nextLink handling
            1.0.2 - (2020-12-04) Added support for testing if authentication token has expired, call Get-MsalToken to refresh. This version and onwards now requires the MSAL.PS module
            1.0.3.Custom - (2020-12-20) Added aditional error handling. Not complete, but more will be added as needed. Christopher Baxter
        #>
        param(
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Switch parameter used to specify the method operation as 'GET'.")]
            [switch]$Get,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST", HelpMessage = "Switch parameter used to specify the method operation as 'POST'.")]
            [switch]$Post,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH", HelpMessage = "Switch parameter used to specify the method operation as 'PATCH'.")]
            [switch]$Patch,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT", HelpMessage = "Switch parameter used to specify the method operation as 'PUT'.")]
            [switch]$Put,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE", HelpMessage = "Switch parameter used to specify the method operation as 'DELETE'.")]
            [switch]$Delete,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Specify the full resource path, e.g. deviceManagement/auditEvents.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [string]$Resource,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Specify a hash-table as the header containing minimum the authentication token.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [System.Collections.Hashtable]$Headers,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST", HelpMessage = "Specify the body construct.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [ValidateNotNullOrEmpty()]
            [System.Object]$Body,
    
            [parameter(Mandatory = $fALSE, ParameterSetName = "GET", HelpMessage = "Specify to use either 'Beta' or 'v1.0' API version.")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "POST")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("Beta", "v1.0")]
            [string]$APIVersion = "v1.0",
    
            [parameter(Mandatory = $fALSE, ParameterSetName = "GET", HelpMessage = "Specify the content type for the graph request.")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "POST")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("application/json", "image/png")]
            [string]$ContentType = "application/json"
        )
        Begin {
            # Construct list as return value for handling both single and multiple instances in response from call
            $GraphResponseList = New-Object -TypeName "System.Collections.ArrayList"
            $Runcount = 1

            # Construct full URI
            $GraphURI = "https://graph.microsoft.com/$($APIVersion)/$($Resource)"
            #Write-Verbose -Message "$($PSCmdlet.ParameterSetName) $($GraphURI)"
        }
        Process {
            # Call Graph API and get JSON response
            do {
                try {
                    # Determine the current time in UTC
                    $UTCDateTime = (Get-Date).ToUniversalTime()
    
                    # Determine the token expiration count as minutes
                    $TokenExpireMins = ([datetime]$Headers["ExpiresOn"] - $UTCDateTime).Minutes
    
                    # Attempt to retrieve a refresh token when token expiration count is less than or equal to 10
                    if ($TokenExpireMins -le 10) {
                        #Write-Verbose -Message "Existing token found but has expired, requesting a new token"
                        $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                        $Headers = New-AuthenticationHeader -AccessToken $AccessToken
                    }
    
                    # Construct table of default request parameters
                    $RequestParams = @{
                        "Uri"         = $GraphURI
                        "Headers"     = $Headers
                        "Method"      = $PSCmdlet.ParameterSetName
                        "ErrorAction" = "Stop"
                        "Verbose"     = $VerbosePreference
                    }
    
                    switch ($PSCmdlet.ParameterSetName) {
                        "POST" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                        "PATCH" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                        "PUT" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                    }
    
                    # Invoke Graph request
                    $GraphResponse = Invoke-RestMethod @RequestParams
                    
                    # Handle paging in response
                    if ($GraphResponse.'@odata.nextLink') {
                        $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                        $GraphURI = $GraphResponse.'@odata.nextLink'
                        #Write-Verbose -Message "NextLink: $($GraphURI)"
                    }
                    else {
                        # NextLink from response was null, assuming last page but also handle if a single instance is returned
                        if (-not([string]::IsNullOrEmpty($GraphResponse.value))) {
                            $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                        }
                        else {
                            $GraphResponseList.Add($GraphResponse) | Out-Null
                        }
                        
                        # Set graph response as handled and stop processing loop
                        $GraphResponseProcess = $fALSE
                    }
                }
                catch [System.Exception] {
                    $ExceptionItem = $PSItem
                    if ($ExceptionItem.Exception.Response.StatusCode -like "429") {
                        # Detected throttling based from response status code
                        $RetryInsecond = $ExceptionItem.Exception.Response.Headers["Retry-After"]
    
                        # Wait for given period of time specified in response headers
                        #Write-Verbose -Message "Graph is throttling the request, will retry in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "Unauthorized") {
                        #Write-Verbose -Message "Your Account does not have the relevent privilege to read this data. Please Elevate your account or contact the administrator"
                        $Script:PIMExpired = $tRUE
                        $GraphResponseProcess = $fALSE
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "GatewayTimeout") {
                        # Detected Gateway Timeout
                        $RetryInsecond = 30
    
                        # Wait for given period of time specified in response headers
                        #Write-Verbose -Message "Gateway Timeout detected, will retry in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "NotFound") {
                        #Write-Verbose -Message "The Device data could not be found"
                        $Script:StatusResult = $ExceptionItem.Exception.Response.StatusCode
                        $GraphResponseProcess = $fALSE
                    }
                    elseif ($PSItem.Exception.Message -like "*Invalid JSON primitive*") {
                        $Runcount++
                        if ($Runcount -eq 5) {
                            $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                            $Headers = New-AuthenticationHeader -AccessToken $AccessToken
                        }
                        if ($Runcount -ge 10) {
                            #Write-Verbose -Message "An Unrecoverable Error occured - Error: Invalid JSON primitive"
                            $GraphResponseProcess = $fALSE
                        }
                        $RetryInsecond = 5
                        #Write-Verbose -Message "Invalid JSON Primitive detected, Trying again in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                        
                    }
                    else {
                        try {
                            # Read the response stream
                            $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList @($ExceptionItem.Exception.Response.GetResponseStream())
                            $StreamReader.BaseStream.Position = 0
                            $StreamReader.DiscardBufferedData()
                            $ResponseBody = ($StreamReader.ReadToEnd() | ConvertFrom-Json)
                            
                            switch ($PSCmdlet.ParameterSetName) {
                                "GET" {
                                    # Output warning message that the request failed with error message description from response stream
                                    Write-Warning -Message "Graph request failed with status code $($ExceptionItem.Exception.Response.StatusCode). Error message: $($ResponseBody.error.message)"
    
                                    # Set graph response as handled and stop processing loop
                                    $GraphResponseProcess = $fALSE
                                }
                                default {
                                    # Construct new custom error record
                                    $SystemException = New-Object -TypeName "System.Management.Automation.RuntimeException" -ArgumentList ("{0}: {1}" -f $ResponseBody.error.code, $ResponseBody.error.message)
                                    $ErrorRecord = New-Object -TypeName "System.Management.Automation.ErrorRecord" -ArgumentList @($SystemException, $ErrorID, [System.Management.Automation.ErrorCategory]::NotImplemented, [string]::Empty)
    
                                    # Throw a terminating custom error record
                                    $PSCmdlet.ThrowTerminatingError($ErrorRecord)
                                }
                            }
    
                            # Set graph response as handled and stop processing loop
                            $GraphResponseProcess = $fALSE
                        }
                        catch [System.Exception] {
                            if ($PSItem.Exception.Message -like "*Invalid JSON primitive*") {
                                $Runcount++
                                if ($Runcount -ge 10) {
                                    #Write-Verbose -Message "An Unrecoverable Error occured - Error: Invalid JSON primitive"
                                    $GraphResponseProcess = $fALSE
                                }
                                $RetryInsecond = 5
                                #Write-Verbose -Message "Invalid JSON Primitive detected, Trying again in $($RetryInsecond) seconds"
                                Start-Sleep -second $RetryInsecond
                                
                            }
                            else {
                                Write-Warning -Message "Unhandled error occurred in function. Error message: $($PSItem.Exception.Message)"
    
                                # Set graph response as handled and stop processing loop
                                $GraphResponseProcess = $fALSE
                            }
                        }
                    }
                }
            }
            until ($GraphResponseProcess -eq $fALSE)
    
            # Handle return value
            return $GraphResponseList
            
        }
    }

    function New-AuthenticationHeader {
        <#
        .SYNOPSIS
            Construct a required header hash-table based on the access token from Get-MsalToken cmdlet.
        .DESCRIPTION
            Construct a required header hash-table based on the access token from Get-MsalToken cmdlet.
        .PARAMETER AccessToken
            Pass the AuthenticationResult object returned from Get-MsalToken cmdlet.
        .NOTES
            Author:      Nickolaj Andersen
            Contact:     @NickolajA
            Created:     2020-12-04
            Updated:     2020-12-04
            Version history:
            1.0.0 - (2020-12-04) Script created
        #>
        param(
            [parameter(Mandatory = $tRUE, HelpMessage = "Pass the AuthenticationResult object returned from Get-MsalToken cmdlet.")]
            [ValidateNotNullOrEmpty()]
            [Microsoft.Identity.Client.AuthenticationResult]$AccessToken
        )
        Process {
            # Construct default header parameters
            $AuthenticationHeader = @{
                "Content-Type"  = "application/json"
                "Authorization" = $AccessToken.CreateAuthorizationHeader()
                "ExpiresOn"     = $AccessToken.ExpiresOn.LocalDateTime
            }
    
            # Amend header with additional required parameters for bitLocker/recoveryKeys resource query
            $AuthenticationHeader.Add("ocp-client-name", "My App")
            $AuthenticationHeader.Add("ocp-client-version", "1.2")
    
            # Handle return value
            return $AuthenticationHeader
        }
    }

    function Clear-ResourceEnvironment {
        # Clear any PowerShell sessions created
        Get-PSSession | Remove-PSSession

        # Release an COM object created
        #$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Shell)

        # Perform garbage collection on session resources 
        [System.GC]::Collect()         
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        # Remove any custom varialbes created
        #Get-Variable -Name MyShell -ErrorAction Silently$VerbosePreference | Remove-Variable
    
    }

    

    # Variables - Customise these for your environment
    $Script:PIMExpired = $null
    $FileDate = Get-Date -Format 'yyyy_MM_dd'
    $FilePath = "C:\Temp\StaleComputers\"
    #$InterimFileLocation = "$($FilePath)InterimFiles\"
    $StaleDeviceReportFileName = "StaleDeviceReport - $($FileDate).xlsx"
    #$StaleDeviceReportRemotePath = "\\RemoteServer\Share\"         # Specify this if you would like to export the report to a share. I added this for remote reporting for another team
    #$StaleDeviceReportFile = "$($StaleDeviceReportRemotePath)$($StaleDeviceReportFileName)"
    $LocalStaleReportFile = "$($FilePath)Output\$($StaleDeviceReportFileName)"
    
    $ADForest = (Get-ADForest).RootDomain                           # Get the name of the Forest Root domain
    $DomainTargets = (Get-ADForest -Identity $ADForest).Domains     # Get the list of domains. This scales out for multidomain AD forests
    $ScriptStartTime = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $StaleDate = (Get-Date).AddDays(-90)                            # number of days before a device is classified as stale\dormant

    [string]$Resource = "deviceManagement/managedDevices"

    # Grab the system proxy for internet access
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    Clear-ResourceEnvironment
    if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    Try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
    catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
        
    #############################################################################################################################################
    # AzureAD Device Data Extraction
    #############################################################################################################################################

    Connect-AzureAD
    Write-Host "Extracting Data from AzureAD. Expected runtime is 7 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $AzureADDevices = [System.Collections.ArrayList]@(Get-AzureADDevice -All:$TRUE | Where-Object { $_.DeviceOSType -like "*Windows*" } | Select-Object @{Name = "AzureADDeviceID"; Expression = { $_.DeviceId.toString() } }, @{Name = "ObjectID"; Expression = { $_.ObjectID.toString() } }, @{Name = "AADEnabled"; Expression = { $_.AccountEnabled } }, @{Name = "AADApproximateLastLogonTimeStamp"; Expression = { (Get-Date -Date $_.ApproximateLastLogonTimeStamp -Format 'yyyy/MM/dd HH:mm') } }, @{Name = "AADDisplayName"; Expression = { $_.DisplayName } }, @{Name = "AADSTALE"; Expression = { if ($_.ApproximateLastLogonTimeStamp -le $StaleDate) { "TRUE" } elseif ($_.ApproximateLastLogonTimeStamp -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } } | Sort-Object azureADDeviceId )
    Write-Host "Collected Data for $($AzureADDevices.count) objects from AzureAD - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    Disconnect-AzureAD

    #$AzureADDevices | Export-Csv -Path "$($InterimFileLocation)STALEAzureADExtract.csv" -Delimiter ";" -NoTypeInformation
    #Remove-Variable -Name AzureADDevices -Force
    #Clear-ResourceEnvironment

    #############################################################################################################################################
    # Intune Managed Device Data Extraction
    #############################################################################################################################################

    Try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
    catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

    Write-Host "Extracting the data from MS Graph Intune. Expected runtime is 4 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $IntuneInterimArray = [System.Collections.ArrayList]::new()
    $IntuneInterimArray = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'" -Headers $AuthenticationHeader -Verbose | Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.azureADDeviceId.toString() } }, @{Name = "IntuneDeviceID"; Expression = { $_.id.ToString() } }, @{Name = "MSGraphDeviceName"; Expression = { $_.deviceName } }, @{Name = "enrolledDateTime"; Expression = { (Get-Date -Date $_.enrolledDateTime -Format "yyyy/MM/dd HH:mm") } }, @{Name = "MSGraphlastSyncDateTime"; Expression = { (Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") } }, @{Name = "MSGraphLastSyncStale"; Expression = { if ((Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") -le $StaleDate) { "TRUE" } elseif ((Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } }, @{Name = "UserUPN"; Expression = { $_.userPrincipalName } } | Sort-Object IntuneDeviceID)

    Write-Host "Collected Data for $($IntuneInterimArray.count) objects from MS Graph Intune - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #$IntuneInterimArray | Export-Csv -Path "$($InterimFileLocation)STALEIntuneInterimArray.csv" -Delimiter ";" -NoTypeInformation
    #Remove-Variable -Name IntuneInterimArray -Force
    #Clear-ResourceEnvironment

    #############################################################################################################################################
    # OnPrem AD Data Extraction
    #############################################################################################################################################

    $AllOPCompsArray = [System.Collections.ArrayList]::new()
    $RAWAllComps = [System.Collections.ArrayList]::new()
    $OPADProcessed = 0
    $OPADCount = $DomainTargets.Count

    Write-Host "Extracting AD OnPrem computer account Data - Expected Runtime is 3 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    foreach ( $DomainTarget in $DomainTargets ) {
        $OPADProcessed++
        [string]$ServerTarget = (Get-ADDomainController -Discover -DomainName $DomainTarget).HostName # Attempt to locate closest domain controller
       
        $OPDisplay = ( $OPADProcessed / $OPADCount ).tostring("P")
        Write-Progress -Activity "Extracting Data" -Status "Collecting Data from OnPrem AD - $($OPADProcessed) of $($OPADCount) - $($OPDisplay) Completed" -CurrentOperation "Extracting from $($DomainTarget) on $($ServerTarget)" -PercentComplete (( $OPADProcessed / $OPADCount ) * 100 )
        try{
            $Comps = [System.Collections.ArrayList]@(Get-ADComputer -Server $ServerTarget -Filter 'operatingsystem -like "Windows 10*"' -Properties CN, CanonicalName, objectGUID, LastLogonDate, Enabled -ErrorAction Stop)# | Select-Object CN, CanonicalName, objectGUID )
            if($?){
                Write-Host "Extracted $($Comps.count) devices from $($ServerTarget) for $($DomainTarget)"
            }
        }
        catch{
            try{
                $Comps = [System.Collections.ArrayList]@(Get-ADComputer -Server $ServerTarget -Filter 'operatingsystem -like "Windows 10*"' -Properties CN, CanonicalName, objectGUID, LastLogonDate, Enabled -ErrorAction Stop)# | Select-Object CN, CanonicalName, objectGUID )
                if($?){
                    Write-Host "Extracted $($Comps.count) devices from $($ServerTarget) for $($DomainTarget)"
                }
            }
            catch{
                Write-Host "Error Extracting from $($ServerTarget) for $($DomainTarget) - Error: $($_.Exception.Message)"
            }
        }
        $RAWAllComps += $Comps
        Remove-Variable -Name Comps -Force
    }
        
    Write-Host "Completed AD OnPrem computer account extraction - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    Write-Host "Standardising OnPrem AD Data - Expected Runtime is 2 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $AllOPCompsArray = [System.Collections.ArrayList]@($RAWAllComps | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.objectGUID.toString() } }, @{Name = "OPDeviceName"; Expression = { $_.CN } }, @{Name = "OPDeviceFQDN"; Expression = { "$($_.CN).$($_.CanonicalName.Split('/')[0])" } }, @{Name = "SourceDomain"; Expression = { "$($_.CanonicalName.Split('/')[0])" } }, @{Name = "OPLastLogonTS"; Expression = { (Get-Date -Date $_.LastLogonDate -Format "yyyy/MM/dd HH:mm") } }, @{Name = "OPSTALE"; Expression = { if ($_.LastLogonDate -le $StaleDate) { "TRUE" } elseif ($_.LastLogonDate -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } }, @{Name = "OPEnabled"; Expression = { $_.Enabled } } | Sort-Object azureADDeviceId )
    Remove-Variable -Name RAWAllComps -Force
    Clear-ResourceEnvironment

    Write-Host "Completed AD OnPrem Data Standardisation - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #$AllOPCompsArray | Export-Csv -Path "$($InterimFileLocation)STALEAllOPCompsArray.csv" -Delimiter ";" -NoTypeInformation
    #Remove-Variable -Name AllOPCompsArray -Force
    #Clear-ResourceEnvironment

    #############################################################################################################################################
    # Blending OnPrem AD data with MSGraph Intune Data
    #############################################################################################################################################

    Write-Host "Blending OnPrem AD Data Array with MS Graph Intune Data - Expected Runtime is 14 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    <#if ($AllOPCompsArray.count -lt 1) {
        $AllOPCompsArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)STALEAllOPCompsArray.csv" -Delimiter ";")
    }
    if ($IntuneInterimArray.count -lt 1) {
        $IntuneInterimArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)STALEIntuneInterimArray.csv" -Delimiter ";")
    }#>
        
    $IntuneInterimArray = [System.Collections.ArrayList]@($IntuneInterimArray | Sort-Object azureADDeviceId)
    $RAWAllDevPreProcArray = [System.Collections.ArrayList]@($IntuneInterimArray | LeftJoin-Object $AllOPCompsArray -On azureADDeviceId)
    $RAWAllPreDevNoIntuneDeviceID = [System.Collections.ArrayList]@($AllOPCompsArray | LeftJoin-Object $IntuneInterimArray -On azureADDeviceId)
    $RAWAllDevNoIntuneDeviceID = [System.Collections.ArrayList]@($RAWAllPreDevNoIntuneDeviceID | Where-Object { $_.IntuneDeviceID -like $null })
    Remove-Variable -Name IntuneInterimArray -Force
    Remove-Variable -Name RAWAllPreDevNoIntuneDeviceID -Force
    Remove-Variable -Name AllOPCompsArray -Force
    Clear-ResourceEnvironment
        
    Write-Host "Completed blending OnPrem AD Data Array with MS Graph Intune Data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    
    #############################################################################################################################################
    # Deduplicating the Blended Data
    #############################################################################################################################################

    Write-Host "Deduplicating blended data (OnPrem AD and MS Graph Intune Data) - Expected Runtime is 40 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    <#if ($AzureADDevices.count -lt 1) {
        $AzureADDevices = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)STALEAzureADExtract.csv" -Delimiter ";")
    }#>
        
    $RAWAllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevPreProcArray + $RAWAllDevNoIntuneDeviceID | Sort-Object AzureADDeviceID)
    $DDAllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevProcArray | Group-Object -Property AzureADDeviceID | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList)
    $AllDevProcArray = [System.Collections.ArrayList]@($AzureADDevices | LeftJoin-Object $DDAllDevProcArray -On AzureADDeviceID)
    Remove-Variable -Name DDAllDevProcArray -Force
    Clear-ResourceEnvironment

    $DDAllDevProcArray = [System.Collections.ArrayList]@($AllDevProcArray | Sort-Object IntuneDeviceID)
    Remove-Variable -Name AzureADDevices -Force
    Remove-Variable -Name RAWAllDevPreProcArray -Force
    Remove-Variable -Name RAWAllDevNoIntuneDeviceID -Force
    Clear-ResourceEnvironment

    #$DDAllDevProcArray | Export-Csv -Path "$($InterimFileLocation)STALEDDAllDevProcArray.csv" -Delimiter ";" -NoTypeInformation
    #Remove-Variable -Name DDAllDevProcArray -Force
    #Clear-ResourceEnvironment

    Write-Host "Completed deduplicating blended data (OnPrem AD and MS Graph Intune Data) - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #############################################################################################################################################
    # Exporting Stale Devices File
    #############################################################################################################################################

    <#if ($DDAllDevProcArray.count -lt 1) {
        $DDAllDevProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)STALEDDAllDevProcArray.csv" -Delimiter ";")
    }#>

    $AllDevices = [System.Collections.ArrayList]@($DDAllDevProcArray | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, @{Name = "TrueStale"; Expression = { if ($_.AADStale -notlike "False" -and $_.OPStale -notlike "False" -and $_.MSGraphLastSyncStale -notlike "False") { "TRUE" } else { "FALSE" } } }, @{Name = "AccountEnabled"; Expression = { if ($_.AADEnabled -notlike "False" -and $_.OPEnabled -notlike "False") { "TRUE" } else { "FALSE" } } })
    #$StaleDevices = [System.Collections.ArrayList]@($AllDevices | Where-Object { ($_.TrueStale -like "TRUE") } | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, TrueStale, AccountEnabled)
    $StaleDevices = [System.Collections.ArrayList]@($AllDevices | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, TrueStale, AccountEnabled)
    #$StaleDevices | Export-Excel -Path $StaleDeviceReportFile -ClearSheet -AutoSize -AutoFilter -Verbose:$VerbosePreference
    $StaleDevices | Export-Excel -Path $LocalStaleReportFile -ClearSheet -AutoSize -AutoFilter -Verbose:$VerbosePreference

    Write-Host "Completed Stale Device Data report - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm') - Start Time was $($ScriptStartTime)" -ForegroundColor Green
    
}