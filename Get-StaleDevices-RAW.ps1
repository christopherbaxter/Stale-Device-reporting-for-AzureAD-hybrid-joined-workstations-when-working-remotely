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
    Updated:     2023-11-24

    I wish to thank Anders Ahl and Nickolaj Andersen as well as the authors of the functions I have used below, without these guys, this script would not exist.

    You will notice that I have hashed out a number of lines in the script. I have done this for a very specific reason. I have had numberous issues with reliable data extraction
    from the various locations. The script will export the data in each section. In my environment, some of the extractions take more than an hour. The lines that have been hashed will allow for faster testing\troubleshooting
    
    You will also notice some strange code for retrying and for token renewal. By default, your token is valid for an hour, depending on the size of your estate,
    this could be WAY too short. Basically, the way the script has been hacked together, you will not have an issue with this, in most cases. I am able to keep the script running for 8 hours (the amount of time my elevated rights remain valid),
    without any trouble.

    Depending on the size of your environment, the RAM util gets quite high, the script also runs a little slowly as it moves along. I have added a number of 'cleanups' into the script, this helps significantly.
    
    In order to configure your Service Principal, follow this guide. https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/
#>
#Requires -Modules "MSAL.PS"
[CmdletBinding(SupportsShouldProcess = $TRUE)]
param(
    #PLEASE make sure you have specified your details below, else edit this and use the switches\variables in command line.
    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the Azure AD tenant ID.")]
    [ValidateNotNullOrEmpty()]
    #[string]$TenantID = "", # Populate this with your TenantID, this will then allow the script to run without asking for the details
    [string]$TenantID,

    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the service principal, also known as app registration, Client ID (also known as Application ID).")]
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
            <#For individual error testing
            $APIVersion = "Beta"
            $Resource = "deviceManagement/managedDevices"/bcf1365d-9a8f-4feb-aa0e-26d3ba23f693"
            $Headers = $AuthenticationHeader
            #>

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
                        #"Method"      = "Get"
                        "ErrorAction" = "Stop"
                        "Verbose"     = $TRUE
                    }
                    <#$RequestParams = @{
                        "Uri"         = $GraphURI
                        "Headers"     = $Headers
                        "Method"      = "Get"
                        "ErrorAction" = "Stop"
                        "Verbose"     = $TRUE
                    }#>
    
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
                        Write-Verbose -Message "NextLink: $($GraphURI)"
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

    # Variables

    $Script:PIMExpired = $null
    $FileDate = Get-Date -Format 'yyyy_MM_dd'
    $ExcelFileName = "FullDeviceReport"

    # Specify the output location of the report. The folder structure must include a folder called InterimFiles and Output
    # For example, if the filepath is left as is, you need to create the following directories:
    # C:\Temp\StaleComputers
    # C:\Temp\StaleComputers\InterimFiles
    # C:\Temp\StaleComputers\Output
    # The InterimFiles directory is used to output data at certain intervals, to allow for troubleshooting. In my estate, the data extract take many hours. This was required.
    $FilePath = "C:\Temp\StaleComputers\"
    $ConsolidatedReportFileName = "$($ExcelFileName)_$($FileDate).xlsx"
    $RemoteReportFileName = $ConsolidatedReportFileName

    # Name these files whatever you like
    $StaleDeviceReportFileName = "StaleDeviceReport - With Exclusions Removed - $($FileDate).xlsx"
    $AllStaleDeviceReportFileName = "ALLWorkstationDeviceReport - $($FileDate).xlsx"
    $LocalAllStaleDeviceReportFile = "$($FilePath)Output\$($AllStaleDeviceReportFileName)"
    $LocalStaleReportFile = "$($FilePath)Output\$($StaleDeviceReportFileName)"

    # if you want the data to be uploaded to a server share, perhaps for another team to consume as well, specify the location here.
    $RemoteFileLocation = "\\server.domain.FQDN\Share\StaleWorkstationReport"
    $RemoteReportExport = "$($RemoteFileLocation)\$($RemoteReportFileName)"

    # The below script will find the forest root domain, then find all the domains in the forest, if you are looking for workstations in a specific domain only, then modfy the Where-Object, you can also exclude domain with -notlike
    $ADForest = (Get-ADForest).RootDomain
    $DomainTargets = (Get-ADForest -Identity $ADForest).Domains #| Where-Object {($_ -like "domain.local")}

    # Specify the number of days you would like to test stale devices against. Some estate classify devices that have not been used for 90 days as stale, some 60 days. In the code below, the (-90) specifies 90 days as the stale measure. Set this to whatever is appropriate.
    $Date = (Get-Date).AddDays(-90)
    $StaleDate = Get-Date -Date $Date -Format "yyyy/MM/dd"

    # You can exclude devices from the report, or more accurately, label devices that are members of the below groups as devices to be excluded from being classified as stale. I found that in some cases (not all, which was confusing), these group names are case sensitive. I code for this later though.
    $ExclusionGroups = @('Win10 Autopilot';'Kiosks';'kiosks';'win10 autopilot')
    
    # This is the MSGraph API resource for Intune data. I believe there is a newer PS module that may be more efficient, but retrofitting my script at this stage may not be appropriate, I run this script in production, with little time to retrofit or update.
    [string]$Resource = "deviceManagement/managedDevices"

    # I was having trouble with extracting data from MSGraph API or AzureAD from time to time. I added the below snippets to allow PS to pull the proxy from the system and authenticate to the proxy.
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    # These lines of code checks for an access token, deletes it, and attempts to refresh the MsalToken, if this fails, it then gets a new token
    if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    Try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
    catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }

    # This will remove the authentication header, then create a new one using the new accesstoken
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
        
    #############################################################################################################################################
    # AzureAD Device Data Extraction
    #############################################################################################################################################

    Connect-AzureAD
    Write-Host "Extracting Data from AzureAD - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    
    # This section extracts data and transforms the data so it can be used later in the script. This is done as a single action. 
    # The following information is extracted from AzureAD: DeviceID, ObjectID, AccountEnabled, ApproximateLastLogonTimeStamp, DisplayName, DeviceOSType, then transformed as follows
    # DeviceID is renamed to AzureADDeviceID and stored as a string value.
    # ObjectID is stored as a string value.
    # AccountEnabled is renamed to AADEnabled.
    # ApproximateLastLogonTimeStamp is renamed to AADApproximateLastLogonTimeStamp and stored as a date\time.
    # DisplayName is renamed to AADDisplayName
    # Now for a little magic. AADStale is a calculated value that tests whether the ApproximateLastLogonTimeStamp date is older than the $Staledate value, or newer. If older, the value is “True”, if newer, the value is “False”.
    # DeviceOSType is stored as is.
    #
    # The array is sorted by AzureADDeviceID

    $AllAzureADDevices = [System.Collections.ArrayList]@()
    $AllAzureADDevices = [System.Collections.ArrayList]@(Get-AzureADDevice -All:$TRUE | Select-Object @{Name = "AzureADDeviceID"; Expression = { $_.DeviceId.toString() } }, @{Name = "ObjectID"; Expression = { $_.ObjectID.toString() } }, @{Name = "AADEnabled"; Expression = { $_.AccountEnabled } }, @{Name = "AADApproximateLastLogonTimeStamp"; Expression = { (Get-Date -Date $_.ApproximateLastLogonTimeStamp -Format 'yyyy/MM/dd') } }, @{Name = "AADDisplayName"; Expression = { $_.DisplayName } }, @{Name = "AADSTALE"; Expression = { if ($_.ApproximateLastLogonTimeStamp -le $StaleDate) { "TRUE" } elseif ($_.ApproximateLastLogonTimeStamp -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } }, DeviceOSType | Sort-Object azureADDeviceId )

    Write-Host "Collected Data for $($AllAzureADDevices.count) objects from AzureAD - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    # Use this if you need to troubleshoot the script.
    #$AllAzureADDevices | Export-Csv -Path "$($FilePath)InterimFiles\ALLAzureADExtract.csv" -Delimiter ";" -NoTypeInformation

    # Troubleshooting - If you need to speed up the process, you can import a previous data extract
    #$AllAzureADDevices = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\AzureADExtract-test.csv" -Delimiter ";")

    ##########################################################################################
    #       Extract Devices that are members of the groups listed in ExcludedGroupList
    ##########################################################################################

    $DExcludedGroupList = [System.Collections.ArrayList]@()
    $ExcludedGroupList = [System.Collections.ArrayList]@()
    foreach ($Excluded in $ExclusionGroups){
        # The only data we extract here is if the device is a member of any of the groups in the list, we take the DeviceID, rename it to AzureADDeviceID and store it as a string value, then we add a field called "StaleExempt" and set it as 'True'.
        $GroupExclusion = [System.Collections.ArrayList]@()
        $GroupExclusion = [System.Collections.ArrayList]@(Get-AzureADDevice -All:$True -Filter "startswith(DevicePhysicalIds,'[OrderId]:$($Excluded)')" | Select-Object @{Name = "AzureADDeviceID"; Expression = { $_.DeviceId.toString() } }, @{Name = "StaleExempt"; Expression = { "TRUE" } } )
        $DExcludedGroupList += $GroupExclusion
    }
    
    Disconnect-AzureAD

    # Here we deduplicate the data. I found that in some cases the group name in the list was case sensitive, and in some cases not, so to cater for this, and to ensure data completeness, I added the group names as they appear, and also in lowercase, meaning that we have duplicates, this is dealt with here.
    $ExcludedGroupList = [System.Collections.ArrayList]@($DExcludedGroupList | Sort-Object AzureADDeviceID -Descending | Group-Object -Property AzureADDeviceID | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList)
    
    # Use this if you need to troubleshoot the script.
    #$ExcludedGroupList | Export-Csv -Path "$($FilePath)InterimFiles\ExcludedDevices.csv" -Delimiter ";" -NoTypeInformation

    # This is used for memory management, more precisely, removing no longer needed data. In large environments, you will quickly overwhelm the ram you have available, these sections help.
    Remove-Variable -Name DExcludedGroupList -Force
    Clear-ResourceEnvironment

    ##########################################################################################
    #       Process Device Exclusions
    ##########################################################################################

    # The below lines are if you are troubleshooting, the data will be imported to speed up processing time
    <#
    if ($AllAzureADDevices.count -lt 1){
        $AllAzureADDevices = @(Import-Csv -Path "$($FilePath)InterimFiles\ALLAzureADExtract.csv" -Delimiter ";")
    }

    if ($ExcludedGroupList.count -lt 1){
        $ExcludedGroupList = @(Import-Csv -Path "$($FilePath)InterimFiles\ExcludedDevices.csv" -Delimiter ";")
    }
    #>
    
    # Here we 'blend' the AzureAD data extract with the excluded group list data extraction. We perform a 'leftjoin' using the azureADDeviceId data to ensure accuracy.
    $AzureADDevices = [System.Collections.ArrayList]@($AllAzureADDevices | LeftJoin-Object $ExcludedGroupList -On azureADDeviceId)
    
    # Use this if you need to troubleshoot the script.
    #$AzureADDevices | Export-Csv -Path "$($FilePath)InterimFiles\AzureADExtract.csv" -Delimiter ";" -NoTypeInformation
    #$AzureADDevices = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\AzureADExtract.csv" -Delimiter ";")

    # This is used for memory management, more precisely, removing no longer needed data. In large environments, you will quickly overwhelm the ram you have available, these sections help.
    Remove-Variable -Name AllAzureADDevices -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Intune Managed Device Data Extraction
    #############################################################################################################################################

    # These lines of code checks for an access token, deletes it, and refreshes the MsalToken
    if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop

    # This will remove the authentication header, then create a new one using the new accesstoken
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

    Write-Host "Extracting the data from MS Graph Intune - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $IntuneInterimArray = [System.Collections.ArrayList]::new()
    $IntuneInterimArray = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices" -Headers $AuthenticationHeader -Verbose | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.azureADDeviceId.toString() } }, @{Name = "IntuneDeviceID"; Expression = { $_.id.ToString() } }, @{Name = "MSGraphDeviceName"; Expression = { $_.deviceName } }, @{Name = "enrolledDateTime"; Expression = { (Get-Date -Date $_.enrolledDateTime -Format "yyyy/MM/dd") } }, @{Name = "MSGraphlastSyncDateTime"; Expression = { (Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd") } }, operatingSystem, osVersion, managementAgent, deviceRegistrationState, complianceState, @{Name = "UserUPN"; Expression = { $_.userPrincipalName } }, @{Name = "DeviceManufacturer"; Expression = { $_.manufacturer } }, @{Name = "DeviceModel"; Expression = { $_.model } }, managedDeviceName, @{Name = "MSGraphEncryptionState"; Expression = { $_.isEncrypted } }, aadRegistered, autopilotEnrolled, joinType )

    # The following data is extracted and transformed: azureADDeviceId, id, deviceName, enrolledDateTime, lastSyncDateTime, operatingSystem, osVersion, managementAgent, deviceRegistrationState, complianceState, userPrincipalName, manufacturer, model, managedDeviceName, isEncrypted, aadRegistered, autopilotEnrolled, joinType
    # Please keep in mind that the data that is extracted from this script is used by other teams to analyse data in an attempt to solve other issues, so there is a lot more here that is actually not really required to determine whether a device is stale. Keep it if you want, or remove it, keep in mind that you will have to modify other sections of the script as well to avoid errors.
    # AzureADDeviceID is stored as a string value.
    # ID is renamed to IntuneDeviceID and stored as a string value.
    # deviceName is renamed to MSGraphDeviceName.
    # EnrolledDateTime is stored as a date\time.
    # lastSyncDateTime is renamed to MSGraphLastSyncDateTime and is stored as a date\time.
    # The magic happens here: MSGraphLastSyncStale is a calculated value that tests whether lastSyncDateTime is older then $StaleDate or newer, if older, the value is stored as “True”, if newer, then “False”.
    # userPrincipalName is renamed to UserUPN.

    Write-Host "Collected Data for $($IntuneInterimArray.count) objects from MS Graph Intune - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    # Use this if you need to troubleshoot the script.
    #$IntuneInterimArray | Export-Csv -Path "$($FilePath)InterimFiles\IntuneInterimArray.csv" -Delimiter ";" -NoTypeInformation
    #$IntuneInterimArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\IntuneInterimArray.csv" -Delimiter ";")

    #############################################################################################################################################
    # OnPrem AD Data Extraction
    #############################################################################################################################################

    $AllOPCompsArray = [System.Collections.ArrayList]::new()
    $RAWAllComps = [System.Collections.ArrayList]::new()
    $OPADProcessed = 0
    $OPADCount = $DomainTargets.Count

    Write-Host "Extracting AD OnPrem computer account Data - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    foreach ( $DomainTarget in $DomainTargets ) {
        # Use this if you need to troubleshoot the script.
        #$DomainTarget = "Specificdomain.FQDN"

        # Selects the 'closest' appropriate Domain Controller, based on AD sites and Services
        [string]$ServerTarget = (Get-ADDomainController -Discover -DomainName $DomainTarget).HostName

        $OPADProcessed++
        # If you are having trouble with a DC that is in the same site as the machine you are running the script on, then you can use the below code snippet to 'hardcode' a specific DC (not recommended), but can solve some specific issues
        <#
        if ( $DomainTarget -eq "Specificdomain.FQDN" ) { [string]$ServerTarget = 'Specificdomain.domaincontroller.FQDN' }
        elseif ( $DomainTarget -eq "Specificdomain1.FQDN" ) { [string]$ServerTarget = 'Specificdomain1.domaincontroller.FQDN' }
        else { [string]$ServerTarget = (Get-ADDomainController -Discover -DomainName $DomainTarget).HostName }
        #>

        $OPDisplay = ( $OPADProcessed / $OPADCount ).tostring("P")
        Write-Progress -Activity "Extracting Data" -Status "Collecting Data from OnPrem AD - $($OPADProcessed) of $($OPADCount) - $($OPDisplay) Completed" -CurrentOperation "Extracting from $($DomainTarget) on $($ServerTarget)" -PercentComplete (( $OPADProcessed / $OPADCount ) * 100 )

        # The below will try to extract the data from the closest Domain Controller, and will handle an error is one occurs. This is not perfect, but could work for you.
        # The following data is extracted: CN, CanonicalName, objectGUID, LastLogonDate, whenChanged, DistinguishedName, OperatingSystem, Enabled
        try{
            $Comps = [System.Collections.ArrayList]@(Get-ADComputer -Server $ServerTarget -Filter 'operatingSystemVersion -like "10.*"' -Properties CN, CanonicalName, objectGUID, LastLogonDate, whenChanged, DistinguishedName, OperatingSystem, Enabled -ErrorAction Stop | Where-Object {($_.OperatingSystem -notlike "*Server*")} | Select-Object CN, CanonicalName, objectGUID, LastLogonDate, whenChanged, DistinguishedName, OperatingSystem, Enabled ) # | Select-Object CN, CanonicalName, objectGUID )
            if($?){
                Write-Host "Extracted $($Comps.count) devices from $($ServerTarget) for $($DomainTarget)"
            }
        }
        catch{
            try{
                # If you are having trouble with a DC that is in the same site as the machine you are running the script on, then you can use the below code snippet to 'hardcode' a specific DC (not recommended), but can solve some specific issues
                <#
                if ( $DomainTarget -eq "Specificdomain.FQDN" ) { [string]$ServerTarget = 'Specificdomain.domaincontroller.FQDN' }
                elseif ( $DomainTarget -eq "Specificdomain1.FQDN" ) { [string]$ServerTarget = 'Specificdomain1.domaincontroller.FQDN' }
                else { [string]$ServerTarget = (Get-ADDomainController -Discover -DomainName $DomainTarget).HostName }
                #>
                $Comps = [System.Collections.ArrayList]@(Get-ADComputer -Server $ServerTarget -Filter 'operatingSystemVersion -like "10.*"' -Properties CN, CanonicalName, objectGUID, LastLogonDate, whenChanged, DistinguishedName, OperatingSystem, Enabled -ErrorAction Stop | Where-Object {($_.OperatingSystem -notlike "*Server*")} | Select-Object CN, CanonicalName, objectGUID, LastLogonDate, whenChanged, DistinguishedName, OperatingSystem, Enabled ) # | Select-Object CN, CanonicalName, objectGUID )
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

    Write-Host "Standardising OnPrem AD Data - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    # This line transforms\normalises the data extracted from on-prem AD
    $AllOPCompsArray = [System.Collections.ArrayList]@($RAWAllComps | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.objectGUID.toString() } }, @{Name = "OPDeviceName"; Expression = { $_.CN } }, @{Name = "OPDeviceFQDN"; Expression = { "$($_.CN).$($_.CanonicalName.Split('/')[0])" } }, @{Name = "SourceDomain"; Expression = { "$($_.CanonicalName.Split('/')[0])" } }, @{Name = "OPLastUpdateTS"; Expression = { (Get-Date -Date $_.whenChanged -Format "yyyy/MM/dd") } }, @{Name = "OPLastLogonTS"; Expression = { (Get-Date -Date $_.LastLogonDate -Format "yyyy/MM/dd") } }, @{Name = "OPSTALE"; Expression = { if ($_.LastLogonDate -le $StaleDate) { "TRUE" } elseif ($_.LastLogonDate -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } }, @{Name = "OPDistinguishedName"; Expression = { $_.DistinguishedName } }, @{Name = "OPOperatingSystem"; Expression = { $_.OperatingSystem } }, @{Name = "OPEnabled"; Expression = { $_.Enabled } } | Sort-Object azureADDeviceId )

    #whenChanged, DistinguishedName, OperatingSystem, Enabled
    # objectGUID is renamed to azureADDeviceId and stored as a string value.
    # CN is renamed to OPDeviceName.
    # OPDeviceFQDN is a calculated value using CN and CanonicalName.
    # SourceDomain is a calculated value extracted from CanonicalName.
    # LastLogonDate is renamed to OPLastLogonTS and stored as date\time.
    # whenChanged is renamed to OPLastupdateTS and stored as date\time.
    # OPSTALE is a calculated value using LastLogonDate and $StaleDate.
    # Enabled is renamed to OPEnabled.
    # The data is sorted by azureADDeviceID.

    # This is used for memory management, more precisely, removing no longer needed data. In large environments, you will quickly overwhelm the ram you have available, these sections help.
    Remove-Variable -Name RAWAllComps -Force
    Clear-ResourceEnvironment

    Write-Host "Completed AD OnPrem Data Standardisation - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    # Use this if you need to troubleshoot the script.
    #$AllOPCompsArray | Export-Csv -Path "$($FilePath)InterimFiles\AllOPCompsArray.csv" -Delimiter ";" -NoTypeInformation
    #$AllOPCompsArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\AllOPCompsArray.csv" -Delimiter ";")

    #############################################################################################################################################
    # Blending OnPrem AD data with MSGraph Intune Data
    #############################################################################################################################################

    Write-Host "Blending OnPrem AD Data Array with MS Graph Intune Data - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    # The below lines are if you are troubleshooting, the data will be imported to speed up processing time.
    <#
    if ($AllOPCompsArray.count -lt 1) {
        $AllOPCompsArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\AllOPCompsArray.csv" -Delimiter ";")
    }
    if ($IntuneInterimArray.count -lt 1) {
        $IntuneInterimArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\IntuneInterimArray.csv" -Delimiter ";")
    }
    #>
    
    # Sorts the Intune data by azureADDeviceId to speed up processing time
    $IntuneInterimArray = [System.Collections.ArrayList]@($IntuneInterimArray | Sort-Object azureADDeviceId)

    # In order to ensure that there are no missing devices in the report, the on-prem and Intune data is joined like the below, then deduplicated later, this ensures no data is lost, or devices are missed.
    $RAWAllDevPreProcArray = [System.Collections.ArrayList]@($IntuneInterimArray | LeftJoin-Object $AllOPCompsArray -On azureADDeviceId)
    $RAWAllPreDevNoIntuneDeviceID = [System.Collections.ArrayList]@($AllOPCompsArray | LeftJoin-Object $IntuneInterimArray -On azureADDeviceId)
    $RAWAllDevNoIntuneDeviceID = [System.Collections.ArrayList]@($RAWAllPreDevNoIntuneDeviceID | Where-Object { $_.IntuneDeviceID -like $null })
    Remove-Variable -Name IntuneInterimArray -Force
    Remove-Variable -Name AllOPCompsArray -Force
    Clear-ResourceEnvironment
    
    Write-Host "Completed blending OnPrem AD Data Array with MS Graph Intune Data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    
    #############################################################################################################################################
    # Deduplicating the Blended Data
    #############################################################################################################################################

    Write-Host "Deduplicating blended data (OnPrem AD and MS Graph Intune Data) - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    
    # The below lines are if you are troubleshooting, the data will be imported to speed up processing time.
    <#
    if ($AzureADDevices.count -lt 1) {
        $AzureADDevices = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\AzureADExtract.csv" -Delimiter ";")
    }
    #>

    # Here we add the 2 arrays from the previous step into a single array, then deduplicate the data to ensure we dont have duplicates, of which there were many.
    $RAWAllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevPreProcArray + $RAWAllDevNoIntuneDeviceID | Sort-Object AzureADDeviceID)
    $DDAllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevProcArray | Group-Object -Property AzureADDeviceID | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList)
    
    # Here we join the AzureAD data with the on-prem and Intune data, joined in the previous step.
    $AllDevProcArray = [System.Collections.ArrayList]@($AzureADDevices | LeftJoin-Object $DDAllDevProcArray -On AzureADDeviceID)

    # This is used for memory management, more precisely, removing no longer needed data. In large environments, you will quickly overwhelm the ram you have available, these sections help.
    Remove-Variable -Name DDAllDevProcArray -Force
    Clear-ResourceEnvironment

    # Here we sort the data into a new array. I did this in this way because I removed some steps that caused the array names to be incorrect later. This goes to solve that, without me needing to correct the code too much.
    $DDAllDevProcArray = [System.Collections.ArrayList]@($AllDevProcArray | Sort-Object IntuneDeviceID)

    # This is used for memory management, more precisely, removing no longer needed data. In large environments, you will quickly overwhelm the ram you have available, these sections help.
    Remove-Variable -Name AzureADDevices -Force
    Remove-Variable -Name RAWAllDevPreProcArray -Force
    Remove-Variable -Name RAWAllDevNoIntuneDeviceID -Force
    Clear-ResourceEnvironment

    # Use this if you need to troubleshoot the script.
    #$DDAllDevProcArray | Export-Csv -Path "$($FilePath)InterimFiles\STALEDDAllDevProcArray.csv" -Delimiter ";" -NoTypeInformation

    Write-Host "Completed deduplicating blended data (OnPrem AD and MS Graph Intune Data) - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #############################################################################################################################################
    # Extracting Stale Devices
    #############################################################################################################################################

    # The below lines are if you are troubleshooting, the data will be imported to speed up processing time.
    <#
    if ($DDAllDevProcArray.count -lt 1) {
        $DDAllDevProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\STALEDDAllDevProcArray.csv" -Delimiter ";")
    }
    #>

    # Now for the magic. This one line is what makes this all work.
    $AllDevices = [System.Collections.ArrayList]@($DDAllDevProcArray | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, operatingSystem, osVersion, managementAgent, deviceRegistrationState, complianceState, DeviceManufacturer, DeviceModel, DeviceSN, OPDistinguishedName, managedDeviceName, MSGraphEncryptionState, aadRegistered, autopilotEnrolled, joinType, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastUpdateTS, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, @{Name = "TrueStale"; Expression = { if ($_.AADStale -notlike "False" -and $_.OPStale -notlike "False" -and $_.MSGraphLastSyncStale -notlike "False") { "TRUE" } else { "FALSE" } } }, @{Name = "AccountEnabled"; Expression = { if ($_.AADEnabled -notlike "False" -and $_.OPEnabled -notlike "False") { "TRUE" } else { "FALSE" } } }, DeviceOSType, OPOperatingSystem, StaleExempt )
    # We take all the data we have collected, and we calculate whether a device is truly stale, or not. 
    # The analysis works like this, if a device is not stale, in either Intune, AzureAD or on-prem AD, then the device is not stale.
    # Here is how I calculate that, using an 'if' statement:
    # If AADStale is not "False" and OPStale is not "False" and MSGraphLastSyncStale is not "False" then "TrueStale" is set to "TRUE". If any of the fields are "False", then the device is NOT stale, and "TrueStale" is set to "FALSE".
    # I calculate it this way as sometimes a device does not exist in one of the datastores\databases, and in this case, the data is blank or null in that field, if I was matching "True", these devices shored as not stale, but they clearly were.
    #
    # Aditionally, I also test if a device is Enabled, in a similar way.
    # If AADEnabled is not "False" and OPEnabled is not "False", then "AccountEnabled" is set to "TRUE" else "AccountEnabled" is set to "FALSE"

    # Here is where things get more interesting, because I also extract the on-prem 'modified' date, I can calculate when a device was disabled, or at least theoretically... lets get into it, and I'll explain.
    $StaleDevices = [System.Collections.ArrayList]@($AllDevices | Where-Object { ($_.TrueStale -like "TRUE") -and ($_.StaleExempt -notlike "TRUE") } | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, operatingSystem, osVersion, managementAgent, deviceRegistrationState, complianceState, DeviceManufacturer, DeviceModel, DeviceSN, OPDistinguishedName, managedDeviceName, MSGraphEncryptionState, aadRegistered, autopilotEnrolled, joinType, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastUpdateTS, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, TrueStale, AccountEnabled, @{Name = "OPDisabledStale"; Expression = { if ($_.OPEnabled -notlike "True" -and $_.OPEnabled -notlike "False" ) { $null } elseif (($_.OPLastUpdateTS -le $StaleDate) -and ($_.OPEnabled -like "False") -and ($_.TrueStale -like "True")) { "TRUE" } else { "FALSE" } } }, DeviceOSType, OPOperatingSystem, StaleExempt )
    # So we have extracted devices from the data that are TrueStale, and also not part of the exclusion groups. We did this by extracting machines that have "True" in the TrueStale field, and do not have "True" in the StaleExempt field.
    # Then I do calculate if the device is disabled, when the last update on the account was, assuming disabling the object was the last update on the account. This is so I can delete computer accounts that are disabled, and have been in this state for more than 90 days.
    
    ####
    # MASSIVE CAVEAT!!! If the domain controller you used to extract the data is newly promoted, the last modification date on the object is the date the object was replicated into the newly promoted DC. Interesting, isnt it?
    ####

    # If OPEnabled is not "True" and OPEnabled is not "False" (dealing with a data extraction error) set "OPDisabledStale" to $null (blank), and moves on to the next device.
    # But, if OPLastUpdateTS older than $StaleDate (date script is run, minus specified days, in this case 90) and OPEnabled is "False" and TrueStale is "True", then set "OPDisabledStale" to "TRUE" else set "OPDisabledStale" to "FALSE"
    # Basically, what this means is that if the device is disabled, and the last update on the on-prem computer account is longer than 90 days ago, then the device is ready for deletion.

    # This array is the same as the one above, but includes all the devices. This data is used to analyse the estate to find data discrepencies. I have seen team use this as the 'golden source'.
    $AllStaleDevices = [System.Collections.ArrayList]@($AllDevices | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, operatingSystem, osVersion, managementAgent, deviceRegistrationState, complianceState, DeviceManufacturer, DeviceModel, DeviceSN, OPDistinguishedName, managedDeviceName, MSGraphEncryptionState, aadRegistered, autopilotEnrolled, joinType, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastUpdateTS, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, TrueStale, AccountEnabled, @{Name = "OPDisabledStale"; Expression = { if ($_.OPEnabled -notlike "True" -and $_.OPEnabled -notlike "False" ) { $null } elseif (($_.OPLastUpdateTS -le $StaleDate) -and ($_.OPEnabled -like "False") -and ($_.TrueStale -like "True")) { "TRUE" } else { "FALSE" } } }, DeviceOSType, OPOperatingSystem, StaleExempt )
    
    $LocalAllStaleDeviceReportFile = "$($FilePath)Output\$($AllStaleDeviceReportFileName)"

    # Here I export the data to files. I have been having troulbe recently with the Excel export, so export to both Excel and CSV, both locally, and to a remote file share.
    $AllStaleDevices | Export-Excel -Path $LocalAllStaleDeviceReportFile -ClearSheet -AutoSize -AutoFilter -Verbose:$VerbosePreference
    $AllStaleDevices | Export-CSV -Path "$($LocalAllStaleDeviceReportFile).csv" -Delimiter ";" -NotypeInformation -Verbose:$VerbosePreference
   
    $StaleDevices | Export-Excel -Path $StaleDeviceReportFile -ClearSheet -AutoSize -AutoFilter -Verbose:$VerbosePreference
    $StaleDevices | Export-CSV -Path "$($StaleDeviceReportFile).csv" -Delimiter ";" -NotypeInformation -Verbose:$VerbosePreference
    $StaleDevices | Export-Excel -Path $LocalStaleReportFile -ClearSheet -AutoSize -AutoFilter -Verbose:$VerbosePreference
    $StaleDevices | Export-CSV -Path "$($LocalStaleReportFile).csv" -Delimiter ";" -NotypeInformation -Verbose:$VerbosePreference

    # Remember to set the appropriate remote path variable
    $StaleDevices | Export-Excel -Path $RemoteReportExport -ClearSheet -AutoSize -AutoFilter -Verbose:$VerbosePreference
    $StaleDevices | Export-CSV -Path "$($RemoteReportExport).csv" -Delimiter ";" -NotypeInformation -Verbose:$VerbosePreference
    
}