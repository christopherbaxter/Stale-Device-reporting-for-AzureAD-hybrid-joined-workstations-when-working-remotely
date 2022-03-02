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
Begin {}
Process {
    
    #############################################################################################################################################
    # Functions
    #############################################################################################################################################

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

    #############################################################################################################################################
    # Variables - Customise these for your environment
    #############################################################################################################################################
    
    $Script:PIMExpired = $null
    $FileDate = Get-Date -Format 'yyyy_MM_dd'
    $FilePath = "C:\Temp\StaleComputers\"
    #$InterimFileLocation = "$($FilePath)InterimFiles\"
    $DeviceExclusions = "$($FilePath)Input\Exclusions.csv"
    $StaleDeviceReportFileName = "StaleDeviceReport"
    $ResultsReportFileName = "StaleDeviceResults - $($FileDate)"
    $ResultsReportFile = "$($FilePath)Output\$($ResultsReportFileName)"
    
    $ADForest = (Get-ADForest).RootDomain                           # Get the name of the Forest Root domain
    $DomainTargets = (Get-ADForest -Identity $ADForest).Domains     # Get the list of domains. This scales out for multidomain AD forests
    
    # Grab the system proxy for internet access
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    #############################################################################################################################################
    # Get Authenticated
    #############################################################################################################################################

    Clear-ResourceEnvironment
    Connect-AzureAD

    #############################################################################################################################################
    # Import Data from the Stale Computer Report
    #############################################################################################################################################

    $MostRecentStaleDeviceReportFileName = @(Get-ChildItem -Path "$($FilePath)Output\" | Where-Object { $_.Name -like "$($StaleDeviceReportFileName)*.xlsx" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object name)
    $MostRecentStaleDeviceReportPath = "$($FilePath)Output\$($MostRecentStaleDeviceReportFileName.Name)"
    $Devices = [System.Collections.ArrayList]@(Import-Excel -Path $MostRecentStaleDeviceReportPath | Sort-Object $_.AzureADDeviceID)
    $UnfilteredStaleDevices = [System.Collections.ArrayList]@($Devices | Where-Object {($_.TrueStale -match "TRUE")})

    #############################################################################################################################################
    # Exclude Specific Devices
    #############################################################################################################################################

    $Exclusions = [System.Collections.ArrayList]@(Import-Csv -Path $DeviceExclusions -Delimiter ";" | Sort-Object $_.AzureADDeviceID)
    $StaleDevices = [System.Collections.ArrayList]@($UnfilteredStaleDevices | Where-Object {$_.AzureADDeviceID -notin $Exclusions.AzureADDeviceID})

    Remove-Variable -Name Devices -Force
    Remove-Variable -Name UnfilteredStaleDevices -Force
    Remove-Variable -Name Exclusions -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Disable On-Prem accounts
    #############################################################################################################################################

    $OPResults = @()
    foreach ( $DomainTarget in $DomainTargets ) {
        
        [string]$ServerTarget = (Get-ADDomainController -Discover -DomainName $DomainTarget).HostName # Attempt to locate closest domain controller
        $StaleOPDevices = [System.Collections.ArrayList]@($StaleDevices | Where-Object {($_.SourceDomain -eq "$($DomainTarget)") -and ($_.OPEnabled -eq "TRUE")} | Select-Object OPDeviceName,AzureADDeviceID,ObjectID)
        Write-Host "Disabling $($StaleOPDevices.count) Stale Devices in domain $($DomainTarget) against server $($ServerTarget)"

        foreach ( $StaleOPDevice in $StaleOPDevices ) {
            try{
                Set-ADComputer -Identity $StaleOPDevice.OPDeviceName -Server $ServerTarget -Enabled:$False -Confirm:$False -ErrorAction Stop
                if($?){
                    $OPResults += @($StaleOPDevice | Select-Object AzureADDeviceID,@{Name = "OPSuccessfullyDisabled";Expression = {"TRUE"}})
                }
            }
            catch{
                $OPResults += @($StaleOPDevice | Select-Object AzureADDeviceID,@{Name = "OPSuccessfullyDisabled";Expression = {"FALSE"}})
            }
        }
    }

    $OPResults = @($OPResults | Sort-Object $_.AzureADDeviceID)

    #############################################################################################################################################
    # Disable AzureAD accounts
    #############################################################################################################################################

    $StaleAZDevices = [System.Collections.ArrayList]@($StaleDevices | Where-Object {($_.AADEnabled -eq "TRUE")} | Select-Object AADDisplayName,AzureADDeviceID,ObjectID)
    $AZResults = @()
    foreach ($Device in $StaleAZDevices) {
        try{
            Set-AzureADDevice -ObjectId $Device.ObjectID -AccountEnabled $False -ErrorAction Stop
            if($?){
                $AZResults += @($Device | Select-Object AzureADDeviceID,@{Name = "AADSuccessfullyDisabled";Expression = {"TRUE"}})
            }
        }
        catch{
            $AZResults += @($Device | Select-Object AzureADDeviceID,@{Name = "AADSuccessfullyDisabled";Expression = {"FALSE"}})
        }
    }
    Disconnect-AzureAD

    $AZResults = @($AZResults | Sort-Object $_.AzureADDeviceID)

    #############################################################################################################################################
    # Create Results Reporting Array
    #############################################################################################################################################

    $OPDeviceResults = [System.Collections.ArrayList]@($StaleDevices | LeftJoin-Object $OPResults -On azureADDeviceId)
    $AllDeviceResults = [System.Collections.ArrayList]@($OPDeviceResults | LeftJoin-Object $AZResults -On azureADDeviceId)

    $AllDeviceResults | Export-Excel -Path $ResultsReportFile -Verbose

}