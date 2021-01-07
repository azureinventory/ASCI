##########################################################################################
#                                                                                        #
#          * Azure Security Center Inventory ( ASCI ) Report Generator *                 #
#                                                                                        #
#       Version: 0.0.8                                                                   #
#       Authors: Claudio Merola <clvieira@microsoft.com>                                 #
#                Renato Gregio <renato.gregio@microsoft.com>                             #
#                                                                                        #
#       Date: 01/07/2020                                                                 #
#                                                                                        #
#           https://github.com/RenatoGregio/AzureSecurityCenterInventory                 #
#                                                                                        #
#                                                                                        #
#        DISCLAIMER:                                                                     #
#        Please note that while being developed by Microsoft employees,                  #
#        Azure Resource Inventory is not a Microsoft service or product.                 #
#                                                                                        #
#        Azure Resource Inventory is a personal driven project, there are none implicit  #
#        or explicit obligations related to this project, it is provided 'as is' with    #
#        no warranties and confer no rights.                                             #
#                                                                                        #
##########################################################################################
<#
.SYNOPSIS  
    This script creates Excel file to Analyze Azure Security Center inside a Tenant
  
.DESCRIPTION  
    Do you want to analyze your security center Advisories in a table format? Document it in xlsx format.
 
.PARAMETER TenantID
    Specify the tenant ID you want to create a Resource Inventory.
    
    >>> IMPORTANT: YOU NEED TO USE THIS PARAMETER FOR TENANTS WITH MULTI-FACTOR AUTHENTICATION. <<< 
 
.PARAMETER SubscriptionID
    Use this parameter to collect a specific Subscription in a Tenant

.PARAMETER AllStatus
    By Default Azure Security Center Inventory catch only "unhealthy" advisory Status. Use this parameter for all Advisories, including "heanthy" and "NotApplicable".
    This option can increase considerably your collect time. 
    See Microsoft Docs for best understandment: https://docs.microsoft.com/en-us/azure/security-center/security-center-recommendations

.PARAMETER Debug
    Execute ASCI in debug mode. 

.EXAMPLE
    Default utilization. Read all tenants you have privileges, select a tenant in menu and collect from all subscriptions:
    PS C:\> .\AzureSecurityCenterInventory.ps1

    Read all tenants you have privileges, select a tenant in menu and collect from All Status all subscriptions:
    PS C:\>.\AzureSecurityCenterInventory.ps1 -AllStatus

    Define the Tenant ID and collect all "unhealthy" Security Advisories:
    PS C:\> .\AzureSecurityCenterInventory.ps1 -TenantID <your-Tenant-Id>

    Define the Tenant ID and collect all "unhealthy" Security Advisories for a specific Subscription:
    PS C:\>.\AzureSecurityCenterInventory.ps1 -TenantID <your-Tenant-Id> -SubscriptionID <your-Subscription-Id>
    
    Define the Tenant ID and collect all Security Advisories:
    PS C:\>.\AzureSecurityCenterInventory.ps1 -TenantID <your-Tenant-Id> -AllStatus
    
    Define the Tenant ID and collect all Security Advisories for a specific Subscription:
    PS C:\>.\AzureSecurityCenterInventory.ps1 -TenantID <your-Tenant-Id> -SubscriptionID <your-Subscription-Id> -AllStatus

.NOTES
    AUTHOR: Claudio Merola and Renato Gregio - Customer Engineer - Customer Success Unit | Azure Infrastucture/Automation/Devops/Governance | Microsoft

.LINK
    https://github.com/azureinventory
    Please note that while being developed by a Microsoft employee, Azure inventory Scripts is not a Microsoft service or product. Azure Inventory Scripts are a personal driven project, there are none implicit or explicit obligations related to this project, it is provided 'as is' with no warranties and confer no rights.
#>
param ($TenantID, $AllStatus, $SubscriptionID) 

$Runtime = Measure-Command -Expression {

    if ($DebugPreference -eq 'Inquire') {
        $DebugPreference = 'Continue'
    }

    $ErrorActionPreference = "silentlycontinue"
    $DesktopPath = "C:\AzureInventory"
    $CSPath = "$HOME/AzureInventory"
    $Global:Subscriptions = ''


    <######################################### Environment #########################################>


    #### Generic Conditional Text rules, Excel style specifications for the spreadsheets and tables:
    $tableStyle = "Light20"
    Write-Debug ('Excel Table Style used: ' + $tableStyle)

    #### Number of Resource Types to be considered in the script:
    $ResourceTypes = 100
    Write-Debug ('Number of Resource Types considered in Excel: ' + $ResourceTypes)

    <######################################### Help ################################################>

    function usageMode() {
        Write-Output "" 
        Write-Output "" 
        Write-Output "Usage: "
        Write-Output "For CloudShell:"
        Write-Output "./AzureSecurityCenterInventory.ps1"      
        Write-Output ""
        Write-Output "For PowerShell Desktop:"      
        Write-Output "./AzureSecurityCenterInventory.ps1 -TenantID <Azure Tenant ID> -SubscriptionID <Subscription ID>"
        Write-Output "" 
        Write-Output "For All Security Advisory Status:"      
        Write-Output "./AzureSecurityCenterInventory.ps1 -AllStatus"
        Write-Output "" 
        Write-Output "This option can increase considerably the amount of time to collect"         
        Write-Output "" 
    }

    <######################################### Environment #########################################>

    function checkAzCli() {
        Write-Host "Validating Az Cli.."
        $azcli = az --version
        if ($null -eq $azcli) {
            Read-Host "Azure CLI Not Found. Press <Enter> to finish script"
            Exit
        }
        Write-Host "Validating Az Cli Extension.."
        $azcliExt = az extension list --output json | ConvertFrom-Json
        if ($azcliExt.name -notin 'resource-graph') {
            Write-Host "Adding Az Cli Extension"
            az extension add --name resource-graph 
        }
        Write-Host "Validating ImportExcel Module.."
        $VarExcel = Get-InstalledModule -Name ImportExcel -ErrorAction silentlycontinue
        if ($null -eq $VarExcel) {
            Write-Host "Trying to install ImportExcel Module.."
            Install-Module -Name ImportExcel -Force
        }
        $VarExcel = Get-InstalledModule -Name ImportExcel -ErrorAction silentlycontinue
        if ($null -eq $VarExcel) {
            Read-Host 'Admininstrator rights required to install ImportExcel Module. Press <Enter> to finish script'
            Exit
        }
    }
    function LoginSession() {
        $Global:DefaultPath = "$DesktopPath\"
        if ($TenantID -eq '' -or $null -eq $TenantID) {
            write-host "Tenant ID not specified. Use -TenantID parameter if you want to specify directly. "        
            write-host "Authenticating Azure"
            write-host ""
            az account clear | Out-Null
            az login | Out-Null
            write-host ""
            write-host ""
            $Tenants = az account list --query [].homeTenantId -o tsv --only-show-errors | Get-Unique
                
            if ($Tenants.Count -eq 1) {
                write-host "You have privileges only in One Tenant "
                write-host ""
                $TenantID = $Tenants
            }
            else { 
                write-host "Select the the Azure Tenant ID that you want to connect : "
                write-host ""
                $SequenceID = 1
                foreach ($TenantID in $Tenants) {
                    write-host "$SequenceID)  $TenantID"
                    $SequenceID ++ 
                }
                write-host ""
                [int]$SelectTenant = read-host "Select Tenant ( default 1 )"
                $defaultTenant = --$SelectTenant
                $TenantID = $Tenants[$defaultTenant]
            }
    
            write-host "Extracting from Tenant $TenantID"
            $Global:Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json
            $Global:Subscriptions = $Subscriptions | Where-Object { $_.tenantID -eq $TenantID }
            if ($SubscriptionID) {
                $Global:Subscriptions = $Subscriptions | Where-Object { $_.ID -eq $SubscriptionID }
            }
        }
    
        else {
            az account clear | Out-Null
            az login -t $TenantID | Out-Null
            $Global:Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json
            $Global:Subscriptions = $Subscriptions | Where-Object { $_.tenantID -eq $TenantID }
            if ($SubscriptionID) {
                $Global:Subscriptions = $Subscriptions | Where-Object { $_.ID -eq $SubscriptionID }
            }
        }
    }
    function checkPS() {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            $Global:PSEnvironment = "Desktop"
            write-host "PowerShell Desktop Identified."
            write-host ""
            LoginSession
        }
        else {
            $Global:PSEnvironment = "CloudShell"
            write-host 'Azure CloudShell Identified.'
            write-host ""
            <#### For Azure CloudShell change your StorageAccount Name, Container and SAS for Grid Extractor transfer. ####>
            $Global:DefaultPath = "$CSPath/" 
            $Global:Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json
        }
    }

    <######################################### Checking PowerShell #########################################>

    checkAzCli
    checkPS

    #### Creating Excel file variable:
    $Global:File = ($DefaultPath + "AzureSecurityCenter_Report_" + (get-date -Format "yyyy-MM-dd_HH_mm") + ".xlsx")
    Write-Debug ('Excel file:' + $File)

    <######################################### Subscriptions #########################################>
    Write-Progress -activity 'Azure Inventory' -Status "1% Complete." -PercentComplete 1 -CurrentOperation 'Discovering Subscriptions..'

    $SubCount = $Subscriptions.count

    Write-Debug ('Number of Subscriptions Found: ' + $SubCount)
    Write-Progress -activity 'Azure Inventory' -Status "3% Complete." -PercentComplete 3 -CurrentOperation "$SubCount Subscriptions found.."

    if ((Test-Path -Path $DefaultPath -PathType Container) -eq $false) {
        New-Item -Type Directory -Force -Path $DefaultPath | Out-Null
    }

    <######################################### Extracting Security Center #########################################>

    Write-Progress -activity 'Azure Inventory' -Status "6% Complete." -PercentComplete 6 -CurrentOperation "Executing Security Advisories extraction jobs.."
    Write-Host " Azure Resource Inventory are collecting Security Center Advisories."
    Write-Host " "
    Write-Host " "
    Write-Debug ('Extracting total number of Security Advisories from Tenant')

    if ($AllStatus.IsPresent -and $SubscriptionID) {
        $SecSize = az graph query  -q  "securityresources | where subscriptionId == '$SubscriptionID' |  summarize count()"  --output json --only-show-errors | ConvertFrom-Json    
    }
    elseif ($AllStatus.IsPresent) {
        $SecSize = az graph query -q  "securityresources | summarize count()" --output json --only-show-errors | ConvertFrom-Json
    }
    elseif ($SubscriptionID) {
        $SecSize = az graph query -q  "securityresources | where properties['status']['code'] == 'Unhealthy' | where subscriptionId == '$SubscriptionID'  | summarize count()" --output json --only-show-errors | ConvertFrom-Json
    }
    else {
        $SecSize = az graph query -q  "securityresources | where properties['status']['code'] == 'Unhealthy' | summarize count()" --output json --only-show-errors | ConvertFrom-Json
    }
    
    $SecSizeNum = $SecSize.'count_'

    if ($SecSizeNum -ge 1) {
        $Loop = $SecSizeNum / 1000
        $Loop = [math]::ceiling($Loop)
        $Looper = 0
        $Limit = 0
        $Global:Sec = @()
        while ($Looper -lt $Loop) {
            $Looper ++
            Write-Progress -activity 'Azure Security Inventory' -Status "$Looper / $Loop" -PercentComplete (($Looper / $Loop) * 100) -Id 1
            if ($AllStatus.IsPresent -and $SubscriptionID) {
                $SecCenter = az graph query  -q  "securityresources | where subscriptionId == '$SubscriptionID'" --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json
            }
            elseif ($AllStatus.IsPresent) {
                $SecCenter = az graph query -q  "securityresources" --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json	
            }
            elseif ($SubscriptionID) {
                $SecCenter = az graph query -q  "securityresources | where properties['status']['code'] == 'Unhealthy'| where subscriptionId == '$SubscriptionID'" --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json
            }
            else {
                $SecCenter = az graph query -q  "securityresources | where properties['status']['code'] == 'Unhealthy'" --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json
            }
            $Global:Sec += $SecCenter
            Start-Sleep 1
            $Limit = $Limit + 1000
        }
        Write-Progress -activity 'Azure Security Inventory' -Status "$Looper / $Loop" -Id 1 -Completed
    }
}
            
<######################################### Importing Data to Excel  #########################################>

function ImportDataExcel {

    $Subs = $Subscriptions

    Write-Progress -activity 'Azure Inventory' -Status "15% Complete." -PercentComplete 15 -CurrentOperation "Starting to process extraction data.."         

    $obj = ''
    $tmp = @()

    $cp = 0

    foreach ($1 in $Sec) {

        $cp ++
        Write-Progress -activity 'Processing Security Inventory' -PercentComplete (($cp / $Sec.count) * 100) -Id 1
        
        $data = $1.PROPERTIES

        $sub1 = $Subs | Where-Object { $_.id -eq $1.properties.resourceDetails.Id.Split("/")[2] }

        $obj = @{
            'Subscription'       = $sub1.Name;
            'Resource Group'     = $1.RESOURCEGROUP;
            'Resource Type'      = $data.resourceDetails.Id.Split("/")[7];
            'Resource Name'      = $data.resourceDetails.Id.Split("/")[8];
            'Categories'         = [string]$data.metadata.categories;
            'Control'            = $data.displayName;
            'Severity'           = $data.metadata.severity;
            'Status'             = $data.status.code;
            'Remediation'        = $data.metadata.remediationDescription;
            'Remediation Effort' = $data.metadata.implementationEffort;
            'User Impact'        = $data.metadata.userImpact;
            'Threats'            = [string]$data.metadata.threats
        }    
        $tmp += $obj
    }

    Write-Progress -activity 'Processing Security Inventory' -Completed -Id 1


    #### Security Center worksheet is always the second sequence:
    Write-Progress -activity 'Azure Inventory' -Status "25% Complete." -PercentComplete 25 -CurrentOperation "Processing Security Center Advisories"         
    Write-Debug ('Generating Security Center sheet.')

    $condtxtsec = $(New-ConditionalText High -Range G:G
        New-ConditionalText High -Range L:L)


    Write-Progress -activity 'Azure Inventory' -Status "30% Complete." -PercentComplete 30 -CurrentOperation "Processing Security Center Advisories"    

    $tmp | 
    ForEach-Object { [PSCustomObject]$_ } | 
    Select-Object 'Subscription',
    'Resource Group',
    'Resource Type',
    'Resource Name',
    'Categories',
    'Control',
    'Severity',
    'Status',
    'Remediation',
    'Remediation Effort',
    'User Impact',
    'Threats' | 
    Export-Excel -Path $File -WorksheetName 'SecurityCenter' -AutoSize -TableName 'SecurityCenter' -TableStyle $tableStyle -ConditionalText $condtxtsec -KillExcel 

    Write-Progress -activity 'Azure Inventory' -Status "50% Complete." -PercentComplete 50 -CurrentOperation "Processing Security Center Advisories"    

    <################################################################### CHARTS ###################################################################>

    Write-Debug ('Generating Overview sheet (Charts).')
    "" | Export-Excel -Path $File -WorksheetName 'Overview' -MoveToStart 
    $excel = Open-ExcelPackage -Path $file -KillExcel

    $PTParams = @{
        PivotTableName    = "P0"
        Address           = $excel.Overview.cells["B3"] # top-left corner of the table
        SourceWorkSheet   = $excel.SecurityCenter
        PivotRows         = @("Severity")
        PivotData         = @{"Severity" = "Count" }
        PivotTableStyle   = $tableStyle
        IncludePivotChart = $true
        ChartType         = "Pie"
        ChartRow          = 11 # place the chart below row 22nd
        ChartColumn       = 0
        Activate          = $true
        PivotFilter       = 'Subscription ID'
        ChartTitle        = 'Security Center Severity'
        ShowPercent       = $true
        ChartHeight       = 250
        ChartWidth        = 400
    }

    Add-PivotTable @PTParams

    $PTParams = @{
        PivotTableName    = "P1"
        Address           = $excel.Overview.cells["I3"] # top-left corner of the table
        SourceWorkSheet   = $excel.SecurityCenter
        PivotRows         = @("Categories")
        PivotData         = @{"Categories" = "Count" }
        PivotTableStyle   = $tableStyle
        IncludePivotChart = $true
        ChartType         = "Pie"
        ChartRow          = 11 # place the chart below row 22nd
        ChartColumn       = 8
        Activate          = $true
        PivotFilter       = 'Subscription ID'
        ChartTitle        = 'Security Center Categories'
        ShowPercent       = $true
        ChartHeight       = 250
        ChartWidth        = 400
    }

    Add-PivotTable @PTParams


    Close-ExcelPackage $excel 

    Get-Job | Remove-Job

    Write-Progress -activity 'Azure Inventory' -Status "90% Complete." -PercentComplete 90 -CurrentOperation "Finishing Security Center Advisories" 
}

ImportDataExcel

$Measure = $Runtime.Totalminutes.ToString('##.##')

Write-Host ('Report Complete. Total Runtime was: ' + $Measure + ' Minutes')
Write-Host ('Total Security Advisories: ' + $Secadvco)


Write-Host ''
Write-Host ('Excel file saved at: ') -NoNewline
write-host $File -ForegroundColor Cyan
Write-Host ''

Write-Progress -activity 'Azure Inventory' -Status "100% Complete." -Completed -CurrentOperation "Security Center Inventory Completed"