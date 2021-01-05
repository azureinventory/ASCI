##########################################################################################
#                                                                                        #
#          * Azure Security Center Inventory ( ASCI ) Report Generator *                 #
#                                                                                        #
#       Version: 0.0.1                                                                   #
#       Authors: Claudio Merola <clvieira@microsoft.com>                                 #
#                Renato Gregio <renato.gregio@microsoft.com>                             #
#                                                                                        #
#       Date: 12/19/2020                                                                 #
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

param ($TenantID, $AdvisoryStatus, $SubscriptionID) 

$Runtime = Measure-Command -Expression {

    if ($DebugPreference -eq 'Inquire') {
        $DebugPreference = 'Continue'
    }

    $ErrorActionPreference = "silentlycontinue"
    $DesktopPath = "C:\AzureInventory"
    $CSPath = "$HOME/AzureInventory"
    $Global:Resources = @()
    $Global:Advisories = ''
    $Global:Security = ''
    $Global:Subscriptions = ''


    <######################################### Environment #########################################>

    #### Creating Excel file variable:
    $Global:File = ($DefaultPath + "AzureSecurityCenterInventory_Report_" + (get-date -Format "yyyy-MM-dd_HH_mm") + ".xlsx")
    Write-Debug ('Excel file:' + $File)

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

    if ($AdvisoryStatus.IsPresent -and $SubscriptionID.IsPresent) {
        $SecSize = az graph query -q  "securityresources | where properties['status']['code'] == '$AdvisoryStatus' | summarize count()" --subscription $SubscriptionID --output json --only-show-errors | ConvertFrom-Json    
    }
    elseif ($AdvisoryStatus.IsPresent) {
        $SecSize = az graph query -q  "securityresources | where properties['status']['code'] == '$AdvisoryStatus' | summarize count()" --output json --only-show-errors | ConvertFrom-Json
    }
    elseif ($SubscriptionID.IsPresent) {
        $SecSize = az graph query  -q  "securityresources | summarize count()" --subscription $SubscriptionID --output json --only-show-errors | ConvertFrom-Json
    }
    else {
        $SecSize = az graph query -q  "securityresources | summarize count()" --output json --only-show-errors | ConvertFrom-Json
    }
    
    $SecSizeNum = $SecSize.'count_'

    if ($SecSizeNum -ge 1) {
        Start-Job -name 'SecAdvisories' -ScriptBlock {
            $Loop = $($args[0]) / 1000
            $Loop = [math]::ceiling($Loop)
            $Looper = 0
            $Limit = 0
            $Sec = @()
            while ($Looper -lt $Loop) {
                $Looper ++
                if ($AdvisoryStatus.IsPresent -and $SubscriptionID.IsPresent) {
                    $SecCenter = az graph query -q "securityresources | order by id asc | where properties['status']['code'] == 'Unhealthy'" --subscription $SubscriptionID --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json
                }
                elseif ($AdvisoryStatus.IsPresent) {
                    $SecCenter = az graph query -q "securityresources | order by id asc | where properties['status']['code'] == 'Unhealthy'" --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json	
                }
                elseif ($SubscriptionID.IsPresent) {
                    $SecCenter = az graph query -q "securityresources | order by id asc" --subscription $SubscriptionID --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json
                }
                else {
                    $SecCenter = az graph query -q "securityresources | order by id asc" --skip $Limit --first 1000 --output json --only-show-errors | ConvertFrom-Json
                }
                $Sec += $SecCenter
                Start-Sleep 3
                $Limit = $Limit + 1000
            }
            $Sec    
        } -ArgumentList $SecSizeNum
    }
}
            
Write-Progress -activity 'Azure Inventory' -Status "10% Complete." -PercentComplete 10 -CurrentOperation "Finishing Security Advisories extraction jobs.."
get-job | Wait-Job

<######################################### Importing Data to Excel  #########################################>

function ImportDataExcel {

    $Subs = $Subscriptions

    Write-Progress -activity 'Azure Inventory' -Status "15% Complete." -PercentComplete 15 -CurrentOperation "Starting to process extraction data.."         
    $Global:Security = Receive-Job -Name 'SecAdvisories'
    get-job | Remove-Job

    Start-Job -Name 'Security' -ScriptBlock {

        $obj = ''
        $tmp = @()

        foreach ($1 in $($args[0])) {
            $data = $1.PROPERTIES

            $sub1 = $($args[1]) | Where-Object { $_.id -eq $1.properties.resourceDetails.Id.Split("/")[2] }

            $obj = @{
                'Subscription ID'    = $sub1.Name;
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

        $tmp

    } -ArgumentList $Security, $Subs | Out-Null

    #### Security Center worksheet is always the second sequence:
    Write-Progress -activity 'Azure Inventory' -Status "25% Complete." -PercentComplete 25 -CurrentOperation "Processing Security Center Advisories"         
    Write-Debug ('Generating Security Center sheet.')
    if ($Security) {

        $condtxtsec = $(New-ConditionalText High -Range G:G
            New-ConditionalText High -Range L:L)

        $Global:Secadvco = $Security.Count

        Write-Progress -activity 'Azure Inventory' -Status "30% Complete." -PercentComplete 30 -CurrentOperation "Processing Security Center Advisories"    

        while (get-job -Name 'Security' | Where-Object { $_.State -eq 'Running' }) {
            Start-Sleep -Seconds 2
        }

        $Sec = Receive-Job -Name 'Security'

        $Sec | 
        ForEach-Object { [PSCustomObject]$_ } | 
        Select-Object 'Subscription ID',
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

    }
    Write-Progress -activity 'Azure Inventory' -Status "50% Complete." -PercentComplete 50 -CurrentOperation "Processing Security Center Advisories"    
    <################################################################### Subscriptions ###################################################################>
    Write-Progress -activity 'Azure Inventory' -Status "75% Complete." -PercentComplete 75 -CurrentOperation "Finishing Security Center Advisories"         

    <#  $ResTable = $sec # | Where-Object { $_.type -ne 'microsoft.advisor/recommendations' }
    $ResTable2 = $ResTable | Select-Object type, resourceGroup, subscriptionId
    $ResTable3 = $ResTable2 | Group-Object -Property type, resourceGroup, subscriptionId 

    Write-Debug ('Generating Subscription sheet for: ' + $SUBs.count + ' Subscriptions.')

    $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

    if ($null -ne $obj) {
        Remove-Variable obj
    }
    $tmp = @()

    foreach ($ResourcesSUB in $ResTable3) {
        $ResourceDetails = $ResourcesSUB.name -split ","
        $SubName = $SUBs | Where-Object { $_.id -eq ($ResourceDetails[2] -replace (" ", "")) }

        $obj = @{
            'Subscription'              = $SubName.Name;
            'Resource Group'            = $ResourceDetails[1];
            'Security Advisory Type'    = $ResourceDetails[0];
            'Total Security Advisories' = $ResourcesSUB.Count
        }
        $tmp += $obj
    }

    $tmp | 
    ForEach-Object { [PSCustomObject]$_ } | 
    Select-Object 'Subscription',
    'Resource Group',
    'Resource Type',
    'Resources' | Export-Excel -Path $File -WorksheetName 'Subscriptions' -AutoSize -TableName 'Subscriptions' -TableStyle $tableStyle -Style $Style -Numberformat '0' -MoveToEnd 

    Remove-Variable tmp
 
    #>
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