function Check-RequiredModules(){

    Write-Host "Checking if all required modules are installed on machine..." -ForegroundColor Yellow
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

    $PowerAppsAdminModule = Get-Module -Name "Microsoft.PowerApps.Administration.PowerShell" -ListAvailable
    if($PowerAppsAdminModule -eq $null){

      Write-host "PowerApps Admin Module not found, Starting install the module" -ForegroundColor Yellow
      Install-Module "Microsoft.PowerApps.Administration.PowerShell" -Force
    }
    
    $PowerAppsModule = Get-Module -Name "Microsoft.PowerApps.PowerShell" -ListAvailable
    if($PowerAppsModule -eq $null){

        Write-host "PowerApps Module not found, Starting install the module" -ForegroundColor Yellow
        Install-Module "Microsoft.PowerApps.PowerShell" -Force -AllowClobber
    }
    
    $AzureADModule = Get-Module -Name "AzureAD" -ListAvailable
    if($AzureADModule -eq $null){

        Write-host "AzureAD Module not found, Starting install the module" -ForegroundColor Yellow
        Install-Module "AzureAD" -Force
    }
    
    $PnPOnlineModule = Get-Module -Name "SharePointPnPPowerShellOnline" -ListAvailable
    if($PnPOnlineModule -eq $null){

        Write-host "PnP PowerShell module not found, Starting install the module" -ForegroundColor Yellow
        Install-Module "SharePointPnPPowerShellOnline" -Force
    }
    
    $AzureADPreviewModule = Get-Module -Name "AzureADPreview" -ListAvailable
    if($AzureADPreviewModule -eq $null){

        Write-host "Azure AD Preview module not found, Starting installation of module" -ForegroundColor Yellow
        Install-Module "AzureADPreview" -Force -AllowClobber
    }
    
    $MicrosoftTeamsModule = Get-Module -Name "MicrosoftTeams" -ListAvailable
    if ($MicrosoftTeamsModule -eq $null){

      Write-host "MicrosoftTeams PowerShell module not found, Starting installation of module" -ForegroundColor Yellow      
      Install-Module "MicrosoftTeams" -Force -Confirm:$false
    }

    $SPOManagementModule = Get-Module -Name "Microsoft.Online.SharePoint.PowerShell" -ListAvailable
    if ($SPOManagementModule -eq $null){

      Write-host "SPO Management PowerShell module not found, Starting installation of module" -ForegroundColor Yellow      
      Install-Module "Microsoft.Online.SharePoint.PowerShell" -Force -Confirm:$false -AllowClobber
    }

    Import-Module Microsoft.PowerApps.Administration.PowerShell
    Import-Module Microsoft.PowerApps.PowerShell
    Import-Module SharePointPnPPowerShellOnline
    Import-Module AzureAD
    Import-Module AzureADPreview
    Import-Module MicrosoftTeams
    Import-Module Microsoft.Online.SharePoint.PowerShell
}
#----------------------------------------------------------------------------------------------------------------
function Connect-PowerPlatformModules(){
    
    Write-Host "Connecting to Azure AD PowerShell" -ForegroundColor Yellow -NoNewline
    $IfConnectedToAzureAD = Connect-AzureAD -Credential $Credential
    if($IfConnectedToAzureAD){

        Write-Host " Connected" -ForegroundColor Green 
    }

    Write-Host "Connecting to Flow and Powerapps PowerShell" -ForegroundColor Yellow
    try{Add-PowerAppsAccount -Username $UserName} catch{Write-Host "Error while executing Add-PowerAppsAccount :" $_.Exception.Message -ForegroundColor Red} 
}
#----------------------------------------------------------------------------------------------------------------
function Generate-OneDriveReport(){
    
    Write-Host "Connecting to SharePoint Online Management PowerShell to generate OneDrive report" -ForegroundColor Yellow -NoNewline
    try{Connect-SPOService -Url "https://$Tenant-admin.sharepoint.com/" -Credential $Credential}catch{Write-Host $_.Exception.Message}
    if(Get-SPOTenant){

        Write-Host " Connected" -ForegroundColor Green     
        $ODFBSites = Get-SPOSite -IncludePersonalSite $True -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" | Select Owner, Title, URL, StorageQuota, StorageUsageCurrent | Sort StorageUsageCurrent -Desc
        $TotalODFBGBUsed = [Math]::Round(($ODFBSites.StorageUsageCurrent | Measure-Object -Sum).Sum /1024,2)
        $OneDriveReportOutput = [System.Collections.Generic.List[Object]]::new()

        ForEach($Site in $ODFBSites){

                    $ReportLine = [PSCustomObject]@{
                                                        Owner       = $Site.Title
                                                        Email       = $Site.Owner
                                                        URL         = $Site.URL
                                                        QuotaGB     = [Math]::Round($Site.StorageQuota/1024,2) 
                                                        UsedGB      = [Math]::Round($Site.StorageUsageCurrent/1024,4)
                                                        PercentUsed = [Math]::Round(($Site.StorageUsageCurrent/$Site.StorageQuota * 100),4) 
                                                    }
                    $OneDriveReportOutput.Add($ReportLine) 
        }

        $OneDriveExportPath = "$ReportExportLocation\OneDriveSitesReport_$DateTime.csv"
        $OneDriveReportOutput | Export-Csv -Path $OneDriveExportPath -Delimiter ',' -NoTypeInformation
        Write-Host "OneDrive report has been generated" -ForegroundColor Green

    }else{ Write-Host " Not connected to SPO Management shell module" -ForegroundColor Red }

}
#----------------------------------------------------------------------------------------------------------------
function Generate-TeamsReport(){

    Write-Host "Connecting to Microsoft Teams PowerShell" -ForegroundColor Yellow -NoNewline
    $IfConnectedToMSTeams = Connect-MicrosoftTeams -Credential $Credential
    if($IfConnectedToMSTeams){

        Write-Host " Connected" -ForegroundColor Green
        $AllTeams = Get-Team
        Write-Host "$($AllTeams.count) teams found. " -ForegroundColor Green -NoNewline
        $TeamsReportOutput = foreach($Team in $AllTeams){
    
                New-Object -TypeName PSObject -Property @{

                    GroupId = $Team.GroupId
                    DisplayName = $Team.DisplayName
                    Description = $Team.Description
                    Visibility = $Team.Visibility
                    MailNickName = $Team.MailNickName
    
                } | Select-Object GroupId, DisplayName, Description, Visibility, MailNickName
        }

        if($TeamsReportOutput -ne $null){
            
            $TeamsExportPath = "$ReportExportLocation\TeamsReport_$DateTime.csv"
            $TeamsReportOutput | Export-Csv -Path $TeamsExportPath -Delimiter ',' -NoTypeInformation            
            Write-Host "Teams report has been generated" -ForegroundColor Green

        }else{Write-Host "No Team found in tenant" -ForegroundColor Cyan}
    
    }else{ Write-Host " Not connected to Microsoft Teams module" -ForegroundColor Red }
}
#----------------------------------------------------------------------------------------------------------------
function Generate-FlowReport(){
    
    Write-Host "Working on Flow report..." -ForegroundColor Yellow
    $FlowEnvironment = Get-FlowEnvironment
    $EnvironmentProgressbarPointer = 0
    $TotalEnvironmentCount = (Measure-Object -InputObject $FlowEnvironment).Count
    $FlowReportOutput = foreach($flowEnv in $FlowEnvironment){
    
    $EnvironmentProgressbarPointer++
    $PercentComplete = (($EnvironmentProgressbarPointer*100)/$TotalEnvironmentCount)
    Write-Progress -Activity "Processing Flow environment" -Status "Completed:$PercentComplete% Processing $($flowEnv.EnvironmentName)" -PercentComplete $PercentComplete -ErrorAction SilentlyContinue
    $flowList = Get-AdminFlow -EnvironmentName $flowEnv.EnvironmentName
    $TotalFlowCountInCurrentEnv = ($flowList | Measure-Object).Count
    $FlowProgressbarPointer = 0
    foreach($flow in $flowList){
        
        $FlowProgressbarPointer++
        Write-Progress -Id 1 -Activity "Processing Flow" -Status $flow.DisplayName -PercentComplete (($FlowProgressbarPointer*100)/$TotalFlowCountInCurrentEnv) -ErrorAction SilentlyContinue
        $flowname = $flow.FlowName       
        $flow1 = Get-AdminFlow -FlowName $flowname -EnvironmentName $flowEnv.EnvironmentName   
        write-host "$($flow.DisplayName)"
        $CreatedBy = try{(Get-AzureADUser -ObjectId $flow.CreatedBy.objectid).Mail} catch{"Not found in AD" }
        $referenceObj = $flow1.Internal.properties.referencedResources.resource
        $StringHavingReferenceListNames = @()
        $StringHavingReferenceListIDs = @()
        If(($referenceObj.site | Select -Unique).Count -eq 1){ Connect-PnPOnline -Url $referenceObj[0].site -UseWebLogin }#connect site only once here if url is same for a flow.
        foreach($refObj in $referenceObj){

            if($refObj.site -ne "[DYNAMIC_VALUE]" -and $refObj.list -ne $null){
                if(($referenceObj.site | Select -Unique).Count -ne 1){ Connect-PnPOnline -Url $refObj.site -UseWebLogin } #connect site again & again if url are different for a flow.

            foreach($o in $refObj){ 

                    $listId = Get-PnPList -Identity $o.list -ErrorAction SilentlyContinue
                    $StringHavingReferenceListNames += $listId.title
                    $StringHavingReferenceListIDs += $listId.Id
            }
        }}
        New-Object -TypeName PSObject -Property @{

                EnvironmentName = $flowEnv.EnvironmentName
                FlowName = $flow.FlowName 
                DisplayName = $flow.DisplayName
                Enabled = $flow.Enabled  
                CreatedDate = $flow.CreatedTime.substring(0,10)
                LastModifiedDate = $flow.LastModifiedTime.substring(0,10) 
                CreatedBy = $CreatedBy
                TriggerAction =  $flow.Internal.properties.definitionSummary.triggers.kind
                Connections = ($flow1.Internal.properties.connectionReferences.PSObject.Properties.Value.displayname | Get-Unique) -join ', '
                References = ($flow1.Internal.properties.referencedResources.resource.site  | Get-Unique) -join ', '
                ListIDs = $StringHavingReferenceListIDs -join ', '
                ListNames = $StringHavingReferenceListNames -join ', '

        }| Select-Object EnvironmentName, FlowName, DisplayName, Enabled, CreatedDate, LastModifiedDate, CreatedBy, TriggerAction, Connections, References, ListNames, ListIDs
    }}
    if($FlowReportOutput -ne $null){

            $FlowExportPath = "$ReportExportLocation\FlowReport_$DateTime.csv"
            $FlowReportOutput | Export-Csv -Path $FlowExportPath -Delimiter ',' -NoTypeInformation            
            Write-Host "Flow report has been generated" -ForegroundColor Green

    }else{Write-Host "No Flow found in tenant" -ForegroundColor Cyan}
}
#----------------------------------------------------------------------------------------------------------------
function Generate-PowerappsReport(){
    
    Write-Host "Working on Powerapps report..." -ForegroundColor Yellow
    $PowerappsEnvironment = Get-AdminPowerAppEnvironment
    $PowerappsReportOutput = foreach($PowerappsEnv in $PowerappsEnvironment){

                    $PowerappsList = Get-AdminPowerApp -EnvironmentName $PowerappsEnv.EnvironmentName
                    if($PowerappsList -ne $null){

                    foreach($powerapp in $PowerappsList){
 
                           Write-Host $powerapp.DisplayName -ForegroundColor Yellow
                           $AllDataSourcesOfPowerapp = $powerapp.Internal.properties.connectionReferences.psobject.Properties.value
                           $Powerapp_FlowDataSources =  $AllDataSourcesOfPowerapp | Where-Object {$_.DisplayName -eq 'Logic Flows'}
                           $Powerapp_OtherDataSources = $AllDataSourcesOfPowerapp | Where-Object {$_.DisplayName -ne 'Logic Flows'}
   
                           $CollectionOfListDataSources = @()
                           foreach($Powerapp_OtherDataSource in $Powerapp_OtherDataSources){

                                    if($Powerapp_OtherDataSource.displayName -eq "SharePoint"){ $CollectionOfListDataSources += $Powerapp_OtherDataSource.dataSources }
                           }

                           New-Object -TypeName PSObject -Property @{

                                        EnvironmentName = $PowerappsEnv.EnvironmentName
                                        AppName = $powerapp.AppName
                                        DisplayName = $powerapp.DisplayName
                                        CreatedDate = $powerapp.CreatedTime.substring(0,10)
                                        LastModifiedDate = $powerapp.LastModifiedTime.substring(0,10) 
                                        Owner = $powerapp.Owner.email
                                        DataSourceTypes = (($Powerapp_OtherDataSources.displayName -join ', '),($Powerapp_FlowDataSources.displayName -join ', ')) -join ', '
                                        FlowDataSourceNames = $Powerapp_FlowDataSources.dataSources -join ', '
                                        OtherDataSourceNames = $Powerapp_OtherDataSources.dataSources -join ', '
                                        ListNames = $CollectionOfListDataSources -join ', '
                                        ListURL = $powerapp.Internal.properties.embeddedApp.listUrl

                            }| Select-Object EnvironmentName, AppName, DisplayName, CreatedDate, LastModifiedDate, Owner, DataSourceTypes, FlowDataSourceNames, OtherDataSourceNames, ListNames, ListURL       
         }}}

        if($PowerappsReportOutput -ne $null){

            $PowerAppsExportPath = "$ReportExportLocation\PowerAppsReport_$DateTime.csv"
            $PowerappsReportOutput | Export-Csv -Path $PowerAppsExportPath -Delimiter ',' -NoTypeInformation            
            Write-Host "PowerApps report has been generated" -ForegroundColor Green

        }else{Write-Host "No Powerapp found in tenant" -ForegroundColor Red}
}
#----------------------------------------------------------------------------------------------------------------
function Generate-Office365GroupsReport(){
    
    Write-Host "Working on Office 365 groups report..." -ForegroundColor Yellow
    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    Connect-EXOPSSession -Credential $Credential
    $AllOffice365Groups = Get-UnifiedGroup
    if($AllOffice365Groups -ne $null){

        $AllOffice365GroupsExportPath = "$ReportExportLocation\AllOffice365GroupsReport_$DateTime.csv"
        $AllOffice365Groups | Select-Object ExternalDirectoryObjectId, DisplayName, PrimarySmtpAddress, ResourceProvisioningOptions  | 
        Export-Csv -Path $AllOffice365GroupsExportPath -Delimiter ',' -NoTypeInformation
        Write-Host "Office 365 group report has been generated" -ForegroundColor Green
    
    }else{Write-Host "No Office 365 group found in tenant" -ForegroundColor Red}
}
#----------------------------------------------------------------------------------------------------------------
$Tenant = Read-Host -Prompt "Please provide tenant name"
$UserName = Read-Host -Prompt "Please provide login account UPN"
$PasswordText = Read-Host -Prompt "Please provide password"
$SecuredPassword = ConvertTo-SecureString $PasswordText -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($UserName,$SecuredPassword)
$DateTime = Get-Date -Format "MMddyyyy_HHmmss"
$ReportExportLocation = "C:\temp"

Check-RequiredModules
Generate-OneDriveReport
Generate-TeamsReport
Connect-PowerPlatformModules
Generate-FlowReport
Generate-PowerappsReport
Generate-Office365GroupsReport