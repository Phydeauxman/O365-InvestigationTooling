function Check-AzureSetup 
{
    $setupBlocked = $false
    #Check Azure
    if(-not(Get-Module -ListAvailable | Where-Object {$_.Name -eq "Azure"}))
    {
        Write-Output "You don't appear to have Azure Powershell Modules installed on this computer. Here are the instructions and download to install. Install, then re-run the investigations tooling." 
        start "https://azure.microsoft.com/en-us/documentation/articles/powershell-install-configure/"
        #start "http://go.microsoft.com/fwlink/p/?linkid=320376&clcid=0x409"
        start "http://go.microsoft.com/?linkid=9811175&clcid=0x409"
        $setupBlocked = $true
    }
    else
    {
        try 
        {
        	$azurePSRoot="C:\Program Files (x86)\Microsoft SDKs\Azure\Powershell"
            $azureServiceManagementModule= $azurePSRoot + "\ServiceManagement\Azure\Azure.psd1"
            Write-Output "Checking Azure Service Management module: $azureServiceManagementModule"
            Write-Output "Looking good!"
            Import-Module $azureServiceManagementModule

        }
        catch
        {
            $setupBlocked = $true
        }
    }

    if ($setupBlocked -eq $true)
    {
        Write-Output "Something with your configuration is amiss, and your Secure Score collection will likely fail. Please review the logs above and correct any issues with your local client configuration."
        break;
    }
    else
    {

        Write-Output "Connecting to Azure"
        $AdminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $globalConfig.AzureAccountName, ($globalConfig.AzureAccountSecureString | ConvertTo-SecureString)
        #Get-Credential -Message "Provide Admin Creds to Connect to Azure Services."
        Add-AzureAccount -Credential $AdminCredential

        Write-Output "Setting the correct environment settings, creating a new storage account, and a new storage container for our blobs."
        Select-AzureSubscription -SubscriptionName $globalConfig.AzureBlobSubscription

        $AzureStorageAccountName = Get-AzureStorageAccount
        if ($AzureStorageAccountName -ne $globalConfig.AzureBlobStorageAccountName) { New-AzureStorageAccount -StorageAccountName $globalconfig.AzureBlobStorageAccountName -Location "West US"; }

        #$AzureSubscription = Get-AzureSubscription
        Set-AzureSubscription -CurrentStorageAccountName $globalConfig.AzureBlobStorageAccountName -SubscriptionName $globalConfig.AzureBlobSubscription
        
        $AzureStorageContainers = Get-AzureStorageContainer
        if ($AzureStorageContainers -ne $globalConfig.AzureBlobContainerName) { New-AzureStorageContainer -Name $globalConfig.AzureBlobContainerName -Permission Off; }
        Write-Output "Everything looks good to go with your azure setup. You've got a storage account and a blob container ready to go."
    }
}

Function Get-GlobalConfig($configFile)
{
    Write-Output "Loading Global Config File"
    
    $config = Get-Content $globalConfigFile -Raw | ConvertFrom-Json
    
    return $config;
}

$globalConfigFile=".\ConfigForO365Investigations.json";
$globalConfig = Get-GlobalConfig $globalConfigFile

#Pre-reqs for REST API calls
Add-Type -AssemblyName System.Web
$clientID = $globalConfig.InvestigationAppId
$ClientSecret = $globalConfig.InvestigationAppSecret
$loginURL = $globalConfig.LoginURL
$tenantdomain = $globalConfig.InvestigationTenantDomain
$resource = $globalConfig.ResourceAPI

# Get an Oauth 2 access token based on client id, secret and tenant domain
$body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
$Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body

#Let's put the oauth token in the header, where it belongs
$headerParams  = @{'Authorization'="$($Authorization.token_type) $($Authorization.access_token)"}

#Instantiate our enumerators
$apiFilters = @()
$rawData = @()
$dayofData = @()
$dateRange = @()
$days = @()

if ($globalConfig.DateToPull){
    $days = $globalConfig.DateToPull
}
else
{
    for ($i = [int]$globalConfig.NumberOfDaysToPull; $i -gt 0; $i--)
    {
        $days += (Get-Date).AddDays(-$i).ToString('yyyy-MM-dd')        
    }
    $days += (Get-Date).ToString('yyyy-MM-dd')

}

$hours = @("00:00Z", "01:00Z", "02:00Z", "03:00Z", "04:00Z", "05:00Z", "06:00Z", "07:00Z", "08:00Z", "09:00Z", "10:00Z", "11:00Z", "12:00Z", "13:00Z", "14:00Z", "15:00Z", "16:00Z", "17:00Z", "18:00Z", "19:00Z", "20:00Z", "21:00Z", "22:00Z", "23:00Z", "23:59:59Z")
foreach ($day in $days) { foreach ($hour in $hours) { $dateRange += $day + "T" + $hour; }}

$workLoads = @("Audit.AzureActiveDirectory", "Audit.Exchange", "Audit.SharePoint", "Audit.General", "DLP.All")
$subs = @()
$subArray = @()
$newSubs = @()
$currentSubs = @()
$global:blobs = @()
$wlCount = @()
$dayCount = @()
$thisBlobdata = @()
$altFormat = @()
$SPOmegaBlob = @()
$EXOmegaBlob = @()
$AADmegaBlob = @()
$Query = @()
$rawRef = @()

#Let's make sure we have the Activity API subscriptions turned on
$subs = Invoke-WebRequest -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantDomain/activity/feed/subscriptions/list" | Select Content
$enabledSubs = $subs.Content | ConvertFrom-Json

foreach($es in $enabledSubs){
    $subArray += $es.contentType
    }

foreach($wl in $workLoads){
    if($subArray -notcontains $wl){
      Write-Host "Looks like we need to turn on your subscriptions now."
      Write-Host "#####################################################"
      Invoke-RestMethod -Method Post -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/start?contentType=$wl"
      #Insert some error handling here
      $newSubs += $wl
      }
    else{
      $currentSubs += $wl
      }
    }

if($newSubs.Count -gt 0){
  Write-Host "#####################################################"
  Write-Host "Enabled the following subscriptions: $newSubs"
  }

Write-Host "`n#####################################################"
Write-Host "The following subscriptions were already enabled: $currentSubs"

#Let's go get some datums! First, let's construct some query parameters
foreach ($wl in $workLoads)
{
    for ($i = 0; $i -lt $dateRange.Length -1; $i++)
    {
        $apiFilters += "?contentType=$wl&startTime=" + $dateRange[$i] + "&endTime=" + $dateRange[$i+1]
    }
}

#Then execute the content enumeration method per workload, per day
foreach ($pull in $apiFilters)
{
    
    do{

        try{
            
            $rawRef = Invoke-WebRequest -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/content$pull" -ErrorVariable errvar
            $response = $true
            #Write-Host "Pulled file #: " -ForegroundColor Green -NoNewline; Write-Host ($i + 1) -ForegroundColor Yellow -NoNewline; Write-Host " from subscription" -ForegroundColor Green
                
           }
        catch{

            #Write-Host $_ -ForegroundColor Red
            $response = $false
            Write-Host "Server to busy, to execute call: " -ForegroundColor Yellow -NoNewline; Write-Host $pull -ForegroundColor Red
            Start-Sleep -Seconds 3
                    
             }
                
      }Until($response -eq $true)
    
    #$rawRef = Invoke-WebRequest -Headers $headerParams -Uri "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/content$pull"

    if ($rawRef.Headers.NextPageUri) 
    {
        $pageTracker = $true
        $thatRabbit = $rawRef
        while ($pageTracker -ne $false)
        {
        	$thisRabbit = Invoke-WebRequest -Headers $headerParams -Uri $thatRabbit.Headers.NextPageUri
            Write-Output "We just called a rabbit: " $thatRabbit.Headers.NextPageUri
			$rawData += $thisRabbit

			If ($thisRabbit.Headers.NextPageUri)
			{
				$pageTrack = $True
			}
			Else
            {
			    $pageTracker = $False
			}
            $thatRabbit = $thisRabbit
        }
    }

    $rawData += $rawRef
    Write-Output "We just called $pull"
    Write-Output "---"

    $timeleft = $AuthorizationExpiration - [datetime]::Now
    if($timeLeft.TotalSeconds -lt 100) 
      {
        Write-Host "Nearing token expiration, acquiring a new one.";
        $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret};
        $Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body; 
        $headerParams  = @{'Authorization'="$($Authorization.token_type) $($Authorization.access_token)"}; 
        $AuthorizationExpiration = [datetime]::Now.AddSeconds($Authorization.expires_in); 
        Write-Host "New token lifespan is $AuthorizationExpiration"; 
      }
}

#Then convert each day's package into discrete blob calls
foreach ($dayofData in $rawData)
{
    $blobs += $dayofData.Content | ConvertFrom-Json
}

Write-Host "#####################################################"
Write-Host "Count of blobs in the Activity API: " -NoNewline; Write-Host $blobs.Count -ForegroundColor Green

foreach ($day in $days)
{
    $dayCount = @($blobs | Where-Object {($_.contentCreated -match $day)})
    Write-Host "`nCount of blobs on " -NoNewline; Write-Host $day -NoNewLine; Write-Host " : " -NoNewLine; Write-Host $dayCount.Count -ForegroundColor Green
}

foreach ($wl in $workLoads)
{
    $wlCount = @($blobs | Where-Object {($_.contentType -eq $wl)})
    Write-Host "`nCount of blobs for " -NoNewline; Write-Host $wl -NoNewLine; Write-Host " : " -NoNewLine; Write-Host $wlCount.Count -ForegroundColor Green

}

#This will write the files from the API to the local data store.
function Export-LocalFiles ($blobs) {

    #Let's make some output directories
    if (! (Test-Path ".\JSON"))
        {
            New-Item -Path .\JSON -ItemType Directory
        }

    if (! (Test-Path ".\CSV"))
        {
            New-Item -Path .\CSV -ItemType Directory
        }

    #Let's build a variable full of the files already in the local store
    $localfiles = @()
    $localFiles = Get-ChildItem $globalConfig.LocalFileStore -Recurse | Select-Object -Property Name

    #Go Get the Content!
    for ($i = 0; $i -le $blobs.Length -1; $i++) 
    { 
        $timeleft = $AuthorizationExpiration - [datetime]::Now
        if ($timeLeft.TotalSeconds -lt 100) 
        {
            Write-Host "Nearing token expiration, acquiring a new one.";
            $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret};
            $Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body; 
            $headerParams  = @{'Authorization'="$($Authorization.token_type) $($Authorization.access_token)"}; 
            $AuthorizationExpiration = [datetime]::Now.AddSeconds($Authorization.expires_in); 
            Write-Host "New token lifespan is $AuthorizationExpiration"; 
        }
            
        if ($localFiles -like "*" + $blobs[$i].contentId + "*") 
        { 
            Write-Output "Looks like we already have this blob locally."; 
        }
        else
        {

            do{

                try{
                    
                    $thisBlobdata = Invoke-WebRequest -Headers $headerParams -Uri $blobs[$i].contentUri -ErrorVariable errvar
                    $response = $true
                    Write-Host "Pulled file #: " -ForegroundColor Green -NoNewline; Write-Host ($i + 1) -ForegroundColor Yellow -NoNewline; Write-Host " from subscription" -ForegroundColor Green
                
                }
                catch{

                    #Write-Host $_ -ForegroundColor Red
                    $response = $false
                    Write-Host "Server to busy to pull file #: " -ForegroundColor Yellow -NoNewline; Write-Host ($i + 1) -ForegroundColor Red
                    Start-Sleep -Seconds 3
                    
                    }
                
            }Until($response -eq $true)
            
            #Write it to JSON
            $thisBlobdata.Content | Out-File (".\JSON\" + $blobs[$i].contentType + $blobs[$i].contentCreated.Substring(0,10) + "--" + $blobs[$i].contentID + ".json")
        
            #Write it to CSV
            $altFormat = $thisBlobdata.Content | ConvertFrom-Json
            $altFormat | Export-Csv -Path (".\CSV\" + $blobs[$i].contentType + $blobs[$i].contentCreated.Substring(0,10) + "--" + $blobs[$i].contentID + ".csv") -NoTypeInformation

            Write-Host "Writing file #: " -NoNewLine; Write-Host ($i + 1) -ForegroundColor Green -NoNewline; Write-Host " out of " -NoNewline; Write-Host $blobs.Length -ForegroundColor Yellow -NoNewline; Write-Host ". You have " -NoNewline; Write-Host ($timeleft.TotalSeconds) -NoNewline; Write-Host " seconds left on your oauth token lifespan.";  

        }

    }

}

function Invoke-AzureSql {
     Param(
      [Parameter(
      Mandatory = $true,
      ParameterSetName = '',
      ValueFromPipeline = $true)]
      [string]$Query
      )

    $AzureSqlAdminUserName = $globalConfig.AzureSqlUsername
    $AzureSqlAdminPassword = $globalConfig.AzureSqlPass
    $AzureSqlDatabase = $globalConfig.AzureSqlDb
    $AzureSqlHost = $globalConfig.AzureSqlHostname
 
    $ConnectionString = "server=" + $AzureSqlHost + ",1433; uid=" + $AzureSqlAdminUserName + "; pwd=" + $AzureSqlAdminPassword + "; database="+$AzureSqlDatabase
 
    Try {
      
      $Connection = New-Object System.Data.SqlClient.SqlConnection
      $Connection.ConnectionString = $ConnectionString
      $Connection.Open()
  
      $Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)
      $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($Command)

      $DataSet = New-Object System.Data.DataSet
      $RecordCount = $dataAdapter.Fill($dataSet, "data")
      $DataSet.Tables[0]
      }

    Catch {
      throw "ERROR : Unable to run query : $query `n$Error[0]"
     }

    Finally {
      $Connection.Close()
      }
 }

function Export-AzureSQL ($blobs) {

    #Make sure we've got tables
    Invoke-AzureSql -Query "if not exists (select * from sysobjects where name ='auditsharepoint' and xtype='U') CREATE TABLE auditsharepoint (CreationTime DATETIME, Id CHAR(36) NOT NULL, Operation TEXT, OrganizationId TEXT, RecordType INT, UserKey TEXT, UserType INT, Version INT, Workload TEXT, ClientIP TEXT, ObjectId TEXT, UserId TEXT, EventSource TEXT, ItemType TEXT, ListId TEXT, ListItemUniqueId TEXT, Site TEXT, UserAgent TEXT, WebId TEXT, SourceFileExtension TEXT, SiteUrl TEXT, SourceFileName TEXT, SourceRelativeUrl TEXT, PRIMARY KEY (Id))"
    Invoke-AzureSql -Query "if not exists (select * from sysobjects where name ='auditgeneral' and xtype='U') CREATE TABLE auditgeneral (CreationTime DATETIME, Id CHAR(36) NOT NULL, Operation TEXT, OrganizationId TEXT, RecordType INT, Version INT, Workload TEXT, ClientIP TEXT, ObjectId TEXT, UserId TEXT, EventSource TEXT, ItemType TEXT, ListId TEXT, ListItemUniqueId TEXT, Site TEXT, UserAgent TEXT, WebId TEXT, SourceFileExtension TEXT, SiteUrl TEXT, SourceFileName TEXT, SourceRelativeUrl TEXT, PRIMARY KEY (Id))"
    Invoke-AzureSql -Query "if not exists (select * from sysobjects where name='auditexchange' and xtype='U') CREATE TABLE auditexchange (CreationTime DATETIME, Id CHAR(36) NOT NULL, Operation TEXT, OrganizationId TEXT, RecordType INT, ResultStatus TEXT, UserKey TEXT, UserType INT, Version INT, Workload TEXT, ObjectId TEXT, UserId TEXT, ExternalAccess TEXT, OrganizationName TEXT, OriginatingServer TEXT, Parameters TEXT, PRIMARY KEY (Id))"
    Invoke-AzureSql -Query "if not exists (select * from sysobjects where name='auditaad' and xtype='U') CREATE TABLE auditaad (CreationTime DATETIME, Id CHAR(36) NOT NULL, Operation TEXT, OrganizationId TEXT, RecordType INT, ResultStatus TEXT, UserKey TEXT, UserType INT, Version INT, Workload TEXT, ClientIP TEXT, ObjectId TEXT, UserId TEXT, AzureActiveDirectoryEventType TEXT, Actor TEXT, ActorContextId TEXT, InterSystemId TEXT, IntraSystemId TEXT, Target TEXT, TargetContextId TEXT, PRIMARY KEY (Id))"
    Invoke-AzureSql -Query "if not exists (select * from sysobjects where name='dlpall' and xtype='U') CREATE TABLE dlpall (CreationTime DATETIME, Id CHAR(36) NOT NULL, Operation TEXT, OrganizationId TEXT, RecordType INT, UserKey TEXT, UserType INT, Version INT, Workload TEXT, ObjectId TEXT, UserId TEXT, IncidentId TEXT, SensitiveInfoDetectionIsIncluded TEXT, PRIMARY KEY (Id))"

    #Go Get the Content!
    
    # Temp Code
    $AADmegaBlob = @()
    $SPOmegaBlob = @()
    $EXOmegaBlob = @()
    $GENmegaBlob = @()
    $DLPmegaBlob = @()

    for ($i = 0; $i -le $blobs.Length -1; $i++) 
    { 
        $timeleft = $AuthorizationExpiration - [datetime]::Now
        if ($timeLeft.TotalSeconds -lt 100) 
        {
           Write-Host "Nearing token expiration, acquiring a new one.";
           $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
           $Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
           $headerParams  = @{'Authorization'="$($Authorization.token_type) $($Authorization.access_token)"}; 
           $AuthorizationExpiration = [datetime]::Now.AddSeconds($Authorization.expires_in);
           Write-Host "New token lifespan is $AuthorizationExpiration";
        }
        
        do{

                try{
                    
                    $thisBlobdata = Invoke-WebRequest -Headers $headerParams -Uri $blobs[$i].contentUri -ErrorVariable errvar
                    $response = $true
                    Write-Host "Pulled file #: " -ForegroundColor Green -NoNewline; Write-Host ($i + 1) -ForegroundColor Yellow -NoNewline; Write-Host " from subscription" -ForegroundColor Green
                
                }
                catch{

                    #Write-Host $_ -ForegroundColor Red
                    $response = $false
                    Write-Host "Server to busy to pull file #: " -ForegroundColor Yellow -NoNewline; Write-Host ($i + 1) -ForegroundColor Red
                    Start-Sleep -Seconds 3
                    
                    }
                
            }Until($response -eq $true)
        
        #$thisBlobdata = Invoke-WebRequest -Headers $headerParams -Uri $blobs[$i].contentUri
        
        #Get it into a more work-able format
        $altFormat = $thisBlobdata.Content | ConvertFrom-Json

        if ($blobs[$i].ContentType -eq "Audit.SharePoint") { $SPOmegaBlob += $altFormat; }
        if ($blobs[$i].ContentType -eq "Audit.Exchange") { $EXOmegaBlob += $altFormat; }
        if ($blobs[$i].ContentType -eq "Audit.AzureActiveDirectory") { $AADmegaBlob += $altFormat; }
        if ($blobs[$i].ContentType -eq "Audit.General") { $GENmegaBlob += $altFormat; }
        if ($blobs[$i].ContentType -eq "DLP.All") { $DLPmegaBlob += $altFormat; }
        Write-Host "Completed blob " ($i +1)
    }

    #Insert the records from the Audit.SharePoint megablob into the SQL database
    foreach ($record in $SPOmegaBlob)
    {
        #Construct the query to a valid string, then execute that sucker
        #Need to handle the parameters object by converting to a string.
        $thisQuery = "if not exists (select Id from auditsharepoint where Id='" + $record.Id + "') BEGIN INSERT INTO auditsharepoint (CreationTime, Id, Operation, OrganizationId, RecordType, UserKey, UserType, Version, Workload, ClientIP, ObjectId, UserId, EventSource, ItemType, ListId, ListItemUniqueID, Site, UserAgent, WebId, SourceFileExtension, SiteUrl, SourceFileName, SourceRelativeUrl) VALUES ('" + $record.CreationTime + "', '" + $record.Id + "', '" + $record.Operation + "', '" + $record.OrganizationId + "', '" + $record.RecordType + "', '" + $record.UserKey + "', '" + $record.UserType + "', '" + $record.Version + "', '" + $record.Workload + "', '" + $record.ClientIP + "', '" + $record.ObjectId + "', '" + $record.UserId + "', '" + $record.EventSource + "', '" + $record.ItemType + "', '" + $record.ListId + "', '" + $record.ListItemUniqueId + "', '" + $record.Site + "', '" + $record.UserAgent + "', '" + $record.WebId + "', '" + $record.SourceFileExtension + "', '" + $record.SiteUrl + "', '" + $record.SourceFileName + "', '" + $record.SourceRelativeUrl + "') END"
        Invoke-AzureSql "$thisQuery"
    }

    Write-Host "#####################################################"
    Write-Host "Successfully updated sharepoint records in SQL."


    #Insert the records from the Audit.Exchange megablob into the SQL database
    foreach ($record in $EXOmegaBlob)
    {
        #Construct the query to a valid string, then execute that sucker
        #Need to handle the parameters object by converting to a string.
        $thisQuery = "if not exists (select Id from auditexchange where Id='" + $record.Id + "') BEGIN INSERT INTO auditexchange (CreationTime, Id, Operation, OrganizationId, RecordType, ResultStatus, UserKey, UserType, Version, Workload, ObjectId, UserId, ExternalAccess, OrganizationName, OriginatingServer, Parameters) VALUES ('" + $record.CreationTime + "', '" + $record.Id + "', '" + $record.Operation + "', '" + $record.OrganizationId + "', '" + $record.RecordType + "', '" + $record.ResultStatus + "', '" + $record.UserKey + "', '" + $record.UserType + "', '" + $record.Version + "', '" + $record.Workload + "', '" + $record.ObjectId + "', '" + $record.UserId + "', '" + $record.ExternalAccess + "', '" + $record.OrganizationName + "', '" + $record.OriginatingServer + "', '" + $record.Parameters + "') END"
        Invoke-AzureSql "$thisQuery"
    }
    Write-Host "#####################################################"
    Write-Host "Successfully updated exchange records in SQL."


    #Insert the records from the Audit.AzureActiveDirectory megablob into the SQL database
    foreach ($record in $AADmegaBlob)
    {
        #Construct the query to a valid string, then execute that sucker
        #Need to handle the parameters object by converting to a string.
        $thisQuery = "if not exists (select Id from auditaad where Id='" + $record.Id + "') BEGIN INSERT INTO auditaad (CreationTime, Id, Operation, OrganizationId, RecordType, ResultStatus, UserKey, UserType, Version, Workload, ClientIP, ObjectId, UserId, AzureActiveDirectoryEventType, Actor, ActorContextId, InterSystemId, IntraSystemId, Target, TargetContextId) VALUES ('" + $record.CreationTime + "', '" + $record.Id + "', '" + $record.Operation + "', '" + $record.OrganizationId + "', '" + $record.RecordType + "', '" + $record.ResultStatus + "', '" + $record.UserKey + "', '" + $record.UserType + "', '" + $record.Version + "', '" + $record.Workload + "', '" + $record.ClientIP + "', '" + $record.ObjectId + "', '" + $record.UserId + "', '" + $record.AzureActiveDirectoryEventType + "', '" + $record.Actor + "', '" + $record.ActorContextId + "', '" + $record.InterSystemsId + "', '" + $record.IntraSystemsId + "', '" + $record.Target + "', '" + $record.TargetContextId + "') END"
        Invoke-AzureSql "$thisQuery"
    }
    Write-Host "#####################################################"
    Write-Host "Successfully updated aad records in SQL."

    #Insert the records from the Audit.General megablob into the SQL database
    foreach ($record in $GENmegaBlob)
    {
        #Construct the query to a valid string, then execute that sucker
        #Need to handle the parameters object by converting to a string.
        $thisQuery = "if not exists (select Id from auditgeneral where Id='" + $record.Id + "') BEGIN INSERT INTO auditgeneral (CreationTime, Id, Operation, OrganizationId, RecordType, Version, Workload, ClientIP, ObjectId, UserId, EventSource, ItemType, ListId, ListItemUniqueId, Site, UserAgent, WebId, SourceFileExtension, SiteUrl, SourceFileName, SourceRelativeUrl) VALUES ('" + $record.CreationTime + "', '" + $record.Id + "', '" + $record.Operation + "', '" + $record.OrganizationId + "', '" + $record.RecordType + "', '" + $record.Version + "', '" + $record.Workload + "', '" + $record.ClientIP + "', '" + $record.ObjectId + "', '" + $record.UserId + "', '" + $record.EventSource + "', '" + $record.ItemType + "', '" + $record.ListId + "', '" + $record.ListItemUniqueId + "', '" + $record.Site + "', '" + $record.UserAgent + "', '" + $record.WebId + "', '" + $record.SourceFileExtension + "', '" + $record.SiteUrl + "', '" + $record.SourceFileName + "', '" + $record.SourceRelativeUrl + "') END"
        Invoke-AzureSql "$thisQuery"
    }

    Write-Host "#####################################################"
    Write-Host "Successfully updated general records in SQL."

    #Insert the records from the DLP megablob into the SQL database
    foreach ($record in $DLPmegaBlob)
    {
        #Construct the query to a valid string, then execute that sucker
        #Need to handle the parameters object by converting to a string.
        $thisQuery = "if not exists (select Id from dlpall where Id='" + $record.Id + "') BEGIN INSERT INTO dlpall (CreationTime, Id, Operation, OrganizationId, RecordType, UserKey, UserType, Version, Workload, ObjectId, UserId, IncidentId, SensitiveInfoDetectionIsIncluded) VALUES ('" + $record.CreationTime + "', '" + $record.Id + "', '" + $record.Operation + "', '" + $record.OrganizationId + "', '" + $record.RecordType + "', '" + $record.UserKey + "', '" + $record.UserType + "', '" + $record.Version + "', '" + $record.Workload + "', '" + $record.ObjectId + "', '" + $record.UserId + "', '" + $record.IncidentId + "', '" + $record.SensitiveInfoDetectionIsIncluded + "') END"
        Invoke-AzureSql "$thisQuery"
    }

    Write-Host "#####################################################"
    Write-Host "Successfully updated DLP records in SQL."

}

function Export-AzureBlob ($blobs) {

    $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    $Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
    $AuthorizationExpiration = [datetime]::Now.AddSeconds($Authorization.expires_in)

    #Let's put the oauth token in the header, where it belongs
    $headerParams  = @{'Authorization'="$($Authorization.token_type) $($Authorization.access_token)"}

    #Go Get the Content!
    for ($i = 0; $i -le $blobs.Length -1; $i++) 
    { 
        #Get the datums
        $timeleft = $AuthorizationExpiration - [datetime]::Now
        if ($timeLeft.TotalSeconds -lt 100) 
        {
           <#Write-Host "Nearing token expiration, acquiring a new one."; 
           $Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $renewbody; 
           $headerParams  = @{'Authorization'="$($Authorization.token_type) $($Authorization.access_token)"}; 
           $AuthorizationExpiration = [datetime]::Now.AddSeconds($Authorization.expires_in); 
           Write-Host "New token lifespan is $AuthorizationExpiration";#>
           $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
           $Authorization = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
        }

        $thisBlobdata = Invoke-WebRequest -Headers $headerParams -Uri $blobs[$i].contentUri
        $thisBlobName = ($blobs[$i].contentType + $blobs[$i].contentCreated.Substring(0,10) + "--" + $blobs[$i].contentID + ".json")
        
        #Write it to JSON
        $thisBlobdata.Content | ConvertTo-Json | Out-File ("c:\temp\" + $thisBlobName)

        Get-ChildItem "c:\temp\$thisBlobName" | Set-AzureStorageBlobContent -Container $globalConfig.AzureBlobContainerName -Force
        Remove-Item "c:\temp\$thisBlobName"
    }

    
}

#This will export the data in the API to a local directory
if ($globalConfig.StoreFilesLocally -eq "True") { Export-LocalFiles $blobs; }

#This will export the data in the API to an Azure SQL instance
if ($globalConfig.StoreDataInAzureSql -eq "True") { Export-AzureSql $blobs; }

#This will export the data in the API to an Azure blob storage account
if ($globalConfig.StoreDataInAzureBlob -eq "True") { Check-AzureSetup; }
if ($globalConfig.StoreDataInAzureBlob -eq "True") { Export-AzureBlob $blobs; }