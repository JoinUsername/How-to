Import-Module -Name "PnP.PowerShell"
# Load SharePoint CSOM Assemblies
# https://www.microsoft.com/en-us/download/details.aspx?id=42038
#Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Add-Type -Path (Resolve-Path ".\dll\Microsoft.SharePoint.Client.dll")
#Add-Type -Path (Resolve-Path ".\dll\Microsoft.SharePoint.Client.Runtime.dll")
#[void]([System.Reflection.Assembly]::LoadFrom(".\dll\Microsoft.SharePoint.Client.dll"))
#[void]([System.Reflection.Assembly]::LoadFrom(".\dll\Microsoft.SharePoint.Client.Runtime.dll"))
# Define Assembly path
#[string]$assemblySPClientPath = '.\dll\Microsoft.SharePoint.Client.dll'
#[string]$assemblySPClienRuntimetPath = '.\dll\Microsoft.SharePoint.Client.Runtime.dll'

# Add assembly DLL
#Add-Type -Path $assemblySPClientPath
#Add-Type -Path $assemblySPClienRuntimetPath

#Add-Type Microsoft.SharePoint.Client;
#Add-Type Microsoft.SharePoint.Client.List;
#Add-Type Microsoft.SharePoint.Client.ListCollection;
#Add-Type Microsoft.SharePoint.Client.Web;
#Add-Type Microsoft.SharePoint.Client.WebCollection;
#Add-Type Microsoft.SharePoint.Client.ContentType;
#Add-Type Microsoft.SharePoint.Client.ContentTypeCollection;

class Process {
    #[string]$Brand
    [Hashtable]$ParentContentTypes = @{}
    $connection = $null
    $ctx = $null;
    $webFields =$null;
    $webContentTypes =$null;
    $webLists =$null;
    [PSCredential]$cred=$null;
    [string]$ClientId=$null;
    [string]$Tenant=$null;
    [string]$CertificatePath=$null;
    [Hashtable]$Connections = @{}
    [string]$Url=$null;

    Process(){
        $this.Init(); 
    }

    Process($clientId, $tenant, $certificatePath){

        $this.ClientId=$clientId;
        $this.Tenant=$tenant;
        $this.CertificatePath=$certificatePath;

        $this.Init();
    }

    hidden Init(){
        
        #$this.ParentContentTypes.add('Name','Id')  # Group; Description
        $this.ParentContentTypes.add('System','0x')  # _Hidden; 
        $this.ParentContentTypes.add('Item','0x01')  # List Content Types; Create a new list item.
        $this.ParentContentTypes.add('Circulation','0x01000F389E14C9CE4CE486270B9D4713A5D6')  # Group Work Content Types; Add a new Circulation.
        $this.ParentContentTypes.add('New Word','0x010018F21907ED4E401CB4F14422ABC65304')  # Group Work Content Types; Add a New Word to this list.
        $this.ParentContentTypes.add('WorkflowServiceDefinition','0x01002A2479FF33DD4BC3B1533A012B653717')  # _Hidden; WorkflowServiceDefinition
        $this.ParentContentTypes.add('Health Analyzer Rule Definition','0x01003A8AA7A4F53046158C5ABD98036A01D5')  # _Hidden; Create a New Health Analyzer Rule
        $this.ParentContentTypes.add('Resource','0x01004C9F4486FBF54864A7B0A33D02AD19B1')  # Group Work Content Types; Add a new resource.
        $this.ParentContentTypes.add('Official Notice','0x01007CE30DD1206047728BAFD1C39A850120')  # Group Work Content Types; Add a new Official Notice.
        $this.ParentContentTypes.add('Phone Call Memo','0x0100807FBAC5EB8A4653B8D24775195B5463')  # Group Work Content Types; Add a new Phone Call Memo.
        $this.ParentContentTypes.add('Holiday','0x01009BE2AB5291BF4C1A986910BD278E4F18')  # Group Work Content Types; Add a new holiday.
        $this.ParentContentTypes.add('What''s New Notification','0x0100A2CA87FF01B442AD93F37CD7DD0943EB')  # Group Work Content Types; Add a new What's New notification.
        $this.ParentContentTypes.add('WorkflowServiceSubscription','0x0100AA27A923036E459D9EF0D18BBD0B9587')  # _Hidden; WorkflowServiceSubscription
        $this.ParentContentTypes.add('Timecard','0x0100C30DDA8EDB2E434EA22D793D9EE42058')  # Group Work Content Types; Add a new timecard data.
        $this.ParentContentTypes.add('Resource Group','0x0100CA13F2F8D61541B180952DFB25E3E8E4')  # Group Work Content Types; Add a new resource group.
        $this.ParentContentTypes.add('Health Analyzer Report','0x0100F95DB3A97E8046B58C6A54FB31F2BD46')  # _Hidden; A report from the health analyzer
        $this.ParentContentTypes.add('Users','0x0100FBEEE6F0C500489B99CDA6BB16C398F7')  # Group Work Content Types; Add new users to this list.
        $this.ParentContentTypes.add('Document','0x0101')  # Document Content Types; Create a new document.
        $this.ParentContentTypes.add('Display Template','0x0101002039C03B61C64EC4A04F5361F3851066')  # _Hidden; Base Content Type containing common Display Template columns. Use one of Control, Group, Item or Filter Display Template content types to create a display type. 
        $this.ParentContentTypes.add('Display Template Code','0x0101002039C03B61C64EC4A04F5361F385106605')  # _Hidden; Display Template Code javascript that registers and executes Display Template rendering logic.
        $this.ParentContentTypes.add('JavaScript Display Template','0x0101002039C03B61C64EC4A04F5361F3851068')  # Display Template Content Types; Create a new custom CSR JavaScript Display Template.
        $this.ParentContentTypes.add('Office Data Connection File','0x010100629D00608F814DD6AC8A86903AEE72AA')  # _Hidden; 
        $this.ParentContentTypes.add('List View Style','0x010100734778F2B7DF462491FC91844AE431CF')  # Document Content Types; Create a new List View Style
        $this.ParentContentTypes.add('Site Page','0x0101009D1CB255DA76424F860D91F20E6C4118')  # Document Content Types; Create a new site page.
        $this.ParentContentTypes.add('Repost Page','0x0101009D1CB255DA76424F860D91F20E6C4118002A50BFCFB7614729B56886FADA02339B')  # Document Content Types; Used to create a News link post. If deleted, the News link option will be disabled for users.
        $this.ParentContentTypes.add('Universal Data Connection File','0x010100B4CBD48E029A4AD8B62CB0E41868F2B0')  # _Hidden; Provide a standard place for applications, such as Microsoft InfoPath, to store data connection information.
        $this.ParentContentTypes.add('Design File','0x010100C5033D6CFB8447359FB795C8A73A2B19')  # _Hidden; HTML, JavaScript, CSS, images, and other supporting files in the Master Page Gallery used by HTML Master Pages, HTML Page Layouts, and Display Templates.
        $this.ParentContentTypes.add('InfoPath Form Template','0x010100F8EF98760CBA4A94994F13BA881038FA')  # _Hidden; A Microsoft InfoPath Form Template.
        $this.ParentContentTypes.add('Form','0x010101')  # Document Content Types; Fill out this form.
        $this.ParentContentTypes.add('Picture','0x010102')  # Document Content Types; Upload an image or a photograph.
        $this.ParentContentTypes.add('Unknown Document Type','0x010104')  # Special Content Types; Allows users to upload documents of any content type to a library. Unknown documents will be treated as their original content type in client applications.
        $this.ParentContentTypes.add('Master Page','0x010105')  # Document Content Types; Create a new master page.
        $this.ParentContentTypes.add('Master Page Preview','0x010106')  # Document Content Types; Create a new master page preview.
        $this.ParentContentTypes.add('User Workflow Document','0x010107')  # _Hidden; Items for use in user defined workflows.
        $this.ParentContentTypes.add('Wiki Page','0x010108')  # Document Content Types; Create a new wiki page.
        $this.ParentContentTypes.add('Basic Page','0x010109')  # Document Content Types; Create a new basic page.
        $this.ParentContentTypes.add('Web Part Page','0x01010901')  # Document Content Types; Create a new Web Part page.
        $this.ParentContentTypes.add('Link to a Document','0x01010A')  # Document Content Types; Create a link to a document in a different location.
        $this.ParentContentTypes.add('Dublin Core Columns','0x01010B')  # Document Content Types; The Dublin Core metadata element set.
        $this.ParentContentTypes.add('Event','0x0102')  # List Content Types; Create a new meeting, deadline or other event.
        $this.ParentContentTypes.add('Reservations','0x0102004F51EFDEA49C49668EF9C6744C8CF87D')  # List Content Types; Reserve resource.
        $this.ParentContentTypes.add('Schedule and Reservations','0x01020072BB2A38F0DB49C3A96CF4FA85529956')  # List Content Types; Create a new appointment and reserve a resource.
        $this.ParentContentTypes.add('Schedule','0x0102007DBDC1392EAF4EBBBF99E41D8922B264')  # List Content Types; Create new appointment.
        $this.ParentContentTypes.add('Issue','0x0103')  # List Content Types; Track an issue or problem.
        $this.ParentContentTypes.add('Announcement','0x0104')  # List Content Types; Create a new news item, status or other short piece of information.
        $this.ParentContentTypes.add('Link','0x0105')  # List Content Types; Create a new link to a Web page or other resource.
        $this.ParentContentTypes.add('Contact','0x0106')  # List Content Types; Store information about a business or personal contact.
        $this.ParentContentTypes.add('Message','0x0107')  # List Content Types; Create a new message.
        $this.ParentContentTypes.add('Task','0x0108')  # List Content Types; Track a work item that you or your team needs to complete.
        $this.ParentContentTypes.add('Workflow Task (SharePoint 2013)','0x0108003365C4474CAE8C42BCE396314E88E51F')  # List Content Types; Create a SharePoint 2013 Workflow Task
        $this.ParentContentTypes.add('Workflow Task','0x010801')  # _Hidden; A work item created by a workflow that you or your team needs to complete.
        $this.ParentContentTypes.add('Administrative Task','0x010802')  # _Hidden; An administrative work item that an administrator needs to complete.
        $this.ParentContentTypes.add('Workflow History','0x0109')  # _Hidden; The history of a workflow.
        $this.ParentContentTypes.add('Person','0x010A')  # _Hidden; 
        $this.ParentContentTypes.add('SharePointGroup','0x010B')  # _Hidden; 
        $this.ParentContentTypes.add('DomainGroup','0x010C')  # _Hidden; 
        $this.ParentContentTypes.add('Post','0x0110')  # List Content Types; Create a new blog post.
        $this.ParentContentTypes.add('Comment','0x0111')  # List Content Types; Create a new blog comment.
        $this.ParentContentTypes.add('East Asia Contact','0x0116')  # List Content Types; Store information about a business or personal contact.
        $this.ParentContentTypes.add('Folder','0x0120')  # Folder Content Types; Create a new folder.
        $this.ParentContentTypes.add('RootOfList','0x012001')  # _Hidden; 
        $this.ParentContentTypes.add('Discussion','0x012002')  # Folder Content Types; Create a new discussion topic.
        $this.ParentContentTypes.add('Summary Task','0x012004')  # Folder Content Types; Group and describe related tasks that you or your team needs to complete.
        $this.ParentContentTypes.add('Document Collection Folder','0x0120D5')  # _Hidden; Create a new Document Collection Folder

    }

    [void] GetVersion(){
        #Show PowerShell version in PowerShell script GetVersion.ps1
        Write-Host "##[debug] PowerShell Version"
        $global:PSVersionTable.PSVersion 
    }
    
    [void] Dispose(){
        # get key value from KeyCollection, KeyCollection to Array
        foreach($key in $($this.Connections.Keys)){ 
            $this.Connections[$key]=$null;    
        }
        $this.Connections.Clear();
        $this.connection =$null;
        $this.ParentContentTypes.Clear();
        $this.connection = $null
        $this.ctx = $null;
        $this.webFields =$null;
        $this.webContentTypes =$null;
        $this.webLists =$null;
        $this.cred=$null;
        $this.ClientId=$null;
        $this.Tenant=$null;
        $this.CertificatePath=$null;
        $this.Connections = $null;
    }
    
    [void] ConnectPipeline($url){
        $this.Url=$url;
        # connection cache
        if($this.Connections.ContainsKey($url)){
            $this.connection = $this.Connections[$url];
            $this.ClientId=$this.connection.ClientId
        }
        else{
            if(-not [string]::IsNullOrEmpty($this.ClientId)){
                $this.connection = Connect-PnPOnline -Url $url -Tenant $this.Tenant -ReturnConnection -ClientId $this.ClientId -CertificatePath $this.CertificatePath
            }
            else{
                $this.connection = Connect-PnPOnline -url $url -Tenant $this.Tenant -CertificatePath $this.CertificatePath -ReturnConnection
                $this.ClientId=$this.connection.ClientId
            }
            $this.Connections.add($url,$this.connection);
        }

        $this.ctx = Get-PnPContext -Connection $this.connection 
        Set-PnPContext -Context $this.ctx -Connection $this.connection

        #Get all Site columns from the site
        $this.webFields = $this.ctx.Web.Fields
        $this.webContentTypes = $this.ctx.Web.ContentTypes # $this.ctx.Web.AvailableContentTypes #
        $this.webLists = $this.ctx.Web.Lists
        $this.ctx.Load($this.webFields)
        $this.ctx.Load($this.webContentTypes)
        $this.ctx.Load($this.webLists)
        $this.ctx.ExecuteQuery()

        # Error: "The Push Notifications feature is not activated on the site"
        # PhonePNSubscriber	41e1d4bf-b1a2-47f7-ab80-d5d6cbba3092	15	Web
        # https://www.techtask.com/sharepoint-foundation-2013-liste-aller-features-und-feature-ids/
        $this.EnableFeature("41e1d4bf-b1a2-47f7-ab80-d5d6cbba3092", "Web")
    }

    [void] Connect($cred, $url){
        $this.Url=$url;
        # connection cache
         if($this.Connections.ContainsKey($url)){
            $this.connection = $this.Connections[$url];
            $this.ClientId=$this.connection.ClientId
        }
        else{
            if($null -eq $cred -and $null -eq $this.cred){
                $this.cred = Get-Credential -Message "Enter SPO credentials"
            }
            elseif($null -ne $cred){
                $this.cred = $cred
            }
            else{ #($cred -eq $null -and $this.cred -ne $null ) {
            # ignore
            }

            if(-not [string]::IsNullOrEmpty($this.ClientId)){
                $this.connection = Connect-PnPOnline -Url $url -ReturnConnection -ClientId $this.ClientId -Credentials $this.cred   #-Interactive #-UseWebLogin
            }
            else{
                $this.connection = Connect-PnPOnline -Url $url -ReturnConnection -Credentials $this.cred  #-Interactive #-UseWebLogin
                $this.ClientId=$this.connection.ClientId
            }
            $this.Connections.add($url,$this.connection);
        }

        $this.ctx = Get-PnPContext -Connection $this.connection 
        Set-PnPContext -Context $this.ctx -Connection $this.connection

        #Get all Site columns from the site
        $this.webFields = $this.ctx.Web.Fields
        $this.webContentTypes = $this.ctx.Web.ContentTypes # $this.ctx.Web.AvailableContentTypes #
        $this.webLists = $this.ctx.Web.Lists
        $this.ctx.Load($this.webFields)
        $this.ctx.Load($this.webContentTypes)
        $this.ctx.Load($this.webLists)
        $this.ctx.ExecuteQuery()

        # Error: "The Push Notifications feature is not activated on the site"
        # PhonePNSubscriber	41e1d4bf-b1a2-47f7-ab80-d5d6cbba3092	15	Web
        # https://www.techtask.com/sharepoint-foundation-2013-liste-aller-features-und-feature-ids/
        $this.EnableFeature("41e1d4bf-b1a2-47f7-ab80-d5d6cbba3092", "Web")
    }

    # return PnP.PowerShell.Commands.Base.PnPConnection
    [object] GetNewConnection($url){
        $con=$null;
        if($this.Connections.ContainsKey($url)){
            $con = $this.Connections[$url];
        }
        else
        {
            #connection with credential parameter
            if($null -ne $this.cred){
                if(-not [string]::IsNullOrEmpty($this.ClientId)){
                    $con = Connect-PnPOnline -Url $url -ReturnConnection -ClientId $this.ClientId -Credentials $this.cred   #-Interactive #-UseWebLogin
                }
                else{
                    $con = Connect-PnPOnline -Url $url -ReturnConnection -Credentials $this.cred  #-Interactive #-UseWebLogin
                }
            } # connection with cert authoization
            else{
                if(-not [string]::IsNullOrEmpty($this.ClientId)){
                    $con = Connect-PnPOnline -Url $url -Tenant $this.Tenant -ReturnConnection -ClientId $this.ClientId -CertificatePath $this.CertificatePath
                }
                else{
                    $con = Connect-PnPOnline -url $url -Tenant $this.Tenant -CertificatePath $this.CertificatePath -ReturnConnection
                }
            }
            $this.Connections.add($url,$con);
        }
        return $con;
    }

    [void] QuickLaunchEnabled([bool]$activate){
        #Get the Web
        $web = Get-PnPWeb -Connection $this.connection
        if($web.QuickLaunchEnabled -eq $activate){
            return;
        }
        #hide left navigation bar in sharepoint online
        $web.QuickLaunchEnabled = $activate
        $web.Update()
        Invoke-PnPQuery -Connection $this.connection
    }

    

    ################
    # before connect sharepoint admin site https://xxxx-admin.sharepoint.com
    # Connect-PnPOnline -Url https://xxxx-admin.sharepoint.com
    # Enable External Sharing for Existing AD Users (Including Guest users!)
    # Options: Disabled, ExistingExternalUserSharingOnly, ExternalUserSharingOnly, ExternalUserAndGuestSharing
    ################
    [void] EnableExternalSharing($url,$option){

        try
        {
            Write-Host "##[debug] EnableExternalSharing to $($url)" -ForegroundColor Yellow
            
            Set-PnPTenantSite -Url $url -SharingCapability $option

            Write-Host "##[section] EnableExternalSharing to $($url)" -ForegroundColor Green 
        }
        catch{
            Write-Host "##[error] EnableExternalSharing $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    ################
    # before connect sharepoint admin site https://xxxx-admin.sharepoint.com
    # Connect-PnPOnline -Url https://xxxx-admin.sharepoint.com
    ################
    [void] CreateCommunicationSite($title, $url, $owner, $lcid){

        try
        {
            Write-Host "##[debug] NewPnPSite to $($title) -> $($url)" -ForegroundColor Yellow
            # lcid: https://learn.microsoft.com/de-de/previous-versions/office/sharepoint-csom/jj167546(v=office.15)
            New-PnPSite -Type CommunicationSite -Title $title -Url $url -Owner $owner -Lcid $lcid

            Write-Host "##[section] NewPnPSite to $($title) -> $($url)" -ForegroundColor Green 
        }
        catch{
            Write-Host "##[error] NewPnPSite $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    [void] GetClientSidePage($title, $url){
        try
        {
            Write-Host "##[debug] GetClientSidePage to $($title) -> $($url)" -ForegroundColor Yellow
            
            $Page = Get-PnPClientSidePage -Identity $title #Home.aspx
            
            $Page.Save();
            
            #Publish the page
            $Page.Publish()

            Write-Host "##[section] GetClientSidePage to $($title) -> $($url)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] GetClientSidePage $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] SetNewPage($title, $url){

        try
        {
            Write-Host "##[debug] NewPnPSite to $($title) -> $($url)" -ForegroundColor Yellow
            
            #Create new page
            $Page = Add-PnPPage -Name "News" -LayoutType Article -Connection $this.connection
            
            #Set Page properties
            Set-PnPPage -Identity $Page -Title "News" -CommentsEnabled:$False -HeaderType Default -Connection $this.connection
            
            #Add Section to the Page
            Add-PnPPageSection -Page $Page -SectionTemplate OneColumn -Connection $this.connection
            
            #Add Text to Page
            Add-PnPPageTextPart -Page $Page -Text "Welcome To News Portal" -Section 1 -Column 1 -Connection $this.connection
            
            #Add News web part to the section
            Add-PnPPageWebPart -Page $Page -DefaultWebPartType News -Section 1 -Column 1 -Connection $this.connection
            
            #Add List to Page
            Add-PnPPageWebPart -Page $Page -DefaultWebPartType List -Section 1 -Column 1 -WebPartProperties -Connection $this.connection @{ selectedListId = "21b99d39-834f-4991-b5f9-bd095fa0633c"}
            
            # configure the page
            $Page.RemovePageHeader();

            $Page.Save();
            
            #Publish the page
            $Page.Publish()



            Write-Host "##[section] NewPnPSite to $($title) -> $($url)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] NewPnPSite $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # DenyAddAndCustomizePagesStatus Default: Enabled
    # $adminUrl='https://xxx-admin.sharepoint.com/', 
    # $siteUrl='https://xxx.sharepoint.com/sites/Dev_Fachanwendung_221028', 
    # $enabled=$false
    [void] SetDenyAddAndCustomizePagesStatus($adminUrl, $siteUrl, $enabled){
        $connectTenant = Connect-PnPOnline -Url $adminUrl -Credentials (Get-Credential) -ReturnConnection
        Get-PnPTenantSite -Detailed -Url $siteUrl -Connection $connectTenant | Select-Object url,DenyAddAndCustomizePages

        Set-PnpTenantSite -Identity $siteUrl -DenyAddAndCustomizePages:$enabled -Connection $connectTenant | Out-Null
        
        Get-PnPTenantSite -Detailed -Url $siteUrl -Connection $connectTenant | Select-Object url,DenyAddAndCustomizePages   

        $connectTenant = $null
    }

    [object] GetSiteCollectionAppCatalog($siteUrl){
        
        Write-host -f Yellow "##[debug] Retrieving all site collection App Catalogs from SharePoint Online";

        $appCatalogsCsom = $this.ctx.Web.TenantAppCatalog.SiteCollectionAppCatalogsSites;
        $this.ctx.Load($appCatalogsCsom);
        $this.ctx.ExecuteQuery();
        
        $appCatalogsCsom
        $appCatalog = $appCatalogsCsom | where {$_.AbsoluteUrl -eq $siteUrl}
        return $appCatalog
    }

    [void] ActivateSiteAppCoatalog($siteUrl){
        
        $catalog = $this.GetSiteCollectionAppCatalog($siteUrl)
        #$catalog = Get-PnPSiteCollectionAppCatalog -Connection $this.connection
        if($null -eq $catalog){
            Add-PnPSiteCollectionAppCatalog -Site $siteUrl -Connection $this.connection
        }
        else{
            Write-host -f Yellow "##[debug] SiteCollection AppCatalog is already active!"
        }
    }

    # Uninstall All Apps in Web
    #
    [object] UninstallSiteApp($identity){
        $apps = Uninstall-PnPApp -Scope Site -Connection $this.connection -Identity $identity | Where-Object {$_.InstalledVersion -ne $null}
        return $apps;
    }
    
    # Get All Apps in Web
    #
    [object] GetSiteApp(){
        $apps = Get-PnPApp -Scope Site -Connection $this.connection | Where-Object {$_.InstalledVersion -ne $null}
        return $apps;
    }
    
    # Remove App in Web by identity 
    [object] RemoveSiteApp($identity){
        $apps = Remove-PnPApp -Scope Site -Connection $this.connection -Identity $identity | Where-Object {$_.InstalledVersion -ne $null}
        return $apps;
    }

    # Get All Apps in Web
    # before Connect to Admin Site
    [object] GetTenantApp(){
        $apps = Get-PnPApp -Scope Tenant -Connection $this.connection
        return $apps;
    }
    
    # Remove App in Web by identity 
    # before Connect to Admin Site
    [object] RemoveTenantApp($identity){
        $apps = Remove-PnPApp -Scope Tenant -Connection $this.connection -Identity $identity
        return $apps;
    }

    # Add an App to SharePoint Online App Catalog “From Your Organization”
    # before Connect to Admin Site
    # AppCatalogURL = "https://xxx.sharepoint.com/sites/Apps" / https://xxx.sharepoint.com/sites/appcatalog
    # Parameters
    # $AppFilePath = "C:\Temp\script-editor.sppkg"
    [void] AddTenantAppToAppCatalog($AppFilePath){
        
        #Connect to SharePoint Online App Catalog site
        #Connect-PnPOnline -Url $AppCatalogURL -Interactive
        
        #Add App to App catalog - upload app to sharepoint online app catalog using powershell
        $App = Add-PnPApp -Path $AppFilePath -Scope Tenant -Connection $this.connection
        
        #Deploy App to the Tenant
        Publish-PnPApp -Identity $App.ID -Scope Tenant -Connection $this.connection
    }

    # Install an Tenant App to SharePoint Online Site “From Your Organization”
    # before Connect to Admin Site
    # Parameters:
    # $AppName  = "Modern Script Editor web part by Puzzlepart"
    [void] InstallTenantApp($AppName){
        
        #Get the App from App Catalog
        $App = Get-PnPApp -Scope Tenant -Connection $this.connection | Where {$_.Title -eq $AppName}
        
        #Install App to the Site
        Install-PnPApp -Identity $App.Id -Connection $this.connection
    }

    # Add an App to SharePoint Online App Site Catalog 
    # before Connect SiteCollection
    # AppCatalogURL = "https://xxx.sharepoint.com/sites/..../AppCatalog"
    # Parameters
    # $AppFilePath = "C:\Temp\script-editor.sppkg"
    [void] AddAppToSiteCollectionCatalog($AppFilePath){
        
        #Connect to SharePoint Online App Catalog site
        #Connect-PnPOnline -Url $AppCatalogURL -Interactive
        
        #Add App to App catalog - upload app to sharepoint online app catalog using powershell
        $App = Add-PnPApp -Path $AppFilePath -Connection $this.connection -Scope Site
        
        #Deploy App to the SiteCollection
        Publish-PnPApp -Identity $App.ID -Scope Site -Connection $this.connection 
    }

    # Install an App to SharePoint Online Site 
    # Parameters:
    # $siteurl = "https://xxx.sharepoint.com/sites/...."
    # $AppName  = "Modern Script Editor web part by Puzzlepart"
    [void] InstallAppOnSite($AppName){
        
        #Get the App from App Catalog
        $App = Get-PnPApp -Scope Site -Connection $this.connection | Where {$_.Title -match $AppName} 
        
        #Install or update App at Site
        if($null -eq $app.InstalledVersion){
            Install-PnPApp -Identity $App.Id -Scope Site -Connection $this.connection
        }
        else{
            if($App.count -gt 1){
                for($i=0;$i -lt $App.count;$i++){
                    Update-PnPApp -Identity $App[$i].Id -Scope Site -Connection $this.connection
                }
            }
            else {
                Update-PnPApp -Identity $App.Id -Scope Site -Connection $this.connection
            }
        }
    }

    # FeatureId guid or string
    # Scope     "Site","Web"
    [void] EnableFeature($FeatureId, $Scope){
    
        #get the Feature
        $Feature = Get-PnPFeature -Scope $Scope -Identity $FeatureId -Connection $this.connection
     
        #Get the Feature status
        If($Feature.DefinitionId -eq $null)
        {   
            #sharepoint online powershell enable feature
            Write-host -f Yellow "##[debug] Activating Feature..."
            Enable-PnPFeature -Scope $Scope -Identity $FeatureId -Connection $this.connection -Force
     
            Write-host -f Green "##[section] Feature Activated Successfully!"
        }
        Else
        {
            Write-host -f Gray "##[debug] Feature is already active!"
        }    
    }

    [void] AddHubSiteAssociation($siteURL, $hubSiteUrl){
        Write-Host "##[debug] add HubSite association" -ForegroundColor Yellow
        Add-PnPHubSiteAssociation -Site $siteURL -HubSite $hubSiteUrl -Connection $this.connection
        Write-Host "##[section] done..." -ForegroundColor Green
    }

    [void] GetTermSet([string]$termGroup, [string]$identity){
        Write-Host "##[debug] verify if termsets exist.." -ForegroundColor Yellow

        try 
        {
            Get-PnPTermSet -Identity $identity -TermGroup $termGroup -ErrorAction Stop -Connection $this.connection   
            Write-Host "##[section] Locations exists $($identity)" -ForegroundColor Green 
        }
        catch{
            Write-Host "##[error] Locations does not exist $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [string] GetGuid(){
        $Id = [GUID]::NewGuid().Guid;
        return $Id;
    }

    [string] GetGuidString(){
        $Id = [guid]::NewGuid().Guid;
        return $Id.Replace("-","");
    }
    
    # Copy a List or Document Library at destination location [string]$sourceUrl == current url
    [void] CopyListToDestination([string]$sourceList,[string]$destinationUrl,[string]$destinationList){
        
        try {
            Write-Host "##[warning] List copy source: $($this.Url); $($sourceList) dest:$($destinationUrl); $($destinationList)" -ForegroundColor Yellow

            #Copy list to another site 
            Copy-PnPList -Identity $sourceList -Title $destinationList -DestinationWebUrl $destinationUrl -Connection $this.connection
        }
        catch {
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy a List or Document Library at source location (backup) [string]$sourceUrl/[string]$destinationUrl == current url
    [void] CopyList([string]$sourceList,[string]$destinationList){
        
        try {
            Write-Host "##[warning] List copy source: $($this.Url); $($sourceList) dest:$($destinationList)" -ForegroundColor Yellow

            #Copy list to another site 
            Copy-PnPList -Identity $sourceList -Title $destinationList -Connection $this.connection
        }
        catch {
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy a List at source location (backup)
    # If your list is using any lookup fields, make sure you copy the parent list first! Otherwise, 
    # you may run into “Invoke-PnPSiteTemplate : Value does not fall within the expected range.” error.
    [void] CopyListToDestination([string]$sourceList,[string]$destinationUrl){
        
        try {
            Write-Host "##[warning] List copy source: $($this.Url); $($sourceList) dest:$($destinationUrl);" -ForegroundColor Yellow

            #Parameters
            $TemplateFile ="$env:TEMP\Template.xml"
            
            #Create the Template
            Get-PnPSiteTemplate -Out $TemplateFile -ListsToExtract $sourceList -Handlers Lists -Connection $this.connection
            
            #Get Data from source List
            Add-PnPDataRowsToSiteTemplate -Path $TemplateFile -List $sourceList -Connection $this.connection
            
            #Connect to destination Site
            $con=$this.GetNewConnection($destinationUrl)
            
            #Apply the Template
            Invoke-PnPSiteTemplate -Path $TemplateFile -Connection $con
        }
        catch {
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy a Page at source location 
    # If your list is using any lookup fields, make sure you copy the parent list first! Otherwise, 
    # you may run into “Invoke-PnPSiteTemplate : Value does not fall within the expected range.” error.
    # $PageName = "About.aspx"
    [void] CopyPage([string]$destinationUrl,[string]$pageName){
        
        try {
            Write-Host "##[warning] CopyPage source: $($this.Url); $($pageName) dest:$($destinationUrl)" -ForegroundColor Yellow
            
            #Export the Source page
            $TempFile = [System.IO.Path]::GetTempFileName()
            Export-PnPPage -Force -Identity $PageName -Out $TempFile -Connection $this.connection
            
            #Import the page to the destination site
            $con = $this.GetNewConnection($destinationUrl);
            Invoke-PnPSiteTemplate -Path $TempFile -Connection $con
        }
        catch {
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Fix Text in all Page at source location 
    # If your list is using any lookup fields, make sure you copy the parent list first! Otherwise, 
    # you may run into “Invoke-PnPSiteTemplate : Value does not fall within the expected range.” error.
    # $type = Uri as additinal option to replace relative Urls
    [void] ReplacePagesText([string]$currentText, [string]$newText, [string]$type){
        
        try {
            Write-Host "##[warning] CopyPages source: $($this.Url); Text:$($currentText) to $($newText)" -ForegroundColor Yellow
            
            # Export all pages from the source
            $TempFile = [System.IO.Path]::GetTempFileName()
            Get-PnPSiteTemplate -Out $TempFile -Handlers PageContents -IncludeAllClientSidePages -Force -Connection $this.connection
            
            $xmldata = $this.GetXml($TempFile);
            # change original to new Text
            $xmldata.InnerXml = $xmldata.InnerXml.Replace($currentText,$newText);
            
            if($type -eq "Uri"){
                # replace relative url /sites/Dev_...
                $sourceUri=[uri]$currentText;
                $destUrl = [uri]$newText;

                $xmldata.InnerXml = $xmldata.InnerXml.Replace($sourceUri.AbsolutePath, $destUrl.AbsolutePath);
            }
            
            $xmldata.Save($TempFile)

            # Import the page to the destination site
            Invoke-PnPSiteTemplate -Path $TempFile -Connection $this.connection
        }
        catch {
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy a Page at source location 
    # If your list is using any lookup fields, make sure you copy the parent list first! Otherwise, 
    # you may run into “Invoke-PnPSiteTemplate : Value does not fall within the expected range.” error.
    # $PageName = "About.aspx"
    [void] CopyPages([string]$destinationUrl){
        
        try {
            Write-Host "##[warning] CopyPages source: $($this.Url); dest:$($destinationUrl)" -ForegroundColor Yellow
            
            # Export all pages from the source
            $TempFile = [System.IO.Path]::GetTempFileName()
            Get-PnPSiteTemplate -Out $TempFile -Handlers PageContents -IncludeAllClientSidePages -Force -Connection $this.connection
            
            $xmldata = $this.GetXml($TempFile);
            # change original source -to target urls
            $xmldata.InnerXml = $xmldata.InnerXml.Replace($this.Url,$destinationUrl);

            # replace relative url /sites/Dev_...
            $sourceUri=[uri]$this.Url;
            $destUrl = [uri]$destinationUrl;

            $xmldata.InnerXml = $xmldata.InnerXml.Replace($sourceUri.AbsolutePath, $destUrl.AbsolutePath);
            $xmldata.Save($TempFile)

            # Import the page to the destination site
            $con = $this.GetNewConnection($destinationUrl);
            Invoke-PnPSiteTemplate -Path $TempFile -Connection $con
        }
        catch {
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # clear target before clear navigation 
    [void] ClearTopNavigationBar([string]$destinationUrl) {
        try 
        {
            Write-Host "##[warning] ClearNavigation target:$($destinationUrl)" -ForegroundColor Yellow
            
            $con = $this.GetNewConnection($destinationUrl);
            # Get Quick Launch Navigation
            $NavigationNodeCollection = Get-PnPNavigationNode -Location TopNavigationBar -Connection $con 

            # Get the Link to Delete
            #$LinkTitle = "Home"
            #$NavigationNode = $NavigationNodeCollection | Where-Object { $_.Title -eq  $LinkTitle}

            # Delete Link from left navigation
            foreach($NavigationNode in $NavigationNodeCollection){
                Remove-PnPNavigationNode -Identity $NavigationNode.Id -Connection $con -Force
            }
        }
        catch {
            Write-Host "##[error] CopyNavigation $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy current navigation to target before clear navigation 
    [void] ClearAndCopyTopNavigationBar([string]$destinationUrl) {
        try 
        {
            Write-Host "##[warning] CopyNavigation source: $($this.Url); dest:$($destinationUrl)" -ForegroundColor Yellow

            # $this.ClearTopNavigationBar($destinationUrl)
            $ImplementedError = New-Object System.NotImplementedException
            throw $ImplementedError
        }
        catch {
            Write-Host "##[error] CopyNavigation $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # clear target before clear navigation 
    [void] ClearQuickLaunch([string]$destinationUrl) {
        try 
        {
            Write-Host "##[warning] ClearNavigation target:$($destinationUrl)" -ForegroundColor Yellow
            
            $con = $this.GetNewConnection($destinationUrl);
            # Get Quick Launch Navigation
            $NavigationNodeCollection = Get-PnPNavigationNode -Location QuickLaunch -Connection $con 

            # Get the Link to Delete
            #$NavigationNode = $NavigationNodeCollection | Where-Object { $_.Title -eq  $LinkTitle}

            # Delete Link from left navigation
            foreach($NavigationNode in $NavigationNodeCollection){
                Remove-PnPNavigationNode -Identity $NavigationNode.Id -Connection $con -Force
            }
        }
        catch {
            Write-Host "##[error] CopyNavigation $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy current navigation to target before clear navigation 
    [void] ClearAndCopyQuickLaunch([string]$destinationUrl) {
        try 
        {
            Write-Host "##[warning] CopyNavigation source: $($this.Url); dest:$($destinationUrl)" -ForegroundColor Yellow
            
            $this.ClearQuickLaunch($destinationUrl)

            $TempFile = [System.IO.Path]::GetTempFileName()
            #-PersistMultiLanguageResources `
            Get-PnPSiteTemplate -Out $TempFile `
                                -Connection $this.connection `
                                -Handlers "Navigation" `
                                -ExcludeHandlers "None, AuditSettings, ComposedLook, CustomActions, ExtensibilityProviders, Features, Fields, Files, Lists, Pages, Publishing, RegionalSettings, SearchSettings, SitePolicy, SupportedUILanguages, TermGroups, Workflows, SiteSecurity, ContentTypes, PropertyBagEntries, PageContents, WebSettings, ImageRenditions, ApplicationLifecycleManagement, Tenant, WebApiPermissions, SiteHeader, SiteFooter, Theme, SiteSettings, All" `
                                -Force

            $xmldata = $this.GetXml($TempFile);

            # change root menu to current site
            $menuStart="Start"
            # replace relative url /sites/Dev_...
            #$destUrl=[uri]$destinationUrl;
            #$site = Get-PnPSite -Connection $this.connection
            $xmldata.Provisioning.Templates.ProvisioningTemplate.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode | Where-Object Title -eq $menuStart |  ForEach-Object { $_.Url = "{sitecollection}" #$site.Url; 
                [pscustomobject]@{ Name = $_.Title; Url = $_.Url }}   
            
            $xmldata.Save($TempFile)

            # Import the page to the destination site
            $con = $this.GetNewConnection($destinationUrl);
            # enable=$true disable=$false the navigation on the SharePoint Online site
            Set-PnPWeb -QuickLaunchEnabled:$true -Connection $con
            Invoke-PnPSiteTemplate -Path $TempFile -Connection $con
        }
        catch {
            Write-Host "##[error] CopyNavigation $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy current navigation to target
    [void] CopyNavigation([string]$destinationUrl) {
        try 
        {
            Write-Host "##[warning] CopyNavigation source: $($this.Url); dest:$($destinationUrl)" -ForegroundColor Yellow
            #$web = Get-PnPWeb -Connection $this.connection 
            #$title = $web.Title;
            $TempFile = [System.IO.Path]::GetTempFileName()
            #-PersistMultiLanguageResources `
            Get-PnPSiteTemplate -Out $TempFile `
                                -Connection $this.connection `
                                -Handlers "Navigation" `
                                -ExcludeHandlers "None, AuditSettings, ComposedLook, CustomActions, ExtensibilityProviders, Features, Fields, Files, Lists, Pages, Publishing, RegionalSettings, SearchSettings, SitePolicy, SupportedUILanguages, TermGroups, Workflows, SiteSecurity, ContentTypes, PropertyBagEntries, PageContents, WebSettings, ImageRenditions, ApplicationLifecycleManagement, Tenant, WebApiPermissions, SiteHeader, SiteFooter, Theme, SiteSettings, All" `
                                -Force

            $xmldata = $this.GetXml($TempFile);

            # Import the page to the destination site
            $con = $this.GetNewConnection($destinationUrl);
            Invoke-PnPSiteTemplate -Path $TempFile -Connection $con
        }
        catch {
            Write-Host "##[error] CopyNavigation $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Update the List View Formatting Definition
    [void] UpdateListViewFormatting([string]$listName, [string]$ViewName, [string]$listViewFormattingJSON) {
        try 
        {
            #$listViewFormattingJSON = Get-Content -Raw -Path "$folderPath\ListFormatting.View.$listName.$($_.Id).json";
            #$listViewFormattingJSON = $listViewFormattingJSON.ToString()
            $listViewColumnDefinition = Get-PnPView -List $listName -Identity $ViewName -Connection $this.connection
            $listViewColumnDefinition | Set-PnPView -Values  @{CustomFormatter = $listViewFormattingJSON} -Connection $this.connection
        }
        catch {
            Write-Host "##[error] CopyNavigation $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Copy current navigation to target
    [void] CopyViews([string]$listName,[string]$targetUrl) {
        try 
        {
            Write-Host "##[warning] CopyView $($listName) source: $($this.Url); dest:$($targetUrl)" -ForegroundColor Yellow
            $views = Get-PnPView -List $listName -Include ViewType,ViewFields,Aggregations,Paged,ViewQuery,RowLimit,ViewJoins,JSLink,ListViewXml, `
                                                                  ViewType,ViewData,ViewType2,StyleId,TabularView,Threaded,Title,Toolbar,ToolbarTemplateName, `
                                                                  RequiresClientIntegration,RowLimit,ServerRelativeUrl,ViewProjectedFields,ViewQuery,ReadOnlyView, `
                                                                  MobileDefaultView,MobileView,ModerationType,NewDocumentTemplates,OrderedView,Paged,PersonalView, `
                                                                  GridLayout,Hidden,HtmlSchemaXml,Id,ImageUrl,IncludeRootFolder,Method, `
                                                                  CustomFormatter,CustomOrder,DefaultView,DefaultViewForContentType,EditorModified,Formats, `
                                                                  Aggregations,AggregationsStatus,AssociatedContentTypeId,BaseViewId,CalendarViewStyles,ColumnWidth, `
                                                                  VisualizationInfo,PageRenderType,Scope,ServerRelativePath,ContentTypeId `
                                                                  -Connection $this.connection
            # Import the page to the destination site
            $con = $this.GetNewConnection($targetUrl);
            #Invoke-PnPSiteTemplate -Path $TempFile -Connection $con
            foreach($view in $views){
                $fields = New-Object System.Collections.ArrayList
                $view.ViewFields.ForEach({ $fields.Add($_) });
                $regex = [regex]'(?:[^\/])+$(?<=.aspx)';
                $fileName=$regex.Match($view.ServerRelativeUrl).value;
                $fileName=$fileName.Replace(".aspx","");
                
                $viewTitle=$view.Title;
                $listView = $this.ExistListView($listName, $viewTitle);
                if($listView -eq $false){
                    $v = Add-PnPView -Connection $con -List $listName -Title $fileName -SetAsDefault:$view.DefaultView -Fields $fields -Query $view.ViewQuery -Aggregations $view.Aggregations
                    $viewTitle=$fileName;
                }

                #cannot set view url when creating the list. the following workaroud required
                $v = Set-PnPView -Connection $con -List $listName -Identity $viewTitle  -Fields $fields -Values @{
                    
                    Title = $view.Title; 
                    Aggregations = $view.Aggregations;
                    ListViewXml = $view.ListViewXml;
                    ViewData = $view.ViewData;
                    ViewQuery = $view.ViewQuery;
                    #ViewType = $view.ViewType;
                    ViewType2 = $view.ViewType2;
                    ViewJoins = $view.ViewJoins;
                   # AggregationsStatus = $view.AggregationsStatus;
                    CalendarViewStyles =$view.CalendarViewStyles;
                    ColumnWidth = $view.ColumnWidth;
                    CustomFormatter  = $view.CustomFormatter;
                    CustomOrder  = $view.CustomOrder;
                    EditorModified  = $view.EditorModified;
                    Formats  = $view.Formats;
                    GridLayout  = $view.GridLayout;
                    Hidden  = $view.Hidden;
                    IncludeRootFolder  = $view.IncludeRootFolder;
                    JSLink  = $view.JSLink;
                    Method  = $view.Method;
                    MobileDefaultView  = $view.MobileDefaultView;
                    MobileView  = $view.MobileView;
                    NewDocumentTemplates  = $view.NewDocumentTemplates;
                    # OrderedView  = $view.OrderedView;
                    Paged  = $view.Paged;
                    RowLimit  = $view.RowLimit;
                    Scope  = $view.Scope;
                    TabularView  = $view.TabularView;
                    Tag  = $view.Tag;
                    Toolbar  = $view.Toolbar;
                    #ViewFields  = $fields;
                    ViewProjectedFields  = $view.ViewProjectedFields;
                    #.AssociatedContentTypeId
                    #.ContentTypeId
                    #DefaultViewForContentType
                    #HtmlSchemaXml 
                    #Id
                    #ImageUrl
                    #ListViewXml
                    #ObjectVersion
                    #.Path
                    #ServerObjectIsNull
                    # BaseViewId = $view.AggregationsStatus;
                    # ModerationType  = $view.ModerationType;
                    # PageRenderType  = $view.PageRenderType;
                    # PersonalView  = $view.PersonalView;
                    # ReadOnlyView  = $view.ReadOnlyView;
                    # RequiresClientIntegration  = $view.RequiresClientIntegration;
                    #ServerRelativePath  = $view.ServerRelativePath;
                    #ServerRelativeUrl  = $view.ServerRelativeUrl;
                    # StyleId  = $view.StyleId;
                    # Threaded   = $view.Threaded;
                    # ToolbarTemplateName  = $view.ToolbarTemplateName;
                    #TypedObject  = $view.TypedObject;
                    # ViewType  = $view.ViewType;
                } 
            }
        }
        catch {
            Write-Host "##[error] CopyView $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] AddLookupAdditionalShowField([System.Xml.XmlElement]$node, $fieldRefId){
        if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
            $ShowFields=$node.AdditionalColumns.Split(",");
            foreach($ShowField in $ShowFields){
                $fieldId=$this.GetGuid();

                Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                    ID='{$($fieldId)}' 
                    DisplayName='$($node.DisplayName):$($ShowField)' 
                    StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                    Name='$($node.InternalName)' 
                    Group='$($node.Group)' 
                    Required='$($node.Required)' 
                    List='$($node.List)' 
                    ShowField='$($ShowField)' 
                    FieldRef='{$($fieldRefId)}'
                    UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                    SourceID='{{siteid}}'></Field>" -Connection $this.connection
            }
        }
    }

    # same site field add web at list
    [void] AddLookupToList([System.Xml.XmlElement]$node)
    {
        try {
                
            $id=$this.GetGuid();

            # Get ID of Source Web
            #$title=$node.WebUrl;
            $sourceWeb=Get-PnPWeb -Connection $this.connection #$this.GetSubWeb($title);

            # Get ID of List where it is
            $listSourceUrl=$node.ListSource;
            $listSource=$this.GetList($listSourceUrl, $sourceWeb);

            # Get ID of List whose items display
            $listDestinationUrl=$node.List;
            $listDestination=$this.GetList($listDestinationUrl, $sourceWeb);

            #Check if the column name exists
            $newField = $listSource.Fields | where {$_.InternalName -eq $node.InternalName}
            if($null -ne $newField) 
            {
                Write-host "##[debug] Site Column $($node.InternalName) already exists!" -f Yellow
            }
            else {
                    
                if($null -ne $sourceWeb){
                    $sourceWebID=$sourceWeb.Id
                
                    if($null -ne $listSource -and $null -ne $listDestination){
                        $listDestinationID= $listDestination.Id

                        switch ($node.Type) {
                            "LookupMulti"{
                        
                                $PrependId = $node.PrependId ? $node.PrependId :'FALSE';
                                #TODO:SourceID=WebId?; Replace the sourceid with the correct webid
                                #SourceID:	Optional Text. Contains the namespace that defines the field, such as http://schemas.microsoft.com/sharepoint/v3 or the GUID of the list in which the custom field was created.
                                #PrependId: Optional Boolean. Used by lookup fields that can have multiple values. Specify TRUE to display the item ID of a target item as well as the value of the target field in Edit and New item forms.
                                Add-PnPFieldFromXml -List $listSource.Id -FieldXml "<Field Type='LookupMulti' Mult='TRUE' PrependId='$($PrependId)' ID='{$($id)}' 
                                                                        DisplayName='$($node.DisplayName)' 
                                                                        StaticName='$($node.InternalName)' 
                                                                        Name='$($node.InternalName)' 
                                                                        Group='$($node.Group)' 
                                                                        Required='$($node.Required)' 
                                                                        List='{$($listDestinationID)}' 
                                                                        WebId='$($sourceWebID)' 
                                                                        ShowField='$($node.ShowField)' 
                                                                        UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                        SourceID='$($sourceWebID)'></Field>" -Connection $this.connection
                            
                                if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
                                    $ShowFields=$node.AdditionalColumns.Split(",");
                                    foreach($ShowField in $ShowFields){
                                        $fieldId=$this.GetGuid();
                                        Add-PnPFieldFromXml -List $listSource.Id -FieldXml "<Field Type='Lookup' 
                                                                            ID='{$($fieldId)}' 
                                                                            DisplayName='$($node.DisplayName):$($ShowField)' 
                                                                            StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Name='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Group='$($node.Group)' 
                                                                            Required='$($node.Required)' 
                                                                            List='{$($listDestinationID)}' 
                                                                            WebId='$($sourceWebID)' 
                                                                            ShowField='$($ShowField)' 
                                                                            FieldRef='{$($id)}'
                                                                            UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                            Mult='FALSE'
                                                                            SourceID='$($sourceWebID)'
                                                                            ></Field>" -Connection $this.connection
                                    }
                                }
                                break;
                            }
                            "Lookup"{
                                # Add fiield
                                Add-PnPFieldFromXml -List $listSource.Id -FieldXml "<Field Type='Lookup' 
                                                                    ID='{$($id)}' 
                                                                    DisplayName='$($node.DisplayName)' 
                                                                    StaticName='$($node.InternalName)' 
                                                                    Name='$($node.InternalName)' 
                                                                    Group='$($node.Group)' 
                                                                    Required='$($node.Required)' 
                                                                    List='{$($listDestinationID)}' 
                                                                    WebId='$($sourceWebID)' 
                                                                    ShowField='$($node.ShowField)' 
                                                                    UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                    Mult='FALSE'
                                                                    SourceID='$($sourceWebID)'
                                                                    ></Field>" -Connection $this.connection
                                                                    <#
                                                                    JSLink='~site/crosssitelookupapplib/clienttemplates.js?field=$($node.InternalName)' 
                                                                        xmlns:csl='Plumsail.CrossSiteLookup'  
                                                                        csl:ShowNew='false' 
                                                                        csl:RetrieveItemsUrlTemplate='function (term, page) {&#xA;  if (!term || term.length == 0) {&#xA;    return &quot;{WebUrl}/_api/web/lists(\'{ListId}\')/items?`$select=Id,{LookupField}&amp;`$orderby=Created desc&amp;`$top=10&quot;;&#xA;  }&#xA;  return &quot;{WebUrl}/_api/web/lists(\'{ListId}\')/items?`$select=Id,{LookupField}&amp;`$orderby={LookupField}&amp;`$filter=startswith({LookupField}, \'&quot; + encodeURIComponent(term) + &quot;\')&amp;`$top=10&quot;;&#xA;}' 
                                                                        csl:ItemFormatResultTemplate='function(item) {&#xA;  return \'&lt;span class=&quot;csl-option&quot;&gt;\' + item[&quot;{LookupField}&quot;] + \'&lt;/span&gt;\'&#xA;}' 
                                                                        csl:NewText='Add new item'
                                                                        csl:NewContentType=''
                                                                        ></Field>" -Connection $this.connection
                                                                        #>

                                if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
                                    $ShowFields=$node.AdditionalColumns.Split(",");
                                    foreach($ShowField in $ShowFields){
                                        $fieldId=$this.GetGuid();
                        
                                        Add-PnPFieldFromXml -List $listSource.Id -FieldXml "<Field Type='Lookup' 
                                                                            ID='{$($fieldId)}' 
                                                                            DisplayName='$($node.DisplayName):$($ShowField)' 
                                                                            StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Name='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Group='$($node.Group)' 
                                                                            Required='$($node.Required)' 
                                                                            List='{$($listDestinationID)}' 
                                                                            WebId='$($sourceWebID)' 
                                                                            ShowField='$($ShowField)' 
                                                                            FieldRef='{$($id)}'
                                                                            UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                            Mult='FALSE'
                                                                            SourceID='$($sourceWebID)'
                                                                            ></Field>" -Connection $this.connection
                                    }
                                }
                            
                                break;
                            }
                            default:{}
                        }
                    }
                    else {
                        Write-Host "##[warning] List dosen't exist source: $($listSource.Title) destList:$($node.List)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "##[warning] Web dosen't exist $($node.WebName)" -ForegroundColor Red
                }
            } 
        }
        catch{
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] AddField([System.Xml.XmlElement]$node){
        try
        {
            if( $node.InternalName -eq "Title" -or
                $node.InternalName -eq "ID" -or
                $node.InternalName -eq "Name" -or
                $node.InternalName -eq "Description" -or
                $node.InternalName -eq "_ExtendedDescription")
            {
                Write-host "##[warning] Site Column $($node.InternalName) already exists!" -f Yellow
                return;
            }

            Write-Host "##[debug] Field to add InternalName: $($node.InternalName) DisplayName: $($node.DisplayName)" -ForegroundColor Yellow
            $xml="";
            $id=$this.GetGuid();

            #Check if the column name exists
            $newField = $this.webFields | where {$_.InternalName -eq $node.InternalName}
            if($null -ne $newField) 
            {
                Write-host "##[warning] Site Column $($node.InternalName) already exists!" -f Yellow
            }
            else {
                #SourceID="{{siteid}}" 
                switch ($node.Type) {
                    "Choice" {  
                        $choices = $node.Choices.Split(',')

                        if($node.Required -eq "TRUE"){
                            Add-PnPField -Type Choice -ID $id -DisplayName $node.DisplayName -InternalName $node.InternalName -Choices $choices -Group $node.Group -Required -Connection $this.connection; 
                        }
                        else{    
                            Add-PnPField -Type Choice -ID $id -DisplayName $node.DisplayName -InternalName $node.InternalName -Choices $choices -Group $node.Group -Connection $this.connection; 
                        }

                        if($node.FillInChoice -eq "True"){
                            $this.ctx = Get-PnPContext -Connection $this.connection 

                            Set-PnPContext -Context $this.ctx -Connection $this.connection

                            #Get all Site columns from the site
                            $this.webFields = $this.ctx.Web.Fields
                            $this.webContentTypes = $this.ctx.Web.ContentTypes # $this.ctx.Web.AvailableContentTypes #
                            $this.webLists = $this.ctx.Web.Lists
                            $this.ctx.Load($this.webFields)
                            $this.ctx.Load($this.webContentTypes)
                            $this.ctx.Load($this.webLists)
                            $this.ctx.ExecuteQuery()
                            
                            #Retrieve field
                            $fieldChoice = $this.webFields | where {$_.InternalName -eq $node.InternalName}
                            #setting the FillInChoice property
                            $fieldChoice.FillInChoice = $true
                            $fieldChoice.Update()
                            $this.ctx.Load($fieldChoice)
                            $this.ctx.ExecuteQuery() 
                        }
                        
                        if($node.DefaultValue){
                            Set-PnPField -Identity $node.InternalName -Values @{DefaultValue=$node.DefaultValue} -Connection $this.connection 
                        }
                        break;
                    }
                    "MultiChoice" {  
                        $choices = $node.Choices.Split(',')
                        if($node.Required -eq "TRUE"){
                            Add-PnPField -Type MultiChoice -ID $id -DisplayName $node.DisplayName -InternalName $node.InternalName -Choices $choices -Group $node.Group -Required $node.Required -Connection $this.connection; 
                        }
                        else{
                            Add-PnPField -Type MultiChoice -ID $id -DisplayName $node.DisplayName -InternalName $node.InternalName -Choices $choices -Group $node.Group -Connection $this.connection; 
                        }
                        if($node.DefaultValue){
                            Set-PnPField -Identity $node.InternalName -Values @{DefaultValue=$node.DefaultValue} -Connection $this.connection 
                        }
                        break;
                    }
                    "Boolean" {  
                        if($node.Required -eq "TRUE"){
                            Add-PnPField -Type Boolean -ID $id -DisplayName $node.DisplayName -InternalName $node.InternalName -Group $node.Group -Required $node.Required -Connection $this.connection; 
                        }
                        else{
                            Add-PnPField -Type Boolean -ID $id -DisplayName $node.DisplayName -InternalName $node.InternalName -Group $node.Group -Connection $this.connection; 
                        }
                        if($node.DefaultValue){
                            Set-PnPField -Identity $node.InternalName -Values @{DefaultValue=$node.DefaultValue} -Connection $this.connection 
                        }
                        break;
                    }
                    "Number"{
                        if($node.Max){
                            Add-PnPFieldFromXml -FieldXml "<Field Type='Number' ID='{$($id)}' CommaSeparator='FALSE' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Required='$($node.Required)' CustomFormatter='$($node.CustomFormatter)' Group='$($node.Group)' Min='$($node.Min)' Max='$($node.Max)' EnforceUniqueValues='FALSE' Percentage='$($node.Percentage)'></Field>" -Connection $this.connection 
                        }
                        else{
                            Add-PnPFieldFromXml -FieldXml "<Field Type='Number' ID='{$($id)}' CommaSeparator='FALSE' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Required='$($node.Required)' CustomFormatter='$($node.CustomFormatter)' Group='$($node.Group)' nforceUniqueValues='FALSE' Percentage='$($node.Percentage)'></Field>" -Connection $this.connection 
                        }    
                        break;
                    }
                    "Currency"{
                        $node.Decimals=$node.Decimals -eq "" ? 0 : $node.Decimals;
                        if($node.Max){
                            Add-PnPFieldFromXml -FieldXml "<Field Type='Currency' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' Decimals='$($node.Decimals)' LCID='$($node.LCID)' Min='$($node.Min)' Max='$($node.Max)' EnforceUniqueValues='FALSE' Percentage='$($node.Percentage)'></Field>" -Connection $this.connection 
                        }
                        else{
                            Add-PnPFieldFromXml -FieldXml "<Field Type='Currency' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' Decimals='$($node.Decimals)' LCID='$($node.LCID)' EnforceUniqueValues='FALSE' Percentage='$($node.Percentage)'></Field>" -Connection $this.connection 
                        }
                        break;
                    }
                    "Calculated"{
                        $attributes="";
                        #Integer, Text, DateTime, Boolean, Number, Currency
                        switch($node.ResultType){
                            "Text" {
                                $attributes="ResultType='$($node.ResultType)'";
                                break;
                            }
                            "DateTime" {
                                $attributes="ResultType='$($node.ResultType)' Format='$($node.Format)'"; # Format: DateTime, DateOnly
                                break;
                            }
                            "Integer" {
                                $attributes="ResultType='$($node.ResultType)' Percentage='$($node.Percentage)'"; # Percentage='False/True' without dezimal
                                break;
                            }
                            "Boolean" {
                                $attributes="ResultType='$($node.ResultType)'";
                                break;
                            }
                            "Number" {
                                $node.Decimals=$node.Decimals -eq "" ? 0 : $node.Decimals;
                                $attributes="ResultType='$($node.ResultType)' Decimals='$($node.Decimals)' Percentage='$($node.Percentage)'"; # Percentage='False/True' Decimals='0/1/2/3/4/5' Empty (automatic) (1 / 1,0 / 100)
                                break;
                            }
                            "Currency" {
                                $node.Decimals=$node.Decimals -eq "" ? 0 : $node.Decimals;
                                $attributes="ResultType='$($node.ResultType)' Decimals='$($node.Decimals)' LCID='$($node.LCID)'"; # LCID="1031"
                                break;
                            }
                            default {
                                $attributes="";
                            }
                        }

                        $xml="";
                        if($node.Required -eq "TRUE"){
                            $xml="<Field Type='Calculated' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' $($attributes)><Formula>$($node.Formula)</Formula></Field>"
                        }
                        else{
                            #<FieldRefs><FieldRef Name='ComplaintEingang' /></FieldRefs>
                            $xml="<Field Type='Calculated' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' $($attributes)><Formula>$($node.Formula)</Formula></Field>"
                        }
                        Add-PnPFieldFromXml -FieldXml $xml -Connection $this.connection

                        break;
                    }
                    "DateTime"{
                        Add-PnPFieldFromXml -FieldXml "<Field Type='DateTime' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='$($node.Format)' FriendlyDisplayFormat='Disabled'></Field>" -Connection $this.connection 
                        break;
                    }
                    "User"{
                        Add-PnPFieldFromXml -FieldXml "<Field Type='User' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' UserSelectionMode='$($node.UserSelectionMode)'></Field>" -Connection $this.connection
                        break;
                    }
                    "UserMulti"{
                        Add-PnPFieldFromXml -FieldXml "<Field Type='UserMulti' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' UserSelectionMode='$($node.UserSelectionMode)' ShowField='$($node.ShowField)' List='UserInfo' Mult='TRUE'></Field>" -Connection $this.connection
                        break;
                    }
                    "Text"{
                        #Add-PnPFieldFromXml -FieldXml "<Field Type='Text' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' />" -Connection $this.connection
                    
                        if($node.Required -eq "TRUE"){
                            Add-PnPField -ID $id -Type Text -DisplayName $node.DisplayName -InternalName $node.InternalName -Group $node.Group -Required -Connection $this.connection
                        }
                        else{
                            Add-PnPField -ID $id -Type Text -DisplayName $node.DisplayName -InternalName $node.InternalName -Group $node.Group -Connection $this.connection
                        }
                        <# #>
                        break;
                    }
                    "Url"{
                        # $Format="" or "Image"
                        # don't use Hyperlink as Format, it must be empty 
                        # default is Hyperlink
                        if($node.Format -eq "Hyperlink"){
                            $node.Format=""
                        }
                        Add-PnPFieldFromXml -FieldXml "<Field Type='URL' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' Format='$($node.Format)' />" -Connection $this.connection
                        break;
                    }
                    "Note"{
                        if($node.RichText -eq "" -or $node.RichText -eq "FALSE")
                        {
                            #Add-PnPField -ID $id -Type Note -DisplayName $node.DisplayName -InternalName $node.InternalName -Group $node.Group -Connection $this.connection

                            #Add-PnPFieldFromXml -FieldXml "<Field Type='Note' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' NumLines='$($node.NumLines)' ></Field>" -Connection $this.connection
                            $xml="<Field Type='Note' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' NumLines='$($node.NumLines)' ></Field>"
                        }
                        else{
                            #Add-PnPFieldFromXml -FieldXml "<Field Type='Note' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' RichText='TRUE' NumLines='$($node.NumLines)' RichTextMode='$($node.RichTextMode)'></Field>" -Connection $this.connection
                            $xml="<Field Type='Note' ID='{$($id)}' DisplayName='$($node.DisplayName)' StaticName='$($node.InternalName)' Name='$($node.InternalName)' Group='$($node.Group)' Required='$($node.Required)' RichText='TRUE' NumLines='$($node.NumLines)' RichTextMode='$($node.RichTextMode)'></Field>"
                        }
                        #$i=[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView
                        #Define XML for Field Schema
                        $newField = $this.webFields.AddFieldAsXml($xml,$True,16) # [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView
                        $this.ctx.ExecuteQuery()   

                        break;
                    }
                    "LookupMulti"{
                        # Get ID of Web
                        $title=$node.WebName;
                        $web=$this.GetSubWeb($title);
                        #$web=$webs[0]; #$web.sCount -gt 1
                        # Get ID of List
                        
                        if($node.List -eq "Self")
                        {
                            Add-PnPFieldFromXml -FieldXml "<Field Type='LookupMulti' 
                                                                    Mult='TRUE' 
                                                                    ID='{$($id)}' 
                                                                    DisplayName='$($node.DisplayName)' 
                                                                    StaticName='$($node.InternalName)' 
                                                                    Name='$($node.InternalName)' 
                                                                    Group='$($node.Group)' 
                                                                    Required='$($node.Required)' 
                                                                    List='$($node.List)' 
                                                                    ShowField='$($node.ShowField)' 
                                                                    UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                    SourceID='{{siteid}}'></Field>" -Connection $this.connection
                           
                            if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
                                $ShowFields=$node.AdditionalColumns.Split(",");
                                foreach($ShowField in $ShowFields){
                                    $fieldId=$this.GetGuid();
                    
                                    Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                                        ID='{$($fieldId)}' 
                                        DisplayName='$($node.DisplayName):$($ShowField)' 
                                        StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                                        Name='$($node.InternalName)_x003a_$($ShowField)' 
                                        Group='$($node.Group)' 
                                        Required='$($node.Required)' 
                                        List='$($node.List)' 
                                        ShowField='$($ShowField)' 
                                        FieldRef='{$($id)}'
                                        UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                        SourceID='{{siteid}}'></Field>" -Connection $this.connection
                                }
                            }
                            return;
                        }

                        $list=$node.List;
                        $list=$this.GetList($list, $web);
                        if($null -ne $web){
                            $lookupWebID=$web.Id
                            #$web=$webs[0]; #$web.sCount -gt 1
                            # Get ID of List
                            $list=$this.GetList($node.List, $web);
                            if($null -ne $list){
                                $lookupListID= $list.Id
                                                

                                    $PrependId = $node.PrependId ? $node.PrependId :'FALSE';
                                    #TODO:SourceID=WebId?; Replace the sourceid with the correct webid
                                    #SourceID:	Optional Text. Contains the namespace that defines the field, such as http://schemas.microsoft.com/sharepoint/v3 or the GUID of the list in which the custom field was created.
                                    #PrependId: Optional Boolean. Used by lookup fields that can have multiple values. Specify TRUE to display the item ID of a target item as well as the value of the target field in Edit and New item forms.
                                    Add-PnPFieldFromXml -FieldXml "<Field Type='LookupMulti' Mult='TRUE' PrependId='$($PrependId)' ID='{$($id)}' 
                                                                            DisplayName='$($node.DisplayName)' 
                                                                            StaticName='$($node.InternalName)' 
                                                                            Name='$($node.InternalName)' 
                                                                            Group='$($node.Group)' 
                                                                            Required='$($node.Required)' 
                                                                            List='{$($lookupListID)}' 
                                                                            WebId='$($lookupWebID)' 
                                                                            ShowField='$($node.ShowField)' 
                                                                            UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                            SourceID='{{siteid}}'></Field>" -Connection $this.connection
                               
                                if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
                                    $ShowFields=$node.AdditionalColumns.Split(",");
                                    foreach($ShowField in $ShowFields){
                                        $fieldId=$this.GetGuid();
                        
                                        Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                                                                            ID='{$($fieldId)}' 
                                                                            DisplayName='$($node.DisplayName):$($ShowField)' 
                                                                            StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Name='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Group='$($node.Group)' 
                                                                            Required='$($node.Required)' 
                                                                            List='{$($lookupListID)}' 
                                                                            WebId='$($lookupWebID)' 
                                                                            ShowField='$($ShowField)' 
                                                                            UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                            Mult='FALSE'
                                                                            FieldRef='{$($id)}'
                                                                            SourceID='$($lookupWebID)'
                                                                            ></Field>" -Connection $this.connection
                                    }
                                }
                            }
                        }
                        break;
                    }
                    "Lookup"{
                        # Get ID of Web
                        $title=$node.WebName;
                        $web=$this.GetSubWeb($title);

                        if($node.List -eq "Self")
                        {
                            Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                                                                    ID='{$($id)}' 
                                                                    DisplayName='$($node.DisplayName)' 
                                                                    StaticName='$($node.InternalName)' 
                                                                    Name='$($node.InternalName)' 
                                                                    Group='$($node.Group)' 
                                                                    Required='$($node.Required)' 
                                                                    List='$($node.List)' 
                                                                    ShowField='$($node.ShowField)' 
                                                                    UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                    SourceID='{{siteid}}'></Field>" -Connection $this.connection
                            
                            if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
                                $ShowFields=$node.AdditionalColumns.Split(",");
                                foreach($ShowField in $ShowFields){
                                    $fieldId=$this.GetGuid();
                    
                                    Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                                        ID='{$($fieldId)}' 
                                        DisplayName='$($node.DisplayName):$($ShowField)' 
                                        StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                                        Name='$($node.InternalName)_x003a_$($ShowField)' 
                                        Group='$($node.Group)' 
                                        Required='$($node.Required)' 
                                        List='$($node.List)' 
                                        ShowField='$($ShowField)' 
                                        FieldRef='{$($id)}'
                                        UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                        SourceID='{{siteid}}'></Field>" -Connection $this.connection
                                }
                            }
                            
                            return;
                        }

                        if($null -ne $web){
                            $lookupWebID=$web.Id
                            #$web=$webs[0]; #$web.sCount -gt 1
                            # Get ID of List
                            $list=$this.GetList($node.List, $web);
                            if($null -ne $list){
                                $lookupListID= $list.Id
                                    
                                # Add fiield
                                Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                                                                    ID='{$($id)}' 
                                                                    DisplayName='$($node.DisplayName)' 
                                                                    StaticName='$($node.InternalName)' 
                                                                    Name='$($node.InternalName)' 
                                                                    Group='$($node.Group)' 
                                                                    Required='$($node.Required)' 
                                                                    List='{$($lookupListID)}' 
                                                                    WebId='$($lookupWebID)' 
                                                                    ShowField='$($node.ShowField)' 
                                                                    UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                    Mult='FALSE'
                                                                    SourceID='$($lookupWebID)'
                                                                    ></Field>" -Connection $this.connection
                                                                    <#
                                                                    JSLink='~site/crosssitelookupapplib/clienttemplates.js?field=$($node.InternalName)' 
                                                                        xmlns:csl='Plumsail.CrossSiteLookup'  
                                                                        csl:ShowNew='false' 
                                                                        csl:RetrieveItemsUrlTemplate='function (term, page) {&#xA;  if (!term || term.length == 0) {&#xA;    return &quot;{WebUrl}/_api/web/lists(\'{ListId}\')/items?`$select=Id,{LookupField}&amp;`$orderby=Created desc&amp;`$top=10&quot;;&#xA;  }&#xA;  return &quot;{WebUrl}/_api/web/lists(\'{ListId}\')/items?`$select=Id,{LookupField}&amp;`$orderby={LookupField}&amp;`$filter=startswith({LookupField}, \'&quot; + encodeURIComponent(term) + &quot;\')&amp;`$top=10&quot;;&#xA;}' 
                                                                        csl:ItemFormatResultTemplate='function(item) {&#xA;  return \'&lt;span class=&quot;csl-option&quot;&gt;\' + item[&quot;{LookupField}&quot;] + \'&lt;/span&gt;\'&#xA;}' 
                                                                        csl:NewText='Add new item'
                                                                        csl:NewContentType=''
                                                                        ></Field>" -Connection $this.connection
                                                                        #>

                                if($node.AdditionalColumns -and $node.AdditionalColumns.length -gt 0){
                                    $ShowFields=$node.AdditionalColumns.Split(",");
                                    foreach($ShowField in $ShowFields){
                                        $fieldId=$this.GetGuid();
                        
                                        Add-PnPFieldFromXml -FieldXml "<Field Type='Lookup' 
                                                                            ID='{$($fieldId)}' 
                                                                            DisplayName='$($node.DisplayName):$($ShowField)' 
                                                                            StaticName='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Name='$($node.InternalName)_x003a_$($ShowField)' 
                                                                            Group='$($node.Group)' 
                                                                            Required='$($node.Required)' 
                                                                            List='{$($lookupListID)}' 
                                                                            WebId='$($lookupWebID)' 
                                                                            ShowField='$($ShowField)' 
                                                                            UnlimitedLengthInDocumentLibrary='$($node.UnlimitedLengthInDocumentLibrary)'
                                                                            Mult='FALSE'
                                                                            FieldRef='{$($id)}'
                                                                            SourceID='$($lookupWebID)'
                                                                            ></Field>" -Connection $this.connection
                                    }
                                }
                            }
                            else {
                                Write-Host "##[warning] List dosen't exist $($node.List)" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "##[warning] Web dosen't exist $($node.WebName)" -ForegroundColor Red
                        }
                        break;
                    }
                    "Taxonomymut"{
                        if($node.Required -eq "TRUE"){
                            Add-PnPTaxonomyField -ID $($id) -DisplayName $node.DisplayName -InternalName $node.InternalName -TermSetPath $node.TermSetPath -Group $node.Group -MultiValue -Required $node.Required -Connection $this.connection 
                        }
                        else{    
                            Add-PnPTaxonomyField -ID $($id) -DisplayName $node.DisplayName -InternalName $node.InternalName -TermSetPath $node.TermSetPath -Group $node.Group -MultiValue -Connection $this.connection 
                        }
                        break;
                    }
                    "Taxonomy"{
                        if($node.Required -eq "TRUE"){
                            Add-PnPTaxonomyField -ID $($id) -DisplayName $node.DisplayName -InternalName $node.InternalName -TermSetPath $node.TermSetPath -Group $node.Group -Required $node.Required -Connection $this.connection 
                        }
                        else{    
                            Add-PnPTaxonomyField -ID $($id) -DisplayName $node.DisplayName -InternalName $node.InternalName -TermSetPath $node.TermSetPath -Group $node.Group -Connection $this.connection 
                        }
                        break;
                    }
                    Default {}
                }
                Write-Host "##[section] Field to added $($node.InternalName)" -ForegroundColor Green
            }
        }
        catch{
            Write-Host "##[error] Field $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [System.Xml.XmlDocument] GetXml([string]$path){
        ## Add argument validation logic here
        #$this.Devices[$slot] = $dev

        ##### Environment Pfade laden
        $transitionFile = [System.IO.Path]::Combine($path)
        [xml]$transitionXml = Get-Content $transitionFile -ErrorAction SilentlyContinue # -Raw 
        # or ($xmlDoc = [xml]::new()).Load((Convert-Path file.xml))

        ##### Verbinden auf die Staging Seite und die Listen und Bibliotheken auflisten
        if ($null -eq $transitionXml) {
            Write-Host "##[warning] no environment config at '$($transitionFile)'"
            return $null;
        }

        return $transitionXml;
    }

    [System.Xml.XmlElement] GetConfiguration([string]$path){
        ## Add argument validation logic here
        #$this.Devices[$slot] = $dev

        ##### Environment Pfade laden
        $transitionFile = [System.IO.Path]::Combine($path)
        [xml]$transitionXml = Get-Content $transitionFile -ErrorAction SilentlyContinue # -Raw 
        # or ($xmlDoc = [xml]::new()).Load((Convert-Path file.xml))

        ##### Verbinden auf die Staging Seite und die Listen und Bibliotheken auflisten
        if ($null -eq $transitionXml) {
            Write-Host "##[warning] no environment config at '$($transitionFile)'"
            return $null;
        }

        return $transitionXml.Configuration;
    }

    [void] UpdateContentType([System.Xml.XmlElement]$node){
        try
        {
            #Check if the ContentType name exists
            $contentType = $this.webContentTypes | where {$_.Name -eq $node.ContentTypeName}
            if($null -ne $contentType) 
            {
                if($node.ComponentIdNew -ne "" -or $node.ComponentIdEdit -ne "" -or $node.ComponentIdDisplay -ne ""){
                    #$this.AddContentTypeCustomForm($node.ContentTypeName, $node.ComponentId);
                    $this.AddContentTypeCustomForm($node.ContentTypeName, $node.ComponentIdNew,$node.ComponentIdEdit,$node.ComponentIdDisplay)
                }
                else{ # clear
                    $this.ClearContentTypeCustomForm($node.ContentTypeName);
                }
            }
        }
        catch{
            Write-Host "##[error] ContentType $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # contenttype add or update properties
    [void] AddContentType([System.Xml.XmlElement]$node){
        try
        {
            #Check if the ContentType name exists
            $contentType = $this.webContentTypes | Where-Object {$_.Name -eq $node.ContentTypeName}
            if($null -ne $contentType) 
            {
                # update CustomForm
                $this.UpdateContentType([System.Xml.XmlElement]$node);

                Write-host "##[warning] ContentType $($node.ContentTypeName) already exists!" -f Yellow
                return;
            }

            [string]$ContentTypeId = $null;
            # first check: standard sharepoint
            [string]$ParentContentTypesId = $this.ParentContentTypes[$node.ParentContentType]
            if($ParentContentTypesId.Length -eq 0){
                # second check: inheritance ContentTypes 
                # get (double digits)
                $ContentTypeId = $this.GetFreeInheritanceContentTypeId($node.ParentContentType);
            }
            else{
                $ContentTypeId = "$($ParentContentTypesId)00$($this.GetGuidString())"
            }

            Write-Host "##[debug] ContentType to add $($node.ContentTypeName)" -ForegroundColor Yellow
            Add-PnPContentType -Name $node.ContentTypeName -Description $node.Description -ContentTypeId $ContentTypeId -Group $node.Group -Connection $this.connection 
            Write-Host "##[section] ContentType to added $($node.ContentTypeName)" -ForegroundColor Green
       
            if($node.ComponentIdNew -ne "" -or $node.ComponentIdEdit -ne "" -or $node.ComponentIdDisplay -ne ""){
               # $this.AddContentTypeCustomForm($node.ContentTypeName, $node.ComponentId);
               $this.AddContentTypeCustomForm($node.ContentTypeName, $node.ComponentIdNew,$node.ComponentIdEdit,$node.ComponentIdDisplay)
            }
        }
        catch{
            Write-Host "##[error] ContentType $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] AddFieldToContentType([System.Xml.XmlElement]$node){
        try
        {
            $CType = $this.webContentTypes | Where {$_.Name -eq $node.ContentType}
            if($null -ne $CType)
            {
                #Get Columns from the content type
                $this.ctx.Load($CType.Fields)
                $this.ctx.ExecuteQuery()

                #Check if the ContentType Field name exists
                $newField = $CType.Fields | where {$_.InternalName -eq $node.Field}
                if($null -ne $newField) 
                {
                    Write-host "##[warning] Field $($node.Field) already exists at $($node.ContentType)!" -f Yellow
                    return;
                }

                Write-Host "##[debug] FieldToContentType to add $($node.Field)" -ForegroundColor Yellow
                Add-PnPFieldToContentType -Field $node.Field -ContentType $node.ContentType -UpdateChildren $true -Connection $this.connection; 
                Write-Host "##[section] FieldToContentType to added $($node.Field)" -ForegroundColor Green
            }
       
        }
        catch{
            Write-Host "##[error] FieldToContentType $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # list name
    # identity = id
    # value = @{"Key" = "value"}@{}
    [void] UpdateListItem($listName, $identity, $values){
        try
        {
            Write-Host "##[debug] AddListItem $($listName)" -ForegroundColor Yellow
                                                                                                # @{"Title" = "Test Title"; "Category"="Test Category"}
            Set-PnPListItem -List $listName -Identity $identity -Connection $this.connection -Values $values
            
            Write-Host "##[section] AddListItem $($listName)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] AddListItem $($_.Exception.Message)" -ForegroundColor Red
        }
    }
                                
    [void] AddListItem($list, $contenttype, $values){
        try
        {
            Write-Host "##[debug] AddListItem $($list)" -ForegroundColor Yellow
                                                                        # @{"Title" = "Test Title"; "Category"="Test Category"}
            Add-PnPListItem -List $list -ContentType $contenttype -Values $values -Connection $this.connection
            
            Write-Host "##[section] AddListItem $($list)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] AddListItem $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    [void] AddList([System.Xml.XmlElement]$node,$hidden){
        try
        {
            #Check if the List name exists
            $newList = $this.webLists | where {$_.Title -eq $node.Title}
            if($null -ne $newList) 
            {
                Write-host "##[warning] List $($node.Title) already exists at $($this.connection.Url)!" -f Yellow
                return;
            }

            Write-Host "##[debug] List to add $($node.Url)" -ForegroundColor Yellow
            if($node.OnQuickLaunch -eq "true"){
                New-PnPList -Url $node.Url -Title $node.Title -Template $node.TemplateType -OnQuickLaunch -EnableContentTypes -Connection $this.connection;
            }
            else{
                New-PnPList -Url $node.Url -Title $node.Title -Template $node.TemplateType -EnableContentTypes -Connection $this.connection
            }

            if($node.EnableVersioning -eq "true"){
                Set-PnPList -Identity $node.Url -EnableVersioning 1 -Connection $this.connection
                Write-Host "##[section] Set versions $($node.Url)" -ForegroundColor Green
            }

            if($node.EnableMinorVersions -eq "true"){
                Set-PnPList -Identity $node.Url -EnableMinorVersions 1 -MajorVersions $node.MaxVersionLimit -MinorVersions $node.MinorVersionLimit -Connection $this.connection
                Write-Host "##[section] Set EnableMinorVersions $($node.Url)" -ForegroundColor Green
            }

            if($hidden -eq "True"){
                Set-PnPList -Identity $node.Url -Hidden 1 -Connection $this.connection
                Write-Host "##[section] Set Hidden $($node.Url)" -ForegroundColor Green
            }
            Write-Host "##[section] List to added $($node.Url)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] List $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] AddContentTypeToList([System.Xml.XmlElement]$node){
        try
        {
            $list=$null;
            if($node.ListTitle){
                Write-Host "##[debug] ContentTypeToList $($node.ListTitle) add $($node.ContentType)" -ForegroundColor Yellow
                $list = Get-PnPList -Identity $node.ListTitle -Connection $this.connection
            }
            else{
                Write-Host "##[debug] ContentTypeToList $($node.List) add $($node.ContentType)" -ForegroundColor Yellow
                $list = Get-PnPList -Identity $node.List -Connection $this.connection
            }
            
            $cts=$this.ctx.Web.Lists.GetByTitle($list.Title).ContentTypes
            $this.ctx.Load($cts)
            $this.ctx.ExecuteQuery()
            #Check if the ContentType name exists
            $newCT = $cts | where {$_.Name -eq $node.ContentType}
            if($null -ne $newCT) 
            {
                Write-host "##[warning] ContentType $($node.ContentType) already exists!" -f Yellow
                return;
            }

            if($node.DefaultContentType -eq "true"){
                Add-PnPContentTypeToList -List $list -ContentType $node.ContentType -DefaultContentType -Connection $this.connection
            }
            else {
                Add-PnPContentTypeToList -List $list -ContentType $node.ContentType -Connection $this.connection
            }
            Write-Host "##[section] ContentTypeToList $($list.Title) added $($node.ContentType)" -ForegroundColor Green
            
        }
        catch{
            Write-Host "##[error] ContentTypeToList $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    #$parentIdentity, $listUrl, $componentId
    [void] AddContentTypeCustomForm($parentIdentity, $componentId){
        try
        {
            # contenttype
            $ctParent = $this.GetContentType($parentIdentity);

            $clientContext = $this.connection.Context; # Get-PnPContext 

            $ctParent.NewFormClientSideComponentId = $componentId;
            $ctParent.EditFormClientSideComponentId = $componentId;
            $ctParent.DisplayFormClientSideComponentId = $componentId;

            # update child contenttypes $true
            $ctParent.Update($true);
            $clientContext.ExecuteQuery();

            Write-Host "##[section] AddContentTypeCustomForm to $($parentIdentity) -> $($componentId) content type form component ID's updated" -ForegroundColor Green

        }
        catch
        {
            Write-Host "##[error] AddContentTypeCustomForm $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    #$parentIdentity, $listUrl, $componentId
    [void] AddContentTypeCustomForm($parentIdentity, $componentIdNew,$componentIdEdit,$componentIdDisplay){
        try
        {
            # contenttype
            $ctParent = $this.GetContentType($parentIdentity);

            $clientContext = $this.connection.Context; # Get-PnPContext 

            $ctParent.NewFormClientSideComponentId = $componentIdNew;
            $ctParent.EditFormClientSideComponentId = $componentIdEdit;
            $ctParent.DisplayFormClientSideComponentId = $componentIdDisplay;

            # update child contenttypes $true
            $ctParent.Update($true);
            $clientContext.ExecuteQuery();

            Write-Host "##[section] AddContentTypeCustomForm to $($parentIdentity) -> $($componentIdNew),$($componentIdEdit),$($componentIdDisplay) content type form component ID's updated" -ForegroundColor Green

        }
        catch
        {
            Write-Host "##[error] AddContentTypeCustomForm $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    #$parentIdentity, $listUrl, $componentId
    [void] ClearContentTypeCustomForm($parentIdentity){
        try
        {
            # contenttype
            $ctParent = $this.GetContentType($parentIdentity);

            $clientContext = $this.connection.Context; # Get-PnPContext 

            $ctParent.NewFormClientSideComponentId = $null;
            $ctParent.EditFormClientSideComponentId = $null;
            $ctParent.DisplayFormClientSideComponentId = $null;

            # update child contenttypes $true
            $ctParent.Update($true);
            $clientContext.ExecuteQuery();

            Write-Host "##[section] ClearContentTypeCustomForm to $($parentIdentity) content type form component ID's updated" -ForegroundColor Green

        }
        catch
        {
            Write-Host "##[error] ClearContentTypeCustomForm $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    #$contentTypeName, $listUrl, $componentId
    [void] AddContentTypeCustomForm($List, $ContentType, $ComponentId){
        try
        {
            $clientContext = $this.connection.Context; # Get-PnPContext 
            # Add view if not exists
            $list = Get-PnPList -Identity $List -Includes Views,Fields,ContentTypes -Connection $this.connection
            if($null -ne $list) 
            {
                $ct = $list.ContentTypes | where-object {$_.Name -eq $ContentType}
                if ($null -ne $ct)
                {
                    $ct.NewFormClientSideComponentId = $ComponentId;
                    $ct.EditFormClientSideComponentId = $ComponentId;
                    $ct.DisplayFormClientSideComponentId = $ComponentId;

                    # List ContentType has no child elements $false
                    $ct.Update($false);
                    $clientContext.ExecuteQuery();

                    Write-Host "##[section] AddContentTypeCustomForm to $($List) -> $($ContentType) content type form component ID's updated" -ForegroundColor Green
                }
            }
        }
        catch
        {
            Write-Host "##[error] AddContentTypeCustomForm $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    #$contentTypeName, $listUrl, $componentId
    [void] AddContentTypeCustomForm([System.Xml.XmlElement]$node){
        try
        {
            $clientContext = $this.connection.Context; # Get-PnPContext 
            # Add view if not exists
            $list = Get-PnPList -Identity $node.List -Includes Views,Fields,ContentTypes -Connection $this.connection
            if($null -ne $list) 
            {
                $ct = $list.ContentTypes | where-object {$_.Name -eq $node.ContentType}
                if ($null -ne $ct)
                {
                    $ct.NewFormClientSideComponentId = $node.ComponentId;
                    $ct.EditFormClientSideComponentId = $node.ComponentId;
                    $ct.DisplayFormClientSideComponentId = $node.ComponentId;

                    $ct.Update($true);
                    $clientContext.ExecuteQuery();

                    Write-Host "##[section] AddContentTypeCustomForm to $($node.List) -> $($node.ContentType) content type form component ID's updated" -ForegroundColor Green
                }
            }
        }
        catch
        {
            Write-Host "##[error] AddContentTypeCustomForm $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Get Listview or $false
    [bool] ExistListView($listName, $viewName){
        try
        {
            Write-Host "##[debug] ExistListView to $($listName) $($viewName)" -ForegroundColor Yellow
            
            # Add view if not exists
            $list = Get-PnPList -Identity $listName -Includes Views,Fields,ContentTypes -Connection $this.connection
            if($null -ne $list) 
            {
                $viewExist = $list.Views | % {$_.Title -eq $viewName}
                if($viewExist -eq $true){ 
                    return $viewExist;
                }
            }
            else{
                Write-Host "##[error] list not found $($listName) " -ForegroundColor Green    
            }
        }
        catch{
            Write-Host "##[error] ExistListView $($_.Exception.Message)" -ForegroundColor Red
        }
        return $false;
    }

    [void] AddListView([System.Xml.XmlElement]$node){
        try
        {
            Write-Host "##[debug] AddListView to $($node.List) -> $($node.Name)" -ForegroundColor Yellow
            
            # Add view if not exists
            $list = Get-PnPList -Identity $node.List -Includes Views,Fields,ContentTypes -Connection $this.connection
            if($null -ne $list) 
            {
                $viewExist = $list.Views | % {$_.Title -eq $node.Name}
                if($viewExist -eq $true){ 
                    # modified view
                    Set-PnPView -List $node.List -Identity $node.Name -Fields $node.Fields.split(",") -Connection $this.connection
                    Write-Host "##[section] UpdateListView to $($node.List) -> $($node.Name)" -ForegroundColor Green
                }
                else {
                    # create view
                    Add-PnPView -List $node.List -Title $node.UrlName -Fields $node.Fields.split(",") -Connection $this.connection
                    Set-PnPView -List $node.List -Identity $node.UrlName -Values @{Title=$node.Name} -Connection $this.connection

                    Write-Host "##[section] AddListView to $($node.List) -> $($node.Name)" -ForegroundColor Green
                }

                if($node.DefaultView -eq "true"){
                    $this.SetDefaultListView($node);
                }
    
                if($node.SetAsTilesView -eq "true"){
                    $this.SetAsTilesView($node);
                }
            }
            else{
                Write-Host "##[error] list not found $($node.List) " -ForegroundColor Green    
            }
        }
        catch{
            Write-Host "##[error] AddListView $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] SetDefaultListView([System.Xml.XmlElement]$node){
        try
        {
            Write-Host "##[debug] SetDefaultListView to $($node.List) -> $($node.Name)" -ForegroundColor Yellow
            
            #Set "xxx xxxxx" as default view
            $targetView = Get-PnPView -List $node.List -Identity $node.Name -Connection $this.connection

            # Set the view as default
            $clientContext = $this.connection.Context; # Get-PnPContext 
            $targetView.DefaultView = 1;
            $targetView.Update();
            $clientContext.ExecuteQuery();
            
            Write-Host "##[section] SetDefaultListView to $($node.List) -> $($node.Name)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] SetDefaultListView $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] SetAsTilesView([System.Xml.XmlElement]$node){
        try
        {
            Write-Host "##[debug] SetAsTilesView to $($node.List) -> $($node.Name)" -ForegroundColor Yellow
            
            #set view to use Tiles style
            Set-PnPView -List $node.List -Identity $node.Name -Values @{"ViewType2" = "TILES"}
            
            Write-Host "##[section] SetAsTilesView to $($node.List) -> $($node.Name)" -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] SetAsTilesView $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    [void] RemoveWeb([string]$identity) {
        try
        {
            Write-Host "##[debug] RemoveWeb to $($this.connection.Url) -> $($identity) ..." -ForegroundColor Yellow
            
            Remove-PnPWeb -Identity $identity -Force -Connection $this.connection
            
            Write-Host "##[section] RemoveWeb $($this.connection.Url) -> $($identity) done." -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] RemoveWeb $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # [Microsoft.SharePoint.Client.Web]
    [object] AddWeb($siteUrl, $url, $title, $description){
        try
        {
            Write-Host "##[debug] processing web $($url)"
            if ($url -eq "" -or $siteUrl -eq "") {
                Write-Host "##[warning] Url empty URL: $($url) or SiteUrl: $($siteUrl)" -ForegroundColor Yellow
                return $null;
            }
            $siteTemplate = "STS#3"
            $subUrl = "$($siteUrl)$($url)";

            # Remove-PnPWeb -Identity $url -Force
            if([string]::IsNullOrEmpty($title)){ 
                return Get-PnPWeb -Connection $this.connection; 
            }
            
            
            $webs = Get-PnPSubWeb -Recurse -IncludeRootWeb -Connection $this.connection | Where Title -eq $title
            if($webs.Count -eq 0){
                #$webs = Get-PnPSubWeb -Connection $this.connection -Recurse  | Where Url -eq $url
                $web = New-PnPWeb -Title $title -Url $url -Description $description -Locale "1031" -Template $siteTemplate -InheritNavigation -Connection $this.connection
                $web.QuickLaunchEnabled = $false
                $web.update()
                Invoke-PnPQuery -Connection $this.connection

                Write-Host "##[section] found web at $($web.Url)" -ForegroundColor Green 
            
                #$fullUrl = $siteUrl + "/" + $url
                #Connect-PnPOnline -Url $fullUrl -Interactive
            
                #foreach ($template in $rootWeb.templates ) {
                #    Write-Host "##[debug] apply template " $template "at url " $fullUrl
                #    Invoke-PnPSiteTemplate -Path $template
                #}     
                
                return $web; 
            }
            return $webs[0];
        }
        catch{
            Write-Host "##[error] " $($_.Exception.Message) -ForegroundColor Red
            return $null;
        }
    }

    #Microsoft.SharePoint.Client.ContentType
    [string] GetFreeInheritanceContentTypeId($parentIdentity){
        try 
        {
            # rendom for new contentType
            $id=1 | % {Get-Random -Minimum 10 -Maximum 99 }
            # contenttype
            $ctParent = $this.GetContentType($parentIdentity);

            $cts = Get-PnPContentType -Connection $this.connection
            $ct = $cts | SELECT Name,ID | WHERE Id -match "$($ctParent.Id.StringValue)$($id)"
            # check exist contenttype with new id
            #$ct=Get-PnPContentType -Identity "$($ct.Id.StringValue)$id" -ErrorAction Stop -Connection $this.connection   
            if($null -eq $ct){
                return "$($ctParent.Id.StringValue)$($id)";
            }
            else {
                # recursive
                $this.GetFreeInheritanceContentTypeId($parentIdentity);
            }
        }
        catch{
            Write-Host "##[error] Content Type $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }  

    #  [Microsoft.SharePoint.Client.ContentType]
    [object] GetContentType($identity){
        try 
        {
            return Get-PnPContentType -Identity $identity -ErrorAction Stop -Connection $this.connection   
        }
        catch{
            Write-Host "##[error] Content Type $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    } 

    #  [Microsoft.SharePoint.Client.ContentTypeCollection]
    [object] GetContentTypeInSite($identity){
        try 
        {
            return Get-PnPContentType -InSiteHierarchy -ErrorAction Stop -Connection $this.connection   
        }
        catch{
            Write-Host "##[error] Content Type $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }

    [void] RemoveContentTypeFromList($listIdentity, $contentTypeIdentity){
        try 
        {
            Write-Host "##[debug] RemoveContentTypeFromList $($listIdentity) -> $($contentTypeIdentity)" -ForegroundColor Yellow

            $list=Get-PnPList -Identity $listIdentity -Connection $this.connection
            Remove-PnPContentTypeFromList -List $list.Id -ContentType $contentTypeIdentity -Connection $this.connection

            Write-Host "##[section] RemoveContentTypeFromList $($listIdentity) -> $($contentTypeIdentity) done." -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] Content Type $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # [Microsoft.SharePoint.Client.List]
    [object] GetList($listIdentity, $web){
        try
        {
            $connect=$this.connection;
            if($web.Url -ne $this.connection.Url)
            {
                if(-not [string]::IsNullOrEmpty($this.Tenant)){
                    # cert auth, pipeline
                    $connect = Connect-PnPOnline -Url $web.Url -Tenant $this.Tenant -ReturnConnection -ClientId $this.ClientId -CertificatePath $this.CertificatePath       
                }
                else{
                    # credential, normal
                    $connect = Connect-PnPOnline -Url $web.Url -ReturnConnection -ClientId $this.ClientId -Credentials $this.cred   #-Interactive #-UseWebLogin          
                }
            }
            Write-Host "##[debug] Get List $($listIdentity)" -ForegroundColor Yellow
            
            $list = Get-PnPList -Identity $listIdentity -Includes Fields -Connection $connect
            
            Write-Host "##[section] Get List $($listIdentity)" -ForegroundColor Green

            # dispose a specific connection
            $connect = $null;

            return $list;
        }
        catch{
            Write-Host "##[error] List $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }

    # [Microsoft.SharePoint.Client.ListCollection]
    [object] GetLists(){
        try
        {
            Write-Host "##[debug] Get Lists $($this.connection.Url) ..." -ForegroundColor Yellow
            
            $lists = Get-PnPList -Includes Fields -Connection $this.connection
            
            Write-Host "##[section] Get Lists $($this.connection.Url) done." -ForegroundColor Green

            return $lists;
        }
        catch{
            Write-Host "##[error] Lists $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }

    # [Microsoft.SharePoint.Client.FieldCollection]
    [object] GetListFields($listObj){
        try
        {
            Write-Host "##[debug] Get GetListFields $($listObj.Title) ..." -ForegroundColor Yellow
            
            $fields = Get-PnPField -List $listObj.Title -ReturnTyped -Connection $this.connection 
            
            Write-Host "##[section] Get GetListFields $($listObj.Title) done." -ForegroundColor Green

            return $fields;
        }
        catch{
            Write-Host "##[error] GetListFields $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }

    [void] SetSiteLogo($assetLibUrl, $logoName, $assetLibConnection) {
        try 
        {
            Write-Host "##[debug] SetSiteLogo $($assetLibUrl) -> $($logoName)" -ForegroundColor Yellow

            #get uploaded file path
            $pathImg = (Get-PnPListItem -List $AssetLibUrl -Connection $assetLibConnection -Fields FileRef).FieldValues | Where-Object {$_.FileRef -match $LogoName}
            #Update site collection logo
            Set-PnPWeb -SiteLogoUrl $pathImg.FileRef -Connection $this.connection

            Write-Host "##[section] SetSiteLogo $($assetLibUrl) -> $($logoName) done." -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] Content Type $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] RemoveSupportedUILanguage($SiteURL, $cultureIgnore) {
        try 
        {
            Write-Host "##[debug] RemoveSupportedUILanguage $($this.connection.Url) ..." -ForegroundColor Yellow
            #$LanguageID2 = 1031;  #Deutsch
            #$LanguageID1 = 1033;  #English

            #Get the Web
            $Web = Get-PnPWeb -Includes RegionalSettings.InstalledLanguages -Connection $this.connection
            $AvailableLanguage = Get-PnPAvailableLanguage -Connection $this.connection;
            $Culture=($cultureIgnore -join ',');
            $AvailableLanguage | Select Lcid | % { 
                #if($_.Lcid -ne $LanguageID1 -and $_.Lcid -ne $LanguageID2){
                if($Culture -notmatch $_.Lcid ) {
                    $Web.RemoveSupportedUILanguage($_.Lcid);     
                }
            }
            $Web.Update();
            Invoke-PnPQuery -Connection $this.connection
                
            Write-Host "##[debug] RemoveSupportedUILanguage $($this.connection.Url) done." -ForegroundColor Yellow
        }
        catch{
            Write-Host "##[error] RemoveSupportedUILanguage $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # [Microsoft.SharePoint.Client.WebCollection]
    [object] GetSubWebs() {
        try 
        {
            Write-Host "##[debug] GetSubWebs $($this.connection.Url) ..." -ForegroundColor Yellow
            $subwebs = Get-PnPSubWeb -Recurse -IncludeRootWeb -Connection $this.connection #| SELECT Title, Url # -Includes "WebTemplate","Description" | Select ServerRelativeUrl, WebTemplate, Description
            
            Write-Host "##[section] GetSubWebs $($this.connection.Url) done." -ForegroundColor Green
            return $subwebs;
        }
        catch{
            Write-Host "##[error] GetSubWebs $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }
    
    # [Microsoft.SharePoint.Client.Web]
    [object] GetSubWeb($title) {
        try 
        {
            # Remove-PnPWeb -Identity $url -Force
            if([string]::IsNullOrEmpty($title)){ 
                return Get-PnPWeb -Connection $this.connection; 
            }

            Write-Host "##[debug] GetSubWeb $($this.connection.Url) ..." -ForegroundColor Yellow
            $subweb = Get-PnPSubWeb -Recurse -IncludeRootWeb -Connection $this.connection | Where Title -eq $title
            #foreach ($subweb in $subwebs) { 
            
            Write-Host "##[section] GetSubWeb $($this.connection.Url) done." -ForegroundColor Green
            return $subweb -ne $null ? $subweb[0] : $null;
        }
        catch{
            Write-Host "##[error] GetSubWeb $($_.Exception.Message)" -ForegroundColor Red
        }
        return $null;
    }

    [void] SetMultilingual([string]$culture) {
        try 
        {
            Write-Host "##[section] SetMultilingual $($this.connection.Url) -> $($culture) ..." -ForegroundColor Green
            #Get the Web
            $Web = Get-PnPWeb -Includes RegionalSettings.InstalledLanguages -Connection $this.connection
            #Get Available Languages
            $Web.RegionalSettings.InstalledLanguages

            #Add Alternate Language
            if($Web.IsMultilingual -ne $true)
            {
                $Web.IsMultilingual = $true;
            }
            
            $Web.AddSupportedUILanguage($culture);     
            $Web.Update();
            Invoke-PnPQuery -Connection $this.connection
            
            Write-Host "##[section] SetMultilingual $($Web.Url) -> $($culture) done." -ForegroundColor Green
        }
        catch{
            Write-Host "##[error] SetMultilingual $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}