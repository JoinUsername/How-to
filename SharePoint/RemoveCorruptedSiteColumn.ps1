#########################################################
#
#   Howto remove corrupted site column in SharePoint Online
#   https://learn.microsoft.com/en-us/graph/api/site-get?view=graph-rest-1.0&tabs=powershell
#
#   Install-Module Microsoft.Graph
#   Install-Module -Name "PnP.PowerShell"
#########################################################

Import-Module Microsoft.Graph.Sites
Import-Module -Name "PnP.PowerShell"

$tenantId= "...."
$fieldname ="...."

# site collection url
Connect-PnPOnline -Url 'https://... .sharepoint.com/sites/xxxx/'
# get site id
$site = Get-PnPSite -Includes Id
$siteId = $site.Id
# Get-PnPSiteTemplate -Out "PnP-Provisioning-File.xml"

# connect graph
Connect-Graph -TenantId $tenantId -Scopes Sites.Manage.All, Sites.FullControl.All
# Welcome To Microsoft Graph!

# find column Id
$data = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/Sites/$($SiteId)/columns"

## Name                           Value
#----                           -----
#@odata.context                 https://graph.microsoft.com/v1.0/$metadata#sites('x SiteID x'…
#value                          {xxxxxx, xxxxxxxxxxx, xxxxxxxx, Categories…}

$field=$data.value | Where-Object {$_.name -match $fieldname }
# ####################################
# Field result example
######################################
#name                           Value
#----                           -----
#text                           {[appendChangesToExistingText, False], [allowMultipleLines, False], [maxLength, 255], […
#readOnly                       False
#required                       False
#indexed                        False
#hidden                         False
#description
#enforceUniqueValues            False
#id                             00000000-0000-0000-0000-000000000000 // Empty
#columnGroup                    X Group Name X
#name                           X Internal Name X
#displayName                    X Display Name X


# Get-MgSiteColumn -SiteId $SiteId -ColumnDefinitionId $field.id
Remove-MgSiteColumn -SiteId $siteId -ColumnDefinitionId $field.id 