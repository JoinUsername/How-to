# SharePoint Online

## Provisioning 
![image](https://user-images.githubusercontent.com/8308659/219971409-40c006dd-d59a-4c34-a430-fa0637711c4e.png)

### DevOps Pipeline
Install-Module -Name "PnP.PowerShell" -RequiredVersion $(PowerShellVersion) -Force #-AllowPrerelease 

<dl>
  <dt>Connection via Certificate</dt>
  <dd>Connect-PnPOnline -Url $(<em>adminUrl</em>) -ClientId $(<em>ClientId</em>) -Tenant $(<em>Tenant</em>) -CertificatePath $(<em>CertificatePath</em>)</dd>

  <dt>Create SharePoint Site Collection</dt>
  <dd>Write-Host 'Create SPSite'

  .\Scripts\Site\CreateCommunicationSite.ps1 -TechnicalRootUrl '$(<em>TechnicalRootUrl</em>)' -Title '$(<em>Title</em>)' -Owner '$(<em>Owner</em>)' -Lcid $(<em>Lcid</em>).</dd>
</dl>

Reference: SharePoint Online Language IDs
This is a liitle :) list of languages that SharePoint Online supports and the codes to represent them:

Name | Language Tag | LCID 
--- | --- | ---
English | en-US	| 1033 
German	| de-DE | 1031 
Chinese | (Simplified) zh-CN | 2052 
Chinese | (Traditional) zh-TW	| 1028 

Deploying fields to SharePoint can lead to errors, for example, if incorrect properties are used on the field being deployed.

Now you have the problem that you can't delete the field in the frontend, rest and also not via PnP PowerShell.

Here the Grah interface helps.
[RemoveCorruptedSiteColumn.ps1](https://github.com/JoinUsername/How-to/tree/main/SharePoint/Scripts/Repairing/RemoveCorruptedSiteColumn.ps1)
