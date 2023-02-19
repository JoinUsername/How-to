Using Module '..\Classes\Process.psm1'
Param(  
    [Parameter(Mandatory=$True)]
    [string]$TechnicalRootUrl,       # https://xxxx.sharepoint.com/sites/xxx without end slash 
    [Parameter(Mandatory=$True)]
    [string]$Title ,                 # name of site
    [Parameter(Mandatory=$True)]
    [string]$Owner,                  # email of user or group
    [Parameter(Mandatory=$True)]
    [int]$lcid                       # langradge code (DE 1031)
)

################
# before connect sharepoint admin site https://xxxx-admin.sharepoint.com
# Connect-PnPOnline -Url https://xxxx-admin.sharepoint.com
################

Import-Module -Name "PnP.PowerShell"
#$PSScriptRoot + "\Script\Process.psm1"

$process = [Process]::New();
$path = $PSScriptRoot 
$currentPath = split-path -parent $MyInvocation.MyCommand.Definition

try
{
    Write-Host "##[debug] NewPnPSite $($TechnicalRootUrl) ..." -ForegroundColor Yellow

    $process.CreateCommunicationSite($Title, $TechnicalRootUrl,$Owner, $lcid);

    Write-Host "##[section] NewPnPSite $($TechnicalRootUrl) done." -ForegroundColor Green
    ################
    # Enable External Sharing for Existing AD Users (Including Guest users!)
    # Options: Disabled, ExistingExternalUserSharingOnly, ExternalUserSharingOnly, ExternalUserAndGuestSharing
    ################
    $process.EnableExternalSharing($TechnicalRootUrl,"ExternalUserSharingOnly");
}
catch [System.Exception]{
    Write-Host "##[error] NewPnPSite $($_.Exception.Message)" -ForegroundColor Red
    Write-Output $_
    throw 
}
