<#
.SYNOPSIS
Set-OwaIMWebConfig - Configures OWA web.config file for IM integration

.DESCRIPTION
Gathers available Exchange Mailbox servers (2013 and 2016), finds the certificate assigned
to UM services, and checks the OWA web.config file to ensure certificate thumbprint and IMServer
are correct.

.PARAMETER IMServer
Used to populate the IMServer key in the web.config. This will be a Lync/Skype for Business
Directory or Front End pool FQDN.

.EXAMPLE
.\Set-OwaIMWebConfig.ps1 -IMServer pool1.domain.com
Checks all mailbox servers and sets the IMServer key to "pool1.domain.com"

.NOTES
Written by Jeff Brown 
Jeff@JeffBrown.tech 
@JeffWBrown 
www.jeffbrown.tech 

Any and all technical advice, scripts, and documentation are provided as is with no guarantee. 
Always review any code and steps before applying to a production system to understand their full impact. 

Requires running from an Exchange Management Shell window.
Tested on Exchange versions 2013 and 2016.

Version Notes
V1.00 - 5/22/2016 - Initial Version
#>

#*************************************************************************************
#************************       Parameters    ****************************************
#*************************************************************************************

[CmdletBinding()]
Param
(
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string]
    $IMServer
)

#*************************************************************************************
#**************************     Variables     ****************************************
#*************************************************************************************

$MailboxServers = @(Get-ExchangeServer | Where-Object {$_.IsMailboxServer -eq $true -and $_.AdminDisplayVersion -like "Version 15.*"})

#*************************************************************************************
#*************************     Module Check     **************************************
#*************************************************************************************

try
{
    Get-ExchangeCertificate -ErrorAction STOP
}
catch
{
    Write-Error "Exchange Module not loaded. Please run from an Exchange Management Shell window."
    RETURN
}

#*************************************************************************************
#**************************     Main Code     ****************************************
#*************************************************************************************

$MailboxServers | ForEach-Object {
    [bool]$IMServerNameExist = $false
    [bool]$IMCertificateThumbprintExist = $false
    
    # Save current server FQDN into variable
    $currentServer = $_.FQDN
    
    Write-Verbose -Message $currentServer

    # Gets the thumprint of the certificate assigned to the UM services
    Write-Verbose -Message "Searching $currentServer for Certificate"
    $certificateThumbprint = (Get-ExchangeCertificate -Server $currentServer | Where-Object {$_.Services -like "*UM*"}).Thumbprint | Select-Object -First 1

    # Checking to see if any valid certificates were found, if not, then exiting current loop
    if ($certificateThumbprint -eq $null)
    {
        Write-Warning -Message "Valid certificate not found on $currentServer - Skipping Server"
        CONTINUE
    }

    # Builds the web.config UNC path
    $webConfigFilePath = "\\$currentServer\c$\Program Files\Microsoft\Exchange Server\V15\ClientAccess\Owa\web.config"
    
    # Makes a backup of the current web.config file
    Write-Verbose -Message "Making backup of web.config file"
    Copy-Item -Path $webConfigFilePath -Destination "$webConfigFilePath-backup" -Force

    # Gets the content of the current web.config file
    $webConfigContent = [XML](Get-Content $webConfigFilePath)
    
    # Checks all the "add keys" in the appSettings section for any existing entries
    $webConfigContent.configuration.appSettings.add | ForEach-Object {
        if ($_.key -eq "IMCertificateThumbprint") # If the key already exists, verify it has the correct thumbprint
        {
            $IMCertificateThumbprintExist = $true
            Write-Verbose -Message "Existing IMCertificateThumbprint key found"
            if ($_.value -ne $certificateThumbprint)
            {
                Write-Verbose -Message "Existing IMCertificateThumbprint key value does not match thumbprint of a valid certificate - Updating IMCertificateThumprint"
                $node = $webConfigContent.SelectSingleNode('configuration/appSettings/add[@key="IMCertificateThumbprint"]')
                $node.Attributes['value'].Value = $certificateThumbprint
                $webConfigContent.Save($webConfigFilePath)

                $output = [PSCustomObject][ordered]@{
                    'Server' = $currentServer
                    'Action' = "Updated Key : IMCertificateThumbprint"
                    'Value' = $certificateThumbprint
                }
                Write-Output $output
            }
            else
            {
                Write-Verbose "Existing IMCertificateThumbprint key value matches a valid certificate"
            }
        }

        if ($_.key -eq "IMServerName") # If the key already exists, verify it matches the server name passed via parameter
        {
            $IMServerNameExist = $true
            Write-Verbose "Existing IMServerName key found"
            if ($_.value -ne $IMServer)
            {
                Write-Verbose "Existing IMServerName key value does not match requested server in parameter - Updating IMServerName"
                $node = $webConfigContent.SelectSingleNode('configuration/appSettings/add[@key="IMServerName"]')
                $node.Attributes['value'].Value = $IMServer
                $webConfigContent.Save($webConfigFilePath)

                $output = [PSCustomObject][ordered]@{
                    'Server' = $currentServer
                    'Action' = "Updated Key : IMServerName"
                    'Value' = $IMServer
                }
                Write-Output $output
            }
            else
            {
                Write-Verbose "Existing IMServerName matches $IMServer"
            }
        }
    }
    
    # Adding IMCertificateThumprint if it was not found in the web.config, adding it here
    if ($IMCertificateThumbprintExist -eq $false)
    {
        Write-Verbose "Adding IMCertificateThumbprint key to webconfig with value of $certificateThumbprint"
        $elementAdd= $webConfigContent.CreateElement("add")
        $key= $webConfigContent.CreateAttribute("key")
        $key.psbase.value = "IMCertificateThumbprint"
        $value= $webConfigContent.CreateAttribute("value")
        $value.psbase.value = $certificateThumbprint
        $elementAdd.SetAttributeNode($key) | Out-Null
        $elementAdd.SetAttributeNode($value) | Out-Null
        $webConfigContent.configuration.appSettings.AppendChild($elementAdd) | Out-Null
        $webConfigContent.Save($webConfigFilePath)
        
        $output = [PSCustomObject][ordered]@{
            'Server' = $currentServer
            'Action' = "Added Key : IMCertificateThumprint"
            'Value' = $certificateThumbprint
        }
        Write-Output $output
    }
    
    # If IMServerName was not found in the web.config, adding it here
    if ($IMServerNameExist -eq $false)
    {
        Write-Verbose "Adding IMServerName key to webconfig with value of $IMServer"
        $elementAdd= $webConfigContent.CreateElement("add")
        $key= $webConfigContent.CreateAttribute("key")
        $key.psbase.value = "IMServerName"
        $value= $webConfigContent.CreateAttribute("value")
        $value.psbase.value = $IMServer
        $elementAdd.SetAttributeNode($key) | Out-Null
        $elementAdd.SetAttributeNode($value) | Out-Null
        $webConfigContent.configuration.appSettings.AppendChild($elementAdd) | Out-Null
        $webConfigContent.Save($webConfigFilePath)
        
        $output = [PSCustomObject][ordered]@{
            'Server' = $currentServer
            'Action' = "Added Key : IMServerName"
            'Value' = $IMServer
        }
        Write-Output $output
    }
}