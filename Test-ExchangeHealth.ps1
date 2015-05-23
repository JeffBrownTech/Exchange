<#
.SYNOPSIS
Tests the health of an Exchange Server 2013 environment
Current Version: 1.00

.DESCRIPTION
This script will test various components of an Exchange Server 2013
environment including health sets, service health (Mailbox or dual role servers only),
virtual directories, DAG health, and database copy health. Only database copies with
errors or large replay or copy queues will be displayed. Currently no test are performed
against Edge servers.

.EXAMPLE
Test-ExchangeHealth
This command will run checks against the entire Exchange environment

.NOTES
Written by Jeff Brown
Jeff@UpStartTech.com
@JeffWBrown
www.upstarttech.com

Any and all technical advice, scripts, and documentation are provided as is with no guarantee.
Always review any code and steps before applying to a production system to understand their full impact.

Version Notes
V1.00 - 5/22/2015 - Initial Version
#>


#*************************************************************************************
#******************************    Variables   ***************************************
#*************************************************************************************

[array]$ExchangeServers     = @(Get-ExchangeServer | Sort-Object Name)
[array]$ActiveDatabases     = @(Get-MailboxDatabase | Where-Object {$_.Recovery -eq $false})
[array]$DAGServers          = @((Get-DatabaseAvailabilityGroup).Servers | Sort-Object)
[array]$VirtualDirs         = @("owa","ecp","oab","rpc","ews","mapi","Microsoft-Server-ActiveSync","Autodiscover")

#*************************************************************************************
#******************************    Main Code   ***************************************
#*************************************************************************************

Write-Host "`n*** Checking Health Sets ***`n" -ForegroundColor Yellow

foreach ($ExchangeServer in $ExchangeServers)
{
    if (($ExchangeServer.IsClientAccessServer -eq $true) -or ($ExchangeServer.IsMailboxServer -eq $true))
    {
	    Write-Host "`n$ExchangeServer : " -NoNewLine
	    $health = @(Get-HealthReport -Server $ExchangeServer.Name | where {$_.AlertValue -eq "Unhealthy"})
	    $healthOutput = $health.HealthSet -join ", "
		
	    if ($health.Count -gt 0)
	    {
		    Write-Host "Not Healthy" -ForegroundColor White -BackgroundColor DarkRed
		    Write-Host "Unhealthy Components: " -NoNewLine
		    Write-Host $healthOutput -ForegroundColor Cyan
	    }
	    else
	    {
		    Write-Host "Healthy" -ForegroundColor Green
	    }
    }
} # End of Checking Health Sets

Write-Host "`n*** Checking Services ***`n" -ForegroundColor Yellow

foreach ($ExchangeServer in $ExchangeServers)
{
	# Checks Exchange services are running on server
	# Test-ServiceHealth only works for Mailbox or Mailbox/Client Access dual role in 2013
	# Client Access Servers are not currently checked for services running
	$servicesGood = $True
	$badServices = @()
				
	If ($ExchangeServer.IsMailboxServer -eq $true)
	{		
		$services = Test-ServiceHealth -Server $ExchangeServer.Name
		Write-Host $ExchangeServer" : " -NoNewLine
			
		foreach ($service in $services)
		{
			if ($service.RequiredServicesRunning -eq $false)
			{
				$servicesGood = $False
				$badServices += $service.ServicesNotRunning
			}
		}
		
		# Removes any duplicated services in array
		$badServices = $badServices | Select-Object -Unique
			
		if ($servicesGood -eq $true)
		{
			Write-Host "All Services Running" -ForegroundColor green
		}
		else
		{
			Write-Host "Services Not Running" -ForegroundColor White -BackgroundColor DarkRed
			Write-Host "Following Services are Not Running:"
			foreach ($badService in $badServices)
			{
				Write-Host $badService
			}
		}
	}
} # End of Checking Service Health

Write-host "`n*** Checking Virtual Directories ***" -ForegroundColor Yellow
foreach ($ExchangeServer in $ExchangeServers)
{
    if ($ExchangeServer.IsClientAccessServer -eq $true)
    {
        Write-Host "`n"$ExchangeServer.Name -ForegroundColor Yellow
        foreach ($VirtualDir in $VirtualDirs)
        {
            try
            {
                [string]$url = "https://" + $ExchangeServer.Fqdn + "/" + $VirtualDir + "/healthcheck.htm"
                Write-Host "`t$url : " -NoNewline
                $webCheck = Invoke-WebRequest $url -ErrorAction STOP
                if (($webCheck.StatusCode -ne "200") -or ($webCheck.StatusDescription -ne "OK"))
                {
                    Write-Host "Error" -ForegroundColor White -BackgroundColor DarkRed
                }
                else
                {
                    Write-Host "Healthy" -ForegroundColor Green
                }
            }
            catch
            {
                Write-Host "Error:" $_.Exception.Message -ForegroundColor White -BackgroundColor DarkRed
            }
        }
    }
} # End of checking virtual directory health

Write-Host "`n*** Checking DAG Health ***`n" -ForegroundColor Yellow		
foreach ($DAGServer in $DAGServers)
{
    [bool]$replHealth = $true
    [array]$badTests = @()
	Write-Host $DAGServer" : " -NoNewline
	$ReplTests = Test-ReplicationHealth -Server $DAGServer
	foreach ($ReplTest in $ReplTests)
	{
		If ($ReplTest.Result -ne "Passed")
		{
			$replHealth = $false
			$obj = New-Object PSObject
			$obj | Add-Member -MemberType NoteProperty -Name "DAG Test" -Value $ReplTest.Check
			$obj | Add-Member -MemberType NoteProperty -Name "Error" -Value $ReplTest.Error
			$badTests += $obj
		}
	}
	if ($replHealth -eq $true)
	{
		Write-Host "Replication Tests Passed" -ForegroundColor Green
	}
	else
	{
		Write-Host "Replication Tests Failed" -ForegroundColor White -BackgroundColor DarkRed
		Write-Host "Following Tests Did Not Pass:"
		$badTests | Format-Table -AutoSize			 	
	}		
}

Write-Host "`n*** Checking Database Copy Status ***`n" -ForegroundColor Yellow

foreach ($ActiveDatabase in $ActiveDatabases)
{
    [array]$copies = @(Get-MailboxDatabaseCopyStatus -Identity $ActiveDatabase.Name)
    foreach ($copy in $copies)
    {
        $badCopy = $null
        if (($copy.Status -ne "Mounted") -and ($copy.Status -ne "Healthy"))
        {
            $badCopy = [ordered]@{
                Name = $copy.Name
                Status = $copy.Status
                CopyQueue = $copy.CopyQueueLength
                ReplayQueue = $copy.ReplayQueueLength
                ContentIndex = $copy.ContentIndexState
            }
        } # End of if ($copy.Status -ne "Mounted")
        elseif ($copy.ContentIndexState -ne "Healthy")
        {
            $badCopy = [ordered]@{
                Name = $copy.Name
                Status = $copy.Status
                CopyQueue = $copy.CopyQueueLength
                ReplayQueue = $copy.ReplayQueueLength
                ContentIndex = $copy.ContentIndexState
            }
        } # End of elseif ($copy.ContentIndexState -ne "Healthy")
        elseif ($copy.CopyQueueLength -ge 100)
        {
           $badCopy = [ordered]@{
                Name = $copy.Name
                Status = $copy.Status
                CopyQueue = $copy.CopyQueueLength
                ReplayQueue = $copy.ReplayQueueLength
                ContentIndex = $copy.ContentIndexState
            }
        } # End of elseif ($copy.CopyQueueLength -ge 100)
        elseif ($copy.ReplayQueueLength -ge 100)
        {
           $badCopy = [ordered]@{
                Name = $copy.Name
                Status = $copy.Status
                CopyQueue = $copy.CopyQueueLength
                ReplayQueue = $copy.ReplayQueueLength
                ContentIndex = $copy.ContentIndexState
            }
        } # End of elseif ($copy.ReplayQueueLength -ge 100)
        
        if ($badCopy -ne $null)
        {
            $outputObj = New-Object -TypeName PSObject -Property $badCopy
            Write-Output $outputObj
        }
    } # End of foreach ($copy in $copies)
} # End of foreach ($ActiveDatabase in $ActiveDatabases)