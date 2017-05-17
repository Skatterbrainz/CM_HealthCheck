#requires -RunAsAdministrator
#requires -version 4
<#
.SYNOPSIS
    Get-CM-Inventory.ps1 collects SCCM hierarchy and site server data

.DESCRIPTION
    Get-CM-Inventory.ps1 collects SCCM hierarchy and site server data
    and stores the information in multiple XML data files which are then
    processed using the Export-CM-Healthcheck.ps1 script to render
    a final MS Word report.

.PARAMETER ReportFolder
    [string] [required] Path to output data folder

.PARAMETER SmsProvider
    [string] [required] FQDN of SCCM site server

.PARAMETER NumberOfDays
    [int] [optional] Number of days to go back for alerts in logs
    default = 7

.PARAMETER HealthcheckFilename
    [string] [optional] Name of configuration file
    default is cmhealthcheck.xml

.PARAMETER Overwrite
    [switch] [optional] Overwrite existing output folder if found
    Folder is named by datestamp, so this only applies when
    running repeatedly on the same date

.PARAMETER NoHotfix
    [switch] [optional] Suppress hotfix inventory
    Can save significant runtime

.NOTES
    Version 0.1 - Raphael Perez - 24/10/2013 - Initial Script
    Version 0.2 - Raphael Perez - 05/11/2014
        - Added Get-MessageInformation and Get-MessageSolution
    Version 0.3 - Raphael Perez - 22/06/2015
        - Added ReportSection
    Version 0.4 - Raphael Perez - 04/02/2016
        - Fixed issue when executing on a Windows 10 machine
    Version 0.5 - David Stein (4/11/2017)
        - Added support for MS Word 2016
        - Changed "cm12R2healthCheck.xml" to "cmhealthcheck.xml"
        - Detailed is now a [switch] not a [boolean]
        - Added params for CoverPage, Author, CustomerName, etc.
        - Bugfixes for Word document builtin properties updates
        - Minor bugfixes throughout
    Version 0.6 - David Stein (4/17/2017)
        - Bugfixes
	Version 0.61 - David Stein (5/16/2017)
		- Added -NoHotFix option
	- Removed Test-PowerShell function dependency
	- Incremented version from 0.51 to 0.6 for consistency

    Thanks to:
    Base script (the hardest part) created by Rafael Perez (www.rflsystems.co.uk)
    Word functions copied from Carl Webster (www.carlwebster.com)
    Word functions copied from David O'Brien (www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/)

    NOTE: This script was tested on SCCM 2012 R2, 1606, 1610, and 1702 Primary sites

    Support: Database name must be CM_<SITECODE> (you need to adapt the queries if not this format)

    Security Rights: user running this tool should have the following rights:
        SQL Server (serveradmin) to be able to see database / cpu stats
        SCCM Database (db_owner) used to create/drop user-defined functions
        msdb Database (db_datareader) used to read backup information
        read-only analyst on the SCCM console
        local administrator on all computer (used to remotely connect to the registry and services)
        firewall allowing 1433 (or equivalent) to all SQL Servers (including SQL Express on Secondary Site)
        Remote WMI/DCOM firewall - http://msdn.microsoft.com/en-us/library/jj980508(v=winembedded.81).aspx
        Remote WUA - http://msdn.microsoft.com/en-us/library/windows/desktop/aa387288%28v=VS.85%29.aspx

    Comments: To get the free disk space, enable the Free Space (MB) for the Logical Disk

.EXAMPLE
    .\Get-CM-Inventory.ps1 -SmsProvider p01.contoso.com -NumberofDays 30

#>

[CmdletBinding(ConfirmImpact="Low")]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the SMS Provider computer name")] 
        [ValidateNotNullOrEmpty()]
        [string] $SmsProvider,
    [Parameter(Mandatory = $false, HelpMessage = "Number of Days for HealthCheck")] 
        [int] $NumberofDays = 7,
	[Parameter(Mandatory = $false, HelpMessage = "HealthCheck query file name")] 
        [string] $HealthcheckFilename = 'cmhealthcheck.xml',
	[Parameter(Mandatory = $false, HelpMessage = "Overwrite existing report?")] 
        [switch] $Overwrite,
	[Parameter(Mandatory=$False, HelpMessage="Skip hotfix inventory")]
		[swtich] $NoHotfix
)
$HealthcheckDebug = $True
$FormatEnumerationLimit = -1
$currentFolder = $PWD.Path
if ($currentFolder.substring($currentFolder.Length-1) -ne '\') { $currentFolder+= '\' }

$logFolder = $currentFolder + "_Logs\"
$reportFolder = $currentFolder + (Get-Date -UFormat "%Y-%m-%d") + "\" + $SmsProvider + "\"

$component = ($MyInvocation.MyCommand.Name -replace '.ps1', '')
$logfile = $logFolder + $component + ".log"
$Error.Clear()
$bLogValidation = $false

#region FUNCTIONS

function Test-Powershell64bit {
    Write-Output ([IntPtr]::size -eq 8)
}

function Set-ReplaceString {
    param (
	    [string] $Value,
	    [string] $SiteCode,
	    [int] $NumberOfDays = "",
		[string] $ServerName = "",
		[bool] $Space = $true
	)
	
	$return = $value
    $date = Get-Date
	
	if ($space) {	
		$return = $return -replace "\r\n", " " 
		$return = $return -replace "\r", " " 
		$return = $return -replace "\n", " " 
		$return = $return -replace "\s", " " 
		$return = $return -replace "\s{2}\b"," "
	}
	$return = $return -replace "@@SITECODE@@",$SiteCode
	$return = $return -replace "@@STARTMONTH@@",$date.ToString("01/MM/yyyy")
	$return = $return -replace "@@TODAYMORNING@@",$date.ToString("yyyy/MM/dd")
	$return = $return -replace "@@NUMBEROFDAYS@@",$NumberOfDays
	$return = $return -replace "@@SERVERNAME@@",$ServerName

	if ($space) {
		while (($return.IndexOf("  ") -ge 0)) { $return = $return -replace "  ", " " }
	}
	Write-Output $return
}

Function Write-Log {
    param (
        [String] $Message,
        [int] $Severity = 1,
        [string] $LogFile = '',
        [bool] $ShowMsg = $true
        
    )
    $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
    $Date  = Get-Date -Format "HH:mm:ss.fff"
    $Date2 = Get-Date -Format "MM-dd-yyyy"
    $type=1
    
    if (($logfile -ne $null) -and ($logfile -ne '')) {    
        "<![LOG[$Message]LOG]!><time=`"$date+$($TimeZoneBias.Bias)`" date=`"$date2`" component=`"$component`" context=`"`" type=`"$severity`" thread=`"`" file=`"`">" | Out-File -FilePath $logfile -Append -NoClobber -Encoding default
    }
    
    if ($showmsg -eq $true) {
        switch ($severity) {
            3 { Write-Host $Message -ForegroundColor Red }
            2 { Write-Host $Message -ForegroundColor Yellow }
            1 { Write-Host $Message }
        }
    }
}

Function Test-Folder {
    param (
        [String] $Path,
        [bool] $Create = $true
    )
    if (Test-Path -Path $Path) { return $true }
    elseif ($Create -eq $true) {
        try {
            New-Item ($Path) -Type Directory -Force | Out-Null
            Write-Output $true        	
        }
        catch {
            Write-Output $false
        }        
    }
    else { Write-Output $false }
}

Function Get-RegistryValue {
    param (
        [String] $ComputerName,
        [string] $LogFile = '' ,
        [string] $KeyName,
        [string] $KeyValue,
        [string] $AccessType = 'LocalMachine'
    )
    if ($healthcheckdebug -eq $true) { 
        Write-Log -Message "Getting registry value from $($ComputerName), $($AccessType), $($keyname), $($keyvalue)" -LogFile $logfile 
    }
    try {
        $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
        $RegKey= $Reg.OpenSubKey($keyname)
	    if ($RegKey -ne $null) { 
		    try { $return = $RegKey.GetValue($keyvalue) }
		    catch { $return = $null }
	    }
	    else { $return = $null }
        if ($healthcheckdebug -eq $true) { Write-Log -Message "Value returned $return" -LogFile $logfile }
    }
    catch {
        $return = "ERROR: Unknown"
        $Error.Clear()
    }
    Write-Output $return    
}

Function ReportSection {
    param (
	    $HealthCheckXML,
		[string] $Section,
		$SqlConn,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		$ReportTable,
		[boolean] $Detailed = $false
	)
	
	Write-Log -Message "Starting Secion $section with detailed as $($detailed.ToString())" -LogFile $logfile
	
	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
        if ($healthCheck.IsTextOnly.ToLower() -eq 'true') { continue }
        if ($healthCheck.IsActive.ToLower() -ne 'true') { continue }
		if ($healthCheck.Section.ToLower() -ne $Section) { continue }
		
		$sqlquery = $healthCheck.SqlQuery
        $tablename = (Set-ReplaceString -Value $healthCheck.XMLFile -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $servername)
        $xmlTableName = $healthCheck.XMLFile

        if ($Section -eq 5) {
            if ($detailed -eq $false) { 
                $tablename += "summary" 
                $xmlTableName += "summary"
                $gbfiels = ""
                foreach ($field in $healthCheck.Fields.Field) {
                    if ($field.groupby -in ("2")) {
                        if ($gbfiels.Length -gt 0) { $gbfiels += "," }
                        $gbfiels += $field.FieldName
                    }
                }
                $sqlquery = "select $($gbfiels), count(1) as Total from ($($sqlquery)) tbl group by $($gbfiels)"
            } 
            else { 
                $tablename += "detail"
                $xmlTableName += "detail"
                $sqlquery = $sqlquery -replace "select distinct", "select"
                $sqlquery = $sqlquery -replace "select", "select distinct"
            }
            
        }
    	$filename = $reportFolder + $tablename + '.xml'
		
		$row = $ReportTable.NewRow()
    	$row.TableName = $xmlTableName
    	$row.XMLFile = $tablename + ".xml"
    	$ReportTable.Rows.Add($row)
		
		Write-Log -Message ("$tablename Information...Starting") -LogFile $logfile
		Write-Log -Message ("Type: $($healthCheck.querytype)") -LogFile $logfile
		
		try {
			switch ($healthCheck.querytype.ToLower()) {
				'mpconnectivity' { Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mplist' | Out-Null}
				'mpcertconnectivity' { Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mpcert' | Out-Null}
				'sql' { Get-SQLData -sqlConn $sqlConn -SQLQuery $sqlquery -FileName $fileName -TableName $tablename -siteCode $siteCode -NumberOfDays $NumberOfDays -servername $servername -healthcheck $healthCheck -logfile $logfile -section $section -detailed $detailed | Out-Null}
				'baseosinfo' { Write-BaseOSInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'diskinfo' { Write-DiskInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'networkinfo' { Write-NetworkInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'rolesinstalled' { Write-RolesInstalled -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile | Out-Null}
				'servicestatus' { Write-ServiceStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'hotfixstatus' { 
                    if (-not $NoHotfix) {
                        Write-HotfixStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null
                    }
                }
           		default {}
			}
		}
		catch {
			$errorMessage = $Error[0].Exception.Message
			$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
			Write-Log -Message "The following error happen" -Severity 3 -LogFile $logfile
			Write-Log -Message "Error $errorCode : $errorMessage connecting to $servername" -LogFile $logfile -Severity 3
			$Error.Clear()
		}
		Write-Log -Message ("$tablename Information...Done") -LogFile $logfile
    }
	Write-Log -Message "End Secion $section" -LogFile $logfile
}

Function Set-FormatedValue {
    param (
	    $Value,
	    [string] $Format,
		[string] $SiteCode
	)
	switch ($format.ToLower()) {
		'schedule' {
			$schedule_Class = [wmiclass]""
			$schedule_class.psbase.path = "\\$($smsprovider)\root\sms\site_$($SiteCodeNamespace):SMS_ScheduleMethods"
			$schedule = ($schedule_class.ReadFromString($value)).TokenData
			if ($schedule.DaySpan -ne 0) { $return = ($schedule.DaySpan * 24 * 60) }
			elseif ($schedule.HourSpan -ne 0) { $return = ($schedule.HourSpan * 60) }
			elseif ($schedule.MinuteSpan -ne 0) { $return = ($schedule.MinuteSpan) }
			return $return
			break
		}
        'alertsname' {
            switch ($value.ToString().ToLower()) {
                '$databasefreespacewarning' {
                    $return = 'Low free space alert for database on site'
                    break
                }
                '$sumcompliance2updategroupdeploymentname' {
                    $return = 'Low deployment success rate alert of update group'
                    break
                }
                default {
                    $return = $value
                    break
                }
            }
            return $return
            break
        }
        'alertsseverity' {
            switch ($value.ToString().ToLower()) {
                '1' {
                    $return = 'Error'
                    break
                }
                '2' {
                    $return = 'Warning'
                    break
                }
                '3' {
                    $return = 'Informational'
                    break
                }
            }
            return $return
            break
        }
        'alertstypeid' {
            switch ($value.ToString().ToLower()) {
                '12' {
                    $return = 'Update group deployment success'
                    break
                }
                '25' {
                    $return = 'Database free space warning'
                    break
                }
                '31' {
                    $return = 'Malware detection'
                    break
                }
                default {
                    $return = $value
                    break
                }
            }
            Write-Output $return
            break
        }
		default {
			Write-Output $value
			break
		}
	}
}

Function Get-SQLData {
    param (
	    $sqlConn,
	    [string] $SQLQuery,
	    [string] $FileName,
	    [string] $TableName,
	    [string] $SiteCode,
	    [int] $NumberOfDays,
	    [string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true,
		$HealthCheck,
        $Section,
        [boolean] $Detailed = $false
	)
    try {
        $SqlCommand = $sqlConn.CreateCommand()
		$logQuery = Set-ReplaceString -value $SQLQuery -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername
		$executionquery = Set-ReplaceString -value $SQLQuery -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -space $false
		
        if ($healthcheckdebug -eq $true) { 
            Write-Log -Message ("SQL Query: $executionquery") -LogFile $logfile -ShowMsg $false
		    Write-Log -Message ("SQL Query: $logQuery") 
        } 

        $SqlCommand.CommandTimeOut = 0
        $SqlCommand.CommandText = $executionquery
        $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCommand
        $dataset = New-Object System.Data.Dataset
        $DataAdapter.Fill($dataset)
		
		if (($dataset.Tables.Count -eq 0) -or ($dataset.Tables[0].Rows.Count -eq 0)) { 
			if ($healthcheckdebug -eq $true) { Write-Log -Message ("SQL Query returned 0 records") -LogFile $logfile }
			Write-Log -Message ("Table $tablename is empty skipping writting file $filename ...") -LogFile $logfile -Severity 2 
		}
		else {
			if ($healthcheckdebug -eq $true) { Write-Log -Message ("SQL Query returned $($dataset.Tables[0].Rows.Count) records") -LogFile $logfile }
			foreach ($field in $healthCheck.Fields.Field) {
                if ($section -eq 5) {
                    if (($detailed -eq $true) -and ($field.groupby -notin ('1','2'))) { continue }
                    elseif (($detailed -eq $false) -and ($field.groupby -notin ('2','3'))) { continue }
                }
				if ($field.format -ne "") {
					foreach ($row in $dataset.Tables[0].Rows) {
						$row.$($field.FieldName) = Set-FormatedValue -value $row.$($field.FieldName) -format $field.format -sitecode $SiteCode
					}
				}
			}				
        	, $dataset.Tables[0] | Export-CliXml -Path $filename
		}
    }
    catch {
        $errorMessage = $Error[0].Exception.Message
        $errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
        if ($continueonerror -eq $false) { Write-Log -message "The following error happen, no futher action taken" -Severity 3 -LogFile $logfile }
        else { Write-Log -message "The following error happen" -Severity 3 -LogFile $logfile }
        Write-Log -message "Error $errorCode : $errorMessage connecting to $ComputerName" -LogFile $logfile -Severity 3
	    $Error.Clear()
        if ($continueonerror -eq $false) { Throw "Error $errorCode : $errorMessage connecting to $ComputerName" }
	}
}

function Create-DataTable {
    param (
	    [string] $TableName,
	    [String[]] $Fields
    )
	$DataTable = New-Object System.Data.DataTable "$tableName"
	foreach ($field in $fields) {
		$col = New-Object System.Data.DataColumn "$field",([string])
		$DataTable.Columns.Add($col)
	}
	,$DataTable
}

Function Write-BaseOSInfo {
    param (
	    [string] $FileName,
	    [string] $TableName,
	    [string] $SiteCode,
	    [int] $NumberOfDays,
	    [string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    $WMIOS = Get-RFLWmiObject -Class "win32_operatingsystem" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    if ($WMIOS -eq $null) { return }	

    $WMICS = Get-RFLWmiObject -Class "win32_computersystem" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
	$WMIProcessor = Get-RFLWmiObject -Class "Win32_processor" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    $WMITimeZone  = Get-RFLWmiObject -Class "Win32_TimeZone" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror

    ##AV Information
    $avInformation = $null
    $AVArray = @("McAfee Security@McShield", "Symantec Endpoint Protection@symantec antivirus", "Sophos Antivirus@savservice", "Avast!@aveservice", "Avast!@avast! antivirus", "Immunet Protect@immunetprotect", "F-Secure@fsma", "AntiVir@antivirservice", "Avira@avguard", "F-Protect@fpavserver", "Panda Security@pshost", "Panda AntiVirus@pavsrv", "BitDefender@bdss", "ArcaBit/ArcaVir@abmainsv", "IKARUS@ikarus-guardx", "ESET Smart Security@ekrn", "G Data Antivirus@avkproxy", "Kaspersky Lab Antivirus@klblmain", "Symantec VirusBlast@vbservprof", "ClamAV@clamav", "Vipre / GFI managed AV@SBAMSvc", "Norton@navapsvc", "Kaspersky@AVP", "Windows Defender@windefend", "Windows Defender/@MsMpSvc", "Microsoft Security Essentials@msmpeng")

    foreach ($av in $AVArray) {
        $info = $av.Split("@")
        if ((Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $info[1]).ToString().Tolower().Indexof("error") -lt 0) {
            $avInformation = $info[0]
            break
        }
    }

    $OSProcessorArch = $WMIOS.OSArchitecture
    
    if ($OSProcessorArch -ne $null) {
	    switch ($OSProcessorArch.ToUpper() ) {
		    "AMD64" {$ProcessorArchDisplay = "64-bit"}
			"i386"  {$ProcessorArchDisplay = "32-bit"}
			"IA64"  {$ProcessorArchDisplay = "64-bit - Itanium"}
			default {$ProcessorArchDisplay = $OSProcessorArch }
	    }
	} 
    else { 
        $ProcessorArchDisplay = "" 
    }
    
    $LastBootUpTime = $WMIOS.ConvertToDateTime($WMIOS.LastBootUpTime)
    $LocalDateTime  = $WMIOS.ConvertToDateTime($WMIOS.LocalDateTime)
    
    $numProcs = 0
	$ProcessorType = ""
	$ProcessorName = ""
	$ProcessorDisplayName= ""

	foreach ($WMIProc in $WMIProcessor) {
		$ProcessorType = $WMIProc.Manufacturer
		switch ($WMIProc.NumberOfCores) {
			1 {$numberOfCores = "single core"}
			2 {$numberOfCores = "dual core"}
			4 {$numberOfCores = "quad core"}
			$null {$numberOfCores = "single core"}
			default { $numberOfCores = $WMIProc.NumberOfCores.ToString() + " core" } 
		}
		
		switch ($WMIProc.Architecture) {
			0 {$CpuArchitecture = "x86"}
			1 {$CpuArchitecture = "MIPS"}
			2 {$CpuArchitecture = "Alpha"}
			3 {$CpuArchitecture = "PowerPC"}
			6 {$CpuArchitecture = "Itanium"}
			9 {$CpuArchitecture = "x64"}
		}
		
		if ($ProcessorDisplayName.Length -eq 0) { 
			$ProcessorDisplayName = " " + $numberOfCores + " $CpuArchitecture processor " + $WMIProc.Name
		}
        else {
			if ($ProcessorName -ne $WMIProc.Name) { 
				$ProcessorDisplayName += "/ " + " " + $numberOfCores + " $CpuArchitecture processor " + $WMIProc.Name
			}
		}
		$numProcs += 1
		$ProcessorName = $WMIProc.name
	}
	$ProcessorDisplayName = "$numProcs" + $ProcessorDisplayName

    if ($WMICS.DomainRole -ne $null) {
		switch ($WMICS.DomainRole) {
			0 {$RoleDisplay = "Workstation"}
			1 {$RoleDisplay = "Member Workstation"}
			2 {$RoleDisplay = "Standalone Server"}
			3 {$RoleDisplay = "Member Server"}
			4 {$RoleDisplay = "Backup Domain Controller"}
			5 {$RoleDisplay = "Primary Domain controller"}
            default: {$RoleDisplay = "unknown, $($WMICS.DomainRole)"}
		}
	}
	
	$Fields = @("ComputerName","OperatingSystem","ServicePack","Version","Architecture","LastBootTime","CurrentTime","TotalPhysicalMemory","FreePhysicalMemory","TimeZone","DaylightInEffect","Domain","Role","Model","NumberOfProcessors","NumberOfLogicalProcessors","Processors","AntiMalware")
	$BaseOSInfoTable = Create-DataTable -TableName $tableName -Fields $Fields

	$row = $BaseOSInfoTable.NewRow()
	$row.ComputerName = $ServerName
	$row.OperatingSystem = $WMIOS.Caption
	$row.ServicePack = $WMIOS.CSDVersion
	$row.Version = $WMIOS.Version
	$row.Architecture = $ProcessorArchDisplay
	$row.LastBootTime = $LastBootUpTime.ToString()
	$row.CurrentTime = $LocalDateTime.ToString()
	$row.TotalPhysicalMemory = ([string]([math]::Round($($WMIOS.TotalVisibleMemorySize/1MB), 2)) + " GB")
	$row.FreePhysicalMemory = ([string]([math]::Round($($WMIOS.FreePhysicalMemory/1MB), 2)) + " GB")
	$row.TimeZone = $WMITimeZone.Description
	$row.DaylightInEffect = $WMICS.DaylightInEffect
	$row.Domain = $WMICS.Domain
	$row.Role = $RoleDisplay
	$row.Model = $WMICS.Model
	$row.NumberOfProcessors = $WMICS.NumberOfProcessors
	$row.NumberOfLogicalProcessors = $WMICS.NumberOfLogicalProcessors
	$row.Processors = $ProcessorDisplayName
	
    if ($avInformation -ne $null) { $row.AntiMalware = $avInformation }
    else { $row.AntiMalware = "Antimalware software not detected" }
	
    $BaseOSInfoTable.Rows.Add($row)
    , $BaseOSInfoTable | Export-CliXml -Path ($filename)
}

Function Write-DiskInfo {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    $DiskList = Get-RFLWmiObject -Class "Win32_LogicalDisk" -Filter "DriveType = 3" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    if ($DiskList -eq $null) { return }
    
	$Fields=@("DeviceID","Size","FreeSpace","FileSystem")
	$DiskDetails = Create-DataTable -TableName $tableName -Fields $Fields

	foreach ($Disk in $DiskList) {
		$row = $DiskDetails.NewRow()
		$row.DeviceID = $Disk.DeviceID
		$row.Size = ([int](($Disk.Size) / 1024 / 1024 / 1024)).ToString()
		$row.FreeSpace = ([int](($Disk.FreeSpace) / 1024 / 1024 / 1024)).ToString()
		$row.FileSystem = $Disk.FileSystem
	    $DiskDetails.Rows.Add($row)
    }
    , $DiskDetails | Export-CliXml -Path ($filename)
}

function Write-NetworkInfo {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    
    $IPDetails = Get-RFLWmiObject -Class "Win32_NetworkAdapterConfiguration" -Filter "IPEnabled = true" -ComputerName $servername -logfile $logfile -continueonerror $continueonerror
    if ($IPDetails -eq $null) { return }

	$Fields = @("IPAddress","DefaultIPGateway","IPSubnet","MACAddress","DHCPEnabled")
	$NetworkInfoTable = Create-DataTable -TableName $tableName -Fields $Fields

	foreach ($IPAddress in $IPDetails) {
		$row = $NetworkInfoTable.NewRow()
		$row.IPAddress = ($IPAddress.IPAddress -join ", ")
		$row.DefaultIPGateway = ($IPAddress.DefaultIPGateway -join ", ")
		$row.IPSubnet = ($IPAddress.IPSubnet -join ", ")
		$row.MACAddress = $IPAddress.MACAddress
		if ($IPAddress.DHCPEnable -eq $true) { $row.DHCPEnabled = "TRUE" } else { $row.DHCPEnabled = "FALSE" }
	    $NetworkInfoTable.Rows.Add($row)
    }
    , $NetworkInfoTable | Export-CliXml -Path ($filename)
}

function Write-RolesInstalled {
    param (
	    [string] $FileName,
	    [string] $TableName,
	    [string] $SiteCode,
	    [int] $NumberOfDays,
	    [string] $LogfFle,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    $WMISMSListRoles = Get-RFLWMIObject -Query "select distinct RoleName from SMS_SCI_SysResUse where NetworkOSPath = '\\\\$Servername'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
    $SMSListRoles = @()
    foreach ($WMIServer in $WMISMSListRoles) { $SMSListRoles += $WMIServer.RoleName }
    $DPProperties = Get-RFLWMIObject -Query "select * from SMS_SCI_SysResUse where RoleName = 'SMS Distribution Point' and NetworkOSPath = '\\\\$Servername' and SiteCode = '$SiteCode'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
  
 	$Fields = @("SiteServer", "IIS", "SQLServer", "DP", "PXE", "MultiCast", "PreStaged", "MP", "FSP", "SSRS", "EP", "SUP", "AI", "AWS", "PWS", "SMP", "Console", "Client")
	$RolesInstalledTable = Create-DataTable -TableName $tableName -Fields $Fields
	
	$row = $RolesInstalledTable.NewRow()
	$row.SiteServer = ($SMSListRoles -contains 'SMS Site Server').ToString()
	$row.SQLServer = ($SMSListRoles -contains 'SMS SQL Server').ToString()
	$row.DP = ($SMSListRoles -contains 'SMS Distribution Point').ToString()
	if ($DPProperties -eq $null) {
		$row.PXE = "False"
		$row.MultiCast = "False"
		$row.PreStaged = "False"
	}
	else {
		$row.PXE = (($DPProperties.Props | Where-Object {$_.PropertyName -eq "IsPXE"}).Value -eq 1).ToString()
		$row.MultiCast = (($DPProperties.Props | Where-Object {$_.PropertyName -eq "IsMulticast"}).Value -eq 1).ToString()
		$row.PreStaged = (($DPProperties.Props | Where-Object {$_.PropertyName -eq "PreStagingAllowed"}).Value -eq 1).ToString()
	}
	$row.MP   = ($SMSListRoles -contains 'SMS Management Point').ToString()
	$row.FSP  = ($SMSListRoles -contains 'SMS Fallback Status Point').ToString()
	$row.SSRS = ($SMSListRoles -contains 'SMS SRS Reporting Point').ToString()
	$row.EP   = ($SMSListRoles -contains 'SMS Endpoint Protection Point').ToString()
	$row.SUP  = ($SMSListRoles -contains 'SMS Software Update Point').ToString()
	$row.AI   = ($SMSListRoles -contains 'AI Update Service Point').ToString()
	$row.AWS  = ($SMSListRoles -contains 'SMS Application Web Service').ToString()
	$row.PWS  = ($SMSListRoles -contains 'SMS Portal Web Site').ToString()
	$row.SMP  = ($SMSListRoles -contains 'SMS State Migration Point').ToString()
	$row.Console = (Test-RegistryExist -ComputerName $servername -Logfile $logfile -KeyName 'SOFTWARE\\Wow6432Node\\Microsoft\\ConfigMgr10\\AdminUI').ToString()
	$row.Client  = (Test-RegistryExist -ComputerName $servername -Logfile $logfile -KeyName 'SOFTWARE\\Microsoft\\CCM\\CCMExec').ToString()
	$row.IIS     = ((Get-RegistryValue -ComputerName $server -Logfile $logfile -KeyName 'SOFTWARE\\Microsoft\\InetStp' -KeyValue 'InstallPath') -ne $null).ToString()
    $RolesInstalledTable.Rows.Add($row)
    , $RolesInstalledTable | Export-Clixml -Path ($filename)
}

Function Get-ServiceStatus {
	param (
		$LogFile,
		[string] $ServerName,
		[string] $ServiceName
    )
	Write-Log -message "Getting service status for $servername, $servicename" -LogFile $logfile
    try {
		$service = Get-Service -ComputerName $servername | Where-Object {$_.Name -eq $servicename}
		if ($service -ne $null) { $return = $service.Status }
		else  { $return = "ERROR: Not Found" }
		Write-Log -Message "Service status $return" -LogFile $logfile
    }
    catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
    }
    Write-Output $return
}

function Write-MPConnectivity {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		$LogFile,
		[string] $Type = 'mplist'
    )
 	$Fields = @("ServerName", "HTTPReturn")
	$MPConnectivityTable = Create-DataTable -TableName $tableName -Fields $Fields

	$MPList = Get-RFLWMIObject -query "select * from SMS_SCI_SysResUse where SiteCode = '$SiteCode' and RoleName = 'SMS Management Point'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
	foreach ($MPInformation in $MPList) {
	    $SSLState = ($MPInformation.Props | Where-Object {$_.PropertyName -eq "SslState"}).Value
		$mpname = $MPInformation.NetworkOSPath -replace '\\', ''
	    if ($SSLState -eq 0) {
			$protocol = 'http'
			$port = $HTTPport 
		} 
		else {
			$protocol = 'https'
			$port = $HTTPSport 
		}
	            
		$web = New-Object -ComObject msxml2.xmlhttp
		$url = $protocol + '://' + $mpname + ':' + $port + '/sms_mp/.sms_aut?' + $type
        if ($healthcheckdebug -eq $true) { Write-Log -Message ("URL Connection: $url") -LogFile $logfile }
		$row = $MPConnectivityTable.NewRow()
		$row.ServerName = $mpname
	    try {   
			$web.open('GET', $url, $false)
			$web.send()
			$row.HTTPReturn = $web.status
	    }
	    catch {
			$row.HTTPReturn = "313 - Unable to connect to host"
			$Error.Clear()
	    }
		Write-Log -Message ("Status: $($web.status)") -LogFile $logfile
		$MPConnectivityTable.Rows.Add($row)
	}
    , $MPConnectivityTable | Export-CliXml -Path ($filename)
}

Function Write-HotfixStatus {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		$LogFile,
		[string] $ServerName,
		$ContinueOnError = $true
    )
    Write-Log -Message "Connecting to server $servername" -LogFile $logfile
    try {         
		$Session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session", $servername))
		$Searcher = $Session.CreateUpdateSearcher()
		$historyCount = $Searcher.GetTotalHistoryCount()
		$return = $Searcher.QueryHistory(0, $historyCount) 
		Write-Log -Message "Hotfix count: $HistoryCount" -logfile $logfile
    }
    catch {
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		Write-Log -Message "The following error happen" -severity 3 -logfile $logfile
		Write-Log -Message "Error $errorCode : $errorMessage connecting to $ComputerName" -logfile $logfile -severity 3
		$Error.Clear()
		return
    }

    $Fields = @("Title", "Date")
	$HotfixTable = Create-DataTable -tablename $tableName -fields $Fields
    foreach ($hotfix in $return) {
		$row = $HotfixTable.NewRow()
		$row.Title = $hotfix.Title
		$row.Date  = $hotfix.Date
		$HotfixTable.Rows.Add($row)
    }
    , $HotfixTable | Export-CliXml -Path ($filename)
}

function Write-ServiceStatus {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		$LogFile,
		[string] $ServerName,
		$ContinueOnError = $true
    )

	$SiteInformation = Get-RFLWmiObject -query "select Type from SMS_Site where ServerName = '$Server'" -namespace "Root\SMS\Site_$SiteCodeNamespace" -computerName $smsprovider -logfile $logfile
    if ($SiteInformation -ne $null) { $SiteType = $SiteInformation.Type }

    $WMISMSListRoles = Get-RFLWMIObject -query "select distinct RoleName from SMS_SCI_SysResUse where NetworkOSPath = '\\\\$Server'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
    $SMSListRoles = @()
    foreach ($WMIServer in $WMISMSListRoles) { $SMSListRoles += $WMIServer.RoleName }
	Write-Log -Message ("Roles discovered: " + $SMSListRoles -join(", ")) -LogFile $logfile
	
 	$Fields = @("ServiceName", "Status")
	$ServicesTable = Create-DataTable -TableName $tableName -Fields $Fields

    if ($SMSListRoles -contains 'AI Update Service Point') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "AI_UPDATE_SERVICE_POINT"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }

    if (($SMSListRoles -contains 'SMS Application Web Service') -or ($SMSListRoles -contains 'SMS Distribution Point') -or ($SMSListRoles -contains 'SMS Fallback Status Point') -or ($SMSListRoles -contains 'SMS Management Point') -or ($SMSListRoles -contains 'SMS Portal Web Site')  ) {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "IISADMIN"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
		
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "W3SVC"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    
    if ($SMSListRoles -contains 'SMS Component Server') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SMS_Executive"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    
    if ($SMSListRoles -contains 'SMS Site Server') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SMS_NOTIFICATION_SERVER"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)

		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SMS_SITE_COMPONENT_MANAGER"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
		
		if ($SiteType -ne 1) {
			$row = $ServicesTable.NewRow()
			$row.ServiceName = "SMS_SITE_VSS_WRITER"
			$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
		    $ServicesTable.Rows.Add($row)
		}
    }
    
    if ($SMSListRoles -contains 'SMS Software Update Point') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "WsusService"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
   
    if ($SMSListRoles -contains 'SMS SQL Server') {
		$row = $ServicesTable.NewRow()
		if ($SiteType -ne 1) {		
			$row.ServiceName = "$SQLServiceName"
		}
		else {
			$row.ServiceName = 'MSSQL$CONFIGMGRSEC'
		}
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
		$ServicesTable.Rows.Add($row)

		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SQLWriter"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    
    if ($SMSListRoles -contains 'SMS SRS Reporting Point') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "ReportServer"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    , $ServicesTable | Export-CliXml -Path ($filename)
}

function Get-RFLCredentials {
    try {
        $cred = Get-Credentials
        if ($healthcheckdebug -eq $true) { Write-Log -Message ("Trying username: $($cred.Username)") -LogFile $logfile }
        Write-Output $cred
    }
    catch {
        Write-Output $null
    }
}

function Get-RFLWmiObject {
    param (
		[String] $Class,
		[string] $Filter = '',
		[string] $Query = '',
		[String] $ComputerName,
		[String] $Namespace = "root\cimv2",
		[String] $LogFile,
		[bool] $ContinueOnError = $false
    )
    if ($query -ne '') { $class = $query }
    if ($healthcheckdebug -eq $true) { Write-Log -Message ("WMI Query: \\$ComputerName\$Namespace, $class, filter: $filter") -LogFile $logfile }

    if ($query -ne '') { $WMIObject = Get-WmiObject -Query $query -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue }
    elseif ($filter -ne '') { $WMIObject = Get-WmiObject -Class $class -Filter $filter -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue }
    else { $WMIObject = Get-WmiObject -Class $class -NameSpace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue }

	if ($WMIObject -eq $null) {
        if ($healthcheckdebug -eq $true) { Write-Log -Message ("WMI Query returned 0) records") -LogFile $logfile -Severity 2 }
	}
	else {
		$i = 1
		foreach ($obj in $wmiobj) { i++ }
		if ($healthcheckdebug -eq $true) { Write-Log -Message ("WMI Query returned $($i) records") -LogFile $logfile }
	}
	
    if ($Error.Count -ne 0) {
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		if ($continueonerror -eq $false) { Write-Log -Message "The following error happen, no futher action taken" -Severity 3 -LogFile $logfile }
		else { Write-Log -Message "The following error happen" -Severity 3 -LogFile $logfile }
		Write-Log -message "Error $errorCode : $errorMessage connecting to $ComputerName" -LogFile $logfile -Severity 3
		$Error.Clear()
		if ($continueonerror -eq $false) { Throw "Error $errorCode : $errorMessage connecting to $ComputerName" }
    }
    
    Write-Output $WMIObject
}

Function Get-SQLServerConnection {
    param (
		[string] $SQLServer,
		[string] $DBName
    )

    Try {
		$conn = New-Object System.Data.SqlClient.SqlConnection
		$conn.ConnectionString = "Data Source=$SQLServer;Initial Catalog=$DBName;Integrated Security=SSPI;"
		return $conn
    }
    Catch {
		$errorMessage = $_.Exception.Message
		$errorCode = "0x{0:X}" -f $_.Exception.ErrorCode
		Write-Log -Message "The following error happen, no futher action taken" -Severity 3 -LogFile $logfile
		Write-Log -Message "Error $errorCode : $errorMessage connecting to $ComputerName" -LogFile $logfile -Severity 3
		$Error.Clear()
		Throw "Error $errorCode : $errorMessage connecting to $SQLServer"
    }
}

Function Test-RegistryExist {
    param (
		[String] $ComputerName,
		[string] $LogFile = '' ,
		[string] $KeyName,
		[string] $AccessType = 'LocalMachine'
    )
	Write-Log -Message "Testing registry key from $($ComputerName), $($AccessType), $($KeyName)" -LogFile $logfile

    try {
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
		$RegKey = $Reg.OpenSubKey($KeyName)
		$return = ($RegKey -ne $null)
    }
    catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
    }
    Write-Output $return
}

function Test-Admin { 
	$identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent() 
	$principal = New-Object System.Security.Principal.WindowsPrincipal($identity) 
	$admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator 
	$principal.IsInRole($admin) 
} 

#endregion FUNCTIONS

try {
	if (-not (Test-Admin)) {
		Write-Host "You are not running PowerShell as Administrator (run as Administrator), , no futher action taken" -ForegroundColor Red
		Exit		
	}

	if (Test-Path -Path $reportFolder) {
		if ($overwrite -eq $true) {
            Write-Verbose "removing previous output folder $($reportFolder)..."
            Remove-Item -Path "$($reportFolder)" -Recurse -Force
        }
        else {
	        Write-Host "Folder $reportFolder already exist, no futher action taken" -ForegroundColor Red
			Exit
        }
    }

    if (Test-Folder -Path $logFolder) {
    	try {
        	New-Item ($logFolder + 'Test.log') -Type File -Force | Out-Null 
        	Remove-Item ($logFolder + 'Test.log') -Force | Out-Null 
    	}
    	catch {
        	Write-Host "Unable to read/write file on $logFolder folder, no futher action taken" -ForegroundColor Red
        	Exit    
    	}
	}
	else {
        Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
        Exit
	}
	$bLogValidation = $true

	if (!(Test-Path -Path ($currentFolder + $healthcheckfilename))) {
        Write-Host "File $($currentFolder)$($healthcheckfilename) does not exist, no futher action taken" -ForegroundColor Red
		Exit
    }
    else { [xml]$HealthCheckXML = Get-Content ($currentFolder + $healthcheckfilename) }

	if (Test-Folder -Path $reportFolder) {
    	try {
        	New-Item ($reportFolder + 'Test.log') -Type file -Force | Out-Null 
        	Remove-Item ($reportFolder + 'Test.log') -Force | Out-Null 
    	}
    	catch {
        	Write-Host "Unable to read/write file on $reportFolder folder, no futher action taken" -ForegroundColor Red
        	Exit    
    	}
	}
	else {
        Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
        Exit
	}
	
	$poshversion = $PSVersionTable.PSVersion.Major
	
    if (!(Test-Powershell64bit)) {
        Write-Log -Message "Powershell is not 64bit, no futher action taken" -Severity 3 -LogFile $logfile
        Exit
    }
 
    Write-Log -Message "==========" -LogFile $logfile -ShowMsg $false
    Write-Log -Message "Starting HealthCheck" -LogFile $logfile
    Write-Log -Message "Running Powershell version $($PSVersionTable.psversion.Major)" -LogFile $logfile
    Write-Log -Message "Running Powershell 64 bits" -LogFile $logfile
    Write-Log -Message "SMS Provider: $smsprovider" -LogFile $logfile
        
    $WMISMSProvider = Get-RFLWmiObject -Class "SMS_ProviderLocation" -NameSpace "Root\SMS" -ComputerName $smsprovider -LogFile $logfile
    $SiteCodeNamespace = $WMISMSProvider.SiteCode
	Write-Log -Message "Site Code: $SiteCodeNamespace" -LogFile $logfile
	
    $WMISMSSite = Get-RFLWmiObject -Class "SMS_Site" -NameSpace "Root\SMS\Site_$SiteCodeNamespace" -Filter "SiteCode = '$SiteCodeNamespace'" -ComputerName $smsprovider -LogFile $logfile
	$SMSSiteServer = $WMISMSSite.ServerName
	Write-Log -Message "Site Server: $($WMISMSSite.ServerName)" -LogFile $logfile
	Write-Log -Message "Site Version: $($WMISMSSite.Version)" -LogFile $logfile

	if (-not ($WMISMSSite.Version -like "5.*")) {
		Write-Log -Message "SCCM Site $($WMISMSSite.Version) not supported. No further action taken" -Severity 3 -LogFile $logfile
		Exit
	}
	
    $SQLServerName  = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Server'
    $SQLServiceName = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server' -KeyValue 'Service Name'
    $SQLPort   = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Port'
    $SQLDBName = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Database Name'
    
	# parse when finding default instance vs named instance
    if ($SQLDBName.IndexOf('\') -ge 0) {
        $SQLDBName = $SQLDBName.Split("\")[1]
    }
    
	Write-Log -Message ("SQLServerName: $SQLServerName") -LogFile $logfile
	Write-Log -Message ("SQLServiceName: $SQLServiceName") -LogFile $logfile
	Write-Log -Message ("SQLPort: $SQLPort") -LogFile $logfile
	Write-Log -Message ("SQLDBName: $SQLDBName") -LogFile $logfile

	$arrServers = @()
	$WMIServers = Get-RFLWMIObject -Query "select distinct NetworkOSPath from SMS_SCI_SysResUse where NetworkOSPath not like '%.microsoft.com' and Type in (1,2,4,8)" -ComputerName $smsprovider -NameSpace "root\sms\site_$SiteCodeNamespace" -LogFile $logfile
	foreach ($WMIServer in $WMIServers) { $arrServers += $WMIServer.NetworkOSPath -replace '\\', '' }
	Write-Log -Message ("Servers discovered: " + $arrServers -join(", ")) -LogFile $logfile

	$Fields = @("TableName", "XMLFile")
	$ReportTable = Create-DataTable -TableName $tableName -Fields $Fields

	$Fields = @("SiteServer", "SQLServer","DBName","SiteCode","NumberOfDays","HealthCheckFileName")
	$ConfigTable = Create-DataTable -TableName $tableName -Fields $Fields

	$row = $ConfigTable.NewRow()
	$row.SiteServer   = $SMSSiteServer
	$row.SQLServer    = $SQLServerName
	$row.DBName       = $SQLDBName
	$row.SiteCode     = $SiteCodeNamespace
	$row.NumberOfDays = [System.Convert]::ToInt32($NumberOfDays)
	$row.HealthCheckFileName = $HealthCheckFileName

	$ConfigTable.Rows.Add($row)
	, $ConfigTable | Export-Clixml -Path ($reportFolder + 'config.xml')

	$sqlConn = Get-SQLServerConnection -SQLServer "$SQLServerName,$SQLPort" -DBName $SQLDBName
	$sqlConn.Open()

	if ($healthcheckdebug -eq $true) { Write-Log -Message ("SQL Query: Creating Functions") -LogFile $logfile }
	$functionsSQLQuery = @"
CREATE FUNCTION [fn_CM12R2HealthCheck_ScheduleToMinutes](@Input varchar(16))
RETURNS bigint
AS
BEGIN
	if (ISNULL(@Input, '') <> '')
	begin
		declare @hex varchar(64), @flag char(3), @minute char(6), @hour char(5), @day char(5), @Cnt tinyint, @Len tinyint, @Output bigint, @Output2 bigint = 0
		
		set @hex = @Input

		SET @HEX=REPLACE (@HEX,'0','0000')
		set @hex=replace (@hex,'1','0001')
		set @hex=replace (@hex,'2','0010')
		set @hex=replace (@hex,'3','0011')
		set @hex=replace (@hex,'4','0100')
		set @hex=replace (@hex,'5','0101')
		set @hex=replace (@hex,'6','0110')
		set @hex=replace (@hex,'7','0111')
		set @hex=replace (@hex,'8','1000')
		set @hex=replace (@hex,'9','1001')
		set @hex=replace (@hex,'A','1010')
		set @hex=replace (@hex,'B','1011')
		set @hex=replace (@hex,'C','1100')
		set @hex=replace (@hex,'D','1101')
		set @hex=replace (@hex,'E','1110')
		set @hex=replace (@hex,'F','1111')
		
		select @Flag = SUBSTRING(@hex,43,3), @minute = SUBSTRING(@hex,46,6), @hour = SUBSTRING(@hex,52,5), @day = SUBSTRING(@hex,57,5)

		if (@flag = '010') --SCHED_TOKEN_RECUR_INTERVAL
		BEGIN
			set @Cnt = 1
			set @Len = LEN(@minute)
			set @Output = CAST(SUBSTRING(@minute, @Len, 1) AS bigint)
			 
			WHILE(@Cnt < @Len) BEGIN
			  SET @Output = @Output + POWER(CAST(SUBSTRING(@minute, @Len - @Cnt, 1) * 2 AS bigint), @Cnt)
			  SET @Cnt = @Cnt + 1
			END
			set @Output2 = @Output
			
			set @Cnt = 1
			set @Len = LEN(@hour)
			set @Output = CAST(SUBSTRING(@hour, @Len, 1) AS bigint)
			 
			WHILE(@Cnt < @Len) BEGIN
			  SET @Output = @Output + POWER(CAST(SUBSTRING(@hour, @Len - @Cnt, 1) * 2 AS bigint), @Cnt)
			  SET @Cnt = @Cnt + 1
			END		
			set @Output2 = @Output2 + (@Output*60)
			
			set @Cnt = 1
			set @Len = LEN(@day)
			set @Output = CAST(SUBSTRING(@day, @Len, 1) AS bigint)
			 
			WHILE(@Cnt < @Len) BEGIN
			  SET @Output = @Output + POWER(CAST(SUBSTRING(@day, @Len - @Cnt, 1) * 2 AS bigint), @Cnt)
			  SET @Cnt = @Cnt + 1
			END		
			set @Output2 = @Output2 + (@Output*60*24)
		END
		ELSE
			set @Output2 = -1
	end
	else
			set @Output2 = -2
		
	return @Output2
END
"@
	$SqlCommand = $sqlConn.CreateCommand()
	$SqlCommand.CommandTimeOut = 0
	$SqlCommand.CommandText = $functionsSQLQuery
	try {
		$SqlCommand.ExecuteNonQuery() | Out-Null
	}
	catch {
		#
	}
	$SqlCommand = $null

	$arrSites = @()
	$SqlCommand = $sqlConn.CreateCommand()

	$executionquery = "select distinct st.SiteCode, (select top 1 srl2.ServerName from v_SystemResourceList srl2 where srl2.RoleName = 'SMS Provider' and srl2.SiteCode = st.SiteCode) as ServerName from v_Site st"
	if ($healthcheckdebug -eq $true) { Write-Log -Message ("SQL Query: $executionquery") -LogFile $logfile }

	$SqlCommand.CommandTimeOut = 0
	$SqlCommand.CommandText = $executionquery

	$DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCommand
	$dataset     = New-Object System.Data.Dataset
	$DataAdapter.Fill($dataset) | Out-Null
	foreach($row in $dataset.Tables[0].Rows) { $arrSites += "$($row.SiteCode)@$($row.ServerName)" }	
	Write-Log -Message ("Sites discovered: " + $arrSites -join(", ")) -LogFile $logfile
	$SqlCommand = $null		

	##section 1 = information that needs be collected against each site
	foreach ($Site in $arrSites) {
		$arrSiteInfo = $Site.Split("@")
		$PortInformation = Get-RFLWmiObject -query "select * from SMS_SCI_Component where FileType=2 and ItemName='SMS_MP_CONTROL_MANAGER|SMS Management Point' and ItemType='Component' and SiteCode='$($arrSiteInfo[0])'" -NameSpace "Root\SMS\Site_$SiteCodeNamespace" -ComputerName $smsprovider -LogFile $logfile
		foreach ($portinfo in $PortInformation) {
			$HTTPport  = ($portinfo.Props | Where-Object {$_.PropertyName -eq "IISPortsList"}).Value1
			$HTTPSport = ($portinfo.Props | Where-Object {$_.PropertyName -eq "IISSSLPortsList"}).Value1
		}
		ReportSection -HealthCheckXML $HealthCheckXML -Section '1' -sqlConn $sqlConn -SiteCode $arrSiteInfo[0] -NumberOfDays $NumberOfDays -ServerName $arrSiteInfo[1] -ReportTable $ReportTable -LogFile $logfile 
	}
	
	##section 2 = information that needs be collected against each computer. should not be site specific. query will run only against the higher site in the hierarchy
    foreach ($server in $arrServers) { 
        ReportSection -HealthCheckXML $HealthCheckXML -Section '2' -sqlConn $sqlConn -siteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ServerName $server -ReportTable $ReportTable -LogFile $logfile 
    }
	
	##section 3 = database analisys information, running on all sql servers in the hierarchy. should not be site specific as it connects to the "master" database
    $DBServers = Get-RFLWMIObject -Query "select distinct NetworkOSPath from SMS_SCI_SysResUse where RoleName = 'SMS SQL Server'" -ComputerName $smsprovider -NameSpace "root\sms\site_$SiteCodeNamespace" -LogFile $logfile
    foreach ($DB in $DBServers) { 
		$DBServerName = $DB.NetworkOSPath.Replace('\',"") 
		Write-Log -Message ("Analysing SQLServer: $DBServerName") -LogFile $logfile
		if ($SQLServerName.ToLower() -eq $DBServerName.ToLower()) { $tmpConnection = $sqlConn }
		else {
			$tmpConnection = Get-SQLServerConnection -SQLServer "$DBServerName,$SQLPort" -DBName "master"
    		$tmpConnection.Open()
		}
		try {
	    	ReportSection -HealthCheckXML $HealthCheckXML -Section '3' -sqlConn $tmpConnection -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ServerName $DBServerName -ReportTable $ReportTable -LogFile $logfile
		}
		finally {
			if ($SQLServerName.ToLower() -ne $DBServerName.ToLower()) { $tmpConnection.Close()  }
		}
	}

    ##Section 4 = Database analysis against whole SCCM infrastructure, query will run only against top SQL Server
	ReportSection -HealthCheckXML $HealthCheckXML -Section '4' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile

    ##Section 5 = summary information against whole SCCM infrastructure. query will run only against the higher site in the hierarchy
	ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile
	
	##Section 5 = detailed information against whole SCCM infrastructure. query will run only against the higher site in the hierarchy
	ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -Detailed $true -LogFile $logfile

	##Section 6 = troubleshooting information
	ReportSection -HealthCheckXML $HealthCheckXML -Section '6' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile
}
catch {
	Write-Log -Message "Something bad happen that I don't know about" -Sverity 3 -LogFile $logfile
	Write-Log -Message "The following error happen, no futher action taken" -Severity 3 -LogFile $logfile
	$errorMessage = $Error[0].Exception.Message
	$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
	Write-Log -Message "Error $errorCode : $errorMessage" -LogFile $logfile -Severity 3
	Write-Log -Message "Full Error Message Error $($error[0].ToString())" -LogFile $logfile -Severity 3
	$Error.Clear()
}
finally {
    #close sql connection
    if ($sqlConn -ne $null) { 		
		if ($healthcheckdebug -eq $true) { Write-Log -Message ("SQL Query: Deleting Functions") -LogFile $logfile }
		$functionsSQLQuery = @"
IF OBJECT_ID (N'fn_CM12R2HealthCheck_ScheduleToMinutes', N'FN') IS NOT NULL
	DROP FUNCTION fn_CM12R2HealthCheck_ScheduleToMinutes;
"@
		$SqlCommand = $sqlConn.CreateCommand()
		$SqlCommand.CommandTimeOut = 0
		$SqlCommand.CommandText = $functionsSQLQuery
		try {
			$SqlCommand.ExecuteNonQuery() | Out-Null 
		}
		catch {
			#
		}
		$SqlCommand = $null		
		$sqlConn.Close() 
	}
	if ($ReportTable -ne $null) { , $ReportTable | Export-CliXml -Path ($reportFolder + 'report.xml') }

	if ($bLogValidation -eq $false) {
		Write-Host "Ending HealthCheck CollectData"
        Write-Host "==========" 
	}
	else {
        Write-Log -Message "Ending HealthCheck CollectData" -LogFile $logfile
        Write-Log -Message "==========" -LogFile $logfile
	}
}