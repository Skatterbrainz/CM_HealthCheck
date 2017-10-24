#requires -version 4
<#
.SYNOPSIS
    Export-CM-Healthcheck.ps1 reads the output from Get-CM-Inventory.ps1 to generate a
    final report using Microsoft Word (2010, 2013, 2016)

.DESCRIPTION
    Export-CM-Healthcheck.ps1 reads the output from Get-CM-Inventory.ps1 to generate a
    final report using Microsoft Word (2010, 2013, 2016)

.PARAMETER ReportFolder
    [string] [required] Path to output data folder

.PARAMETER Detailed
    [switch] [optional]

.PARAMETER Healthcheckfilename
    [string] [optional] healthcheck configuration file name
	default   = "https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml"
	alternate = ".\cmhealthcheck.xml"

.PARAMETER Healthcheckdebug
    [boolean] [optional] Enable verbose output (or use -Verbose)

.PARAMETER CoverPage
    [string] [optional] 
    default = "Slice (Light)"

.PARAMETER CustomerName
    [string] [optional] Name of customer
    default = "Company"

.PARAMETER AuthorName
    [string] [optional] Name of report author
    default = "Author"

.PARAMETER Overwrite
    [switch] [optional] Overwrite existing report file if found

.NOTES
    Thanks to:
    Base script (the hardest part) created by Rafael Perez (www.rflsystems.co.uk)
    Word functions copied from Carl Webster (www.carlwebster.com)
    Word functions copied from David O'Brien (www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/)

.EXAMPLE
    Option 1: powershell.exe -ExecutionPolicy Bypass .\Export-CM-Healthcheck.ps1 [Parameters]
    Option 2: Open Powershell and execute .\Export-CM-Healthcheck.ps1 [Parameters]
    Option 3: .\Export-CM-HealthCheck.ps1 -ReportFolder "2017-05-17\cm1.contoso.com" -Detailed -CustomerName "ACME" -AuthorName "David Stein" -Overwrite -Verbose

#>

[CmdletBinding()]
PARAM (
    [Parameter (Mandatory = $True, HelpMessage = "Collected data folder")] 
        [ValidateNotNullOrEmpty()]
        [string] $ReportFolder,
	[Parameter (Mandatory = $False, HelpMessage = "Export full data, not only summary")] 
        [switch] $Detailed,
    [parameter (Mandatory = $False, HelpMessage = "Word Template cover page name")] 
        [string] $CoverPage = "Slice (Light)",
    [parameter (Mandatory = $False, HelpMessage = "Customer company name")] 
        [string] $CustomerName = "Customer Name",
    [parameter (Mandatory = $False, HelpMessage = "Author's full name")] 
        [string] $AuthorName = "Your Name",
	[parameter (Mandatory = $False, HelpMessage = "Footer text")]
		[string] $CopyrightName  = "Your Company Name",
	[Parameter (Mandatory = $False, HelpMessage = "HealthCheck query file name")] 
		[string] $Healthcheckfilename = 'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml',
	[Parameter (Mandatory = $False, HelpMessage = "HealthCheck messages file name")]
		[string] $MessagesFilename = 'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/Messages.xml',
	[Parameter (Mandatory = $False, HelpMessage = "Debug more?")] 
        $Healthcheckdebug = $False,
    [parameter (Mandatory = $False, HelpMessage = "Overwrite existing report file")]
        [switch] $Overwrite
)
$time1 = Get-Date -Format "hh:mm:ss"
Start-Transcript -Path ".\_logs\export-reportfile.log" -Append
$ScriptVersion  = "1710.01"
$FormatEnumerationLimit = -1
$bLogValidation = $False
$bAutoProps     = $True
$currentFolder  = $PWD.Path
$NormalFontSize = 10

if ($healthcheckdebug -eq $true) { $PSDefaultParameterValues = @{"*:Verbose"=$True}; $currentFolder = "C:\Temp\CMHealthCheck\" }
$logFolder = $currentFolder + "_Logs\"
if ($reportFolder.substring($reportFolder.length-1) -ne '\') { $reportFolder+= '\' }
$component = ($MyInvocation.MyCommand.Name -replace '.ps1', '')
$logfile = $logFolder + $component + ".log"
$Error.Clear()

Write-Verbose "Current Folder: $currentFolder"
Write-Verbose "Log Folder: $logFolder"
Write-Verbose "Log File: $logfile" 
Write-Verbose "Healthcheck Data File: $Healthcheckfilename"

#region functions

function Test-Powershell64bit {
    Write-Output ([IntPtr]::size -eq 8)
}

Function Write-Log {
    param (
        [String]$Message,
        [int]$Severity = 1,
        [string]$LogFile = '',
        [bool]$ShowMsg = $true
    )
    $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
    $Date  = Get-Date -Format "HH:mm:ss.fff"
    $Date2 = Get-Date -Format "MM-dd-yyyy"
    $type  = 1
    
    if (($logfile -ne $null) -and ($logfile -ne '')) {    
        "<![LOG[$Message]LOG]!><time=`"$date+$($TimeZoneBias.bias)`" date=`"$date2`" component=`"$component`" context=`"`" type=`"$severity`" thread=`"`" file=`"`">" | 
            Out-File -FilePath $logfile -Append -NoClobber -Encoding default
    }
    Write-Verbose "write-log: $Message"
    
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
        [String]$Path,
        [bool]$Create = $true
    )
    if (Test-Path -Path $Path) { Write-Output $true }
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

function Get-XmlUrlContent {
    param (
        [parameter(Mandatory=$True, HelpMessage="Target URL")]
        [ValidateNotNullOrEmpty()]
        [string] $Url
	)
	Write-Verbose "reading data from remote file: $Url"
    $content = ""
    try {
        $content = Invoke-WebRequest -Uri $Url -ErrorAction Stop
    }
    catch {
    }
    if ($content -ne "") {
        $lines = $content -split "`n"
        $result = ""
        for ($i = 1; $i -lt $lines.count; $i++) {
            $result += $lines[$i] + "`n"
        }
    }
    Write-Output $result
}

function Get-MessageInformation {
    param (
		$MessageID
	)
	$msg = $MessagesXML.dtsHealthCheck.Message | Where-Object {$_.MessageId -eq $MessageID}
	if ($msg -eq $null) {
        Write-Output "Unknown Message ID $MessageID" 
    }
	else { 
        Write-Output $msg.Description 
    }
}

function Get-MessageSolution {
    param (
		$MessageID
	)
	$msg = $MessagesXML.dtsHealthCheck.MessageSolution | Where-Object {$_.MessageId -eq $MessageID}
	if ($msg -eq $null)	{ 
        Write-Output "There is no known possible solution for Message ID $MessageID" 
    }
	else { 
        Write-Output $msg.Description 
    }
}

function Write-WordText {
    param (
		$WordSelection,
		[string] $Text    = "",
		[string] $Style   = "No Spacing",
		$Bold    = $false,
		$NewLine = $false,
		$NewPage = $false
	)
	
	$texttowrite = ""
	$wordselection.Style = $Style

    if ($bold) { $wordselection.Font.Bold = 1 } else { $wordselection.Font.Bold = 0 }
	$texttowrite += $text 
	$wordselection.TypeText($text)
	If ($newline) { $wordselection.TypeParagraph() }	
	If ($newpage) { $wordselection.InsertNewPage() }	
}

Function Set-WordDocumentProperty {
    param (
		$Document,
		$Name,
		$Value
	)
    Write-Verbose "info: document property [$Name] set to [$Value]"
    $document.BuiltInDocumentProperties($Name) = $Value
}

Function Write-ReportSection {
    param (
		$HealthCheckXML,
        $Section,
		$Detailed = $false,
        $Doc,
		$Selection,
        $LogFile
	)

	Write-Log -Message "Starting Section $section with detailed as $($detailed.ToString())" -LogFile $logfile

	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
		if ($healthCheck.Section.tolower() -ne $Section) { continue }
		$Description = $healthCheck.Description -replace("@@NumberOfDays@@", $NumberOfDays)
		if ($healthCheck.IsActive.tolower() -ne 'true') { continue }
        if ($healthCheck.IsTextOnly.tolower() -eq 'true') {
            if ($Section -eq 5) {
                if ($detailed -eq $false) { 
                    $Description += " - Overview" 
                } 
                else { 
                    $Description += " - Detailed"
                }            
            }
			Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
			Continue;
		}
		
		Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
        $bFound = $false
        $tableName = $healthCheck.XMLFile
        if ($Section -eq 5) {
            if (!($detailed)) { 
                $tablename += "summary" 
            } 
            else { 
                $tablename += "detail"
            }            
        }

		foreach ($rp in $ReportTable) {
			if ($rp.TableName -eq $tableName) {
                $bFound = $true
				Write-Log -Message (" - Exporting $($rp.XMLFile) ...") -LogFile $logfile
				$filename = $rp.XMLFile				
				if ($filename.IndexOf("_") -gt 0) {
					$xmltitle = $filename.Substring(0,$filename.IndexOf("_"))
					$xmltile = ($rp.TableName.Substring(0,$rp.TableName.IndexOf("_")).Replace("@","")).Tolower()
					switch ($xmltile) {
						"sitecode"   { $xmltile = "Site Code: "; break; }
						"servername" { $xmltile = "Server Name: "; break; }
					}
					switch ($healthCheck.WordStyle) {
						"Heading 1" { $newstyle = "Heading 2"; break; }
						"Heading 2" { $newstyle = "Heading 3"; break; }
						"Heading 3" { $newstyle = "Heading 4"; break; }
						default { $newstyle = $healthCheck.WordStyle; break }
					}
					
					$xmltile += $filename.Substring(0,$filename.IndexOf("_"))

					Write-WordText -WordSelection $selection -Text $xmltile -Style $newstyle -NewLine $true
				}				
				
	            if (!(Test-Path ($reportFolder + $rp.XMLFile))) {
					Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
					Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
					$selection.TypeParagraph()
				}
				else {
					Write-Verbose "importing XML file: $filename"
					$datatable = Import-CliXml -Path ($reportFolder + $filename)
					$count = 0
					$datatable | Where-Object { $count++ }
					
		            if ($count -eq 0) {
						Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
						Write-Log -Message ("Table: 0 rows") -LogFile $logfile -Severity 2
						$selection.TypeParagraph()
						continue
		            }

					switch ($healthCheck.PrintType.ToLower()) {
						"table" {
                            Write-Verbose "writing table type: table"
							$Table = $Null
					        $TableRange = $Null
					        $TableRange = $doc.Application.Selection.Range
                            $Columns = 0
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            }

							$Table = $doc.Tables.Add($TableRange, $count+1, $Columns)
							$table.Style = $TableStyle

							$i = 1;
							Write-Log -Message ("Table: $count rows and $Columns columns") -LogFile $logfile

							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }

								$Table.Cell(1, $i).Range.Font.Bold = $True
								$Table.Cell(1, $i).Range.Text = $field.Description
								$i++
	                        }
							$xRow = 2
							$records = 1
							$y=0
							foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
								$i = 1;
								foreach ($field in $HealthCheck.Fields.Field) {
                                    if ($section -eq 5) {
                                        if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                        elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                    }

									$Table.Cell($xRow, $i).Range.Font.Bold = $false
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($row.$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($row.$($field.FieldName))
											break ;
										}										
										default {
											$TextToWord = $row.$($field.FieldName);
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($xRow, $i).Range.Text = $TextToWord.ToString()
									$i++
		                        }
								$xRow++
								$records++
							}

							$selection.EndOf(15) | Out-Null
							$selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.view.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							$selection.TypeParagraph()
							break
						}
						"simpletable" {
							Write-Verbose "writing table type: simpletable"
							$Table = $Null
							$TableRange = $Null
							$TableRange = $doc.Application.Selection.Range
							$Columns = 0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            }

							$Table = $doc.Tables.Add($TableRange, $Columns, 2)
							$table.Style = $TableSimpleStyle
							$i = 1;
							Write-Log -Message ("Table: $Columns rows and 2 columns") -LogFile $logfile
							$records = 1
							$y=0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }

								if ($records -ge 500) {
									Write-Log -Message ("Exported $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
								$Table.Cell($i, 1).Range.Font.Bold = $true
								$Table.Cell($i, 1).Range.Text = $field.Description
								$Table.Cell($i, 2).Range.Font.Bold = $false

								if ($poshversion -ne 3) { 
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.Rows[0].$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.Rows[0].$($field.FieldName))
											break ;
										}											
										default {
											$TextToWord = $datatable.Rows[0].$($field.FieldName)
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
								}
								else {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.$($field.FieldName))
											break ;
										}											
										default {
											$TextToWord = $datatable.$($field.FieldName) 
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
								}
								$i++
								$records++
							}

					        $selection.EndOf(15) | Out-Null
					        $selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.View.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							$selection.TypeParagraph()
							break
							break
						}
						default {
                            Write-Verbose "writing table type: default"
							$records = 1
							$y=0
		                    foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
		                        foreach ($field in $HealthCheck.Fields.Field) {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = ($field.Description + " : " + (Get-MessageInformation -MessageID ($row.$($field.FieldName))))
											break ;
										}
										"messagesolution" {
											$TextToWord = ($field.Description + " : " + (Get-MessageSolution -MessageID ($row.$($field.FieldName))))
											break ;
										}												
										default {
											$TextToWord = ($field.Description + " : " + $row.$($field.FieldName))
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									Write-WordText -WordSelection $selection -Text ($TextToWord.ToString()) -NewLine $true
		                        }
								$selection.TypeParagraph()
								$records++
		                    }
						}
                	}
				}
			}
		}
        if ($bFound -eq $false) {
		    Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
		    Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
		    $selection.TypeParagraph()
		}
	}
}

function Write-RevisionTable (){
	Write-Verbose "inserting revision history table"
	$Table = $Null
	$TableRange = $Null
	$TableRange = $selection.Range
	$Table = $doc.Tables.Add($TableRange, 4, 4)
	$Table.Style = $TableStyle
	for ($i -eq 1; $i -le 4; $i++) {
		$Table.Cell($i, 1).Range.Font.Bold = $True
		$Table.Cell($i, 2).Range.Font.Bold = $False
	}
	$Table.Cell(1, 1).Range.Text = "Version"
	$Table.Cell(1, 2).Range.Text = "Date"
	$Table.Cell(1, 3).Range.Text = "Description"
	$Table.Cell(1, 4).Range.Text = "Author"
	$Table.Cell(2, 1).Range.Text = "1.0"
	$Table.Cell(2, 2).Range.Text = $(Get-Date).ToShortDateString()
	$Table.Cell(2, 3).Range.Text = "Initial release"
	$Table.Cell(2, 4).Range.Text = $($AuthorName -split " " | ForEach-Object {$_.Substring(0,1)}) -join ""

	$selection.EndOf(15) | Out-Null
	$selection.MoveDown() | Out-Null
	$doc.ActiveWindow.ActivePane.view.SeekView = 0
	$selection.EndKey(6, 0) | Out-Null
	$selection.TypeParagraph()
}

function Write-TableGrid {
    param (
        [parameter(Mandatory=$True)]
        [string] $Caption,
        [parameter(Mandatory=$True)]
        [int] $Rows,
        [parameter(Mandatory=$True)]
        [string[]] $ColumnHeadings
	)
	Write-Verbose "inserting custom table: $Caption"
    $Selection.TypeText($Caption)
    $Selection.Style = "Heading 1"
    $Selection.TypeParagraph()
    $Cols  = $ColumnHeadings.Length
    $Table = $doc.Tables.Add($Selection.Range, $rows, $cols)
    $Table.Style = "Grid Table 4 - Accent 1"
    for ($col = 1; $col -le $cols; $col++) {
        $Table.Cell(1, $col).Range.Text = $ColumnHeadings[$col-1]
    }
    for ($row = 1; $row -lt $rows; $row++) {
        $Table.Cell($row+1, 1).Range.Text = $row.ToString()
    }
    $Table.PreferredWidthType = 2
    $Table.PreferredWidth = 100
    $Table.Columns.First.PreferredWidthType = 2
    $Table.Columns.First.PreferredWidth = 7
    if ($Cols -gt 2) {
        $Table.Columns(2).PreferredWidthType = 2
        $Table.Columns(2).PreferredWIdth = 7
    }
    $Selection.EndOf(15) | Out-Null
	$Selection.MoveDown() | Out-Null
	$doc.ActiveWindow.ActivePane.view.SeekView = 0
	$Selection.EndKey(6, 0) | Out-Null
	$Selection.TypeParagraph()
}
#endregion

Write-Output "script version: $ScriptVersion"
$poshversion = $PSVersionTable.PSVersion.Major

if ($Healthcheckfilename.StartsWith('http')) {
	Write-Verbose "importing xml from remote URI: $healthcheckfilename"
	try {
		[xml]$HealthCheckXML = Invoke-RestMethod -Uri $Healthcheckfilename
	}
	catch {
		Write-Error "Failed to import data from Uri: $HealthcheckFilename"
		Write-Warning "If no Internet access is allowed, use -HealthcheckFilename '.\cmhealthcheck.xml'"
		break
	}
	Write-Verbose "configuration XML data loaded successfully"
}
else {
	Write-Verbose "importing Configuration xml from local file: $healthcheckfilename"
	if (!(Test-Path -Path $healthcheckfilename)) {
        Write-Warning "File $healthcheckfilename does not exist, no futher action taken"
		break
    }
    else { 
        try {
			[xml]$HealthCheckXML = Get-Content ($healthcheckfilename) 
		}
		catch {
			Write-Error "Failed to import data from local file: $HealthcheckFilename"
			break
		}
		Write-Verbose "configuration XML data loaded successfully"
    }
}

if ($MessagesFilename.StartsWith('http')) {
	Write-Verbose "importing Messages xml from remote URL: $MessagesFilename"
	try {
		[xml]$MessagesXML = Get-XmlUrlContent -Url $MessagesFilename
	}
	catch {
		Write-Error "Failed to import data from Uri: $MessagesFilename"
		Write-Warning "If no Internet access is allowed, use -MessagesFilename '.\messages.xml'"
		break
	}
	Write-Verbose "Messages XML data loaded successfully"
}
else {
	if (!(Test-Path -Path ".\Messages.xml")) {
        Write-Warning "File Messages.xml does not exist, no futher action taken"
		break
    }
    else { 
        Write-Verbose "reading messages.xml data"
        try {
			[xml]$MessagesXML = Get-Content '.\Messages.xml'
		}
		catch {
			Write-Error "Failed to import data from local file: $MessagesFilename"
			break
		}
	}
	Write-Verbose "Messages XML data loaded successfully"
}

if ($HealthCheckXML -and $MessagesXML) {
    if (Test-Folder -Path $logFolder) {
    	try {
        	New-Item ($logFolder + 'Test.log') -Type File -Force | Out-Null 
        	Remove-Item ($logFolder + 'Test.log') -Force | Out-Null 
    	}
    	catch {
        	Write-Warning "Unable to read/write file on $logFolder folder, no futher action taken"
        	Exit    
    	}
	}
	else {
        Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
        Exit
	}
	$bLogValidation = $true

	if (Test-Folder -Path $reportFolder -Create $false) {
		if (!(Test-Path -Path ($reportFolder + "config.xml"))) {
        	Write-Log -Message "File $($reportFolder)config.xml does not exist, no futher action taken" -Severity 3 -LogFile $logfile
        	Exit
		}
		else { 
            Write-Verbose "reading config.xml data"
            $ConfigTable = Import-CliXml -Path ($reportFolder + "config.xml") 
        }
		
		if ($poshversion -ne 3) { $NumberOfDays = $ConfigTable.Rows[0].NumberOfDays }
		else { $NumberOfDays = $ConfigTable.NumberOfDays }
		
		if (!(Test-Path -Path ($reportFolder + "report.xml"))) {
        	Write-Log -Message "File $($reportFolder)report.xml does not exist, no futher action taken" -Severity 3 -LogFile $logfile
        	Exit
		}
		else {
	 		$ReportTable = New-Object System.Data.DataTable 'ReportTable'
	        $ReportTable = Import-CliXml -Path ($reportFolder + "report.xml")
		}
	}
	else {
        Write-Warning "Folder: $reportFolder does not exist, no futher action taken"
        Exit
	}
	
    if (!(Test-Powershell64bit)) {
        Write-Log -Message "Powershell is not 64bit, no futher action taken" -Severity 3 -LogFile $logfile
        Exit
    }

	Write-Log -Message "==========" -LogFile $logfile -ShowMsg $false
    Write-Log -Message "Starting HealthCheck report" -LogFile $logfile
    Write-Log -Message "Script Version: $ScriptVersion" -LogFile $logfile
    Write-Log -Message "Running Powershell version $poshversion" -LogFile $logfile
    Write-Log -Message "Running Powershell 64 bits" -LogFile $logfile
    Write-Log -Message "Report Folder: $reportFolder" -LogFile $logfile
    Write-Log -Message "Detailed Report: $detailed" -LogFile $logfile
	Write-Log -Message "Number Of days: $NumberOfDays" -LogFile $logfile

	Write-Verbose "info: connecting to Microsoft Word..."
    try {
        $Word = New-Object -ComObject "Word.Application" -ErrorAction Stop
    }
    catch {
        Write-Warning "Error: This script requires Microsoft Word"
        break
    }
    $wordVersion = $Word.Version
	Write-Log -Message "Word Version: $WordVersion" -LogFile $logfile	
	Write-Verbose "info: Microsoft Word version: $WordVersion"
	if ($WordVersion -ge "16.0") {
		$TableStyle = "Grid Table 4 - Accent 1"
		$TableSimpleStyle = "Grid Table 4 - Accent 1"
	}
	elseif ($WordVersion -eq "15.0") {
		$TableStyle = "Grid Table 4 - Accent 1"
		$TableSimpleStyle = "Grid Table 4 - Accent 1"
	}
	elseif ($WordVersion -eq "14.0") {
		$TableStyle = "Medium Shading 1 - Accent 1"
		$TableSimpleStyle = "Light Grid - Accent 1"
	}
	else { 
		Write-Log -Message "This script requires Word 2010 to 2016 version, no further action taken" -Severity 3 -LogFile $logfile 
		Exit
	}

    $Word.Visible = $True
	$Doc = $Word.Documents.Add()
	$Selection = $Word.Selection
	
    Write-Verbose "info: disabling real-time spelling and grammar check"
	$Word.Options.CheckGrammarAsYouType  = $False
	$Word.Options.CheckSpellingAsYouType = $False
	$Doc.Styles("Normal").Font.Size = $NormalFontSize
	
    Write-Verbose "info: loading default building blocks template"
	$word.Templates.LoadBuildingBlocks() | Out-Null	
	$BuildingBlocks = $word.Templates | Where-Object {$_.name -eq "Built-In Building Blocks.dotx"}
	$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
	
    if ($doc -eq $null) {
        Write-Error "Failed to obtain handle to Word document"
        break
    }
    if ($bAutoProps -eq $True) {
        Write-Verbose "info: setting document properties"
        $doc.BuiltInDocumentProperties("Title")    = "System Center Configuration Manager HealthCheck"
        $doc.BuiltInDocumentProperties("Subject")  = "Prepared for $CustomerName"
	    $doc.BuiltInDocumentProperties("Author")   = $AuthorName
	    $doc.BuiltInDocumentProperties("Company")  = $CopyrightName
        $doc.BuiltInDocumentProperties("Category") = "HEALTHCHECK"
        $doc.BuiltInDocumentProperties("Keywords") = "sccm,healthcheck,systemcenter,configmgr,$CustomerName"
	}

    Write-Verbose "info: inserting document parts"
	$part.Insert($selection.Range,$True) | Out-Null
	$selection.InsertNewPage()
	
	Write-Verbose "info: inserting table of contents"
    $toc = $BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
	$toc.Insert($selection.Range,$True) | Out-Null

	$selection.InsertNewPage()

	$currentview = $doc.ActiveWindow.ActivePane.view.SeekView
	$doc.ActiveWindow.ActivePane.view.SeekView = 4
	$selection.HeaderFooter.Range.Text= "Copyright $([char]0x00A9) $((Get-Date).Year) - $CopyrightName"
	$selection.HeaderFooter.PageNumbers.Add(2) | Out-Null
	$doc.ActiveWindow.ActivePane.view.SeekView = $currentview
	$selection.EndKey(6,0) | Out-Null

    $absText = "This document provides a point-in-time inventory and analysis of the "
    $absText += "System Center Configuration Manager site environment for $CustomerName. "
	$absText += "For questions, concerns or comments, please consult the $CopyrightName "
	$absText += "architect or engineer who provided this document."
	
	Write-WordText -WordSelection $selection -Text "Abstract" -Style "Heading 1" -NewLine $true
	Write-WordText -WordSelection $selection -Text $absText -NewLine $true
	
	# insert REVISION HISTORY table
	Write-WordText -WordSelection $selection -Text "Revision History" -Style "Heading 1" -NewLine $true

	Write-RevisionTable
	
	$selection.InsertNewPage()

	Write-TableGrid -Caption "Summary of Findings" -Rows 4 -ColumnHeadings ("Item", "Explanation")
	Write-TableGrid -Caption "Summary of Recommendations" -Rows 4 -ColumnHeadings ("Item", "Severity", "Explanation")

	$selection.InsertNewPage()

    Write-ReportSection -HealthCheckXML $HealthCheckXML -section '1' -Doc $doc -Selection $selection -LogFile $logfile 
    Write-ReportSection -HealthCheckXML $HealthCheckXML -section '2' -Doc $doc -Selection $selection -LogFile $logfile 
    Write-ReportSection -HealthCheckXML $HealthCheckXML -section '3' -Doc $doc -Selection $selection -LogFile $logfile 
    Write-ReportSection -HealthCheckXML $HealthCheckXML -section '4' -Doc $doc -Selection $selection -LogFile $logfile 
    Write-ReportSection -HealthCheckXML $HealthCheckXML -section '5' -Doc $doc -Selection $selection -LogFile $logfile 

    if ($detailed -eq $true) {
        Write-ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -Detailed $true -Doc $doc -Selection $selection -LogFile $logfile 
    }

    Write-ReportSection -HealthCheckXML $HealthCheckXML -Section '6' -Doc $doc -Selection $selection -LogFile $logfile 
}
else {
	Write-Log -Message "unable to load Healthcheck or Messages XML data" -Severity 3 -LogFile $logfile
	Write-Error "failed to load configuration data from XML files"
	$error.Clear()
}
if ($toc -ne $null) {
	$doc.TablesOfContents.Item(1).Update()
	if ($bLogValidation -eq $False) {
		Write-Host "ending healthcheck report"
		Write-Host "================="
	}
	else {
		Write-Log -Message "Ending HealthCheck Export" -LogFile $logfile
		Write-Log -Message "=================" -LogFile $logfile
	}
}
$time2   = Get-Date -Format "hh:mm:ss"
$RunTime = New-TimeSpan $time1 $time2
$Difference = "{0:g}" -f $RunTime
Write-Output "completed in (HH:MM:SS) $Difference"
Stop-Transcript
