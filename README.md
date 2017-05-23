# cm_healthcheck 0.64
Scripts for auditing and reporting a Configuration Manager site environment.

There are two (2) primary scripts: Get-CM-Inventory.ps1 and Export-CM-HealthCheck.ps1.

## Get-CM-Inventory.ps1

Data collection process script

* -SmsProvider [FQDN of top-level SCCM site server]
* -Overwrite [switch]

### Examples

* Get-CM-Inventory.ps1 -SmsProvider "cm01.contoso.com" -Verbose
* Get-CM-Inventory.ps1 -SmsProvider "cm01.contoso.com" -Verbose -Overwrite -NoHotfix

--
## Export-CM-HealthCheck.ps1

Report generation script

* -ReportFolder [path to server data files]
* -CustomerName [name]
* -AuthorName [name]
* -CompanyName [name]
* -Overwrite [switch]

### Examples

* Export-CM-HealthCheck.ps1 -ReportFolder "2017-05-22\cm01.contoso.com" -CustomerName "Contoso" -AuthorName "Frank Zappa" -CompanyName "Fubar Tech" -Overwrite -Verbose

### Change Log

Version 0.1 - Raphael Perez - 24/10/2013 
- Initial Script

Version 0.2 - Raphael Perez - 05/11/2014
- Added Get-MessageInformation and Get-MessageSolution

Version 0.3 - Raphael Perez - 22/06/2015
- Added ReportSection

Version 0.4 - Raphael Perez - 04/02/2016
- Fixed issue when executing on a Windows 10 machine

Version 0.5 - David Stein (4/10/2017)
- Added support for MS Word 2016
- Changed "cm12R2healthCheck.xml" to "cmhealthcheck.xml"
- Detailed is now a [switch] not a [boolean]
- Added params for CoverPage, Author, CustomerName, etc.
- Bugfixes for Word document builtin properties updates
- Minor bugfixes throughout

Version 0.6 - David Stein (4/18/2017)
- Set table styles to be consistent

Version 0.6.1 - David Stein (4/23/2017)
- Added CmdletBinding() and some other additions

Version 0.6.2 - David Stein (5/16/2017)
- Minor formatting updates
