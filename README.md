# cm_healthcheck 0.64
Scripts for auditing and reporting a Configuration Manager site environment.

There are two (2) primary scripts: Get-CM-Inventory.ps1 and Export-CM-HealthCheck.ps1.  Get-CM-Inventory.ps1 is usually executed on the CAS or standalone primary site server.  Export-CM-HealthCheck.ps1 is usually executed on a desktop computer which has Microsoft Office Word installed (2010, 2013, or 2016 required)

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

## Errata

* Portions of the scripts in this repository are based on the work of Raphael Perez.  Authorship is indicated within each file.
