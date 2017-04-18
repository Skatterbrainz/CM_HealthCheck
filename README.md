# cm_healthcheck
Scripts for auditing and reporting a Configuration Manager site environment.

There are two (2) primary scripts: Get-CM-Inventory.ps1 and Export-CM-HealthCheck.ps1.

## Data Collection: Get-CM-Inventory.ps1

* -SmsProvider [FQDN of top-level SCCM site server]
* -Overwrite [switch]

### Examples

* Get-CM-Inventory.ps1 -SmsProvider "cm1.contoso.com" -Verbose

## Create Report: Export-CM-HealthCheck.ps1

* -ReportFolder [path to server data files]
* -CustomerName [name]
* -AuthorName [name]
* -CompanyName [name]
* -Overwrite [switch]

### Examples

* Export-CM-HealthCheck.ps1 -ReportFolder "2017-04-17\cm1.contoso.com" -CustomerName "Contoso" -AuthorName "Frank Zappa" -CompanyName "Fubar Tech" -Overwrite -Verbose

## Generate AD Summary Report: Get-AD-Inventory.ps1

### Examples

* Get-AD-Inventory.ps1 -ReportFile "contoso.com.htm" -AdminGroups -MoreGroups "sccm_admins,sql_admins"
