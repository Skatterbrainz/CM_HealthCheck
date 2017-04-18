# cm_healthcheck
Scripts for auditing and reporting a Configuration Manager site environment.

There are two (2) primary scripts: Get-CM-Inventory.ps1 and Export-CM-HealthCheck.ps1.

## Data Collection: Get-CM-Inventory.ps1

-SmsProvider <FQDN of top-level SCCM site server>
-Overwrite <switch>

## Create Report: Export-CM-HealthCheck.ps1

-ReportFolder <path to server data files>
-CustomerName <name>
-AuthorName <name>
-CompanyName <name>
-Overwrite <switch>

## Generate AD Summary Report: Get-AD-Inventory.ps1
