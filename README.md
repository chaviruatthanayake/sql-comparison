# SQL Twin Inspector

SQL Twin Inspector is a PowerShell-based graphical tool for comparing and inspecting two Microsoft SQL Server instances — for example, a local on-prem server and an EC2 (cloud) SQL Server.
It gathers key system metrics, configuration parameters, and security details and displays them in a simple Windows UI, with exportable CSV and HTML reports.

##  Features

* Connect to two SQL Servers simultaneously (on-prem + cloud, or any two hosts)

* Supports Windows Authentication and SQL Authentication

* Collects the following details:

    * Overview: Edition, version, CPU layout, total RAM, SQL memory usage

    * CPU (Now): Real-time SQL CPU % and estimated system CPU %

    * Configuration: All sys.configurations parameters

    * Databases: State, recovery model, encryption, collation, etc.

    * Schemas: Schema-to-owner mappings

    * DB Sizes: Data/log file usage per database

    * Logins: SQL & Windows logins with policy/expiration flags

    * Auth Mode: Windows-only or Mixed Mode

* Built-in Test Connections button

* Export all tabs to CSV files + a consolidated HTML report

* 100% read-only — does not modify any server settings

##  Requirements

* Windows PowerShell 5.1 or later

* .NET Framework (included with Windows)

* Connectivity to both SQL Servers via TCP (default port 1433)

* SQL Server logins with the following minimum permissions:

``` pgsl
GRANT VIEW SERVER STATE;
GRANT VIEW SERVER SECURITY DEFINITION;
```

(For SQL Server 2022+, use GRANT VIEW SERVER PERFORMANCE STATE; instead of the first line.)

##  How to Run

1. Clone or download this repository.

2. Open PowerShell as Administrator.

3. (Optional) Allow script execution for this session:

```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

4. Run the script:

```
.\SqlTwinInspector.ps1
```

5. In the UI:

    * Server 1: your local SQL instance (.\MSSQLSERVER or localhost)

    * Server 2: your EC2 SQL host (e.g. ec2-hostname,1433)

    * Choose Auth (Windows or SQL)

    * If SQL Auth, enter the User and Password

6. Click Test Connections → ensure both show OK.

7. Click Collect → view data in tabs.

8. Click Export (CSV + HTML) → choose a folder to save the report.

##  Output

After export, you’ll find a folder like:

```
SqlTwinInspector_YYYYMMDD_HHMMSS/
│
├─ index.html
├─ Overview.csv
├─ CPU__Now_.csv
├─ Configuration.csv
├─ Databases.csv
├─ Schemas.csv
├─ DB_Sizes.csv
├─ Logins.csv
└─ Auth_Mode.csv
```

Open ```index.html``` in your browser for a clean consolidated report.

## Security Notes

* Credentials are never stored — they are entered at runtime only.

* The script performs read-only queries (SELECT only).

* Use a dedicated SQL login (report_reader) with least-privilege access.

## Example Login Setup

On each SQL Server:

```
CREATE LOGIN [report_reader] WITH PASSWORD = 'StrongPasswordHere!';
GRANT VIEW SERVER STATE TO [report_reader];
GRANT VIEW SERVER SECURITY DEFINITION TO [report_reader];
```

##  License

This project is provided “as is” for administrative and educational use.
Modify freely for your environment.