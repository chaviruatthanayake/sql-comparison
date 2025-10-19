# SQL Server Metrics Dashboard

A PowerShell-based real-time monitoring dashboard for SQL Server that collects and displays metrics from multiple SQL Server instances in a single, user-friendly interface.


## Features

 **Multi-Server Monitoring** - Monitor both local and remote SQL Server instances simultaneously  
 **Real-Time Auto-Refresh** - Automatically updates metrics every 60 seconds  
 **Comprehensive Metrics** - Displays 7 key metric categories:
- Server Information (Version, Edition, Instance)
- CPU & Memory Utilization
- Configuration Parameters
- Databases (with sizes and states)
- User Logins
- Database Schemas
- Authentication Mode

 **User-Friendly GUI** - Tabbed interface with organized data grids  
 **CSV Export** - Export all metrics to CSV for reporting and analysis  
 **No Installation Required** - Pure PowerShell script, no dependencies

---

## Requirements

### System Requirements
- **Windows OS** with PowerShell 5.1 or later
- **.NET Framework** 4.5 or later (usually pre-installed on Windows)
- **Network Access** to SQL Server instances (port 1433)

### SQL Server Requirements
- **SQL Server 2008 or later** (any edition)
- **SQL Server Authentication** or **Windows Authentication** enabled
- **User account** with appropriate permissions (see Permissions section)

---

## Installation

1. **Download the Script**
   ```powershell
   # Save the script as: SQLServerDashboard.ps1
   ```

2. **Configure Server Connections**
   
   Edit the script and update the `$script:servers` section with your server details:

   ```powershell
   $script:servers = @(
       @{ 
           Name='Local SQL Server'
           Instance='.\SQLEXPRESS'           # Your local SQL Server instance
           UseWindowsAuth=$true              # Use Windows Authentication
           Username=''
           Password=''
       },
       @{ 
           Name='EC2 SQL Server'
           Instance='public_ip,1433'    # Remote server IP/hostname
           UseWindowsAuth=$false             # Use SQL Authentication
           Username='username'          # SQL Server username
           Password='your_password'          # SQL Server password
       }
   )
   ```

---

## Permissions Setup

### Option 1: Read-Only User (Recommended for Production)

This creates a dedicated monitoring user with minimal permissions:

**Step 1: Create Login and User**

```sql
-- Connect to SQL Server with admin privileges (sa or sysadmin)

USE master;
GO

-- Create SQL Server login
CREATE LOGIN [username] WITH PASSWORD = 'YourStrongPassword123!';
GO

-- Grant server-level permissions
GRANT VIEW SERVER STATE TO [username];
GRANT VIEW ANY DEFINITION TO [username];
GRANT CONNECT SQL TO [username];
GO
```

**Step 2: Grant Database Access (Optional but Recommended)**

```sql
-- Allow user to see all databases
EXEC sp_MSforeachdb 'USE [?]; IF DB_NAME() NOT IN (''tempdb'') CREATE USER [username] FOR LOGIN [username]';
GO
```

### Option 2: Windows Authentication (For Local Servers)

```sql
-- Connect to SQL Server with admin privileges

USE master;
GO

-- Create login for Windows user (replace with your Windows username)
CREATE LOGIN [DOMAIN\Username] FROM WINDOWS;
GO

-- Grant permissions
GRANT VIEW SERVER STATE TO [DOMAIN\Username];
GRANT VIEW ANY DEFINITION TO [DOMAIN\Username];
GO
```

### Option 3: Full Admin Access (Testing Only)

 **Warning:** Only use for testing. Not recommended for production.

```sql
USE master;
GO

-- Grant sysadmin role (full access)
ALTER SERVER ROLE sysadmin ADD MEMBER [username];
GO
```

---

## Configuration

### 1. Local SQL Server (Windows Authentication)

```powershell
@{ 
    Name='Local SQL Server'
    Instance='.\SQLEXPRESS'       # Or 'localhost' or '(local)'
    UseWindowsAuth=$true
    Username=''
    Password=''
}
```

### 2. Remote SQL Server (SQL Authentication)

```powershell
@{ 
    Name='Remote SQL Server'
    Instance='sql_ip,1433'  # Server IP and port
    UseWindowsAuth=$false
    Username='username'
    Password='YourPassword123!'
}
```

### 3. Named Instance

```powershell
@{ 
    Name='Named Instance'
    Instance='ServerName\InstanceName'
    UseWindowsAuth=$true
    Username=''
    Password=''
}
```

---

## Usage

### Running the Dashboard

1. **Open PowerShell**
   ```powershell
   # Right-click PowerShell and select "Run as Administrator" (if needed)
   ```

2. **Set Execution Policy** (First time only)
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
   ```

3. **Navigate to Script Location**
   ```powershell
   cd "C:\Path\To\Script"
   ```

4. **Run the Script**
   ```powershell
   .\SqlTwinInspector.ps1
   ```

5. **Dashboard Opens Automatically**
   - The GUI will launch and display metrics from all configured servers
   - Data refreshes automatically every 60 seconds
   - Status bar shows last refresh time

### Exporting Data

1. Click **"Export All to CSV"** button at the bottom
2. Choose save location and filename
3. CSV file contains all metrics from all servers with timestamps

---

## Troubleshooting

### Issue: "Login failed for user"

**Solution:** Check credentials and grant permissions

```sql
-- Verify login exists
SELECT name, type_desc FROM sys.server_principals WHERE name = 'username';

-- Grant required permissions
GRANT VIEW SERVER STATE TO [username];
GRANT VIEW ANY DEFINITION TO [username];
GO
```

### Issue: CPU & Memory shows "DMV not accessible"

**Solution:** Grant VIEW SERVER STATE permission

```sql
USE master;
GO
GRANT VIEW SERVER STATE TO [username];
GO
```

Then restart the PowerShell script.

### Issue: "Cannot connect to server"

**Possible Causes:**
1. **Firewall blocking** - Ensure port 1433 is open
2. **SQL Server not listening on TCP/IP**
   - Open SQL Server Configuration Manager
   - Enable TCP/IP protocol
   - Restart SQL Server service
3. **Wrong server name/IP** - Verify connection string
4. **SQL Server Browser not running** - Start SQL Server Browser service for named instances

**Test Connection:**
```powershell
# Test from PowerShell
Test-NetConnection -ComputerName "ServerIP" -Port 1433
```

### Issue: Script execution is disabled

**Solution:**
```powershell
# Run this command first
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

### Issue: Authentication tab is blank

**Solution:** This is normal if the query returns empty. The Authentication Mode will show in the tab if available.

---

## Firewall Configuration

### Windows Firewall (Local Server)

```powershell
# Allow SQL Server port
New-NetFirewallRule -DisplayName "SQL Server" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow
```

### AWS EC2 Security Group (Remote Server)

1. Go to **EC2 Console** → **Security Groups**
2. Select your SQL Server security group
3. Add **Inbound Rule**:
   - Type: Custom TCP
   - Port: 1433
   - Source: Your local IP address or `0.0.0.0/0` (all IPs - use cautiously)

### Azure VM

1. Go to **Virtual Machines** → **Networking**
2. Add **Inbound Port Rule**:
   - Source: Your IP
   - Port: 1433
   - Protocol: TCP

---

## Customization

### Change Refresh Interval

Edit this line in the script:

```powershell
$timer.Interval = 60000  # Time in milliseconds (60000 = 1 minute)

# Examples:
# 30 seconds: $timer.Interval = 30000
# 5 minutes:  $timer.Interval = 300000
```

### Add More Servers

Add additional server configurations to the `$script:servers` array:

```powershell
$script:servers = @(
    @{ Name='Server1'; Instance='server1'; UseWindowsAuth=$true; Username=''; Password='' },
    @{ Name='Server2'; Instance='server2'; UseWindowsAuth=$true; Username=''; Password='' },
    @{ Name='Server3'; Instance='server3'; UseWindowsAuth=$false; Username='user'; Password='pass' }
)
```

### Disable Debug Messages

Comment out or remove lines starting with:
```powershell
Write-Host "  [DEBUG]" ...
```

---

## Metrics Explained

### Server Info
- **MachineName**: Physical server name
- **ServerName**: SQL Server instance name
- **ProductVersion**: SQL Server version number
- **ProductLevel**: Service pack level (RTM, SP1, etc.)
- **Edition**: SQL Server edition (Express, Standard, Enterprise)
- **InstanceName**: Named instance or Default

### CPU & Memory
- **LogicalCPUs**: Number of logical processors
- **HyperthreadRatio**: Hyperthreading ratio
- **PhysicalMemoryMB**: Total physical RAM in MB
- **CommittedMemoryMB**: Memory committed by SQL Server
- **CommittedTargetMB**: Target memory allocation

### Configuration
- All SQL Server configuration parameters
- Shows current value, running value, min/max values
- Includes dynamic and advanced settings

### Databases
- **DatabaseName**: Database name
- **SizeMB**: Total database size in megabytes
- **State**: ONLINE, OFFLINE, etc.
- **RecoveryModel**: FULL, SIMPLE, BULK_LOGGED
- **CompatibilityLevel**: Database compatibility level

### Logins
- All SQL Server logins (SQL and Windows)
- Login type, creation date, disabled status
- Default database

### Schemas
- Database schemas and their owners
- Useful for security auditing

### Authentication
- Shows if server uses Windows Authentication only or Mixed Mode

---

## Security Best Practices

 **Use read-only accounts** - Don't use 'sa' or sysadmin accounts  
 **Strong passwords** - Use complex passwords for SQL authentication  
 **Encrypt credentials** - Consider using PowerShell SecureString for passwords  
 **Restrict network access** - Use firewall rules to limit access  
 **Regular audits** - Review who has access to monitoring accounts  
 **Secure the script** - Store script in a protected location  

---

## Performance Considerations

- **Network Bandwidth**: Each refresh queries both servers
- **SQL Server Load**: Minimal impact, queries use DMVs and system views
- **Refresh Interval**: 60 seconds is optimal for most scenarios
- **Number of Servers**: Tested with up to 10 servers without issues

---

## Known Limitations

- Requires PowerShell on Windows (not compatible with PowerShell Core on Linux/Mac)
- DMV access requires appropriate SQL Server permissions
- Some metrics may not be available on SQL Server Express edition
- Large numbers of databases (100+) may slow refresh times

---

## Troubleshooting Commands

### Test SQL Server Connection
```powershell
# Test connection
$conn = "Server=YourServer;Database=master;User Id=report_reader;Password=YourPass;TrustServerCertificate=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection($conn)
$connection.Open()
Write-Host "Connection successful!"
$connection.Close()
```

### Check SQL Server Services
```powershell
# Check if SQL Server is running
Get-Service -Name "*SQL*" | Select-Object Name, Status, DisplayName
```

### Test Network Connectivity
```powershell
# Test if port 1433 is accessible
Test-NetConnection -ComputerName "YourServerIP" -Port 1433
```

---

## License

This script is provided as-is for monitoring and administrative purposes. Use at your own risk.

---