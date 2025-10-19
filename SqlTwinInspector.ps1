# SQL Server Metrics Dashboard - All Metrics in One Window
# Uses proven data conversion that works

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell -ArgumentList @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`"") -Wait
    return
}

Write-Host "=== SQL Server Metrics Dashboard ===" -ForegroundColor Cyan

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:servers = @(
    @{ Name='Local SQL Server'; Instance='.\SQLEXPRESS'; UseWindowsAuth=$true; Username=''; Password='' },
    @{ Name='EC2 SQL Server'; Instance='EC2_public_ip,1433'; UseWindowsAuth=$false; Username='usera=name'; Password='password' }
)

function Get-SqlConnectionString {
    param($ServerConfig)
    if ($ServerConfig.UseWindowsAuth) {
        "Server=$($ServerConfig.Instance);Database=master;Integrated Security=True;Connection Timeout=10;TrustServerCertificate=True;"
    } else {
        "Server=$($ServerConfig.Instance);Database=master;User Id=$($ServerConfig.Username);Password=$($ServerConfig.Password);Connection Timeout=10;TrustServerCertificate=True;"
    }
}

function Exec-Query {
    param([string]$Conn, [string]$Sql)
    try {
        $cn = New-Object System.Data.SqlClient.SqlConnection($Conn)
        $cn.Open()
        $cm = $cn.CreateCommand()
        $cm.CommandText = $Sql
        $cm.CommandTimeout = 30
        $da = New-Object System.Data.SqlClient.SqlDataAdapter($cm)
        $dt = New-Object System.Data.DataTable
        $rowCount = $da.Fill($dt)
        $cn.Close()
        
        Write-Host "  [DEBUG] Query returned $rowCount rows, $($dt.Columns.Count) columns" -ForegroundColor Gray
        
        # Convert to PSObjects - THIS METHOD WORKS
        $results = @()
        if ($rowCount -gt 0) {
            foreach ($row in $dt.Rows) {
                $obj = [ordered]@{}
                for ($i = 0; $i -lt $dt.Columns.Count; $i++) {
                    $obj[$dt.Columns[$i].ColumnName] = $row[$i]
                }
                $results += New-Object PSObject -Property $obj
            }
        }
        return $results
    } catch {
        Write-Host "  [Warning] Query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}

function Get-ServerMetrics {
    param($ServerConfig)

    $conn = Get-SqlConnectionString -ServerConfig $ServerConfig
    Write-Host "`nCollecting metrics from $($ServerConfig.Name)..." -ForegroundColor Cyan

    $test = New-Object System.Data.SqlClient.SqlConnection($conn)
    try { 
        $test.Open()
        Write-Host "  [OK] Connection successful" -ForegroundColor Green
        $test.Close() 
    } catch { 
        Write-Host "  [ERROR] Connection failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null 
    }

    $metrics = [ordered]@{}

    # Server Info
    try {
        $cn = New-Object System.Data.SqlClient.SqlConnection($conn)
        $cn.Open()
        
        $cm = $cn.CreateCommand()
        $cm.CommandText = "SELECT CAST(SERVERPROPERTY('MachineName') AS NVARCHAR(128))"
        $mach = $cm.ExecuteScalar()
        
        $cm.CommandText = "SELECT CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(128))"
        $srv = $cm.ExecuteScalar()
        
        $cm.CommandText = "SELECT CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128))"
        $ver = $cm.ExecuteScalar()
        
        $cm.CommandText = "SELECT CAST(SERVERPROPERTY('ProductLevel') AS NVARCHAR(128))"
        $lvl = $cm.ExecuteScalar()
        
        $cm.CommandText = "SELECT CAST(SERVERPROPERTY('Edition') AS NVARCHAR(128))"
        $edt = $cm.ExecuteScalar()
        
        $cm.CommandText = "SELECT CAST(ISNULL(SERVERPROPERTY('InstanceName'), 'Default') AS NVARCHAR(128))"
        $inst = $cm.ExecuteScalar()
        
        $cn.Close()
        
        $serverInfo = [PSCustomObject][ordered]@{
            MachineName = if ($mach) { $mach } else { "Unknown" }
            ServerName = if ($srv) { $srv } else { $ServerConfig.Instance }
            ProductVersion = if ($ver) { $ver } else { "Unknown" }
            ProductLevel = if ($lvl) { $lvl } else { "Unknown" }
            Edition = if ($edt) { $edt } else { "Unknown" }
            InstanceName = if ($inst) { $inst } else { "Default" }
        }
        $metrics['ServerInfo'] = @($serverInfo)
        Write-Host "  [OK] ServerInfo: 1 row" -ForegroundColor Green
    } catch {
        Write-Host "  [Warning] ServerInfo failed" -ForegroundColor Yellow
        $metrics['ServerInfo'] = @()
    }

    # CPU and Memory - with better error handling
    Write-Host "  [DEBUG] Querying CPU and Memory..." -ForegroundColor Gray
    $cpuMemory = Exec-Query -Conn $conn -Sql "SELECT cpu_count AS LogicalCPUs, hyperthread_ratio AS HyperthreadRatio, CAST(physical_memory_kb/1024 AS BIGINT) AS PhysicalMemoryMB, CAST(committed_kb/1024 AS BIGINT) AS CommittedMemoryMB, CAST(committed_target_kb/1024 AS BIGINT) AS CommittedTargetMB FROM sys.dm_os_sys_info"
    
    # Force array type
    if ($cpuMemory -ne $null -and $cpuMemory.GetType().Name -ne 'Object[]') {
        $cpuMemory = @($cpuMemory)
    }
    
    Write-Host "  [DEBUG] CPU Memory result count: $($cpuMemory.Count)" -ForegroundColor Gray
    Write-Host "  [DEBUG] CPU Memory type: $($cpuMemory.GetType().Name)" -ForegroundColor Gray
    
    if ($cpuMemory.Count -eq 0 -or $cpuMemory -eq $null) {
        Write-Host "  [DEBUG] No data from sys.dm_os_sys_info, trying alternative method..." -ForegroundColor Gray
        
        # Try alternative method using xp_msver
        try {
            $cn = New-Object System.Data.SqlClient.SqlConnection($conn)
            $cn.Open()
            
            # Get CPU count
            $cm = $cn.CreateCommand()
            $cm.CommandText = "SELECT cpu_count FROM sys.dm_os_sys_info"
            $cpuCount = $cm.ExecuteScalar()
            
            # Get memory
            $cm.CommandText = "SELECT physical_memory_kb FROM sys.dm_os_sys_info"
            $physMem = $cm.ExecuteScalar()
            
            $cn.Close()
            
            Write-Host "  [DEBUG] Direct scalar - CPU: $cpuCount, Memory: $physMem" -ForegroundColor Gray
            
            if ($cpuCount -ne $null) {
                $cpuMemory = @([PSCustomObject][ordered]@{
                    LogicalCPUs = $cpuCount
                    HyperthreadRatio = "N/A"
                    PhysicalMemoryMB = if ($physMem) { [math]::Round($physMem/1024, 2) } else { "N/A" }
                    CommittedMemoryMB = "N/A"
                    CommittedTargetMB = "N/A"
                })
            }
        } catch {
            Write-Host "  [DEBUG] Alternative method also failed: $($_.Exception.Message)" -ForegroundColor Gray
        }
    }
    
    if ($cpuMemory.Count -eq 0 -or $cpuMemory -eq $null) {
        $cpuMemory = @([PSCustomObject][ordered]@{
            LogicalCPUs = "DMV not accessible"
            HyperthreadRatio = "DMV not accessible"
            PhysicalMemoryMB = "DMV not accessible"
            CommittedMemoryMB = "DMV not accessible"
            CommittedTargetMB = "DMV not accessible"
        })
    }
    
    Write-Host "  [DEBUG] Final CPU Memory count: $($cpuMemory.Count)" -ForegroundColor Gray
    $metrics['CPUMemory'] = $cpuMemory
    Write-Host "  [OK] CPU and Memory: $($cpuMemory.Count) rows" -ForegroundColor Green

    # Configuration
    $config = Exec-Query -Conn $conn -Sql "SELECT name, value AS CurrentValue, value_in_use AS RunningValue, CAST(minimum AS NVARCHAR(50)) AS MinValue, CAST(maximum AS NVARCHAR(50)) AS MaxValue, LEFT(description, 100) AS description FROM sys.configurations ORDER BY name"
    $metrics['Configuration'] = $config
    Write-Host "  [OK] Configuration: $($config.Count) rows" -ForegroundColor Green

    # Databases
    $databases = Exec-Query -Conn $conn -Sql "SELECT d.name AS DatabaseName, d.database_id AS DatabaseID, d.create_date AS CreatedDate, d.compatibility_level AS CompatibilityLevel, d.state_desc AS State, d.recovery_model_desc AS RecoveryModel, CAST(ISNULL(SUM(mf.size), 0) * 8.0 / 1024 AS DECIMAL(10,2)) AS SizeMB FROM sys.databases d LEFT JOIN sys.master_files mf ON d.database_id = mf.database_id GROUP BY d.name, d.database_id, d.create_date, d.compatibility_level, d.state_desc, d.recovery_model_desc ORDER BY d.name"
    $metrics['Databases'] = $databases
    Write-Host "  [OK] Databases: $($databases.Count) rows" -ForegroundColor Green

    # Logins
    $logins = Exec-Query -Conn $conn -Sql "SELECT name AS LoginName, type_desc AS LoginType, create_date AS CreatedDate, default_database_name AS DefaultDatabase, is_disabled AS IsDisabled FROM sys.server_principals WHERE type IN ('S', 'U', 'G') ORDER BY name"
    $metrics['Logins'] = $logins
    Write-Host "  [OK] Logins: $($logins.Count) rows" -ForegroundColor Green

    # Schemas
    $schemas = Exec-Query -Conn $conn -Sql "SELECT s.name AS SchemaName, u.name AS Owner FROM sys.schemas s INNER JOIN sys.database_principals u ON s.principal_id = u.principal_id ORDER BY s.name"
    $metrics['Schemas'] = $schemas
    Write-Host "  [OK] Schemas: $($schemas.Count) rows" -ForegroundColor Green

    # Authentication Mode
    $auth = Exec-Query -Conn $conn -Sql "SELECT CASE SERVERPROPERTY('IsIntegratedSecurityOnly') WHEN 1 THEN 'Windows Authentication' ELSE 'Mixed Mode' END AS AuthenticationMode"
    if ($auth.Count -eq 0) {
        $auth = @([PSCustomObject]@{ AuthenticationMode = "Unknown" })
    }
    $metrics['Authentication'] = $auth
    Write-Host "  [OK] Authentication: $($auth.Count) rows" -ForegroundColor Green

    return $metrics
}

function Show-MetricsGUI {
    param($AllMetrics)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SQL Server Metrics Dashboard"
    $form.Size = New-Object System.Drawing.Size(1400, 900)
    $form.StartPosition = "CenterScreen"
    $form.WindowState = "Maximized"

    # Main tab control for servers
    $mainTabs = New-Object System.Windows.Forms.TabControl
    $mainTabs.Dock = "Fill"
    $form.Controls.Add($mainTabs)

    foreach ($serverName in $AllMetrics.Keys) {
        $serverTab = New-Object System.Windows.Forms.TabPage
        $serverTab.Text = $serverName
        $mainTabs.Controls.Add($serverTab)

        # Inner tab control for metrics
        $innerTabs = New-Object System.Windows.Forms.TabControl
        $innerTabs.Dock = "Fill"
        $serverTab.Controls.Add($innerTabs)

        $m = $AllMetrics[$serverName]
        
        if ($m -eq $null) {
            $errorLabel = New-Object System.Windows.Forms.Label
            $errorLabel.Text = "Failed to collect metrics"
            $errorLabel.Dock = "Fill"
            $errorLabel.TextAlign = "MiddleCenter"
            $errorLabel.ForeColor = "Red"
            $serverTab.Controls.Add($errorLabel)
            continue
        }

        $tabOrder = @(
            @{Key='ServerInfo'; Name='Server Info'},
            @{Key='CPUMemory'; Name='CPU and Memory'},
            @{Key='Configuration'; Name='Configuration'},
            @{Key='Databases'; Name='Databases'},
            @{Key='Logins'; Name='Logins'},
            @{Key='Schemas'; Name='Schemas'},
            @{Key='Authentication'; Name='Authentication'}
        )

        foreach ($tabDef in $tabOrder) {
            $key = $tabDef.Key
            $name = $tabDef.Name
            
            if ($m[$key] -and $m[$key].Count -gt 0) {
                $tab = New-Object System.Windows.Forms.TabPage
                $count = $m[$key].Count
                $tab.Text = "$name ($count)"
                $innerTabs.Controls.Add($tab)

                $grid = New-Object System.Windows.Forms.DataGridView
                $grid.Dock = "Fill"
                $grid.ReadOnly = $true
                $grid.AllowUserToAddRows = $false
                $grid.BackgroundColor = [System.Drawing.Color]::White
                $grid.RowHeadersVisible = $false
                $grid.AutoSizeColumnsMode = "DisplayedCells"
                $grid.SelectionMode = "FullRowSelect"
                $grid.MultiSelect = $true
                
                # THIS IS THE KEY - Direct ArrayList binding works!
                $arrayList = New-Object System.Collections.ArrayList
                $arrayList.AddRange($m[$key])
                $grid.DataSource = $arrayList
                
                $tab.Controls.Add($grid)
            }
        }
    }

    # Export button
    $exportBtn = New-Object System.Windows.Forms.Button
    $exportBtn.Text = "Export All to CSV"
    $exportBtn.Dock = "Bottom"
    $exportBtn.Height = 40
    $exportBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $exportBtn.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "CSV files (*.csv)|*.csv"
        $dlg.FileName = "SQLMetrics_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($dlg.ShowDialog() -eq "OK") {
            $allData = @()
            foreach ($serverName in $AllMetrics.Keys) {
                $m = $AllMetrics[$serverName]
                if ($m -ne $null) {
                    foreach ($metricType in $m.Keys) {
                        foreach ($item in $m[$metricType]) {
                            $obj = $item.PSObject.Copy()
                            $obj | Add-Member -NotePropertyName "ServerName" -NotePropertyValue $serverName -Force
                            $obj | Add-Member -NotePropertyName "MetricType" -NotePropertyValue $metricType -Force
                            $allData += $obj
                        }
                    }
                }
            }
            if ($allData.Count -gt 0) {
                $allData | Export-Csv -Path $dlg.FileName -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Exported $($allData.Count) rows to:`n$($dlg.FileName)", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    })
    $form.Controls.Add($exportBtn)

    [System.Windows.Forms.Application]::Run($form)
}

Write-Host ""
$allMetrics = @{}
foreach ($s in $script:servers) {
    $allMetrics[$s.Name] = Get-ServerMetrics -ServerConfig $s
}

Write-Host "`nLaunching GUI..." -ForegroundColor Green
Write-Host ""
Show-MetricsGUI -AllMetrics $allMetrics