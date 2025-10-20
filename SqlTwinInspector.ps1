# SQL Server Metrics Dashboard - Enhanced Version
# Features: Better UI, Auto-refresh, Improved Excel export with sheets

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
    @{ Name='EC2 SQL Server'; Instance='EC2_public_ip,1433'; UseWindowsAuth=$false; Username='USERNAME'; Password='PASSWORD' }
)

$script:refreshTimer = $null
$script:currentForm = $null
$script:allMetrics = @{}

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
            ServerSource = $ServerConfig.Name
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

    # CPU and Memory
    Write-Host "  [DEBUG] Querying CPU and Memory..." -ForegroundColor Gray
    $cpuMemory = Exec-Query -Conn $conn -Sql "SELECT cpu_count AS LogicalCPUs, hyperthread_ratio AS HyperthreadRatio, CAST(physical_memory_kb/1024 AS BIGINT) AS PhysicalMemoryMB, CAST(committed_kb/1024 AS BIGINT) AS CommittedMemoryMB, CAST(committed_target_kb/1024 AS BIGINT) AS CommittedTargetMB FROM sys.dm_os_sys_info"
    
    if ($cpuMemory -ne $null -and $cpuMemory.GetType().Name -ne 'Object[]') {
        $cpuMemory = @($cpuMemory)
    }
    
    if ($cpuMemory.Count -eq 0 -or $cpuMemory -eq $null) {
        Write-Host "  [DEBUG] No data from sys.dm_os_sys_info, trying alternative method..." -ForegroundColor Gray
        
        try {
            $cn = New-Object System.Data.SqlClient.SqlConnection($conn)
            $cn.Open()
            
            $cm = $cn.CreateCommand()
            $cm.CommandText = "SELECT cpu_count FROM sys.dm_os_sys_info"
            $cpuCount = $cm.ExecuteScalar()
            
            $cm.CommandText = "SELECT physical_memory_kb FROM sys.dm_os_sys_info"
            $physMem = $cm.ExecuteScalar()
            
            $cn.Close()
            
            if ($cpuCount -ne $null) {
                $cpuMemory = @([PSCustomObject][ordered]@{
                    ServerSource = $ServerConfig.Name
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
            ServerSource = $ServerConfig.Name
            LogicalCPUs = "DMV not accessible"
            HyperthreadRatio = "DMV not accessible"
            PhysicalMemoryMB = "DMV not accessible"
            CommittedMemoryMB = "DMV not accessible"
            CommittedTargetMB = "DMV not accessible"
        })
    }
    
    $metrics['CPUMemory'] = $cpuMemory
    Write-Host "  [OK] CPU and Memory: $($cpuMemory.Count) rows" -ForegroundColor Green

    # Configuration
    $config = Exec-Query -Conn $conn -Sql "SELECT name, value AS CurrentValue, value_in_use AS RunningValue, CAST(minimum AS NVARCHAR(50)) AS MinValue, CAST(maximum AS NVARCHAR(50)) AS MaxValue, LEFT(description, 100) AS Description FROM sys.configurations ORDER BY name"
    foreach ($item in $config) { $item | Add-Member -NotePropertyName "ServerSource" -NotePropertyValue $ServerConfig.Name }
    $metrics['Configuration'] = $config
    Write-Host "  [OK] Configuration: $($config.Count) rows" -ForegroundColor Green

    # Databases
    $databases = Exec-Query -Conn $conn -Sql "SELECT d.name AS DatabaseName, d.database_id AS DatabaseID, d.create_date AS CreatedDate, d.compatibility_level AS CompatibilityLevel, d.state_desc AS State, d.recovery_model_desc AS RecoveryModel, CAST(ISNULL(SUM(mf.size), 0) * 8.0 / 1024 AS DECIMAL(10,2)) AS SizeMB, d.is_read_committed_snapshot_on AS ReadCommittedSnapshot, d.snapshot_isolation_state AS SnapshotIsolation FROM sys.databases d LEFT JOIN sys.master_files mf ON d.database_id = mf.database_id GROUP BY d.name, d.database_id, d.create_date, d.compatibility_level, d.state_desc, d.recovery_model_desc, d.is_read_committed_snapshot_on, d.snapshot_isolation_state ORDER BY d.name"
    foreach ($item in $databases) {
        $item | Add-Member -NotePropertyName "ServerSource" -NotePropertyValue $ServerConfig.Name
        $item.ReadCommittedSnapshot = if ($item.ReadCommittedSnapshot -eq $true) { "Read Committed Snapshot" } else { "Off" }
        $item.SnapshotIsolation = switch ($item.SnapshotIsolation) {
            0 { "Off" }
            1 { "In Transition to On" }
            2 { "On" }
            default { "Unknown" }
        }
    }
    $metrics['Databases'] = $databases
    Write-Host "  [OK] Databases: $($databases.Count) rows" -ForegroundColor Green

    # Logins
    $logins = Exec-Query -Conn $conn -Sql "SELECT name AS LoginName, type_desc AS LoginType, create_date AS CreatedDate, default_database_name AS DefaultDatabase, is_disabled AS IsDisabled FROM sys.server_principals WHERE type IN ('S', 'U', 'G') ORDER BY name"
    foreach ($item in $logins) { $item | Add-Member -NotePropertyName "ServerSource" -NotePropertyValue $ServerConfig.Name }
    $metrics['Logins'] = $logins
    Write-Host "  [OK] Logins: $($logins.Count) rows" -ForegroundColor Green

    # Schemas
    $schemas = Exec-Query -Conn $conn -Sql "SELECT s.name AS SchemaName, u.name AS Owner FROM sys.schemas s INNER JOIN sys.database_principals u ON s.principal_id = u.principal_id ORDER BY s.name"
    foreach ($item in $schemas) { $item | Add-Member -NotePropertyName "ServerSource" -NotePropertyValue $ServerConfig.Name }
    $metrics['Schemas'] = $schemas
    Write-Host "  [OK] Schemas: $($schemas.Count) rows" -ForegroundColor Green

    # Authentication Mode
    $auth = Exec-Query -Conn $conn -Sql "SELECT CASE SERVERPROPERTY('IsIntegratedSecurityOnly') WHEN 1 THEN 'Windows Authentication' ELSE 'Mixed Mode' END AS AuthenticationMode"
    if ($auth.Count -eq 0) {
        $auth = @([PSCustomObject]@{ AuthenticationMode = "Unknown"; ServerSource = $ServerConfig.Name })
    } else {
        foreach ($item in $auth) { $item | Add-Member -NotePropertyName "ServerSource" -NotePropertyValue $ServerConfig.Name }
    }
    $metrics['Authentication'] = $auth
    Write-Host "  [OK] Authentication: $($auth.Count) rows" -ForegroundColor Green

    return $metrics
}

function Update-MetricsDisplay {
    param($Form, $AllMetrics)
    
    # Update all grids in all tabs
    foreach ($tab in $Form.Controls[0].TabPages) {
        $splitContainer = $tab.Controls[0]
        
        # Update left panel (first server)
        if ($splitContainer.Panel1.Controls.Count -gt 0) {
            $panel1 = $splitContainer.Panel1.Controls[0]
            $grid1 = $panel1.Controls | Where-Object { $_ -is [System.Windows.Forms.DataGridView] } | Select-Object -First 1
            if ($grid1 -and $AllMetrics.Count -gt 0) {
                $serverName1 = ($AllMetrics.Keys | Sort-Object)[0]
                $metricKey = $tab.Text -replace ' ', ''
                $keyMap = @{
                    'ServerInfo'='ServerInfo'
                    'CPUandMemory'='CPUMemory'
                    'Configuration'='Configuration'
                    'Databases'='Databases'
                    'Logins'='Logins'
                    'Schemas'='Schemas'
                    'Authentication'='Authentication'
                }
                $metricKey = $keyMap[$metricKey]
                
                if ($AllMetrics[$serverName1] -and $AllMetrics[$serverName1][$metricKey]) {
                    $arrayList1 = New-Object System.Collections.ArrayList
                    $data1 = @($AllMetrics[$serverName1][$metricKey])
                    $arrayList1.AddRange($data1)
                    $grid1.DataSource = $arrayList1
                    $grid1.Refresh()
                }
            }
        }
        
        # Update right panel (second server)
        if ($splitContainer.Panel2.Controls.Count -gt 0) {
            $panel2 = $splitContainer.Panel2.Controls[0]
            $grid2 = $panel2.Controls | Where-Object { $_ -is [System.Windows.Forms.DataGridView] } | Select-Object -First 1
            if ($grid2 -and $AllMetrics.Count -gt 1) {
                $serverName2 = ($AllMetrics.Keys | Sort-Object)[1]
                $metricKey = $tab.Text -replace ' ', ''
                $keyMap = @{
                    'ServerInfo'='ServerInfo'
                    'CPUandMemory'='CPUMemory'
                    'Configuration'='Configuration'
                    'Databases'='Databases'
                    'Logins'='Logins'
                    'Schemas'='Schemas'
                    'Authentication'='Authentication'
                }
                $metricKey = $keyMap[$metricKey]
                
                if ($AllMetrics[$serverName2] -and $AllMetrics[$serverName2][$metricKey]) {
                    $arrayList2 = New-Object System.Collections.ArrayList
                    $data2 = @($AllMetrics[$serverName2][$metricKey])
                    $arrayList2.AddRange($data2)
                    $grid2.DataSource = $arrayList2
                    $grid2.Refresh()
                }
            }
        }
    }
    
    # Update status bar
    $statusLabel = $Form.Controls | Where-Object { $_.Name -eq 'StatusLabel' } | Select-Object -First 1
    if ($statusLabel) {
        $statusLabel.Text = "Last Updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }
}

function Refresh-Data {
    Write-Host "`n[REFRESH] Collecting fresh data..." -ForegroundColor Yellow
    $script:allMetrics = @{}
    foreach ($s in $script:servers) {
        $script:allMetrics[$s.Name] = Get-ServerMetrics -ServerConfig $s
    }
    
    if ($script:currentForm) {
        $script:currentForm.Invoke([Action]{
            Update-MetricsDisplay -Form $script:currentForm -AllMetrics $script:allMetrics
        })
    }
    Write-Host "[REFRESH] Data refresh completed" -ForegroundColor Green
}

function Show-MetricsGUI {
    param($AllMetrics)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SQL Server Metrics Dashboard - $(($AllMetrics.Keys | Sort-Object) -join ' | ') (Auto-refresh: 60s)"
    $form.Size = New-Object System.Drawing.Size(1600, 900)
    $form.StartPosition = "CenterScreen"
    $form.WindowState = "Maximized"
    $script:currentForm = $form

    # Main panel to hold everything
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = "Fill"
    $form.Controls.Add($mainPanel)

    # Tab control for metric types
    $metricTabs = New-Object System.Windows.Forms.TabControl
    $metricTabs.Dock = "Fill"
    $mainPanel.Controls.Add($metricTabs)

    $tabOrder = @(
        @{Key='ServerInfo'; Name='Server Info'},
        @{Key='CPUandMemory'; Name='CPU and Memory'},
        @{Key='Configuration'; Name='Configuration'},
        @{Key='Databases'; Name='Databases'},
        @{Key='Logins'; Name='Logins'},
        @{Key='Schemas'; Name='Schemas'},
        @{Key='Authentication'; Name='Authentication'}
    )

    foreach ($tabDef in $tabOrder) {
        $key = $tabDef.Key
        $name = $tabDef.Name
        
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $name
        $metricTabs.Controls.Add($tab)

        $splitContainer = New-Object System.Windows.Forms.SplitContainer
        $splitContainer.Dock = "Fill"
        $splitContainer.Orientation = "Vertical"
        $splitContainer.SplitterWidth = 5
        $tab.Controls.Add($splitContainer)

        $serverNames = $AllMetrics.Keys | Sort-Object

        # Left panel - First server
        if ($serverNames.Count -gt 0) {
            $serverName1 = $serverNames[0]
            $panel1 = New-Object System.Windows.Forms.Panel
            $panel1.Dock = "Fill"
            $splitContainer.Panel1.Controls.Add($panel1)

            $grid1 = New-Object System.Windows.Forms.DataGridView
            $grid1.Dock = "Fill"
            $grid1.ReadOnly = $true
            $grid1.AllowUserToAddRows = $false
            $grid1.BackgroundColor = [System.Drawing.Color]::White
            $grid1.RowHeadersVisible = $true
            $grid1.RowHeadersWidth = 30
            $grid1.ColumnHeadersVisible = $true
            $grid1.ColumnHeadersHeight = 60
            $grid1.ColumnHeadersHeightSizeMode = "DisableResizing"
            $grid1.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::DarkBlue
            $grid1.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
            $grid1.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $grid1.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
            $grid1.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5)
            $grid1.ColumnHeadersDefaultCellStyle.WrapMode = "True"
            $grid1.AutoSizeColumnsMode = "None"
            $grid1.SelectionMode = "FullRowSelect"
            $grid1.ScrollBars = "Both"
            $grid1.Add_DataBindingComplete({
                $this.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
                foreach ($col in $this.Columns) {
                    $col.MinimumWidth = 120
                    if ($col.Width -lt 120) { $col.Width = 120 }
                    if ($col.Width -gt 300) { $col.Width = 300 }
                }
            })
            $grid1.MultiSelect = $true
            $grid1.EnableHeadersVisualStyles = $false
            $grid1.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $grid1.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
            $panel1.Controls.Add($grid1)

            if ($AllMetrics[$serverName1] -and $AllMetrics[$serverName1][$key]) {
                $arrayList1 = New-Object System.Collections.ArrayList
                $data1 = @($AllMetrics[$serverName1][$key])
                $arrayList1.AddRange($data1)
                $grid1.DataSource = $arrayList1
            }
        }

        # Right panel - Second server
        if ($serverNames.Count -gt 1) {
            $serverName2 = $serverNames[1]
            $panel2 = New-Object System.Windows.Forms.Panel
            $panel2.Dock = "Fill"
            $splitContainer.Panel2.Controls.Add($panel2)

            $grid2 = New-Object System.Windows.Forms.DataGridView
            $grid2.Dock = "Fill"
            $grid2.ReadOnly = $true
            $grid2.AllowUserToAddRows = $false
            $grid2.BackgroundColor = [System.Drawing.Color]::White
            $grid2.RowHeadersVisible = $true
            $grid2.RowHeadersWidth = 30
            $grid2.ColumnHeadersVisible = $true
            $grid2.ColumnHeadersHeight = 60
            $grid2.ColumnHeadersHeightSizeMode = "DisableResizing"
            $grid2.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::DarkGreen
            $grid2.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
            $grid2.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $grid2.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
            $grid2.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5)
            $grid2.ColumnHeadersDefaultCellStyle.WrapMode = "True"
            $grid2.AutoSizeColumnsMode = "None"
            $grid2.SelectionMode = "FullRowSelect"
            $grid2.ScrollBars = "Both"
            $grid2.Add_DataBindingComplete({
                $this.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
                foreach ($col in $this.Columns) {
                    $col.MinimumWidth = 120
                    if ($col.Width -lt 120) { $col.Width = 120 }
                    if ($col.Width -gt 300) { $col.Width = 300 }
                }
            })
            $grid2.MultiSelect = $true
            $grid2.EnableHeadersVisualStyles = $false
            $grid2.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $grid2.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
            $panel2.Controls.Add($grid2)

            if ($AllMetrics[$serverName2] -and $AllMetrics[$serverName2][$key]) {
                $arrayList2 = New-Object System.Collections.ArrayList
                $data2 = @($AllMetrics[$serverName2][$key])
                $arrayList2.AddRange($data2)
                $grid2.DataSource = $arrayList2
            }
        }
        
        $splitContainer.SplitterDistance = [int]($splitContainer.Width / 2)
    }

    # Status Strip for better button and status visibility
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $form.Controls.Add($statusStrip)

    # Status label
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Name = "StatusLabel"
    $statusLabel.Text = "Last Updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $statusLabel.Spring = $true
    $statusStrip.Items.Add($statusLabel) | Out-Null

    # Export button in status strip
    $exportBtn = New-Object System.Windows.Forms.ToolStripDropDownButton
    $exportBtn.Text = "Export Excel"
    $exportBtn.DisplayStyle = "Text"
    $exportBtn.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDlg.Description = "Select folder to save Excel file"
        
        if ($folderDlg.ShowDialog() -eq "OK") {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $fileName = Join-Path $folderDlg.SelectedPath "SQLMetrics_$timestamp.xlsx"
            
            $excelData = @{}
            foreach ($serverName in $script:allMetrics.Keys | Sort-Object) {
                $m = $script:allMetrics[$serverName]
                if ($m -ne $null) {
                    foreach ($metricType in $m.Keys) {
                        $sheetName = "$serverName - $metricType"
                        if (-not $excelData.ContainsKey($sheetName)) {
                            $excelData[$sheetName] = @()
                        }
                        $excelData[$sheetName] += $m[$metricType]
                    }
                }
            }
            
            $excelData.GetEnumerator() | ForEach-Object {
                $_.Value | Export-Excel -Path $fileName -WorksheetName $_.Name -AutoSize -Append
            }
            
            if (Test-Path $fileName) {
                [System.Windows.Forms.MessageBox]::Show("Successfully exported metrics to: `n$fileName", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    })
    $statusStrip.Items.Add($exportBtn) | Out-Null

    # Refresh button in status strip
    $refreshBtn = New-Object System.Windows.Forms.ToolStripDropDownButton
    $refreshBtn.Text = "Refresh"
    $refreshBtn.DisplayStyle = "Text"
    $refreshBtn.Add_Click({
        $refreshBtn.Enabled = $false
        $refreshBtn.Text = "Refreshing..."
        Refresh-Data
        $refreshBtn.Text = "Refresh"
        $refreshBtn.Enabled = $true
    })
    $statusStrip.Items.Add($refreshBtn) | Out-Null

    # Setup auto-refresh timer (60 seconds)
    $script:refreshTimer = New-Object System.Windows.Forms.Timer
    $script:refreshTimer.Interval = 60000  # 60 seconds
    $script:refreshTimer.Add_Tick({
        Refresh-Data
    })
    $script:refreshTimer.Start()

    # Update status label function
    function Update-StatusLabel {
        $statusLabel.Text = "Last Updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }

    # Modify Refresh-Data to update status
    $oldRefreshData = ${function:Refresh-Data}.Definition
    ${function:Refresh-Data} = {
        Write-Host "`n[REFRESH] Collecting fresh data..." -ForegroundColor Yellow
        $script:allMetrics = @{}
        foreach ($s in $script:servers) {
            $script:allMetrics[$s.Name] = Get-ServerMetrics -ServerConfig $s
        }
        
        if ($script:currentForm) {
            $script:currentForm.Invoke([Action]{
                Update-MetricsDisplay -Form $script:currentForm -AllMetrics $script:allMetrics
                Update-StatusLabel
            })
        }
        Write-Host "[REFRESH] Data refresh completed" -ForegroundColor Green
    }

    # Cleanup on form close
    $form.Add_FormClosing({
        if ($script:refreshTimer) {
            $script:refreshTimer.Stop()
            $script:refreshTimer.Dispose()
        }
    })

    [System.Windows.Forms.Application]::Run($form)
}

# Initial data collection
Write-Host ""
foreach ($s in $script:servers) {
    $script:allMetrics[$s.Name] = Get-ServerMetrics -ServerConfig $s
}

Write-Host "`nLaunching GUI with auto-refresh..." -ForegroundColor Green
Write-Host ""
Show-MetricsGUI -AllMetrics $script:allMetrics