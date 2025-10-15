Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- UI helpers ----------
function New-TextBox([string]$text='', [int]$x=0, [int]$y=0, [int]$w=250) {
  $tb = New-Object System.Windows.Forms.TextBox
  $tb.Location = New-Object System.Drawing.Point($x,$y)
  $tb.Width = $w; $tb.Text = $text; return $tb
}
function New-Label([string]$text, [int]$x, [int]$y) {
  $lbl = New-Object System.Windows.Forms.Label
  $lbl.Text = $text; $lbl.AutoSize = $true
  $lbl.Location = New-Object System.Drawing.Point($x,$y); return $lbl
}
function New-Combo([string[]]$items, [int]$x, [int]$y, [int]$w=140) {
  $cb = New-Object System.Windows.Forms.ComboBox
  $cb.DropDownStyle = 'DropDownList'
  $cb.Items.AddRange($items); $cb.SelectedIndex = 0
  $cb.Location = New-Object System.Drawing.Point($x,$y); $cb.Width = $w; return $cb
}
function New-Grid() {
  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = 'Fill'; $grid.ReadOnly = $true
  $grid.AutoSizeColumnsMode = 'AllCells'
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.RowHeadersVisible = $false; return $grid
}
function New-Tab([string]$name) {
  $tab = New-Object System.Windows.Forms.TabPage
  $tab.Text = $name; return $tab
}

# ---------- SQL helpers ----------
function Invoke-SqlQuery {
  param(
    [Parameter(Mandatory=$true)][string]$Server,
    [string]$Database = 'master',
    [Parameter(Mandatory=$true)][string]$AuthMode,   # 'Windows' or 'SQL'
    [string]$Username,
    [string]$Password,
    [Parameter(Mandatory=$true)][string]$Query
  )
  try {
    if ($AuthMode -eq 'Windows') {
      $connString = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    } else {
      $connString = "Server=$Server;Database=$Database;User ID=$Username;Password=$Password;TrustServerCertificate=True;"
    }
    $table = New-Object System.Data.DataTable
    $conn  = New-Object System.Data.SqlClient.SqlConnection $connString
    $cmd   = $conn.CreateCommand(); $cmd.CommandText = $Query; $cmd.CommandTimeout = 60
    $conn.Open(); $r = $cmd.ExecuteReader(); $table.Load($r); $r.Close(); $conn.Close()
    return $table
  } catch {
    $err = New-Object System.Data.DataTable
    $err.Columns.Add("ERROR") | Out-Null
    $row = $err.NewRow(); $row.ERROR = $_.Exception.Message
    $err.Rows.Add($row) | Out-Null; return $err
  }
}

function Get-Combined {
  param([string]$Query, [hashtable]$S1, [hashtable]$S2)
  $t1 = Invoke-SqlQuery -Server $S1.Server -AuthMode $S1.Auth -Username $S1.User -Password $S1.Pass -Query $Query
  if (-not $t1.Columns.Contains('Server')) { [void]$t1.Columns.Add('Server') }
  foreach ($r in $t1.Rows) { $r['Server'] = $S1.Server }

  $t2 = Invoke-SqlQuery -Server $S2.Server -AuthMode $S2.Auth -Username $S2.User -Password $S2.Pass -Query $Query
  if (-not $t2.Columns.Contains('Server')) { [void]$t2.Columns.Add('Server') }
  foreach ($r in $t2.Rows) { $r['Server'] = $S2.Server }

  $combined = $t1.Clone()
  foreach ($col in $t2.Columns) { if (-not $combined.Columns.Contains($col.ColumnName)) { [void]$combined.Columns.Add($col.ColumnName) } }
  foreach ($r in $t1.Rows) { [void]$combined.ImportRow($r) }
  foreach ($r in $t2.Rows) { [void]$combined.ImportRow($r) }
  return $combined
}

function Test-SqlConnection { param([hashtable]$S)
  try { $null = Invoke-SqlQuery -Server $S.Server -AuthMode $S.Auth -Username $S.User -Password $S.Pass -Query "SELECT @@SERVERNAME AS ServerName;"; return $true } catch { return $false }
}

# ---------- Export helpers ----------
function Get-DataTableFromGrid([System.Windows.Forms.DataGridView]$grid) {
  # DataSource can be DataTable or DataView
  if ($null -eq $grid.DataSource) { return $null }
  if ($grid.DataSource -is [System.Data.DataView]) { return $grid.DataSource.ToTable() }
  return [System.Data.DataTable]$grid.DataSource
}

function Export-TableCsv([System.Data.DataTable]$dt, [string]$path) {
  if ($null -eq $dt) { return }
  $dt | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
}

function Export-ReportHtml([hashtable]$tables, [string]$path) {
  $css = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
h1 { margin-bottom: 0; }
small { color: #666; }
h2 { border-bottom: 1px solid #ddd; padding-bottom: 4px; margin-top: 28px; }
table { border-collapse: collapse; width: 100%; margin-top: 10px; }
th, td { border: 1px solid #e5e5e5; padding: 6px 8px; text-align: left; font-size: 12px; }
th { background: #f7f7f7; }
</style>
"@
  $parts = @()
  $parts += "<!DOCTYPE html><html><head><meta charset='utf-8'><title>SQL Twin Inspector Report</title>$css</head><body>"
  $parts += "<h1>SQL Twin Inspector Report</h1><small>Generated: $(Get-Date)</small>"

  foreach ($k in $tables.Keys) {
    $dt = $tables[$k]
    if ($null -eq $dt -or $dt.Rows.Count -eq 0) { continue }
    $html = $dt | ConvertTo-Html -Fragment
    $parts += "<h2>$k</h2>$html"
  }
  $parts += "</body></html>"
  [IO.File]::WriteAllText($path, ($parts -join "`n"), [Text.UTF8Encoding]::new($false))
}

# ---------- T-SQL queries ----------
$Queries = @{
  Overview = @"
SELECT
  @@SERVERNAME AS ServerName,
  CAST(SERVERPROPERTY('MachineName') AS nvarchar(128)) AS MachineName,
  CAST(SERVERPROPERTY('Edition') AS nvarchar(128)) AS Edition,
  CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(128)) AS ProductVersion,
  si.cpu_count,
  si.scheduler_count,
  si.hyperthread_ratio,
  (si.physical_memory_kb/1024) AS HostPhysicalMemoryMB,
  pm.physical_memory_in_use_kb/1024 AS SqlMemoryUsedMB
FROM sys.dm_os_sys_info AS si
CROSS APPLY (SELECT physical_memory_in_use_kb FROM sys.dm_os_process_memory) pm;
"@
  CpuNow = @"
WITH rb AS (
  SELECT TOP (1) CONVERT(xml, record) AS x
  FROM sys.dm_os_ring_buffers
  WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
    AND record LIKE '%<SystemHealth>%'
  ORDER BY [timestamp] DESC
)
SELECT 
  CAST(x.value('(//SystemHealth/ProcessUtilization)[1]', 'int') AS int)      AS SqlProcessCpuPercent,
  CAST(100 - x.value('(//SystemHealth/SystemIdle)[1]', 'int') AS int)         AS SystemCpuPercent_Estimate,
  GETDATE() AS SampleTimeLocal
FROM rb;
"@
  Config = @"
SELECT name, value, value_in_use, description
FROM sys.configurations
ORDER BY name;
"@
  Databases = @"
SELECT
  name,
  state_desc,
  recovery_model_desc,
  containment_desc,
  is_encrypted,
  compatibility_level,
  collation_name,
  is_hadr_enabled,
  page_verify_option_desc,
  is_read_only
FROM sys.databases
ORDER BY name;
"@
  Schemas = @"
SELECT s.name AS schema_name, u.name AS owner
FROM sys.schemas s
JOIN sys.sysusers u ON s.principal_id = u.uid
ORDER BY s.name;
"@
  DbSizes = @"
SELECT
  DB_NAME(database_id) AS database_name,
  CAST(SUM(size)*8.0/1024.0 AS DECIMAL(18,2)) AS sizeMB,
  type_desc
FROM sys.master_files
GROUP BY database_id, type_desc
ORDER BY database_name, type_desc;
"@
  Logins = @"
SELECT
  sp.name,
  sp.type_desc,
  sp.is_disabled,
  ISNULL(sl.default_database_name,'') AS default_database_name,
  ISNULL(sl.is_policy_checked,0) AS is_policy_checked,
  ISNULL(sl.is_expiration_checked,0) AS is_expiration_checked,
  sp.create_date,
  sp.modify_date
FROM sys.server_principals sp
LEFT JOIN sys.sql_logins sl ON sp.principal_id = sl.principal_id
WHERE sp.type IN ('S','U','G')
  AND sp.name NOT LIKE '##%'
ORDER BY sp.name;
"@
  AuthMode = @"
SELECT CASE SERVERPROPERTY('IsIntegratedSecurityOnly')
         WHEN 1 THEN 'Windows Authentication'
         WHEN 0 THEN 'Mixed Mode'
       END AS AuthenticationMode;
"@
}

# ---------- Build UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SQL Twin Inspector"
$form.Width = 1200; $form.Height = 780
$form.StartPosition = 'CenterScreen'

# Server 1
$form.Controls.Add((New-Label "Server 1 (.\MSSQLSERVER or host,port)" 10 10))
$tbS1 = New-TextBox '' 10 30 280; $form.Controls.Add($tbS1)
$form.Controls.Add((New-Label "Auth" 300 10))
$cbS1Auth  = New-Combo @('Windows','SQL') 300 30 120; $form.Controls.Add($cbS1Auth)
$form.Controls.Add((New-Label "User" 430 10))
$tbS1User  = New-TextBox '' 430 30 160; $form.Controls.Add($tbS1User)
$form.Controls.Add((New-Label "Password" 600 10))
$tbS1Pass  = New-TextBox '' 600 30 160; $tbS1Pass.PasswordChar='*'; $form.Controls.Add($tbS1Pass)

# Server 2
$form.Controls.Add((New-Label "Server 2 (EC2 SQL host,port)" 10 65))
$tbS2 = New-TextBox '' 10 85 280; $form.Controls.Add($tbS2)
$form.Controls.Add((New-Label "Auth" 300 65))
$cbS2Auth  = New-Combo @('Windows','SQL') 300 85 120; $form.Controls.Add($cbS2Auth)
$form.Controls.Add((New-Label "User" 430 65))
$tbS2User  = New-TextBox '' 430 85 160; $form.Controls.Add($tbS2User)
$form.Controls.Add((New-Label "Password" 600 65))
$tbS2Pass  = New-TextBox '' 600 85 160; $tbS2Pass.PasswordChar='*'; $form.Controls.Add($tbS2Pass)

# Buttons
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Test Connections"; $btnTest.Location = New-Object System.Drawing.Point(780, 30); $btnTest.Width=130
$form.Controls.Add($btnTest)

$btnCollect = New-Object System.Windows.Forms.Button
$btnCollect.Text = "Collect"; $btnCollect.Location = New-Object System.Drawing.Point(780, 80); $btnCollect.Width=130
$form.Controls.Add($btnCollect)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export (CSV + HTML)"; $btnExport.Location = New-Object System.Drawing.Point(930, 80); $btnExport.Width=180
$form.Controls.Add($btnExport)

$status = New-Label "" 930 36; $status.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($status)

# Tabs & grids
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(10, 120)
$tabs.Width = 1160; $tabs.Height = 600

$tabOverview = New-Tab "Overview";  $gridOverview = New-Grid;  $tabOverview.Controls.Add($gridOverview)
$tabCpuNow   = New-Tab "CPU (Now)"; $gridCpuNow   = New-Grid;  $tabCpuNow.Controls.Add($gridCpuNow)
$tabConfig   = New-Tab "Configuration"; $gridConfig = New-Grid; $tabConfig.Controls.Add($gridConfig)
$tabDatabases= New-Tab "Databases"; $gridDatabases= New-Grid;  $tabDatabases.Controls.Add($gridDatabases)
$tabSchemas  = New-Tab "Schemas"; $gridSchemas    = New-Grid;  $tabSchemas.Controls.Add($gridSchemas)
$tabDbSizes  = New-Tab "DB Sizes"; $gridDbSizes   = New-Grid;  $tabDbSizes.Controls.Add($gridDbSizes)
$tabLogins   = New-Tab "Logins"; $gridLogins     = New-Grid;  $tabLogins.Controls.Add($gridLogins)
$tabAuth     = New-Tab "Auth Mode"; $gridAuth     = New-Grid;  $tabAuth.Controls.Add($gridAuth)

$tabs.TabPages.AddRange(@($tabOverview,$tabCpuNow,$tabConfig,$tabDatabases,$tabSchemas,$tabDbSizes,$tabLogins,$tabAuth))
$form.Controls.Add($tabs)

# Enable/disable user/pass for Windows vs SQL auth
$cbS1Auth.Add_SelectedIndexChanged({ $isSql = ($cbS1Auth.SelectedItem -eq 'SQL'); $tbS1User.Enabled=$isSql; $tbS1Pass.Enabled=$isSql })
$cbS2Auth.Add_SelectedIndexChanged({ $isSql = ($cbS2Auth.SelectedItem -eq 'SQL'); $tbS2User.Enabled=$isSql; $tbS2Pass.Enabled=$isSql })
$tbS1User.Enabled=$false; $tbS1Pass.Enabled=$false; $tbS2User.Enabled=$false; $tbS2Pass.Enabled=$false

# Build server hash tables
function Get-Servers {
  $S1 = @{ Server = $tbS1.Text.Trim(); Auth = $cbS1Auth.SelectedItem; User = $tbS1User.Text.Trim(); Pass = $tbS1Pass.Text }
  $S2 = @{ Server = $tbS2.Text.Trim(); Auth = $cbS2Auth.SelectedItem; User = $tbS2User.Text.Trim(); Pass = $tbS2Pass.Text }
  return @($S1,$S2)
}

# Test connections
$btnTest.Add_Click({
  $status.Text = "Testing..."; $form.Refresh()
  $S = Get-Servers
  $S1,$S2 = $S[0],$S[1]
  if ([string]::IsNullOrWhiteSpace($S1.Server) -or [string]::IsNullOrWhiteSpace($S2.Server)) {
    [System.Windows.Forms.MessageBox]::Show("Enter both Server 1 and Server 2.") | Out-Null; $status.Text = ""; return
  }
  $ok1 = Test-SqlConnection -S $S1; $ok2 = Test-SqlConnection -S $S2
  [System.Windows.Forms.MessageBox]::Show("Server 1: " + ($(if($ok1){"OK"}else{"FAILED"}) ) + "`nServer 2: " + ($(if($ok2){"OK"}else{"FAILED"}) )) | Out-Null
  $status.Text = "Ready."
})

# Collect
$btnCollect.Add_Click({
  $status.Text = "Collecting..."; $form.Refresh()
  $S = Get-Servers; $S1,$S2 = $S[0],$S[1]
  if ([string]::IsNullOrWhiteSpace($S1.Server) -or [string]::IsNullOrWhiteSpace($S2.Server)) {
    [System.Windows.Forms.MessageBox]::Show("Enter both Server 1 and Server 2.") | Out-Null; $status.Text = ""; return
  }
  try {
    $gridOverview.DataSource  = (Get-Combined -Query $Queries.Overview  -S1 $S1 -S2 $S2)
    $gridCpuNow.DataSource    = (Get-Combined -Query $Queries.CpuNow    -S1 $S1 -S2 $S2)
    $gridConfig.DataSource    = (Get-Combined -Query $Queries.Config    -S1 $S1 -S2 $S2)
    $gridDatabases.DataSource = (Get-Combined -Query $Queries.Databases -S1 $S1 -S2 $S2)
    $gridSchemas.DataSource   = (Get-Combined -Query $Queries.Schemas   -S1 $S1 -S2 $S2)
    $gridDbSizes.DataSource   = (Get-Combined -Query $Queries.DbSizes   -S1 $S1 -S2 $S2)
    $gridLogins.DataSource    = (Get-Combined -Query $Queries.Logins    -S1 $S1 -S2 $S2)
    $gridAuth.DataSource      = (Get-Combined -Query $Queries.AuthMode  -S1 $S1 -S2 $S2)
    $status.Text = "Done."
  } catch {
    $status.Text = "Error."
    [System.Windows.Forms.MessageBox]::Show("Error while collecting: $($_.Exception.Message)") | Out-Null
  }
})

# Export
$btnExport.Add_Click({
  try {
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Choose a folder to save CSVs and HTML report"
    if ($fb.ShowDialog() -ne "OK") { return }
    $root = Join-Path $fb.SelectedPath ("SqlTwinInspector_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    New-Item -Path $root -ItemType Directory -Force | Out-Null

    # Gather tables from grids
    $tables = [ordered]@{}
    $dtOverview  = Get-DataTableFromGrid $gridOverview
    $dtCpuNow    = Get-DataTableFromGrid $gridCpuNow
    $dtConfig    = Get-DataTableFromGrid $gridConfig
    $dtDatabases = Get-DataTableFromGrid $gridDatabases
    $dtSchemas   = Get-DataTableFromGrid $gridSchemas
    $dtDbSizes   = Get-DataTableFromGrid $gridDbSizes
    $dtLogins    = Get-DataTableFromGrid $gridLogins
    $dtAuth      = Get-DataTableFromGrid $gridAuth

    $tables['Overview']    = $dtOverview
    $tables['CPU (Now)']   = $dtCpuNow
    $tables['Configuration']= $dtConfig
    $tables['Databases']   = $dtDatabases
    $tables['Schemas']     = $dtSchemas
    $tables['DB Sizes']    = $dtDbSizes
    $tables['Logins']      = $dtLogins
    $tables['Auth Mode']   = $dtAuth

    # CSVs
    foreach ($k in $tables.Keys) {
      $dt = $tables[$k]
      if ($null -eq $dt -or $dt.Rows.Count -eq 0) { continue }
      $safe = ($k -replace '[^a-zA-Z0-9_-]','_')
      Export-TableCsv -dt $dt -path (Join-Path $root "$safe.csv")
    }

    # HTML report
    Export-ReportHtml -tables $tables -path (Join-Path $root "index.html")

    [System.Windows.Forms.MessageBox]::Show("Export complete:`n$root`n`nOpen index.html for a full report.") | Out-Null
  } catch {
    [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)") | Out-Null
  }
})

# Show
[void]$form.ShowDialog()
