<#
.SYNOPSIS
    扫描"脚本自身所在目录"下所有顶层文件夹的 NTFS 权限，生成可交互 HTML 报表
    （文件夹 x 用户 权限矩阵，可勾选排除某些用户，也可勾选排除某些文件夹）。

.说明
    - 默认扫描的目录 = 脚本文件所在的目录（与运行 PowerShell 时的当前工作目录无关，
      也就是说不管你在哪个路径下打开 PowerShell 执行这个脚本，它扫描的都是脚本
      自己所在的那个文件夹）。
    - 只检查顶层文件夹本身的权限（不递归进入子文件夹）。
    - 每个单元格显示该用户对该文件夹的：读 / 写 / 完全控制 三个维度。
    - 用户名只在表头出现一次（不重复），文件夹名只在第一列出现一次（不重复）。
    - 不依赖任何第三方模块，纯 PowerShell 自带命令实现。
    - 生成的 HTML 报表默认也保存在脚本所在目录。
    - HTML 报表顶部有两个"排除"区域：
        1) Exclude Users  - 勾选某个用户，把该用户列从表格中隐藏。
        2) Exclude Folders - 勾选某个文件夹，把该文件夹行从表格中隐藏。
      默认打开报表时所有用户和文件夹都显示（不预先勾选任何人/任何文件夹），
      如果需要默认就排除某些账户或文件夹，用 -DefaultHiddenUsers /
      -DefaultHiddenFolders 指定。
    - 如果想在 Excel 里看数据：直接用 Excel 打开生成的 .html 文件即可
      （Excel 能识别 HTML 表格），再另存为 .xlsx 也可以。

.用法
    把脚本放到要检查的目录里，双击运行，或在该目录下打开 PowerShell 执行:
        .\Get-FolderPermissions.ps1
    如果需要改成扫描/输出到别的目录，可以显式指定:
        .\Get-FolderPermissions.ps1 -TargetPath "D:\Share" -OutputPath "D:\Share\Report"
    如果想调整"默认就勾选排除"的用户或文件夹列表（都支持通配符，如 *SYSTEM*）:
        .\Get-FolderPermissions.ps1 -DefaultHiddenUsers @('NT AUTHORITY\SYSTEM','CodexSandboxUsers') -DefaultHiddenFolders @('#Recycle*','_temp*')
#>

param(
    # 要扫描其"顶层子文件夹"的目录，默认是脚本自身所在目录（不是运行时所在的当前目录）
    [string]$TargetPath = $PSScriptRoot,

    # 报表输出目录，默认也是脚本自身所在目录
    [string]$OutputPath = $PSScriptRoot,

    # 生成的 HTML 报表中，默认预先勾选为"排除/隐藏"的用户名或用户组（支持通配符）。
    # 默认是空列表，即打开报表时所有用户都显示；需要的话可以自己传入，例如：
    # -DefaultHiddenUsers @('NT AUTHORITY\SYSTEM','CodexSandboxUsers')
    # 用户在报表里仍可以随时手动勾选/取消勾选来隐藏或重新显示某个用户。
    [string[]]$DefaultHiddenUsers = @(),

    # 生成的 HTML 报表中，默认预先勾选为"排除/隐藏"的文件夹名（支持通配符）。
    # 默认是空列表，即打开报表时所有文件夹都显示。
    [string[]]$DefaultHiddenFolders = @()
)

# 兼容某些直接用 F5/ISE 运行、$PSScriptRoot 可能为空的情况
if ([string]::IsNullOrWhiteSpace($TargetPath)) {
    $TargetPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# 0. 准备工作
# ---------------------------------------------------------------
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$reportTitle     = "SHARE Folder Permissions Report"
$generatedAt     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$htmlPath        = Join-Path $OutputPath "$reportTitle.html"

Add-Type -AssemblyName System.Web

Write-Host "正在扫描目录: $TargetPath" -ForegroundColor Cyan

# ---------------------------------------------------------------
# 1. 权限归类函数：把 FileSystemRights 归类为 读/写/完全控制
# ---------------------------------------------------------------
function Get-PermLabel {
    param([System.Security.AccessControl.FileSystemRights]$Rights)

    $fsr = [System.Security.AccessControl.FileSystemRights]

    $hasFull  = $Rights.HasFlag($fsr::FullControl)
    $hasWrite = $hasFull -or $Rights.HasFlag($fsr::Write) -or $Rights.HasFlag($fsr::Modify)
    $hasRead  = $hasFull -or $Rights.HasFlag($fsr::Read) -or $Rights.HasFlag($fsr::ReadAndExecute)

    $parts = @()
    if ($hasRead)  { $parts += "Read" }
    if ($hasWrite) { $parts += "Write" }
    if ($hasFull)  { $parts += "Full Control" }

    if ($parts.Count -eq 0) { return $null }
    return ($parts -join ",")
}

# 判断某个用户名是否匹配 -DefaultHiddenUsers 列表（支持通配符，如 *SYSTEM*）
function Test-IsDefaultHidden {
    param([string]$UserName)
    foreach ($pattern in $DefaultHiddenUsers) {
        if ($UserName -like $pattern) { return $true }
    }
    return $false
}

# 判断某个文件夹名是否匹配 -DefaultHiddenFolders 列表（支持通配符）
function Test-IsFolderDefaultHidden {
    param([string]$FolderName)
    foreach ($pattern in $DefaultHiddenFolders) {
        if ($FolderName -like $pattern) { return $true }
    }
    return $false
}

# ---------------------------------------------------------------
# 2. 获取顶层文件夹列表（不递归）
# ---------------------------------------------------------------
$topFolders = Get-ChildItem -Path $TargetPath -Directory -ErrorAction SilentlyContinue |
    Sort-Object Name -Unique

if (-not $topFolders -or $topFolders.Count -eq 0) {
    Write-Warning "在 $TargetPath 下未找到任何顶层文件夹。"
    return
}

# ---------------------------------------------------------------
# 3. 遍历每个文件夹的 ACL，构建 "文件夹 -> 用户 -> 权限" 的字典
#    同时收集所有出现过的用户名（去重）
# ---------------------------------------------------------------
$matrix    = [ordered]@{}   # 文件夹名 -> @{ 用户 -> 权限字符串 }
$userSet   = New-Object 'System.Collections.Generic.SortedSet[string]'

foreach ($folder in $topFolders) {
    $folderName = $folder.Name
    $matrix[$folderName] = @{}

    try {
        $acl = Get-Acl -Path $folder.FullName
    }
    catch {
        Write-Warning "无法读取权限: $($folder.FullName) - $($_.Exception.Message)"
        continue
    }

    foreach ($ace in $acl.Access) {
        # 只统计允许(Allow)的规则，跳过拒绝(Deny)规则
        if ($ace.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) {
            continue
        }

        $userName = $ace.IdentityReference.Value
        $label    = Get-PermLabel -Rights $ace.FileSystemRights
        if (-not $label) { continue }

        [void]$userSet.Add($userName)

        # 同一用户可能有多条规则（例如显式+继承），做合并
        if ($matrix[$folderName].ContainsKey($userName)) {
            $existing = $matrix[$folderName][$userName] -split ","
            $combined = ($existing + ($label -split ",")) | Select-Object -Unique
            # 保持 Read/Write/Full Control 的固定顺序
            $order = @("Read","Write","Full Control")
            $sorted = $order | Where-Object { $combined -contains $_ }
            $matrix[$folderName][$userName] = ($sorted -join ",")
        }
        else {
            $matrix[$folderName][$userName] = $label
        }
    }
}

$userList = $userSet | Sort-Object
$folderList = $topFolders.Name

Write-Host "共发现 $($folderList.Count) 个顶层文件夹, $($userList.Count) 个不同用户/组。" -ForegroundColor Cyan

if ($userList.Count -eq 0) {
    Write-Warning "未提取到任何用户权限信息，请检查是否有权限读取这些文件夹的 ACL。"
}

# ---------------------------------------------------------------
# 4. 构建统一的行数据（每个文件夹一行，每个用户一列）
# ---------------------------------------------------------------
$rows = foreach ($folderName in $folderList) {
    $rowObj = [ordered]@{ "Folder" = $folderName }
    foreach ($userName in $userList) {
        $val = $matrix[$folderName][$userName]
        $rowObj[$userName] = if ($val) { $val } else { "-" }
    }
    [PSCustomObject]$rowObj
}

# ---------------------------------------------------------------
# 5. 生成可交互 HTML 报表（下拉框只看单一用户）
# ---------------------------------------------------------------

# 5.1 表头
$headerCells = "<th>Folder</th>"
foreach ($u in $userList) {
    $uEsc = [System.Web.HttpUtility]::HtmlEncode($u)
    $headerCells += "<th class='userCol' data-user='$uEsc'>$uEsc</th>"
}

# 5.2 表体
$bodyRows = ""
foreach ($row in $rows) {
    $folderEsc = [System.Web.HttpUtility]::HtmlEncode($row."Folder")
    $cells = "<td class='folderCol'>$folderEsc</td>"
    foreach ($u in $userList) {
        $val = $row.$u
        $cssClass = if ($val -eq "-") { "noperm" } else { "hasperm" }
        $cells += "<td class='userCol $cssClass' data-user='$([System.Web.HttpUtility]::HtmlEncode($u))'>$([System.Web.HttpUtility]::HtmlEncode($val))</td>"
    }
    $bodyRows += "<tr class='folderRow' data-folder='$folderEsc'>$cells</tr>`n"
}

# 5.3 排除用户的复选框列表（默认根据 -DefaultHiddenUsers 预先勾选）
$excludeCheckboxes = ""
foreach ($u in $userList) {
    $uEsc = [System.Web.HttpUtility]::HtmlEncode($u)
    $checkedAttr = if (Test-IsDefaultHidden -UserName $u) { "checked" } else { "" }
    $excludeCheckboxes += "<label class='exclude-item'><input type='checkbox' class='excludeChk' value='$uEsc' onchange='applyFilters()' $checkedAttr> $uEsc</label>`n"
}

# 5.4 排除文件夹的复选框列表（默认根据 -DefaultHiddenFolders 预先勾选）
$excludeFolderCheckboxes = ""
foreach ($f in $folderList) {
    $fEsc = [System.Web.HttpUtility]::HtmlEncode($f)
    $checkedAttr = if (Test-IsFolderDefaultHidden -FolderName $f) { "checked" } else { "" }
    $excludeFolderCheckboxes += "<label class='exclude-item'><input type='checkbox' class='excludeFolderChk' value='$fEsc' onchange='applyFilters()' $checkedAttr> $fEsc</label>`n"
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>$reportTitle</title>
<style>
    body { font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; background:#f7f7f9; }
    h1 { font-size: 20px; }
    h2.section-title { font-size: 14px; margin: 0 0 8px 0; color: #333; }
    .panel { background:#fff; border: 1px solid #ddd; border-radius: 6px; padding: 12px 16px; margin-bottom: 16px; }
    .exclude-list { display: flex; flex-wrap: wrap; gap: 6px 18px; max-height: 160px; overflow-y: auto; padding: 4px 0; }
    .exclude-item { font-size: 13px; white-space: nowrap; }
    .exclude-item input { margin-right: 4px; }
    .panel-buttons { margin-bottom: 10px; }
    .panel-buttons button { font-size: 12px; padding: 4px 10px; margin-right: 8px; cursor: pointer; }
    table { border-collapse: collapse; width: 100%; background: #fff; }
    th, td { border: 1px solid #ddd; padding: 6px 10px; font-size: 13px; text-align: center; }
    th { background: #2f5597; color: #fff; position: sticky; top: 0; }
    td.folderCol { text-align: left; font-weight: bold; background: #f0f4fa; }
    td.hasperm { color: #1a7f37; }
    td.noperm { color: #999; }
    tr:nth-child(even) td.folderCol { background: #e8eef7; }
    .meta { color: #666; font-size: 12px; margin-bottom: 10px; }
    .signature-line { display: flex; flex-wrap: wrap; gap: 24px; margin-bottom: 20px; font-size: 13px; color: #333; }
    .sig-item { display: flex; align-items: flex-end; white-space: nowrap; }
    .sig-blank { display: inline-block; width: 140px; border-bottom: 1px solid #333; margin-left: 6px; height: 16px; }
    @media print {
        .panel { display: none; }
        .sig-blank { width: 160px; }
    }
</style>
</head>
<body>
<h1>$reportTitle</h1>
<div class="meta">Scan Directory: $([System.Web.HttpUtility]::HtmlEncode($TargetPath)) &nbsp;|&nbsp; Generated: $generatedAt</div>
<div class="signature-line">
    <span class="sig-item">GM:<span class="sig-blank"></span></span>
    <span class="sig-item">DGM:<span class="sig-blank"></span></span>
    <span class="sig-item">CISO:<span class="sig-blank"></span></span>
    <span class="sig-item">ISO:<span class="sig-blank"></span></span>
</div>

<div class="panel">
    <h2 class="section-title">Exclude Users (checked users will be hidden from the table below)</h2>
    <div class="panel-buttons">
        <button type="button" onclick="setAllExclude('.excludeChk', true)">Exclude All</button>
        <button type="button" onclick="setAllExclude('.excludeChk', false)">Show All</button>
    </div>
    <div class="exclude-list" id="excludeList">
        $excludeCheckboxes
    </div>
</div>

<div class="panel">
    <h2 class="section-title">Exclude Folders (checked folders will be hidden from the table below)</h2>
    <div class="panel-buttons">
        <button type="button" onclick="setAllExclude('.excludeFolderChk', true)">Exclude All</button>
        <button type="button" onclick="setAllExclude('.excludeFolderChk', false)">Show All</button>
    </div>
    <div class="exclude-list" id="excludeFolderList">
        $excludeFolderCheckboxes
    </div>
</div>

<table id="permTable">
<thead>
<tr>$headerCells</tr>
</thead>
<tbody>
$bodyRows
</tbody>
</table>

<script>
function applyFilters() {
    var excludedUsers = {};
    document.querySelectorAll('.excludeChk').forEach(function(chk) {
        if (chk.checked) { excludedUsers[chk.value] = true; }
    });

    var excludedFolders = {};
    document.querySelectorAll('.excludeFolderChk').forEach(function(chk) {
        if (chk.checked) { excludedFolders[chk.value] = true; }
    });

    document.querySelectorAll('.userCol').forEach(function(el) {
        var u = el.getAttribute('data-user');
        el.style.display = excludedUsers[u] === true ? 'none' : '';
    });

    document.querySelectorAll('.folderRow').forEach(function(row) {
        var f = row.getAttribute('data-folder');
        row.style.display = excludedFolders[f] === true ? 'none' : '';
    });
}

function setAllExclude(selector, checked) {
    document.querySelectorAll(selector).forEach(function(chk) {
        chk.checked = checked;
    });
    applyFilters();
}

// Apply filters on page load based on default checkbox state
document.addEventListener('DOMContentLoaded', applyFilters);
</script>
</body>
</html>
"@

$html | Out-File -FilePath $htmlPath -Encoding UTF8

Write-Host "HTML 报表已生成: $htmlPath" -ForegroundColor Green
Write-Host "完成。" -ForegroundColor Cyan