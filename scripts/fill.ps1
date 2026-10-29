<#
.SYNOPSIS
    通过 ADB + 输入法跨设备剪贴板，批量填充抖音「屏蔽词」。

.DESCRIPTION
    背景：抖音的屏蔽词只能在手机 App 内逐条手工输入，网页端无入口。
    adb shell input text 不支持中文（只能发 ASCII），
    Android 10+ 的 service call clipboard 接口也已失效。

    可行路径是三层接力：
        电脑写入 Windows 剪贴板
          → 输入法（如微信输入法）跨设备同步到手机剪贴板
          → adb shell input keyevent 279（粘贴键）送进输入框

    本脚本只做第 1 步和第 3 步，第 2 步依赖你手机上的输入法已开启跨设备粘贴。

.PARAMETER File
    词表文件路径，一行一个词，# 开头为注释。相对路径按脚本所在目录解析。

.PARAMETER Words
    直接给词，逗号分隔，例如 "蓝鲸,窒息挑战,约死群"。

.PARAMETER Adb
    adb.exe 路径。默认 "adb"，即依赖 PATH。Windows 建议写成 "D:\platform-tools\adb.exe"。

.PARAMETER Verify
    逐词校验模式：每填一个词都回读界面，比对输入框内容，失败自动关弹窗重试。
    慢一倍，但不会在连续失败时污染列表。

.PARAMETER Calibrate
    校准模式：读一次当前界面，打印「添加关键词 / 输入框 / 提交」的实际坐标。
    换手机、换分辨率、换抖音版本后先跑一次这个。

.PARAMETER DryRun
    只打印将要填入的词和坐标，不操作手机。

.EXAMPLE
    .\fill.ps1 -File ..\data\douyin-recommended.txt -Adb "D:\platform-tools\adb.exe"

.EXAMPLE
    .\fill.ps1 -Words "蓝鲸,窒息挑战" -Verify

.EXAMPLE
    .\fill.ps1 -Calibrate -Adb "D:\platform-tools\adb.exe"

.NOTES
    仅用于管理自己或自己监护的未成年人的设备。使用前请确认已获得设备所有者同意。
#>
[CmdletBinding()]
param(
    [string]$File,
    [string]$Words,
    [string]$Adb = 'adb',
    [int]$Start = 1,
    [int]$Count = 0,
    [switch]$Verify,
    [switch]$Calibrate,
    [switch]$DryRun,
    # 以下坐标默认值按 1080x2400 逻辑分辨率标定（Android 16 实测）
    [int]$AddX = 575, [int]$AddY = 2319,
    [int]$BoxX = 492, [int]$BoxY = 1696,
    [int]$SubmitX = 539, [int]$SubmitY = 1412
)

# adb 输出是 UTF-8。不设这一行，PowerShell 会按 GBK 解码，中文全变乱码，
# 脚本就永远「找不到按钮」——这是最容易踩的坑。
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Continue'

$script:dumpPath = '/sdcard/_blockfill.xml'

function Invoke-Adb([string[]]$Arguments) {
    (& $Adb @Arguments 2>$null) -join "`n"
}

function Get-UiXml {
    Invoke-Adb @('shell', 'uiautomator', 'dump', $script:dumpPath) | Out-Null
    Invoke-Adb @('shell', 'cat', $script:dumpPath)
}

function Tap([int]$X, [int]$Y) {
    Invoke-Adb @('shell', 'input', 'tap', "$X", "$Y") | Out-Null
}

function Find-NodeCenter($Xml, [string]$Text) {
    $pattern = '<node[^>]*text="' + [regex]::Escape($Text) + '"[^>]*bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"'
    $m = [regex]::Match($Xml, $pattern)
    if (-not $m.Success) { return $null }
    return @{
        X = [int](([int]$m.Groups[1].Value + [int]$m.Groups[3].Value) / 2)
        Y = [int](([int]$m.Groups[2].Value + [int]$m.Groups[4].Value) / 2)
    }
}

function Get-EditTextValue($Xml) {
    $m = [regex]::Match($Xml, '<node[^>]*class="android\.widget\.EditText"[^>]*bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"')
    if (-not $m.Success) { return $null }
    [regex]::Match($m.Value, 'text="([^"]*)"').Groups[1].Value
}

function Close-Dialog {
    Invoke-Adb @('shell', 'input', 'keyevent', '4') | Out-Null
    Start-Sleep -Milliseconds 600
    if ((Get-UiXml) -match '填写屏蔽词') {
        Invoke-Adb @('shell', 'input', 'keyevent', '4') | Out-Null
        Start-Sleep -Milliseconds 600
    }
}

function Get-WordsFromFile([string]$Path) {
    $full = if ([System.IO.Path]::IsPathRooted($Path)) { $Path }
            else { Join-Path $PSScriptRoot $Path }
    if (-not (Test-Path $full)) { throw "词表不存在：$full" }
    Get-Content $full -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' -and -not $_.StartsWith('#') }
}

# ---------- 校准 ----------
if ($Calibrate) {
    Write-Host '正在读取当前界面…' -ForegroundColor Cyan
    $xml = Get-UiXml
    $add = Find-NodeCenter $xml '添加关键词'
    $submit = Find-NodeCenter $xml '提交'
    $edit = [regex]::Match($xml, '<node[^>]*class="android\.widget\.EditText"[^>]*bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"')

    if ($add) { Write-Host ("添加关键词  -AddX {0} -AddY {1}" -f $add.X, $add.Y) -ForegroundColor Green }
    else { Write-Host '没找到「添加关键词」——请先进入 屏蔽设置 页面' -ForegroundColor Yellow }

    if ($edit.Success) {
        $bx = [int](([int]$edit.Groups[1].Value + [int]$edit.Groups[3].Value) / 2)
        $by = [int](([int]$edit.Groups[2].Value + [int]$edit.Groups[4].Value) / 2)
        Write-Host ("输入框      -BoxX {0} -BoxY {1}" -f $bx, $by) -ForegroundColor Green
    }
    else { Write-Host '没找到输入框——请先点开「添加关键词」弹窗' -ForegroundColor Yellow }

    if ($submit) { Write-Host ("提交按钮    -SubmitX {0} -SubmitY {1}" -f $submit.X, $submit.Y) -ForegroundColor Green }
    else { Write-Host '没找到「提交」——请先点开弹窗并把键盘调出来' -ForegroundColor Yellow }

    Write-Host ''
    Write-Host '注意：键盘收起时和弹出时，输入框与提交按钮的坐标不同。' -ForegroundColor DarkGray
    Write-Host '建议按「弹窗已打开 + 键盘已弹出」的状态校准。' -ForegroundColor DarkGray
    return
}

# ---------- 组装词表 ----------
if ($Words) {
    $all = $Words -split '[,，、]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
}
elseif ($File) {
    $all = Get-WordsFromFile $File
}
else {
    throw '必须用 -File 或 -Words 指定词表，或用 -Calibrate 校准坐标。'
}

# 抖音规则：只支持中文，单个词最多 5 个字（实测）
$valid = @()
$rejected = @()
foreach ($w in $all) {
    if ($w -match '^[\u4e00-\u9fa5]+$' -and $w.Length -le 5) { $valid += $w }
    else { $rejected += $w }
}

if ($rejected.Count -gt 0) {
    Write-Host ('已跳过 {0} 个不符合抖音规则的词（非纯中文或超过 5 字）：{1}' -f $rejected.Count, ($rejected -join ' ')) -ForegroundColor Yellow
}
if ($valid.Count -eq 0) { throw '没有可用的词。' }

if ($Count -gt 0) { $valid = $valid | Select-Object -Skip ($Start - 1) -First $Count }
elseif ($Start -gt 1) { $valid = $valid | Select-Object -Skip ($Start - 1) }

Write-Host ('共 {0} 个词待填入，模式：{1}' -f $valid.Count, $(if ($Verify) { '逐词校验（稳）' } else { '快速（约 4 秒/词）' })) -ForegroundColor Cyan

if ($DryRun) {
    Write-Host ''
    Write-Host ('坐标：添加({0},{1}) 输入框({2},{3}) 提交({4},{5})' -f $AddX, $AddY, $BoxX, $BoxY, $SubmitX, $SubmitY)
    Write-Host '待填入：'
    $valid | ForEach-Object { "  $_" }
    return
}

# ---------- 主循环 ----------
$ok = 0; $fail = 0; $i = 0

# 预置第一个词，让同步提前开始
Set-Clipboard -Value $valid[0]
Start-Sleep -Milliseconds 250

for ($n = 0; $n -lt $valid.Count; $n++) {
    $w = $valid[$n]
    $i++

    if ($Verify) {
        # 逐词校验：慢但可控
        Invoke-Adb @('shell', 'uiautomator', 'dump', $script:dumpPath) | Out-Null
        $xml = Invoke-Adb @('shell', 'cat', $script:dumpPath)

        $addBtn = Find-NodeCenter $xml '添加关键词'
        if (-not $addBtn) {
            Write-Host ("[{0}/{1}] {2} 失败：不在屏蔽设置页面" -f $i, $valid.Count, $w) -ForegroundColor Red
            $fail++; continue
        }
        Tap $addBtn.X $addBtn.Y
        Start-Sleep -Milliseconds 1400

        Tap $BoxX $BoxY
        Start-Sleep -Milliseconds 1400

        Invoke-Adb @('shell', 'input', 'keyevent', '279') | Out-Null
        Start-Sleep -Milliseconds 1000

        $xml = Get-UiXml
        $val = Get-EditTextValue $xml
        if ($val -ne $w) {
            Invoke-Adb @('shell', 'input', 'keyevent', '279') | Out-Null
            Start-Sleep -Milliseconds 900
            $xml = Get-UiXml
            $val = Get-EditTextValue $xml
        }
        if ($val -ne $w) {
            Write-Host ("[{0}/{1}] {2} 失败：输入框实际为 [{3}]" -f $i, $valid.Count, $w, $val) -ForegroundColor Red
            $fail++; Close-Dialog; continue
        }

        $submit = Find-NodeCenter $xml '提交'
        if (-not $submit) {
            Write-Host ("[{0}/{1}] {2} 失败：未找到提交按钮" -f $i, $valid.Count, $w) -ForegroundColor Red
            $fail++; Close-Dialog; continue
        }
        Tap $submit.X $submit.Y
        if ($n + 1 -lt $valid.Count) { Set-Clipboard -Value $valid[$n + 1] }
        Start-Sleep -Milliseconds 1800

        Write-Host ("[{0}/{1}] {2} OK" -f $i, $valid.Count, $w) -ForegroundColor Green
        $ok++
    }
    else {
        # 快速模式：整段点击在手机端一次执行，只有两次 adb 往返
        $chainA = "input tap $AddX $AddY; sleep 0.8; input tap $BoxX $BoxY; sleep 0.9; input keyevent 279"
        Invoke-Adb @('shell', $chainA) | Out-Null

        if ($n + 1 -lt $valid.Count) { Set-Clipboard -Value $valid[$n + 1] }

        $chainB = "sleep 0.9; input tap $SubmitX $SubmitY; sleep 0.6"
        Invoke-Adb @('shell', $chainB) | Out-Null

        Write-Host ("[{0}/{1}] {2}" -f $i, $valid.Count, $w)
        $ok++
    }
}

Write-Host ''
Write-Host ("完成：提交 {0} 个，失败 {1} 个" -f $ok, $fail) -ForegroundColor Cyan
if (-not $Verify) {
    Write-Host '快速模式不做逐词校验。填完后请打开手机上的屏蔽列表翻一遍，确认没有漏词或错词。' -ForegroundColor DarkGray
}
