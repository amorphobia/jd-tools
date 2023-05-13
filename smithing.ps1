#
# 铸键术 - 为小狼毫下载和更新键道配方
#
# Copyright © 2023 Xuesong Peng <pengxuesong.cn@gmail.com>
# ConvertTo-ImageSource: Copyright © 2016 Chris Carter
# WPF GUI support: Copyright © 2019-2023 Pete Batard <pete@akeo.ie>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# 本程序是免费软件：您可以根据自由软件基金会发布的 GNU 通用公共许可证（第三版或
# 任何您指定的更新版本）中的条款重新分发和／或修改它。
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# 分发该程序是希望它有用，但不提供任何保证；甚至没有针对特定用途的适销性或适用性的
# 默示保证。有关详细信息，请参阅 GNU 通用公共许可证。
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
# 您应该已随本程序收到 GNU 通用公共许可证的副本。如果没有，请参阅 <http://www.gnu.org/licenses/>
#

# 注意：含有非 ASCII 编码的文字，需要将编码格式设置为 ANSI (GB18030) 或带 BOM 的 UTF-8

param(
    [string]$RimeUserDir = "$env:APPDATA\Rime\"
)

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

Write-Host "请稍候……"

$AppTitle = "铸键术"

$defaultContent = @"
patch:
  schema_list:
    - schema: xkjd6
"@

function Get-Jiandao {
    param(
        [string]$Url,
        [string]$DestPath,
        [string]$Branch = "plum",
        [switch]$SystemGit,
        [switch]$DirectDownload
    )

    if (Test-Path $DestPath) {
        Remove-Item -Recurse -Force $DestPath
    }

    if ($DirectDownload) {
        $f = New-TemporaryFile
        $tempFile = Move-Item -Path (Convert-Path $f.PSPath) -Destination ((Convert-Path $f.PSParentPath) + "\" + ($Url.Split('/')[-1])) -PassThru -Force
        Invoke-WebRequest -Uri $Url -OutFile $tempFile
        Expand-Archive -Path $tempFile.FullName -DestinationPath $DestPath -Force
        $DestPath = $DestPath + "\" + (Get-ChildItem $DestPath)[0].Name
        Remove-Item -Recurse -Force $tempFile
    } else {
        if ($SystemGit) {
            git clone --depth 1 $Url $DestPath --branch $Branch
        } else {
            if (!("LibGit2Sharp.Repository" -as [type])) {
                $env:PATH = "" + $PSScriptRoot + ";" + $env:PATH
                Import-Module $PSScriptRoot\LibGit2Sharp.dll
            }
            $Opt = [LibGit2Sharp.CloneOptions]::new()
            $Opt.BranchName = $Branch
            $Opt.OnProgress = {
                param(
                    $output
                )

                Write-Host "$output"
                return $true
            }
            $Opt.OnTransferProgress = {
                Param(
                    [LibGit2Sharp.TransferProgress]$progress
                )

                $ratio = [Int32](100 * $progress.IndexedObjects / $progress.TotalObjects)

                $XMLForm.Title = "传输中…… $ratio% 已完成"
                Write-Progress -Activity "传输中……" -status "$ratio% 已完成" -PercentComplete $ratio

                return $true
            }
            [LibGit2Sharp.Repository]::Clone($Url, $DestPath, $Opt)
        }
    }

    $DestPath
}

function Copy-Schema {
    param(
        [string]$filePath,
        [bool]$overwrite,
        [string]$luaOp
    )
    $excludeFiles = @(".git", "recipe.yaml")
    if (-not $overwrite -and (Test-Path "$RimeUserDir\xkjd6.user.dict.yaml")) {
        $excludeFiles += "xkjd6.user.dict.yaml"
    }
    if ($luaOp -ne "overwrite" -and (Test-Path "$RimeUserDir\rime.lua")) {
        $excludeFiles += "rime.lua"
    }
    if ($luaOp -eq "append") {
        [IO.File]::AppendAllText("$RimeUserDir\rime.lua", "-- 星空键道`n" + [System.IO.File]::ReadAllText("$filePath\rime.lua"))
    }
    Copy-Item -Recurse -Force "$filePath\*" $RimeUserDir -Exclude $excludeFiles
}

function Set-DefaultContent {
    [IO.File]::WriteAllLines("$RimeUserDir\default.custom.yaml", $defaultContent)
}

# From https://www.powershellgallery.com/packages/IconForGUI/1.5.2
# Copyright © 2016 Chris Carter. All rights reserved.
# License: https://creativecommons.org/licenses/by-sa/4.0/
function ConvertTo-ImageSource {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [System.Drawing.Icon]$Icon
    )

    Process {
        foreach ($i in $Icon) {
            [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
                $i.Handle,
                (New-Object System.Windows.Int32Rect -Args 0,0,$i.Width, $i.Height),
                [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions()
            )
        }
    }
}

# WPF GUI support From https://github.com/pbatard/Fido
# Copyright © 2019-2023 Pete Batard <pete@akeo.ie>
# License： https://github.com/pbatard/Fido/blob/master/LICENSE.txt
$Drawing_Assembly = "System.Drawing"
# PowerShell 7 altered the name of the Drawing assembly...
if ($host.version -ge "7.0") {
    $Drawing_Assembly += ".Common"
}

$Signature = @{
    Namespace            = "WinAPI"
    Name                 = "Utils"
    Language             = "CSharp"
    UsingNamespace       = "System.Runtime", "System.IO", "System.Text", "System.Drawing", "System.Globalization"
    ReferencedAssemblies = $Drawing_Assembly
    ErrorAction          = "Stop"
    WarningAction        = "Ignore"
    MemberDefinition     = @"
        [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        internal static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr handle, int state);
        // Extract an icon from a DLL
        public static Icon ExtractIcon(string file, int number, bool largeIcon) {
            IntPtr large, small;
            ExtractIconEx(file, number, out large, out small, 1);
            try {
                return Icon.FromHandle(largeIcon ? large : small);
            } catch {
                return null;
            }
        }
"@
}

if (!("WinAPI.Utils" -as [type])) {
    Add-Type @Signature
}
Add-Type -AssemblyName PresentationFramework

# Hide the powershell window: https://stackoverflow.com/a/27992426/1069307
[WinAPI.Utils]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0) | Out-Null

function Update-Control([object]$Control) {
    $Control.Dispatcher.Invoke("Render", [Windows.Input.InputEventHandler] { $Confirm.UpdateLayout() }, $null, $null) | Out-Null
}

[xml]$XAML = @"
<Window xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation" Height = "352" Width = "384" ResizeMode = "NoResize">
    <Grid Name = "XMLGrid">
        <Button Name = "Confirm" FontSize = "16" Height = "26" Width = "160" HorizontalAlignment = "Left" VerticalAlignment = "Top" Margin = "14,266,0,0" Content = "确定" />
        <Button Name = "Cancel" FontSize = "16" Height = "26" Width = "160" HorizontalAlignment = "Left" VerticalAlignment = "Top" Margin = "194,266,0,0" Content = "取消" />
        <TextBlock Name = "Source" FontSize = "16" Width="340" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "16,8,0,0" Text = "选择源" />
        <RadioButton Name = "GitHub" FontSize = "16" Height = "26" Width = "80" HorizontalAlignment = "Left" VerticalAlignment = "Top" Margin = "14,38,0,0" GroupName="Source" Content="GitHub" IsChecked = "True" />
        <RadioButton Name = "Gitee" FontSize = "16" Height = "26" Width = "80" HorizontalAlignment = "Left" VerticalAlignment = "Top" Margin = "194,38,0,0" GroupName="Source" Content="Gitee" />
        <TextBlock Name = "LuaTitle" FontSize = "16" Width="340" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "16,76,0,0" Text = "如何处理 rime.lua" />
        <ComboBox Name = "LuaOps" FontSize = "14" Height = "24" Width="340" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "16,108,0,0" SelectedIndex = "0" />
        <CheckBox Name = "Overwrite" FontSize = "16" Width="340" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "14,152,0,0" Content = "覆盖用户词典 (xkjd6.user.dict.yaml)" />
        <CheckBox Name = "OverwriteDefault" FontSize = "16" Width="340" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "14,190,0,0" Content = "覆盖 default.custom.yaml" />
        <TextBlock Name = "Book" xml:space="preserve" FontSize = "16" Width="340" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "14,228,0,0"><Hyperlink Name = "JDRepo" NavigateUri="https://github.com/xkinput/Rime_JD">键道官方仓库</Hyperlink>    <Hyperlink Name = "BookLink" NavigateUri="https://pingshunhuangalex.gitbook.io/rime-xkjd/">键道详尽操作指南</Hyperlink></TextBlock>
    </Grid>
</Window>
"@

$LuaOperations = @(
    @("不作修改", "untouch"),
    @("追加", "append"),
    @("覆盖", "overwrite")
)

$XMLForm = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML))
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $XMLForm.FindName($_.Name) -Scope Script }
$XMLForm.Title = $AppTitle

$XMLForm.Icon = [WinAPI.Utils]::ExtractIcon("shell32.dll", 43, $true) | ConvertTo-ImageSource

$i = 0
$ops = @()
foreach($op in $LuaOperations) {
    $ops += @(New-Object PsObject -Property @{ Text = $op[0]; Operation = $op[1]; Index = $i })
    $i++
}
$LuaOps.ItemsSource = $ops
$LuaOps.DisplayMemberPath = "Text"

$Cancel.add_click({
    $XMLForm.Close()
})

$Confirm.add_click({
    $GitHub.IsEnabled = $false
    $Gitee.IsEnabled = $false
    $Confirm.IsEnabled = $false
    $Cancel.IsEnabled = $false
    $Overwrite.IsEnabled = $false
    $OverwriteDefault.IsEnabled = $false
    Update-Control($Confirm)
    Update-Control($Cancel)

    $XMLForm.Title = "请稍候……"
    Update-Control($XMLForm)

    $RepoHost = if ($GitHub.IsChecked) { "github.com" } else { "gitee.com" }
    $RepoUrl = "https://" + $RepoHost + "/xkinput/Rime_JD.git"
    $DestPath = "$env:TEMP\jd"
    $RepoBranch = "plum"
    $gitPresent = (Get-Command "git.exe" -ErrorAction SilentlyContinue) -ne $null

    $filePath = ""
    $cancelled = $false
    if ($gitPresent) {
        $filePath = Get-Jiandao -Url $RepoUrl -DestPath $DestPath -Branch $RepoBranch -SystemGit
    } elseif ($GitHub.IsChecked) {
        $url = "https://github.com/xkinput/Rime_JD/archive/refs/heads/plum.zip"
        $filePath = Get-Jiandao -Url $url -DestPath $DestPath -DirectDownload
    } else {
        $msgBoxInput = [System.Windows.MessageBox]::Show("将使用本工具自带的 git 克隆，可能会造成窗口假死，请耐心等待", "警告：未在系统路径中找到 git", "OKCancel")
        if ($msgBoxInput -eq "OK") {
            $filePath = Get-Jiandao -Url $RepoUrl -DestPath $DestPath -Branch $RepoBranch
        } else {
            $cancelled = $true
        }
    }
    if ($filePath.GetType().Name -ne "String") {
        $filePath = $filePath[-1]
    }

    if (-not $cancelled) {
        Copy-Schema $filePath -overwrite $Overwrite.IsChecked -luaOp $LuaOps.SelectedValue.Operation
        Remove-Item -Recurse -Force $DestPath
        if ($OverwriteDefault.IsChecked -or -not (Test-Path "$RimeUserDir\default.custom.yaml")) {
            Set-DefaultContent
        }
        $deployBox = [System.Windows.MessageBox]::Show("是否要现在重新部署？", "更新完毕", "OKCancel")
        if ($deployBox -eq "OK") {
            $weaselPath = (Get-ChildItem "${env:ProgramFiles(x86)}\Rime\" | Where-Object Name -match '^weasel-[\d\.]+$')[0].FullName
            if (Test-Path "$weaselPath\WeaselDeployer.exe") {
                & $weaselPath\WeaselDeployer.exe /deploy
            }
        }
        $XMLForm.Close()
    }

    $XMLForm.Title = $AppTitle

    $GitHub.IsEnabled = $true
    $Gitee.IsEnabled = $true
    $Confirm.IsEnabled = $true
    $Cancel.IsEnabled = $true
    $Overwrite.IsEnabled = $true
    $OverwriteDefault.IsEnabled = $true
})

$JDRepo.add_click({
    $RepoWebUrl = if ($GitHub.IsChecked) { 'https://github.com/xkinput/Rime_JD' } else { 'https://gitee.com/xkinput/Rime_JD' }
    [system.Diagnostics.Process]::start($RepoWebUrl)
})

$BookLink.add_click({
    [system.Diagnostics.Process]::start('https://pingshunhuangalex.gitbook.io/rime-xkjd/')
})

$XMLForm.Add_Loaded({$XMLForm.Activate()})
$XMLForm.ShowDialog() | Out-Null
