$productId = "Microsoft.VisualStudio.Product.Community"
$channelUri = "https://aka.ms/vs/17/release/channel"
$channelUri_all = "https://aka.ms/vs/channels"
$channelUriPattern = "https://aka\.ms/vs/(\d+)/release/channel"
$vsinstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vs_installer.exe";
$VSWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$NativeDesktopWorkload = "Microsoft.VisualStudio.Workload.NativeDesktop"

$InstallerProcesses = @("vs_installer", "VSIXInstaller", "InstallCleanup", "setup")
$VS_download = "https://c2rsetup.officeapps.live.com/c2r/downloadVS.aspx?sku=community&channel=Release"
$VS_downloadPath = "vs_community.exe"
$VS_isDoenlaed = $false
$VSC_download = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64"
$VSC_downloadPath = "VSCodeUserSetup.exe"
$VSC_isDoenlaed = $false
$VSC_settingsPath = "$env:appdata\Code\User\settings.json"
$shell_setup = "`&(Join-Path (`&`"`${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe`" -latest -products * -requires Microsoft.VisualStudio.Workload.NativeDesktop -property installationPath) `"Common7\Tools\Launch-VsDevShell.ps1`")"
$codeRunner_c = "cd `$dir `&`& `$chcpCode = [System.Text.Encoding]::GetEncoding([int]((chcp) -replace '[^\d]')) `&`& [Console]::OutputEncoding = `$chcpCode;cl.exe /nologo /source-charset:([System.Text.Encoding]`$chcpCode).WebName /execution-charset:([System.Text.Encoding]`$chcpCode).WebName `$fileName && echo `"`" ; .\`"`$fileNameWithoutExt`""
$codeRunner_cpp = "cd `$dir `&`& `$chcpCode = [System.Text.Encoding]::GetEncoding([int]((chcp) -replace '[^\d]')) `&`& [Console]::OutputEncoding = `$chcpCode;cl.exe /nologo /EHsc /source-charset:([System.Text.Encoding]`$chcpCode).WebName /execution-charset:([System.Text.Encoding]`$chcpCode).WebName `$fileName && echo `"`" ; .\`"`$fileNameWithoutExt`""
[string[]]$ps_settings = @("-NoExit", "-Command", "chcp 65001;", $shell_setup)

function Read-YN([ref]$ans, $mes)
{
    while($true)
    {
        [void]($yn = Read-Host $mes)
        if($yn -imatch '^(y|yes)$')
        {
            $ans.Value = $true
            return
        }
        if($yn -imatch '^(n|no)$')
        {
            $ans.Value = $false
            return
        }
    }
}
function exitScript() {
    pause
    exit
}
function exit-Q {
    param (
        $Question,
        $exitMes = $null
    )
    $yn = $false
    Read-YN ([ref]$yn) $Question
    if (!$yn) {
        if(![System.String]::IsNullOrWhiteSpace($exitMes))
        {
            [void](Write-Host $exitMes)
        }
        exitScript
    }
}
function ChackRunningVS()
{
    $IsRunning = $false
    do{
        foreach ($procName in $InstallerProcesses) {
            if (Get-Process -Name $procName -ErrorAction SilentlyContinue) {
                $IsRunning = $true
                break
            }
        }
        if($IsRunning)
        {
            Write-Host "VisualStudio関連のプロセスを検知しました" -ForegroundColor Yellow
            Write-Host "VisualStudio関連アプリを終了後、再試行してください"
            Write-Host ""
            
            exit-Q "再試行しますか？ (y/n)" "インストールをキャンセルします"
        }
    } while ($IsRunning)
    Write-Host "インストールを開始します。インストーラが起動したら指示に従ってください" -ForegroundColor Yellow
}

function download([ref]$filename, $url)
{
    $dldir = Split-Path -Path $filename.Value -Parent
    if([System.String]::IsNullOrWhiteSpace($dldir))
    {
        $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($filename.Value)
    }
    else
    {
        $BaseName = Join-Path $dldir ([System.IO.Path]::GetFileNameWithoutExtension($filename.Value))
    }

    $Extension = [System.IO.Path]::GetExtension($filename.Value)
    $tempname = $BaseName
    for ($i = 2; (Test-Path ("$($tempname)$($Extension)")); $i++)
    {
        $tempname = "$($BaseName)($i)"
    }
    $filename.Value = "$($tempname)$($Extension)"
    while($true)
    {
        try {
            Invoke-WebRequest -Uri $url -outfile $filename.Value
            break
        }
        catch {
            Write-Host "ダウンロード処理中にエラーが発生しました"
            Write-Host $_
            Write-Host ""
            exit-Q "再試行しますか？ (y/n)" "インストールをキャンセルします"
        }
    }
}
function get-latestVS()
{
    $resultUri = $channelUri
    if (!($resultUri -match $channelUriPattern))
    {
        return $resultUri
    }
    $version = [int]($matches[1])
    while($true)
    {
        try {
            $allChannels = Invoke-RestMethod -Uri $channelUri_all
            break
        }
        catch {
            [void](Write-Host "VisualStudioの最新バージョンの取得に失敗しました")
            [void](Write-Host $_)
            [void](Write-Host "")
            exit-Q "再試行しますか？ (y/n)" "インストールをキャンセルします"
        }
    }

    foreach ($channel in $allChannels.channels)
    {
        if($channel.channelUri -match $channelUriPattern)
        {
            if($version -lt [int]($matches[1]))
            {
                $version = [int]($matches[1])
                $resultUri = $channel.channelUri
            }
        }
    }
    return $resultUri
}
Write-Host "**********************************************************************"
Write-Host "** CodeRunner auto installer v1.4"
Write-Host "**********************************************************************"
Write-Host ""

$channelUri = get-latestVS
$VSPath = @()
if(Test-Path $VSWhere)
{
    $VSPath = &$VSWhere -products * -format json | ConvertFrom-Json
    if($VSPath.Count -gt 0)
    {
        if((&$VSWhere -latest -products * -requires Microsoft.VisualStudio.Workload.NativeDesktop -format json | ConvertFrom-Json).Count -gt 0)
        {
            Write-Host "C++ツールがインストール済みのVisualStudioを検知しました"
        }
        else
        {
            Write-Host "C++用のツールがインストール済みのVisualStudioから検出できませんでした"
            Write-Host "C++用のツールをVisualStudioに追加してください"
            Write-Host "インストールを開始したら、このウィンドウへ戻って操作を続けてください" -ForegroundColor Yellow
            Write-Host ""
            if($VSPath.Count -gt 1)
            {
                Write-Host "複数のVisualStudioのインスタンスを検出しました"
                Write-Host
                for($i=0; $i -lt $VSPath.Count; $i++){
                    Write-Host "$($i + 1) : $($VSPath[$i].displayName) ($($VSPath[$i].installationPath))"
                }
                Write-Host ""
                while ($true) {
                    $Choice = Read-Host "C++ツールを追加したいインスタンスの番号を入力してください (1 - $($VSPath.Count))"
                    
                    # 入力値の検証
                    if ($Choice -match '^\d+$' -and [int]$Choice -ge 1 -and [int]$Choice -le $VSPath.Count) {
                        $SelectedIndex = [int]$Choice - 1
                        $SelectedInstance = $VSPath[$SelectedIndex]
                        break
                    } else {
                        Write-Host "無効な入力です。リストの番号を入力してください" -ForegroundColor Red
                    }
                }
            }
            else
            {
                $SelectedInstance = $VSPath[0]
                Write-Host "VisualStudioのインスタンスが1つしか無いため、自動的にインストール先を選択します"
            }
            Write-Host "$($SelectedInstance.displayName) ($($SelectedInstance.installationPath)が選択されました"
            exit-Q "C++用のツールのインストールを開始しますか？(y/n)" "インストールをキャンセルします"
            ChackRunningVS
            Start-Process $vsinstaller "modify" "--installPath" $SelectedInstance.installationPath "--add" $NativeDesktopWorkload "--force"
        }
    }
}
if($VSPath.Count -le 0)
{
    Write-Host "VisualStudioが検出できませんでした" -ForegroundColor Yellow
    Write-Host "VisualStudio Communityをインストールしてください"
    Write-Host "インストールを開始したら、このウィンドウに戻って操作を続けてください"
    exit-Q "VisualStudio CommunityとC++用のツールのインストールを開始しますか？(y/n)" "インストールをキャンセルします"
    if(!(Test-Path $VSWhere))
    {
        Write-Host "VisualStudio Installerが検出できないため、インターネットからダウンロードします"
        $VS_isDoenlaed = $true
        download ([ref]$VS_downloadPath) $VS_download
        $vsinstaller = $VS_downloadPath
    }
    ChackRunningVS
    Start-Process $vsinstaller -ArgumentList @("install", "--productId", $productId, "--channelUri", $channelUri, "--add", $NativeDesktopWorkload, "--includeRecommended", "--force")
}

Write-Host ""
if(Get-Command code -ea SilentlyContinue)
{
    Write-Host "VisualStudio Codeを検出しました"
}
else {
    Write-Host "VisualStudio Codeを検出できませんでした" 
    Write-Host "VisualStudio Codeをインストールしてください"
    Write-Host ""
    Write-Host "必ず「PATHへの追加」を有効にしてください" -ForegroundColor Yellow
    Write-Host "インストールが完了したら、このウィンドウへ戻って操作を続けてください" -ForegroundColor Yellow
    exit-Q "VisualStudio Codeをインストールしますか？ (y/n)"  "インストールをキャンセルします"

    $VSC_isDoenlaed = $true
    download ([ref]$VSC_downloadPath) $VSC_download
    &("./$VSC_downloadPath")
    Write-Host ""
    Write-Host "インストールの完了を待機しています" -ForegroundColor Yellow
    Write-Host "インストールを完了したにも関わらず次に進まない場合は、システムを再起動してもう一度この.ps1ファイルを実行してください" -ForegroundColor Yellow
    while(!(Get-Command code -ea SilentlyContinue))
    {
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }
    Write-Host "VisualStudio Codeを検出しました"
}
Write-Host ""
Write-Host "これからインストールするCode Runnerには、プログラムの実行時に自動的にファイルを保存する機能があります"
Write-Host "その機能を有効にしますか？"
Write-Host "1 : プログラムの実行時に、すべてのファイルを保存する"
Write-Host "2 : プログラムの実行時に、今開いているファイルのみ保存する"
Write-Host "3 : 自動保存機能を使用しない"
while ($true) {
    $Choice = Read-Host "どのオプションを使用しますか？ (1 - 3)"
    # 入力値の検証
    if ($Choice -match '^\d+$' -and [int]$Choice -ge 1 -and [int]$Choice -le 3) {
        break
    } else {
        Write-Host "無効な入力です。リストの番号を入力してください" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "これより、以下の変更をVisualStudio Codeへ適用します。" -ForegroundColor Yellow
Write-Host "・拡張機能の追加 C/C++ Extension Pack (ms-vscode.cpptools-extension-pack)"
Write-Host "・拡張機能の追加 Code Runner (formulahendry.code-runner)"
Write-Host "・統合ターミナルを起動時にchcp 65001を実行後、Developer PowerShell for VSが使えるように上書きする"
Write-Host "・C/C++をCode Runnerで実行するとき、VisualStudioのコンパイラ(cl.exe)でコンパイルするように上書きする"
Write-Host "・Code Runnerの「Run In Terminal」を有効化"
switch ($Choice) {
    "1"
    {
        Write-Host "・Code Runnerの「Save All File Before Run」を有効化"
    }
    "2"
    {
        Write-Host "・Code Runnerの「Save File Before Run」を有効化"
    }
}
Write-Host ""
exit-Q "変更を適用してもよろしいですか？(y/n)" "インストールをキャンセルします"
code --install-extension ms-vscode.cpptools-extension-pack
code --install-extension formulahendry.code-runner

$VSC_settings = [PSCustomObject]@{}
try
{
    $VSC_settings = ConvertFrom-Json (Get-Content -Path $VSC_settingsPath -Raw)
}
catch
{
    $VSC_settings = [PSCustomObject]@{}
}

Add-Member -InputObject $VSC_settings -Force "terminal.integrated.defaultProfile.windows" "PowerShell"
if($null -eq ($VSC_settings."terminal.integrated.profiles.windows"))
{
    Add-Member -InputObject $VSC_settings -Force "terminal.integrated.profiles.windows" ([PSCustomObject]@{})
}
if($null -eq ($VSC_settings."terminal.integrated.profiles.windows"."PowerShell"))
{
    Add-Member -InputObject ($VSC_settings."terminal.integrated.profiles.windows") -Force "PowerShell" ([PSCustomObject]@{})
}
Add-Member -InputObject $VSC_settings."terminal.integrated.profiles.windows"."PowerShell" -Force "source" "PowerShell"
Add-Member -InputObject $VSC_settings."terminal.integrated.profiles.windows"."PowerShell" -Force "args" ([string[]]$ps_settings)
switch ($Choice) {
    "1"
    {
        Add-Member -InputObject $VSC_settings -Force "code-runner.saveAllFilesBeforeRun" $true 
        Add-Member -InputObject $VSC_settings -Force "code-runner.saveFileBeforeRun" $false 
    }
    "2"
    {
        Add-Member -InputObject $VSC_settings -Force "code-runner.saveAllFilesBeforeRun" $false 
        Add-Member -InputObject $VSC_settings -Force "code-runner.saveFileBeforeRun" $true 
    }
}

Add-Member -InputObject $VSC_settings -Force "code-runner.runInTerminal" $true
$CodeRunner_cmd = [PSCustomObject]@{}
if ($null -ne $CodeRunner_cmd."code-runner.executorMap")
{
    $CodeRunner_cmd = $VSC_settings."code-runner.executorMap"
}
Add-Member -InputObject $CodeRunner_cmd "c" $codeRunner_c -Force 
Add-Member -InputObject $CodeRunner_cmd "cpp" $codeRunner_cpp -Force
Add-Member -InputObject $VSC_settings "code-runner.executorMap" $CodeRunner_cmd -Force
ConvertTo-Json -Depth 100 $VSC_settings | Set-Content -Path $VSC_settingsPath -Encoding UTF8

Write-Host ""
Write-Host "変更を適用しました"
if($VS_isDoenlaed -or $VSC_isDoenlaed)
{
    $VS_isDoenlaed = $true
    $VSC_isDoenlaed = $true
    $yn = $false
    [string[]]$downloaded_list = @()
    if($VS_isDoenlaed)
    {
        $downloaded_list += "VisualStudio"
    }
    if($VSC_isDoenlaed)
    {
        $downloaded_list += "VisualStudio Code"
    }
    Read-YN ([ref]$yn) "ダウンロードした$($downloaded_list -Join "/")のインストーラーを削除しますか？ (y/n)"
    if($yn)
    {
        if($VS_isDoenlaed)
        {
            Remove-Item $VS_downloadPath
        }
        if($VSC_isDoenlaed)
        {
            Remove-Item $VSC_downloadPath
        }
    }
}

Write-Host ""
Write-Host "インストールが完了しました" -ForegroundColor Yellow
Write-Host "設定等の反映のため、コンピュータを再起動することをおすすめします"
exitScript
