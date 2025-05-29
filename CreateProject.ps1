# Microsoft.Graph PowerShellモジュールが必要です: Install-Module Microsoft.Graph -Scope CurrentUser
# TeamsドライブアイテムのインタラクティブなGraphエクスプローラー（プロジェクトJSON出力機能付き）

# 必要なモジュールのインポート
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Files

function Connect-Graph {
    try {
        Connect-MgGraph -Scopes "Team.ReadBasic.All","Files.Read.All"
    } catch {
        Write-Error "Graphの認証に失敗しました。Microsoft.Graphモジュールと権限が設定されていることを確認してください。"
        exit 1
    }
}

function Read-Selection {
    param([string]$Message, [int]$MaxIndex)
    while ($true) {
        $selection = Read-Host -Prompt $Message
        if ($selection -match '^[0-9]+$' -and [int]$selection -ge 1 -and [int]$selection -le $MaxIndex) {
            return [int]$selection
        } else {
            Write-Host "1から$MaxIndexまでの数字を入力してください。"
        }
    }
}

function Get-MyUserInfo {
    $me = Get-MgUser -UserId "me"
    if (-not $me) {
        Write-Error "ユーザー情報の取得に失敗しました。"
        exit 1
    }
    return $me
}

function Get-MyTeams {
    $teams = Get-MgUserJoinedTeam -UserId "me"
    if (-not $teams) { return @() }
    
    $indexed = @(); $i = 1
    foreach ($t in $teams) {
        Write-Host "[$i] $($t.Id) : $($t.DisplayName)"
        $indexed += $t; $i++
    }
    return $indexed
}

function Show-DriveItems {
    param([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDriveItem]$Parent)
    $items = if ($Parent) { Get-MgGroupDriveItemChild -GroupId $global:GroupId -DriveId $global:DriveId -ItemId $Parent.Id -All } else { Get-MgGroupDriveItemChild -GroupId $global:GroupId -DriveId $global:DriveId -All }
    $folders = $items | Where-Object { $null -ne $_.Folder }
    $files   = $items | Where-Object { $null -ne $_.File }
    if ($folders) {
        Write-Host "フォルダ:" -ForegroundColor Cyan
        $i = 1; foreach ($f in $folders) { Write-Host "  [$i] $($f.Id) : $($f.Name)"; $i++ }
    }
    if ($files) {
        Write-Host "ファイル:" -ForegroundColor Yellow
        foreach ($f in $files) { Write-Host "    $($f.Id) : $($f.Name)" }
    }
    return @{ Folders = $folders; Files = $files }
}

# 認証とチームの選択
Connect-Graph
$teams = Get-MyTeams; if (-not $teams) { Write-Error "チームが見つかりません。"; exit }
$sel = Read-Selection -Message "番号を入力してチームを選択してください" -MaxIndex $teams.Count
$team = $teams[$sel - 1]
$drive = Get-MgTeamDrive -TeamId $team.Id
$global:DriveId = $drive.Id
$global:GroupId = $team.Id

# ナビゲーションとXLSXファイルのURL収集
$stack = @($null); $lastXlsx = @()
while ($true) {
    $current = $stack[-1]
    $listing = Show-DriveItems -Parent $current
    $folders = $listing.Folders
    $selection = Read-Host -Prompt "フォルダ番号を入力して移動 / 'select'でこのフォルダに決定 / 'back'で戻る / 'exit'で終了"
    switch ($selection.ToLower()) {
        'back' { if ($stack.Count -gt 1) { $stack = $stack[0..($stack.Count-2)] } }
        'select' {
            $items = if ($current) { 
                Get-MgGroupDriveItemChild -GroupId $GroupId -DriveId $DriveId -ItemId $current.Id -All 
            } else { 
                Get-MgGroupDriveItemChild -GroupId $GroupId -DriveId $DriveId -All 
            }
            $xlsx = $items | Where-Object { $_.File -and $_.Name -like '*.xlsx' }
            if (-not $xlsx) { Write-Host "Excelファイル（.xlsx）が見つかりません。"; continue }
            $lastXlsx = @()
            Write-Host "見つかったExcelファイル:" -ForegroundColor Green
            foreach ($f in $xlsx) {
                # $url = (Get-MgGroupDriveItem -GroupId $GroupId -DriveId $DriveId -ItemId $f.Id -Select '@microsoft.graph.downloadUrl')."@microsoft.graph.downloadUrl"
                Write-Host "$($f.Name) : $($f.Id)"
                $lastXlsx += @{ type = 'sharepoint'; identifier = $f.Name; path = $f.Id }
            }
        }
        'exit' { break }
        default {
            if ($selection -match '^[0-9]+$') {
                $idx = [int]$selection
                if ($idx -ge 1 -and $idx -le $folders.Count) { $stack += $folders[$idx-1] } else { Write-Host "無効なフォルダ番号です。" }
            } else { Write-Host "不明なコマンドです。" }
        }
    }
}

# プロジェクトJSONの作成
while ($true) {
    $yn = Read-Host -Prompt "このリストからプロジェクトファイルを作成しますか？ (y/n)"
    if ($yn -eq 'y') {
        $projName = "NewProject_$(Get-Date -Format yyyy-MM-dd)"
        $project = @{ project = @{ project_name = $projName; files = $lastXlsx } }
        $json = $project | ConvertTo-Json -Depth 5
        if (-not (Test-Path 'projects')) { New-Item -ItemType Directory -Path 'projects' | Out-Null }
        $filePath = "projects\$projName.json"
        $json | Out-File -FilePath $filePath -Encoding UTF8
        Write-Host "プロジェクトを $filePath に保存しました"
        break
    } elseif ($yn -eq 'n') {
        continue
    } else {
        Write-Host "'y'または'n'を入力してください。"
    }
}
