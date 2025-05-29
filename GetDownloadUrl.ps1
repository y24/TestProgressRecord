# Microsoft.Graph PowerShellモジュールが必要です: Install-Module Microsoft.Graph -Scope CurrentUser
# SharePointファイルのダウンロードURLを取得するスクリプト

param(
    [Parameter(Mandatory=$true)]
    [string]$ItemId
)

# 必要なモジュールのインポート
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Files

# Graph APIに接続
try {
    Connect-MgGraph -Scopes "Team.ReadBasic.All","Files.Read.All"
} catch {
    Write-Error "Graphの認証に失敗しました。Microsoft.Graphモジュールと権限が設定されていることを確認してください。"
    exit 1
}

# ファイルのダウンロードURLを取得
try {
    # ファイルが所属するチームとドライブの情報を取得
    $driveItem = Get-MgDriveItem -DriveItem $ItemId
    if (-not $driveItem) {
        Write-Error "指定されたItemIdのファイルが見つかりません。"
        exit 1
    }

    # ダウンロードURLを取得して出力
    $downloadUrl = (Get-MgDriveItem -DriveId $driveItem.ParentReference.DriveId -ItemId $ItemId -Select '@microsoft.graph.downloadUrl').'@microsoft.graph.downloadUrl'
    Write-Output $downloadUrl

} catch {
    Write-Error "ダウンロードURLの取得に失敗しました: $_"
    exit 1
} 