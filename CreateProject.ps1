# Requires Microsoft.Graph PowerShell SDK: Install-Module Microsoft.Graph -Scope CurrentUser
# Interactive Graph Explorer for Teams Drive Items with Project JSON Output

function Connect-Graph {
    try {
        Connect-MgGraph -Scopes "Group.Read.All","Files.Read.All","Team.ReadBasic.All"
    } catch {
        Write-Error "Graph authentication failed. Ensure Microsoft.Graph module and permissions are set."
        exit 1
    }
}

function Prompt-Selection {
    param([string]$Message, [int]$MaxIndex)
    while ($true) {
        $input = Read-Host -Prompt $Message
        if ($input -match '^[0-9]+$' -and [int]$input -ge 1 -and [int]$input -le $MaxIndex) {
            return [int]$input
        } else {
            Write-Host "Please enter a number between 1 and $MaxIndex."
        }
    }
}

function Get-MyTeams {
    $teams = Get-MgUserTeam -All
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
    $items = if ($Parent) { Get-MgDriveItemChild -DriveId $global:DriveId -ItemId $Parent.Id -All } else { Get-MgDriveItemChild -DriveId $global:DriveId -All }
    $folders = $items | Where-Object { $_.Folder -ne $null }
    $files   = $items | Where-Object { $_.File   -ne $null }
    if ($folders) {
        Write-Host "Folders:" -ForegroundColor Cyan
        $i = 1; foreach ($f in $folders) { Write-Host "  [$i] $($f.Id) : $($f.Name)"; $i++ }
    }
    if ($files) {
        Write-Host "Files:" -ForegroundColor Yellow
        foreach ($f in $files) { Write-Host "    $($f.Id) : $($f.Name)" }
    }
    return @{ Folders = $folders; Files = $files }
}

# Authenticate and select Team
Connect-Graph
$teams = Get-MyTeams; if (-not $teams) { Write-Error "No teams found."; exit }
$sel = Prompt-Selection -Message "Select a team by number" -MaxIndex $teams.Count
$team = $teams[$sel - 1]
$drive = Get-MgTeamDrive -TeamId $team.Id; $global:DriveId = $drive.Id

# Navigation and collecting XLSX URLs
$stack = @($null); $lastXlsx = @()
while ($true) {
    $current = $stack[-1]
    $listing = Show-DriveItems -Parent $current
    $folders = $listing.Folders
    $input = Read-Host -Prompt "Enter folder number to navigate, 'back', 'get', or 'exit'"
    switch ($input.ToLower()) {
        'back' { if ($stack.Count -gt 1) { $stack = $stack[0..($stack.Count-2)] } }
        'get' {
            $items = if ($current) { Get-MgDriveItemChild -DriveId $DriveId -ItemId $current.Id -All } else { Get-MgDriveItemChild -DriveId $DriveId -All }
            $xlsx = $items | Where-Object { $_.File -and $_.Name -like '*.xlsx' }
            if (-not $xlsx) { Write-Host "No .xlsx files found."; continue }
            $lastXlsx = @()
            Write-Host "Found .xlsx files:" -ForegroundColor Green
            foreach ($f in $xlsx) {
                $url = (Get-MgDriveItem -DriveId $DriveId -ItemId $f.Id -Select '@microsoft.graph.downloadUrl')."@microsoft.graph.downloadUrl"
                Write-Host "$($f.Name) : $url"
                $lastXlsx += @{ type = 'sharepoint'; identifier = $f.Name; path = $url }
            }
        }
        'exit' { break }
        default {
            if ($input -match '^[0-9]+$') {
                $idx = [int]$input
                if ($idx -ge 1 -and $idx -le $folders.Count) { $stack += $folders[$idx-1] } else { Write-Host "Invalid folder number." }
            } else { Write-Host "Unknown command." }
        }
    }
}

# Project JSON creation
while ($true) {
    $yn = Read-Host -Prompt "Create project file from this list? (y/n)"
    if ($yn -eq 'y') {
        $projName = "NewProject_$(Get-Date -Format yyyy-MM-dd)"
        $project = @{ project = @{ project_name = $projName; files = $lastXlsx } }
        $json = $project | ConvertTo-Json -Depth 5
        if (-not (Test-Path 'projects')) { New-Item -ItemType Directory -Path 'projects' | Out-Null }
        $filePath = "projects\$projName.json"
        $json | Out-File -FilePath $filePath -Encoding UTF8
        Write-Host "Project saved to $filePath"
        break
    } elseif ($yn -eq 'n') {
        continue
    } else {
        Write-Host "Please enter 'y' or 'n'."
    }
}
