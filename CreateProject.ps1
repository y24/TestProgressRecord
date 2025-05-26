# Requires Microsoft.Graph PowerShell SDK: Install-Module Microsoft.Graph -Scope CurrentUser
# Interactive Graph Explorer for Teams Drive Items

# Connect and authenticate
function Connect-Graph {
    try {
        Connect-MgGraph -Scopes "Group.Read.All","Files.Read.All","Team.ReadBasic.All"
    } catch {
        Write-Error "Graph authentication failed. Please ensure you have the Microsoft.Graph module and proper permissions."
        exit 1
    }
}

# Prompt user and return selection index
function Prompt-Selection {
    param(
        [string]$Message,
        [int]$MaxIndex
    )
    while ($true) {
        $input = Read-Host -Prompt $Message
        if ($input -match '^[0-9]+$' -and [int]$input -ge 1 -and [int]$input -le $MaxIndex) {
            return [int]$input
        } else {
            Write-Host "Please enter a number between 1 and $MaxIndex."
        }
    }
}

# List teams the user is a member of
function Get-MyTeams {
    $teams = Get-MgUserTeam -All
    $indexed = @()
    $i = 1
    foreach ($t in $teams) {
        Write-Host "[$i] $($t.Id) : $($t.DisplayName)"
        $indexed += $t
        $i++
    }
    return $indexed
}

# List drive items (folders & files) for a given drive or folder
function Show-DriveItems {
    param(
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDriveItem]$ParentItem
    )
    # If no parent specified, root of drive
    if (-not $ParentItem) {
        $items = get-mgdriveitemchild -DriveId $global:DriveId -All
    } else {
        $items = get-mgdriveitemchild -DriveId $global:DriveId -ItemId $ParentItem.Id -All
    }
    # Separate folders and files
    $folders = $items | Where-Object { $_.Folder -ne $null }
    $files   = $items | Where-Object { $_.File   -ne $null }

    # Display folders with index
    $i = 1
    if ($folders) {
        Write-Host "Folders:" -ForegroundColor Cyan
        foreach ($f in $folders) {
            Write-Host "  [$i] $($f.Id) : $($f.Name)"
            $i++
        }
    }
    # Display files
    if ($files) {
        Write-Host "Files:" -ForegroundColor Yellow
        foreach ($f in $files) {
            Write-Host "    $($f.Id) : $($f.Name)"
        }
    }
    return @{ Folders = $folders; Files = $files }
}

# Main interactive loop
Connect-Graph
# Team selection
$teams = Get-MyTeams
if (-not $teams) { Write-Error "No teams found."; exit }
$selIndex = Prompt-Selection -Message "Select a team by number" -MaxIndex $teams.Count
$team = $teams[$selIndex - 1]

# Get drive for selected team
try {
    $drive = Get-MgTeamDrive -TeamId $team.Id
    $global:DriveId = $drive.Id
} catch {
    Write-Error "Failed to get drive for team $($team.DisplayName)."
    exit
}

# Navigation stack
$stack = @($null)  # root represented by $null

while ($true) {
    $current = $stack[-1]
    $listing = Show-DriveItems -ParentItem $current
    $folders = $listing.Folders

    $input = Read-Host -Prompt "Enter folder number to navigate, 'back', 'get', or 'exit'"
    switch ($input.ToLower()) {
        'back' {
            if ($stack.Count -gt 1) { $stack = $stack[0..($stack.Count-2)] } # pop
        }
        'get' {
            # Filter xlsx files in current folder
            $items = if ($current) { Get-MgDriveItemChild -DriveId $DriveId -ItemId $current.Id -All } else { Get-MgDriveItemChild -DriveId $DriveId -All }
            $xlsx = $items | Where-Object { $_.File -and $_.Name -like '*.xlsx' }
            if (-not $xlsx) { Write-Host "No .xlsx files found."; continue }
            Write-Host "Downloading URLs for .xlsx files:" -ForegroundColor Green
            foreach ($f in $xlsx) {
                # Get download URL
                $url = (Get-MgDriveItem -DriveId $DriveId -ItemId $f.Id -Select '@microsoft.graph.downloadUrl')."@microsoft.graph.downloadUrl"
                Write-Host "$($f.Name) : $url"
            }
        }
        'exit' { break }
        default {
            if ($input -match '^[0-9]+$') {
                $idx = [int]$input
                if ($idx -ge 1 -and $idx -le $folders.Count) {
                    $stack += $folders[$idx-1]
                } else {
                    Write-Host "Invalid folder number.";
                }
            } else {
                Write-Host "Unknown command.";
            }
        }
    }
}

# Project creation
while ($true) {
    $yn = Read-Host -Prompt "Create project file from this list? (y/n)"
    if ($yn -eq 'y') {
        $project = @{
            TeamId = $team.Id;
            TeamName = $team.DisplayName;
            Path = $stack | ForEach-Object { if ($_) { @{ Id = $_.Id; Name = $_.Name } } }
        }
        $json = $project | ConvertTo-Json -Depth 5
        $folder = "projects"
        if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder }
        $fileName = "NewProject_$(Get-Date -Format yyyy-MM-dd).json"
        $path = Join-Path $folder $fileName
        $json | Out-File -FilePath $path -Encoding UTF8
        Write-Host "Project saved to $path"
        break
    } elseif ($yn -eq 'n') {
        # redisplay current folder
        continue
    } else {
        Write-Host "Please enter 'y' or 'n'."
    }
}
