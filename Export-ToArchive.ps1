# Export-ToArchive.ps1
# Incremental exporter for Outlook (Classic) that saves mail categorized "ToArchive" to .msg files.
# Designed for a single user environment (Windows 11 / Classic Outlook).
# Author: Mike Harris
# Version: 0.0.2

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ================== CONFIG ==================
$CategoryToArchive       = "ToArchive"
$CategoryArchived        = "Archived"    # set to "" to disable auto-re-categorize
$AlsoRequireOlderThanDays = 0            # set to 0 to ignore age
$ExportRoot              = "$env:USERPROFILE\Documents\Outlook Archive"
# If your path differs, change $ExportRoot to an explicit path.

$IndexPath               = Join-Path $ExportRoot "_index.json"
$LogPath                 = Join-Path $ExportRoot "_export.log"
$ComputeHash             = $false          # set to $true to compute SHA256 of saved .msg (costly)
$FoldersToScan           = @("Inbox","Inbox\Resolved","Sent Items","Archive","Deleted Items")  # change to @() to scan ALL folders (slower)

# ================== HELPERS ==================
function Write-Log($msg) {
	# Write a message to the _export.log file with associated timestamp
    $line = "{0}  {1}" -f (Get-Date).ToString("s"), $msg
    Add-Content -Path $LogPath -Value $line
}

function Sanitize-FileName([string]$name) {
	# Remove invalid filename characters from string and trim whitespace
    if (-not $name) { return "(no-subject)" }
    $bad = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $bad) { 
		$name = $name.Replace($c, '-') 
	}
    $name = $name.Trim()
    if ($name.Length -eq 0) { return "(empty)" }
    return $name
}

function Load-Index {
	# Load the JSON data from the _index.json file
	# Returns an empty object if _index.json doesn't exist
	# Writes a log failure if JSON cannot be extracted
    if (Test-Path $IndexPath) {
        try {
            $raw = Get-Content $IndexPath -Raw
            if ($raw.Trim().Length -gt 0) { 
				return ($raw | ConvertFrom-Json) 
			}
        } catch {
            Write-Log ("WARN: Index load failed - recreating. {0}" -f $_)
        }
    }
    return @{}  # empty hashtable-like object
}

function Save-Index($idx) {
	# Save the index to JSON as _index.json
    $idx | ConvertTo-Json -Depth 5 | Set-Content -Path $IndexPath -Encoding UTF8
}

function Get-OutlookApp {
	# Return the Outlook Application if it is open
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        return $null
    }
}

function Get-UniqueKeyForMailItem($mail) {
	# Return the InternetMessageID or the EntryID for the item
    try { 
		$imid = $mail.InternetMessageID 
	} catch { 
		$imid = $null 
	}
    if ($imid -and $imid.Trim().Length -gt 0) {
        return "imid:" + $imid.Trim().ToLowerInvariant()
    }
    return "eid:" + $mail.EntryID
}

function Save-MailItemAsMsg($item, $path) {
    # olMSG Unicode = 3
    $item.SaveAs($path, 3)
}

function Compute-FileHash($path) {
	# compute SHA256 hash of item at $path
    if (-not (Test-Path $path)) { 
		return $null 
	}
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = [System.IO.File]::OpenRead($path)
    try {
        $hash = $sha.ComputeHash($fs)
        return [System.BitConverter]::ToString($hash) -replace '-', ''
    } finally {
        $fs.Close()
        $sha.Dispose()
    }
}

function Get-FolderByPath($ns, [string]$path) {
    # Supports paths like:
		# "Inbox"
		# "Inbox\Resolved"
		# "Archive\SomeSubfolder"
    $parts = $path -split "\\"
    if ($parts.Count -eq 0) { 
		return $null 
	}

    # First segment: allow default folders by common names
    $current = $null
    switch ($parts[0]) {
        "Inbox" { 
			$current = $ns.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox) 
		} "Sent Items" { 
			$current = $ns.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderSentMail) 
		} "Deleted Items" {
			$current = $ns.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderDeletedItems)
		} default {
            # Try as a top-level folder under the store root
            foreach ($f in $ns.DefaultStore.GetRootFolder().Folders) {
                if ($f.Name -eq $parts[0]) { 
					$current = $f 
					break 
				}
            }
        }
    }
	
    if (-not $current) { 
		return $null 
	}

    # Walk subfolders
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $next = $null
        foreach ($sf in $current.Folders) {
            if ($sf.Name -eq $parts[$i]) { 
				$next = $sf
				break 
			}
        }
        if (-not $next) { 
			return $null 
		}
        $current = $next
    }
	
    return $current
}

function Process-Folder($folder, $filter, [hashtable]$index) {
	# Traverse the folder and create an item to save for each "ToArchive" message
	
    try {
        $items = $folder.Items
        if (-not $items) { 
			return 0 
		}
        # Restrict is faster than iterating everything.
        $restricted = $items.Restrict($filter)
    } catch {
        Write-Log ("WARN: cannot restrict folder {0}: {1}" -f $folder.Name, $_)
        return 0
    }
	
    $count = 0
	# Iterate through the list of items in the folder
    for ($i = $restricted.Count; $i -ge 1; $i--) {
		$item = $restricted.Item($i)
        if (-not $item) { 
			continue 
		}
        if ($item.MessageClass -notlike "IPM.Note*") { 
			continue 
		}
		
		# Skip items that already have been archived
        $key = Get-UniqueKeyForMailItem $item
        if ($index.ContainsKey($key)) { 
			continue 
		}

		# Gather information for the item
        $dt = $item.ReceivedTime.ToString("yyyy-MM-dd")
        $subj = Sanitize-FileName($item.Subject)
        $sender = Sanitize-FileName($item.SenderName)
        $folderForDate = Join-Path $ExportRoot $dt # Create a folder for the date of the email/message item
        New-Item -ItemType Directory -Path $folderForDate -Force | Out-Null

        $base = "{0} - {1} - {2}" -f $dt, $sender, $subj
        if ($base.Length -gt 160) { 
			$base = $base.Substring(0,160) # Trim the information to 160 characters
		}
        $path = Join-Path $folderForDate ($base + ".msg")
        $n = 1
        while (Test-Path $path) {
            $path = Join-Path $folderForDate ($base + " ($n).msg")
            $n++
        }

		# Save the item
        try {
            Save-MailItemAsMsg $item $path
            if ($ComputeHash) {
                $h = Compute-FileHash $path
            } else { 
				$h = $null 
			}
            $index[$key] = @{
                ExportedAt = (Get-Date).ToString("s")
                FilePath   = $path
                Hash       = $h
            }
            $count++

            # Optionally, remove ToArchive and add Archived to avoid re-matching
            if ($CategoryArchived -and $CategoryArchived.Trim().Length -gt 0) {
                try {
                    $cats = $item.Categories
                    $catsArray = @()
                    if ($cats) { 
						$catsArray = ($cats -split ",\s*") | Where-Object { $_ -and $_ -ne $CategoryToArchive } 
					}
                    $catsArray += $CategoryArchived
                    $item.Categories = ($catsArray | Select-Object -Unique) -join ", "
                    $item.Save()
                } catch {
                    Write-Log ("WARN: failed to update categories for {0}: {1}" -f $key, $_)
                }
            }

        } catch {
            Write-Log ("ERROR: failed saving item '{0}' -> {1} : {2}" -f $item.Subject, $path, $_)
        }
    } # for
    return $count
}



# ================== MAIN ==================

# Ensure the output path is writable
New-Item -ItemType Directory -Path $ExportRoot -Force | Out-Null
if (-not (Test-Path $ExportRoot)) {
    Write-Error "Export root $ExportRoot not writable."
    exit 1
}

# Get the outlook application and close if Outlook is not open (designed behavior)
$ol = Get-OutlookApp
if (-not $ol) {
    Write-Log "Outlook not running; exiting (by design)."
    exit 0
}

# Rebuild table of archived messages from _index.json
$ns = $ol.GetNameSpace("MAPI")
$index = Load-Index
if ($index -is [System.Management.Automation.PSCustomObject]) {
    $ht = @{}
    foreach ($p in $index.PSObject.Properties) { 
		$ht[$p.Name] = $p.Value 
	}
    $index = $ht
} elseif ($null -eq $index) {
    $index = @{}
}


# Build a DASL filter for categories (Keywords) and optional date restriction
$catEsc = $CategoryToArchive.Replace("'", "''")
$sqlCat = '"' + "urn:schemas-microsoft-com:office:office#Keywords" + '"' + " LIKE '%$catEsc%'"

# Optional setting to only archive messages older than a specified age
if ($AlsoRequireOlderThanDays -gt 0) {
    $cutoff = (Get-Date).AddDays(-1 * $AlsoRequireOlderThanDays).ToString("yyyy-MM-ddTHH:mm:ss")
    $sqlDate = '"' + "urn:schemas:httpmail:datereceived" + '"' + " <= '$cutoff'"
    $filter = "@SQL=" + "(" + $sqlCat + ") AND (" + $sqlDate + ")"
	
# Default archiving behavior, independent of message age
} else {
    $filter = "@SQL=" + "(" + $sqlCat + ")"
}
Write-Log ("Filter: {0}" -f $filter)


# Decide which folders to scan
$totalExported = 0
if ($FoldersToScan.Count -eq 0) {
    # Walk entire store recursively (careful: can be slower).
    function Recurse-Folders([Microsoft.Office.Interop.Outlook.MAPIFolder]$fld) {
        $exported = Process-Folder $fld $filter $index
        $GLOBALS:totalExported += $exported
        foreach ($sub in $fld.Folders) { Recurse-Folders $sub }
    }
    $root = $ns.DefaultStore.GetRootFolder()
    Recurse-Folders $root
} else {
    foreach ($name in $FoldersToScan) {
		try {
			$found = Get-FolderByPath $ns $name
			if (-not $found) {
				Write-Log ("WARN: folder '{0}' not found; skipping." -f $name)
				continue
			}
			$exported = Process-Folder $found $filter $index
			$totalExported += $exported
		} catch {
			Write-Log ("WARN: error accessing folder {0}: {1}" -f $name, $_)
		}
	}
}

Save-Index $index
Write-Log ("Completed: exported {0} item(s)." -f $totalExported)
# Release COM objects to avoid Outlook locks
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
