<#
.SYNOPSIS
    Command-line add-on manager for World of Warcraft: Forever (add-ons from CurseForge).

.DESCRIPTION
    Installs, updates, lists and removes CurseForge add-ons for the World of Warcraft: Forever
    client, i.e. the "_forever_" folder inside the main World of Warcraft folder.

    Everything is contained in this single file and no CurseForge API key is needed:
      * api.cfwidget.com                           resolves add-on URLs to CurseForge project IDs
      * api.curse.tools                            fallback for URLs that CFWidget does not know
                                                   (public proxy of the official CurseForge API)
      * www.curseforge.com/api/v1/mods/<id>/...    WoW Forever file lists (game flavor 88568),
                                                   required dependencies and file downloads
      * mediafilez.forgecdn.net / edge.forgecdn.net  CDN fallback for downloads

    State (WoW folder + managed add-ons) is stored as JSON in
      %APPDATA%\ForeverAddonManager\forever-addon-manager.json   (change with -StateFile)

    Commands:
      set-path [folder]     Set the World of Warcraft folder (auto-detects when omitted)
      config                Show the current configuration
      add <url> [...]       Install add-on(s) from CurseForge URLs
      remove <addon> [...]  Uninstall add-on(s): files and saved settings
      list                  List installed add-ons and their update status
      update [addon ...]    Update all add-ons, or only the given ones
      help                  Show the help screen

    <addon> can be the number shown by 'list', the CurseForge URL, the slug (e.g. eqol),
    the project ID or (part of) the add-on name.

    Set the environment variable FAM_COLOR to ansi, console or none to force a color mode.

.PARAMETER Command
    The command to run: help, set-path, config, add, remove, list or update.

.PARAMETER Arguments
    Arguments of the command: CurseForge URLs (add), add-ons (remove, update) or a folder (set-path).

.PARAMETER Channel
    add/update: release channel to follow - release (default), beta or alpha.
    Switching the channel with update installs the newest file of the new channel right away.

.PARAMETER Force
    add/update: reinstall even when up to date. remove: do not ask for confirmation.
    set-path: save the folder even if no WoW Forever installation is found in it.

.PARAMETER KeepSettings
    remove: keep the add-on's saved settings (WTF\...\SavedVariables).

.PARAMETER SkipDependencies
    add: do not install required dependencies automatically.

.PARAMETER Offline
    list: do not contact CurseForge, show the status of the last check.

.PARAMETER NoColor
    Disable colored output.

.PARAMETER StateFile
    Use another state file than the default one.

.NOTES
    WoW Forever Add-on Manager by phunqe.

    If Windows blocks the script ("running scripts is disabled on this system"), either run
      Unblock-File .\forever-addon-manager.ps1   (for a downloaded file) and/or
      Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
    or start it with: powershell -ExecutionPolicy Bypass -File .\forever-addon-manager.ps1 <command>

    Works with Windows PowerShell 5.1 and PowerShell 7+. Restart WoW after installing add-ons.

.EXAMPLE
    .\forever-addon-manager.ps1 set-path "C:\Program Files (x86)\World of Warcraft"

.EXAMPLE
    .\forever-addon-manager.ps1 add https://www.curseforge.com/wow/addons/eqol

.EXAMPLE
    .\forever-addon-manager.ps1 list

.EXAMPLE
    .\forever-addon-manager.ps1 update

.EXAMPLE
    .\forever-addon-manager.ps1 update eqol -Channel beta

.EXAMPLE
    .\forever-addon-manager.ps1 remove eqol
#>
#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string] $Command = 'help',

    [Parameter(Position = 1, ValueFromRemainingArguments = $true)]
    [string[]] $Arguments = @(),

    [ValidateSet('release', 'beta', 'alpha')]
    [string] $Channel,

    [switch] $Force,
    [switch] $KeepSettings,
    [switch] $SkipDependencies,
    [switch] $Offline,
    [switch] $NoColor,
    [string] $StateFile
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

#region Constants ---------------------------------------------------------------------------------

$script:AppName         = 'WoW Forever Add-on Manager'
$script:AppVersion      = '1.0.0'
$script:AppAuthor       = 'phunqe'
$script:ScriptName      = '.\forever-addon-manager.ps1'
$script:FlavorDir       = '_forever_'
$script:FlavorTypeId    = 88568   # CurseForge game flavor / gameVersionTypeId of "WoW Forever"
$script:CfSite          = 'https://www.curseforge.com'
$script:CfWidgetApi     = 'https://api.cfwidget.com'
$script:CfApiProxy      = 'https://api.curse.tools/v1/cf'   # public proxy of the official CurseForge API
$script:CfPageSize      = 50
$script:UserAgent       = "ForeverAddonManager/$($script:AppVersion) (PowerShell $($PSVersionTable.PSVersion))"
$script:ChannelRank     = @{ release = 1; beta = 2; alpha = 3 }
$script:ReleaseTypeName = @{ 1 = 'release'; 2 = 'beta'; 3 = 'alpha' }
$script:Http            = $null
$script:ZipReady        = $false
$script:HadErrors       = $false
$script:InstalledCount  = 0
$script:InstanceLock    = $null

# Glyphs are built from code points so the file stays pure ASCII (Windows PowerShell 5.1 reads
# BOM-less scripts as ANSI). All of them are part of the WGL4 set, i.e. Consolas/Cascadia safe.
# Plain ASCII replacements are used when output is redirected to a non-Unicode encoding.
function Set-GlyphSet {
    param([switch] $Ascii)
    if ($Ascii) {
        $script:Glyph = @{
            Diamond = '*'; Dot = '*'; Ring = 'o'; Up = '^'; Right = '->'; Cross = 'x'; Check = '+'
            Ellipsis = '~'; Line = '-'; Arrow = '>'; Middot = '|'; Dash = '-'
        }
        $unknownGlyph = '-'
    }
    else {
        $script:Glyph = @{
            Diamond  = [string][char]0x2666
            Dot      = [string][char]0x25CF
            Ring     = [string][char]0x25CB
            Up       = [string][char]0x2191
            Right    = [string][char]0x2192
            Cross    = [string][char]0x00D7
            Check    = [string][char]0x221A
            Ellipsis = [string][char]0x2026
            Line     = [string][char]0x2500
            Arrow    = [string][char]0x00BB
            Middot   = [string][char]0x00B7
            Dash     = [string][char]0x2014
        }
        $unknownGlyph = $script:Glyph.Middot
    }
    $script:StatusInfo = @{
        ok        = @{ Glyph = $script:Glyph.Dot;   Label = 'Up to date';       Style = 'Ok' }
        update    = @{ Glyph = $script:Glyph.Up;    Label = 'Update available'; Style = 'Update' }
        missing   = @{ Glyph = $script:Glyph.Cross; Label = 'Files missing';    Style = 'Error' }
        noforever = @{ Glyph = $script:Glyph.Ring;  Label = 'No Forever build'; Style = 'Warn' }
        error     = @{ Glyph = '!';                 Label = 'Check failed';     Style = 'Error' }
        unknown   = @{ Glyph = $unknownGlyph;       Label = 'Not checked';      Style = 'Dim' }
    }
}
Set-GlyphSet

# Dark theme. Each style: 24-bit color for VT terminals, ConsoleColor fallback, bold.
$script:Palette = @{
    Brand   = @('FFD100', 'Yellow',     $true)    # WoW UI gold
    Sub     = @('8B98A9', 'Gray',       $false)
    Head    = @('7AA7DC', 'DarkCyan',   $true)
    Rule    = @('3A4350', 'DarkGray',   $false)
    Dim     = @('6E7A8A', 'DarkGray',   $false)
    Text    = @('C9D1D9', 'Gray',       $false)
    Name    = @('ECF0F4', 'White',      $false)
    Version = @('AEB9C6', 'Gray',       $false)
    Newer   = @('FFD100', 'Yellow',     $true)
    Ok      = @('3FB950', 'Green',      $false)
    Update  = @('FFC940', 'Yellow',     $true)
    Error   = @('F85149', 'Red',        $false)
    Warn    = @('FF9F1C', 'DarkYellow', $false)   # legendary orange
    Info    = @('58A6FF', 'Cyan',       $false)
    Release = @('8B98A9', 'Gray',       $false)
    Beta    = @('4DA3FF', 'Cyan',       $false)   # rare blue, brightened for dark backgrounds
    Alpha   = @('B862FF', 'Magenta',    $false)   # epic purple, brightened for dark backgrounds
    Key     = @('7AA7DC', 'DarkCyan',   $false)
    Cmd     = @('FFD100', 'Yellow',     $false)
}

#endregion

#region Console output ----------------------------------------------------------------------------

function Initialize-Console {
    $mode = 'console'
    $override = ([string]$env:FAM_COLOR).ToLowerInvariant()
    if ($script:NoColor -or $env:NO_COLOR) {
        $mode = 'none'
    }
    elseif ($override -in @('ansi', 'console', 'none')) {
        $mode = $override
    }
    elseif ($Host.UI.SupportsVirtualTerminal) {
        $redirected = $false
        try { $redirected = [Console]::IsOutputRedirected } catch { }
        $psStyle = Get-Variable -Name PSStyle -ValueOnly -ErrorAction SilentlyContinue
        if ($redirected -or ($psStyle -and [string]$psStyle.OutputRendering -eq 'PlainText')) { $mode = 'none' }
        else { $mode = 'ansi' }
    }
    $script:ColorMode = $mode
    $redirectedOutput = $false
    $codePage = 65001
    try { $redirectedOutput = [Console]::IsOutputRedirected; $codePage = [Console]::OutputEncoding.CodePage } catch { }
    if ($redirectedOutput -and $codePage -notin @(65001, 1200, 1201, 12000)) { Set-GlyphSet -Ascii }
    $script:AnsiReset = "$([char]27)[0m"
    $script:AnsiCodes = @{}
    foreach ($name in @($script:Palette.Keys)) {
        $p = $script:Palette[$name]
        $r = [Convert]::ToInt32($p[0].Substring(0, 2), 16)
        $g = [Convert]::ToInt32($p[0].Substring(2, 2), 16)
        $b = [Convert]::ToInt32($p[0].Substring(4, 2), 16)
        $bold = ''
        if ($p[2]) { $bold = '1;' }
        $script:AnsiCodes[$name] = "$([char]27)[$($bold)38;2;$($r);$($g);$($b)m"
    }
}

# Writes text segments: Write-Styled @('text', 'Style', 'more text', 'OtherStyle', ...)
function Write-Styled {
    param([object[]] $Parts = @(), [switch] $NoNewline)
    if ($script:ColorMode -eq 'console') {
        for ($i = 0; $i -lt $Parts.Count; $i += 2) {
            $text = [string]$Parts[$i]
            if ($text.Length -eq 0) { continue }
            $style = ''
            if ($i + 1 -lt $Parts.Count) { $style = [string]$Parts[$i + 1] }
            if ($style -and $script:Palette.ContainsKey($style)) {
                Write-Host $text -NoNewline -ForegroundColor $script:Palette[$style][1]
            }
            else {
                Write-Host $text -NoNewline
            }
        }
        if (-not $NoNewline) { Write-Host '' }
        return
    }
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Parts.Count; $i += 2) {
        $text = [string]$Parts[$i]
        if ($text.Length -eq 0) { continue }
        $style = ''
        if ($i + 1 -lt $Parts.Count) { $style = [string]$Parts[$i + 1] }
        if ($script:ColorMode -eq 'ansi' -and $style -and $script:AnsiCodes.ContainsKey($style)) {
            [void]$sb.Append($script:AnsiCodes[$style]).Append($text).Append($script:AnsiReset)
        }
        else {
            [void]$sb.Append($text)
        }
    }
    Write-Host $sb.ToString() -NoNewline:$NoNewline
}

function Write-Msg {
    param([string] $Level, [string] $Text)
    switch ($Level) {
        'Ok'    { Write-Styled @("  $($script:Glyph.Check) ", 'Ok', $Text, 'Text') }
        'Info'  { Write-Styled @("  $($script:Glyph.Arrow) ", 'Info', $Text, 'Text') }
        'Step'  { Write-Styled @("  $($script:Glyph.Arrow) ", 'Dim', $Text, 'Dim') }
        'Warn'  { Write-Styled @('  ! ', 'Warn', $Text, 'Warn') }
        'Error' { Write-Styled @("  $($script:Glyph.Cross) ", 'Error', $Text, 'Error') }
        default { Write-Styled @("    $Text", 'Dim') }
    }
}

function Write-Banner {
    param([string] $Subtitle, [switch] $Credit)
    $creditText = ''
    $authorText = ''
    if ($Credit) { $creditText = "  $($script:Glyph.Middot)  by "; $authorText = $script:AppAuthor }
    Write-Host ''
    Write-Styled @("  $($script:Glyph.Diamond) ", 'Brand', $script:AppName, 'Brand', "  v$($script:AppVersion)", 'Dim', $creditText, 'Dim', $authorText, 'Text')
    if ($Subtitle) { Write-Styled @("    $Subtitle", 'Sub') }
    Write-Host ''
}

function Write-KeyValue {
    param([string] $Key, [string] $Value, [string] $ValueStyle = 'Text', [string] $Note = '', [string] $NoteStyle = 'Dim')
    $noteText = ''
    if ($Note) { $noteText = "  $Note" }
    Write-Styled @(('    ' + $Key.PadRight(20)), 'Key', $Value, $ValueStyle, $noteText, $NoteStyle)
}

function Get-ConsoleWidth {
    $width = 0
    try { $width = [int]$Host.UI.RawUI.WindowSize.Width } catch { }
    if ($width -le 0) { try { $width = [Console]::WindowWidth } catch { } }
    if ($width -le 0) { $width = 120 }
    if ($width -lt 40) { $width = 40 }
    return $width
}

function Limit-Text {
    param([string] $Text, [int] $Width)
    if ($Width -le 0) { return '' }
    if ($Text.Length -le $Width) { return $Text }
    if ($Width -eq 1) { return $script:Glyph.Ellipsis }
    return $Text.Substring(0, $Width - 1) + $script:Glyph.Ellipsis
}

#endregion

#region Utilities --------------------------------------------------------------------------------

function Get-ErrorMessage {
    param($ErrorObject)
    $ex = $ErrorObject
    if ($ex -is [System.Management.Automation.ErrorRecord]) { $ex = $ex.Exception }
    if ($null -eq $ex) { return 'Unknown error' }
    if ($ex -isnot [System.Exception]) { return [string]$ex }
    while ($null -ne $ex.InnerException) { $ex = $ex.InnerException }
    if ($ex -is [System.OperationCanceledException]) { return 'The request timed out.' }
    return $ex.Message
}

function ConvertTo-UtcDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    $text = [string]$Value
    if (-not $text) { return $null }
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) { return $parsed }
    return $null
}

function ConvertTo-IsoDate {
    param($Value)
    $date = ConvertTo-UtcDate $Value
    if ($null -eq $date) { return $null }
    return $date.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'", [Globalization.CultureInfo]::InvariantCulture)
}

function Format-ShortDate {
    param($Value)
    $date = ConvertTo-UtcDate $Value
    if ($null -eq $date) { return '' }
    return $date.ToLocalTime().ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

# Returns a [string[]] without empty entries. The leading comma keeps PowerShell from unrolling it.
function ConvertTo-StringList {
    param($Value)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($Value)) {
        if ($null -ne $item -and [string]$item -ne '') { $list.Add([string]$item) }
    }
    return , $list.ToArray()
}

function Format-Size {
    param([long] $Bytes)
    if ($Bytes -le 0) { return 'unknown size' }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return "$Bytes bytes"
}

# Add-on folder names must be plain names directly inside Interface\AddOns. Windows strips trailing
# dots and spaces, so e.g. "..." would point at the AddOns folder itself.
function Test-SafeFolderName {
    param([string] $Name)
    if (-not $Name -or $Name -eq '.' -or $Name -eq '..' -or $Name.Trim() -ne $Name -or $Name.EndsWith('.')) { return $false }
    if ($Name.Contains('\') -or $Name.Contains('/') -or $Name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { return $false }
    return $true
}

function Get-AddonFolderPath {
    param([string] $AddOnsDir, [string] $Name)
    if (-not (Test-SafeFolderName $Name)) { throw "Refusing to touch the unsafe add-on folder name '$Name'." }
    return [IO.Path]::Combine($AddOnsDir, $Name)
}

function Format-FolderList {
    param($Folders, [int] $Max = 4)
    $all = ConvertTo-StringList $Folders
    if ($all.Count -eq 0) { return '-' }
    if ($all.Count -le $Max) { return ($all -join ', ') }
    return (($all[0..($Max - 1)] -join ', ') + " +$($all.Count - $Max) more")
}

# Picks a compact version out of a CurseForge display name, e.g. "DBM-Core-12.0.5" -> "12.0.5".
function Get-ShortVersion {
    param([string] $DisplayName, [string] $FileName)
    $text = $DisplayName
    if (-not $text) { $text = $FileName }
    if (-not $text) { return '' }
    $text = ($text.Trim() -replace '\.zip$', '')
    $match = [regex]::Match($text, '(?<![0-9A-Za-z])[vV]?\d+(?:\.\d+)+(?:[-+_][0-9A-Za-z][0-9A-Za-z.\-+]*)?')
    if ($match.Success -and $match.Index -gt 0) { return $match.Value }
    return $text
}

#endregion

#region WoW installation ---------------------------------------------------------------------------

function Get-FlavorPath    { param($State) return [IO.Path]::Combine([string]$State.wowPath, $script:FlavorDir) }
function Get-AddOnsPath    { param($State) return [IO.Path]::Combine([string]$State.wowPath, $script:FlavorDir, 'Interface', 'AddOns') }
function Get-WtfPath       { param($State) return [IO.Path]::Combine([string]$State.wowPath, $script:FlavorDir, 'WTF') }

function Test-WowRootFolder {
    param([string] $Path)
    if (-not $Path) { return $false }
    return ((Test-Path -LiteralPath ([IO.Path]::Combine($Path, 'Data')) -PathType Container) -or
            (Test-Path -LiteralPath ([IO.Path]::Combine($Path, '.build.info')) -PathType Leaf))
}

# Accepts the main World of Warcraft folder, the _forever_ folder (or another client folder such as
# _retail_), a folder below it (e.g. Interface\AddOns) or Wow.exe, and returns the main folder.
function Resolve-WowRoot {
    param([string] $Path)
    $p = ([string]$Path).Trim().Trim('"', "'").Trim()
    if (-not $p) { throw 'No folder was given.' }
    $p = [Environment]::ExpandEnvironmentVariables($p)
    try { $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }
    catch { $full = [IO.Path]::GetFullPath($p) }   # e.g. a drive that is not connected
    if ($full.Length -gt 3) { $full = $full.TrimEnd('\', '/') }
    if (Test-Path -LiteralPath $full -PathType Leaf) { $full = Split-Path -Parent $full }
    if (Test-Path -LiteralPath ([IO.Path]::Combine($full, $script:FlavorDir)) -PathType Container) { return $full }
    $probe = $full
    for ($level = 0; $level -lt 3 -and $probe; $level++) {
        $leaf = Split-Path -Leaf $probe
        $parent = Split-Path -Parent $probe
        if ($parent -and $leaf -match '^_[A-Za-z0-9]+(_[A-Za-z0-9]+)*_$' -and
            ($leaf -ieq $script:FlavorDir -or (Test-WowRootFolder $parent))) {
            return $parent
        }
        if (-not $parent -or $parent -eq $probe) { break }
        $probe = $parent
    }
    return $full
}

function Find-WowInstallation {
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($key in @('HKLM:\SOFTWARE\WOW6432Node\Blizzard Entertainment\World of Warcraft',
                       'HKLM:\SOFTWARE\Blizzard Entertainment\World of Warcraft')) {
        try {
            $installPath = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).InstallPath
            if ($installPath) { $candidates.Add([string]$installPath) }
        }
        catch { }
    }
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, 'C:\', 'D:\', 'E:\', 'F:\')) {
        if (-not $base) { continue }
        foreach ($relative in @('World of Warcraft', 'Games\World of Warcraft', 'Blizzard\World of Warcraft')) {
            $candidates.Add([IO.Path]::Combine($base, $relative))
        }
    }
    $candidates.Add('/Applications/World of Warcraft')
    foreach ($candidate in $candidates) {
        try {
            if (-not (Test-Path -LiteralPath $candidate)) { continue }
            $root = Resolve-WowRoot $candidate
            if (Test-Path -LiteralPath ([IO.Path]::Combine($root, $script:FlavorDir)) -PathType Container) { return $root }
        }
        catch { }
    }
    return $null
}

function Assert-WowPath {
    param($State)
    if (-not $State.wowPath) {
        throw "The World of Warcraft folder is not set yet. Run:  $($script:ScriptName) set-path ""C:\Program Files (x86)\World of Warcraft""  (or 'set-path' without a folder to auto-detect it)"
    }
    $flavorPath = Get-FlavorPath $State
    if (-not (Test-Path -LiteralPath $flavorPath -PathType Container)) {
        throw "The WoW Forever folder was not found: $flavorPath  - check it with 'config' or change it with 'set-path'."
    }
}

function Test-WowRunning {
    param($State)
    if (-not $State.wowPath) { return $false }
    $prefix = (Get-FlavorPath $State) + [IO.Path]::DirectorySeparatorChar
    try {
        foreach ($process in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like 'Wow*' -or $_.ProcessName -like 'World of Warcraft*' })) {
            $exe = $null
            try { $exe = $process.Path } catch { }
            if ($exe -and $exe.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
    }
    catch { }
    return $false
}

function Write-RunningGameHint {
    param($State, [string] $Text = 'WoW Forever is running: restart the game to load the changes (/reload is not enough for new add-ons).')
    if (Test-WowRunning $State) { Write-Msg Warn $Text }
}

#endregion

#region State (JSON) ------------------------------------------------------------------------------

function Get-DefaultStatePath {
    $base = [Environment]::GetFolderPath('ApplicationData')
    if (-not $base) { $base = $HOME }
    return [IO.Path]::Combine($base, 'ForeverAddonManager', 'forever-addon-manager.json')
}

function New-State {
    return [pscustomobject]@{ schemaVersion = 1; wowPath = ''; addons = @() }
}

function ConvertTo-AddonRecord {
    param($Data)
    $channelName = ([string]$Data.channel).ToLowerInvariant()
    if (-not $script:ChannelRank.ContainsKey($channelName)) { $channelName = 'release' }
    $latest = $null
    if ($Data.latest -and $Data.latest.fileId) {
        $latest = [pscustomobject]@{
            fileId      = [long]$Data.latest.fileId
            displayName = [string]$Data.latest.displayName
            version     = [string]$Data.latest.version
            fileDate    = ConvertTo-IsoDate $Data.latest.fileDate
            releaseType = [string]$Data.latest.releaseType
        }
    }
    return [pscustomobject]@{
        projectId    = [long]$Data.projectId
        slug         = [string]$Data.slug
        name         = [string]$Data.name
        authors      = [string]$Data.authors
        url          = [string]$Data.url
        channel      = $channelName
        fileId       = [long]$Data.fileId
        fileName     = [string]$Data.fileName
        displayName  = [string]$Data.displayName
        version      = [string]$Data.version
        releaseType  = [string]$Data.releaseType
        fileDate     = ConvertTo-IsoDate $Data.fileDate
        gameVersions = ConvertTo-StringList $Data.gameVersions
        folders      = ConvertTo-StringList $Data.folders
        installedAt  = ConvertTo-IsoDate $Data.installedAt
        updatedAt    = ConvertTo-IsoDate $Data.updatedAt
        checkedAt    = ConvertTo-IsoDate $Data.checkedAt
        checkResult  = [string]$Data.checkResult
        latest       = $latest
    }
}

function Read-State {
    $state = New-State
    if (-not (Test-Path -LiteralPath $script:StatePath -PathType Leaf)) { return $state }
    $hint = "Fix or delete it, or restore the previous version from '$($script:StatePath).bak'."
    $raw = [IO.File]::ReadAllText($script:StatePath)
    if (-not $raw.Trim()) { throw "The state file '$($script:StatePath)' is empty. $hint" }
    try { $data = ConvertFrom-Json -InputObject $raw }
    catch {
        $reason = (Get-ErrorMessage $_) -replace '\s+', ' '
        throw "The state file '$($script:StatePath)' could not be read ($reason). $hint"
    }
    if ($data -isnot [System.Management.Automation.PSCustomObject]) { throw "The state file '$($script:StatePath)' has an unexpected format. $hint" }
    if ([int]$data.schemaVersion -gt 1) { throw "The state file '$($script:StatePath)' was written by a newer version of this tool." }
    if ($data.wowPath) { $state.wowPath = [string]$data.wowPath }
    $list = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $number = 0
    foreach ($addon in @($data.addons)) {
        $number++
        if ($null -eq $addon) { continue }
        if (-not $addon.projectId -or [long]$addon.projectId -le 0) { throw "Add-on entry #$number in the state file has no project ID. $hint" }
        $record = ConvertTo-AddonRecord $addon
        foreach ($folder in $record.folders) {
            if (-not (Test-SafeFolderName $folder)) { throw "Add-on entry '$($record.name)' in the state file has an invalid folder name '$folder'. $hint" }
        }
        if ($seen.ContainsKey([string]$record.projectId)) { continue }
        $seen[[string]$record.projectId] = $true
        $list.Add($record)
    }
    $state.addons = $list.ToArray()
    return $state
}

function Save-State {
    param($State)
    $dir = Split-Path -Parent $script:StatePath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $State.addons = @(Get-SortedAddons $State)
    $json = ConvertTo-Json -InputObject $State -Depth 8
    $tmp = "$($script:StatePath).tmp"
    [IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $script:StatePath -PathType Leaf) {
        try { [IO.File]::Replace($tmp, $script:StatePath, "$($script:StatePath).bak") }
        catch {
            Copy-Item -LiteralPath $script:StatePath -Destination "$($script:StatePath).bak" -Force
            Move-Item -LiteralPath $tmp -Destination $script:StatePath -Force
        }
    }
    else {
        [IO.File]::Move($tmp, $script:StatePath)
    }
}

function Get-SortedAddons {
    param($State)
    return @(@($State.addons) | Where-Object { $null -ne $_ } |
        Sort-Object -Property @{ Expression = { ([string]$_.name).ToLowerInvariant() } }, @{ Expression = { [long]$_.projectId } })
}

# Only one instance at a time may change a state file and its AddOns folder (named mutex per state file).
function Enter-InstanceLock {
    $script:InstanceLock = $null
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($script:StatePath.ToLowerInvariant()))
        $hash = [BitConverter]::ToString($bytes).Replace('-', '').Substring(0, 16)
        $mutex = New-Object System.Threading.Mutex -ArgumentList $false, "Local\ForeverAddonManager-$hash"
    }
    catch {
        Write-Verbose "Running without an instance lock: $(Get-ErrorMessage $_)"
        return
    }
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne(0)
        if (-not $acquired) {
            Write-Msg Step 'Another instance of the add-on manager is running, waiting for it to finish...'
            $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(10))
        }
    }
    catch {
        $inner = $_.Exception
        while ($inner.InnerException) { $inner = $inner.InnerException }
        if ($inner -is [System.Threading.AbandonedMutexException]) { $acquired = $true } else { throw }
    }
    if (-not $acquired) {
        $mutex.Dispose()
        throw 'Another instance of the add-on manager is still running. Please try again later.'
    }
    $script:InstanceLock = $mutex
}

function Exit-InstanceLock {
    if (-not $script:InstanceLock) { return }
    try { $script:InstanceLock.ReleaseMutex() } catch { }
    $script:InstanceLock.Dispose()
    $script:InstanceLock = $null
}

function Get-AddonById {
    param($State, [long] $ProjectId)
    return (@($State.addons) | Where-Object { $_ -and [long]$_.projectId -eq $ProjectId } | Select-Object -First 1)
}

function Set-AddonRecord {
    param($State, $Record)
    $others = @(@($State.addons) | Where-Object { $_ -and [long]$_.projectId -ne [long]$Record.projectId })
    $State.addons = @($others + $Record)
}

function Remove-AddonRecord {
    param($State, [long] $ProjectId)
    $State.addons = @(@($State.addons) | Where-Object { $_ -and [long]$_.projectId -ne $ProjectId })
}

# Maps lower-case folder name -> names of the managed add-ons using it (optionally excluding one).
function Get-FolderOwners {
    param($State, [long] $ExcludeProjectId = 0)
    $owners = @{}
    foreach ($addon in @($State.addons)) {
        if (-not $addon -or [long]$addon.projectId -eq $ExcludeProjectId) { continue }
        foreach ($folder in @($addon.folders)) {
            if (-not $folder) { continue }
            $key = ([string]$folder).ToLowerInvariant()
            if (-not $owners.ContainsKey($key)) { $owners[$key] = New-Object System.Collections.Generic.List[string] }
            $owners[$key].Add([string]$addon.name)
        }
    }
    return $owners
}

function Get-MissingFolders {
    param($State, $Addon)
    $addOnsDir = Get-AddOnsPath $State
    return @(@($Addon.folders) | Where-Object { $_ -and -not (Test-Path -LiteralPath (Join-Path $addOnsDir $_) -PathType Container) })
}

function Find-ManagedAddon {
    param($State, [string] $Query)
    $sorted = @(Get-SortedAddons $State)
    $q = ([string]$Query).Trim().Trim('"', "'").Trim()
    if (-not $q) { throw 'Empty add-on name.' }
    if ($q -match '^\d+$') {
        $number = [long]$q
        if ($number -ge 1 -and $number -le $sorted.Count) { return $sorted[$number - 1] }
        $byId = @($sorted | Where-Object { [long]$_.projectId -eq $number })
        if ($byId.Count -gt 0) { return $byId[0] }
        throw "There is no installed add-on with number or project ID $q (see 'list')."
    }
    $reference = $null
    try { $reference = ConvertFrom-AddonReference $q } catch { }
    if ($reference -and $reference.ProjectId) {
        $byId = @($sorted | Where-Object { [long]$_.projectId -eq [long]$reference.ProjectId })
        if ($byId.Count -gt 0) { return $byId[0] }
    }
    if ($reference -and $reference.Slug) {
        $bySlug = @($sorted | Where-Object { $_.slug -eq $reference.Slug })
        if ($bySlug.Count -gt 0) { return $bySlug[0] }
    }
    $exact = @($sorted | Where-Object { $_.name -eq $q -or $_.slug -eq $q })
    if ($exact.Count -eq 1) { return $exact[0] }
    $partial = @($sorted | Where-Object {
            ([string]$_.name).IndexOf($q, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            ([string]$_.slug).IndexOf($q, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
    if ($partial.Count -eq 1) { return $partial[0] }
    if ($partial.Count -gt 1) {
        throw "'$q' matches several add-ons: $(($partial | ForEach-Object { $_.name }) -join ', '). Please be more specific."
    }
    throw "No installed add-on matches '$q'. Run 'list' to see the installed add-ons."
}

function Resolve-ManagedAddons {
    param($State, [string[]] $Items)
    $found = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($item in $Items) {
        $addon = Find-ManagedAddon -State $State -Query $item
        $key = [string]$addon.projectId
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $found.Add($addon) }
    }
    return $found.ToArray()
}

#endregion

#region HTTP --------------------------------------------------------------------------------------

function Initialize-Http {
    if ($script:Http) { return }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
    try { if ([Net.ServicePointManager]::DefaultConnectionLimit -lt 8) { [Net.ServicePointManager]::DefaultConnectionLimit = 8 } } catch { }
    try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch { }
    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    $client = New-Object System.Net.Http.HttpClient -ArgumentList $handler
    $client.Timeout = [TimeSpan]::FromMinutes(5)
    [void]$client.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', $script:UserAgent)
    $script:Http = $client
}

# GETs all URLs in parallel (with retries for transient errors). Returns a hashtable url -> result.
function Invoke-HttpGetMany {
    param([string[]] $Urls, [int] $MaxAttempts = 3, [int] $TimeoutSeconds = 60)
    Initialize-Http
    $results = @{}
    $pending = @($Urls | Where-Object { $_ } | Select-Object -Unique)
    for ($attempt = 1; $attempt -le $MaxAttempts -and $pending.Count -gt 0; $attempt++) {
        if ($attempt -gt 1) { Start-Sleep -Milliseconds (1000 * ($attempt - 1)) }
        $tasks = @{}
        foreach ($url in $pending) { $tasks[$url] = $script:Http.GetAsync($url) }
        try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($tasks.Values), [TimeSpan]::FromSeconds($TimeoutSeconds)) } catch { }
        $retry = New-Object System.Collections.Generic.List[string]
        foreach ($url in $pending) {
            $task = $tasks[$url]
            $result = [pscustomobject]@{ Url = $url; Ok = $false; Status = 0; Content = $null; Error = $null }
            if ($task.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
                $response = $task.Result
                try {
                    $result.Status = [int]$response.StatusCode
                    $result.Content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $result.Ok = [bool]$response.IsSuccessStatusCode
                }
                catch { $result.Error = Get-ErrorMessage $_ }
                finally { $response.Dispose() }
                if ($result.Status -in @(408, 429, 500, 502, 503, 504) -or $result.Error) { $retry.Add($url) }
            }
            else {
                if ($task.Exception) { $result.Error = Get-ErrorMessage $task.Exception } else { $result.Error = 'The request timed out.' }
                $retry.Add($url)
            }
            $results[$url] = $result
        }
        $pending = $retry.ToArray()
    }
    return $results
}

function Invoke-HttpGet {
    param([string] $Url)
    return (Invoke-HttpGetMany -Urls @($Url))[$Url]
}

function Format-HttpFailure {
    param($Result)
    if (-not $Result) { return 'no response' }
    if ($Result.Error) { return $Result.Error }
    switch ([int]$Result.Status) {
        403 { return 'CurseForge refused the request (HTTP 403, bot protection). Please try again later.' }
        404 { return 'not found (HTTP 404)' }
        429 { return 'too many requests (HTTP 429). Please wait a moment and try again.' }
        default { return "HTTP $($Result.Status)" }
    }
}

function Invoke-HttpGetJson {
    param([string] $Url)
    $result = Invoke-HttpGet $Url
    if (-not $result.Ok) { throw (Format-HttpFailure $result) }
    return (ConvertFrom-Json -InputObject $result.Content)
}

function Save-HttpFile {
    param([string] $Url, [string] $Path)
    Initialize-Http
    $response = $script:Http.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
    try {
        if (-not $response.IsSuccessStatusCode) { throw "HTTP $([int]$response.StatusCode)" }
        $inStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        try {
            $outStream = [IO.File]::Open($Path, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try { $inStream.CopyTo($outStream, 81920) } finally { $outStream.Dispose() }
        }
        finally { $inStream.Dispose() }
    }
    finally { $response.Dispose() }
}

#endregion

#region CurseForge --------------------------------------------------------------------------------

# Parses a CurseForge URL, slug or project ID into @{ Slug; ProjectId }.
function ConvertFrom-AddonReference {
    param([string] $Text)
    $t = ([string]$Text).Trim().Trim('"', "'").Trim()
    if ($t -match '^\d+$') { return [pscustomobject]@{ Slug = $null; ProjectId = [long]$t } }
    $u = $t -replace '^https?://', ''
    if ($u -match '^(?:www\.|beta\.|legacy\.)?curseforge\.com/projects/(?<id>\d+)') {
        return [pscustomobject]@{ Slug = $null; ProjectId = [long]$Matches['id'] }
    }
    if ($u -match '^wow\.curseforge\.com/projects/(?<slug>[^/?#\s]+)') {
        return [pscustomobject]@{ Slug = $Matches['slug'].ToLowerInvariant(); ProjectId = $null }
    }
    if ($u -match '^(?:www\.|beta\.|legacy\.)?curseforge\.com/(?<game>[^/?#\s]+)/(?<kind>[^/?#\s]+)/(?<slug>[^/?#\s]+)') {
        if ($Matches['game'] -ne 'wow' -or $Matches['kind'] -ne 'addons') {
            throw "Only World of Warcraft add-on URLs (https://www.curseforge.com/wow/addons/<name>) are supported: $Text"
        }
        return [pscustomobject]@{ Slug = $Matches['slug'].ToLowerInvariant(); ProjectId = $null }
    }
    if ($t -match '^https?://' -or $t.Contains('/')) { throw "Not a CurseForge add-on URL: $Text" }
    if ($t -match '^[A-Za-z0-9][A-Za-z0-9_\-]*$') { return [pscustomobject]@{ Slug = $t.ToLowerInvariant(); ProjectId = $null } }
    throw "Not a valid CurseForge add-on URL: $Text"
}

# Project data from CFWidget (www.curseforge.com itself is behind bot protection for scripts).
function Get-WidgetProject {
    param([string] $Slug, [long] $ProjectId)
    if ($Slug) { $url = "$($script:CfWidgetApi)/wow/addons/$([Uri]::EscapeDataString($Slug))" }
    else { $url = "$($script:CfWidgetApi)/$ProjectId" }
    $result = $null
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        $result = Invoke-HttpGet $url
        if ($result.Status -ne 202) { break }   # 202 = CFWidget is still fetching the project
        if ($attempt -eq 1) { Write-Msg Step 'CurseForge project data is being prepared, waiting a few seconds...' }
        Start-Sleep -Seconds 3
    }
    if ($result.Status -eq 404) {
        if ($Slug) { throw "No CurseForge add-on named '$Slug' was found. Please check the URL." }
        throw "No CurseForge project with ID $ProjectId was found."
    }
    if ($result.Status -eq 202) { throw 'CurseForge project data is still being prepared. Please try again in a minute.' }
    if (-not $result.Ok) { throw "Could not look up the add-on: $(Format-HttpFailure $result)" }
    $data = ConvertFrom-Json -InputObject $result.Content
    if (-not $data.id) { throw 'Unexpected answer while looking up the add-on.' }
    return $data
}

# Fallback lookup through api.curse.tools, a public proxy of the official CurseForge API. CFWidget does
# not know every project by its slug (it answers 404 for some young projects), the official API does.
function Get-ProxyProject {
    param([string] $Slug, [long] $ProjectId)
    if ($Slug) {
        $data = Invoke-HttpGetJson "$($script:CfApiProxy)/mods/search?gameId=1&classId=1&slug=$([Uri]::EscapeDataString($Slug))"
        $mod = @($data.data) | Where-Object { $_ -and [string]$_.slug -eq $Slug } | Select-Object -First 1
    }
    else {
        $result = Invoke-HttpGet "$($script:CfApiProxy)/mods/$ProjectId"
        if ($result.Status -eq 404) { return $null }
        if (-not $result.Ok) { throw (Format-HttpFailure $result) }
        $mod = (ConvertFrom-Json -InputObject $result.Content).data
    }
    if (-not $mod -or -not $mod.id) { return $null }
    if ([int]$mod.gameId -ne 1) { throw "CurseForge project $($mod.id) ($($mod.name)) is not a World of Warcraft add-on." }
    $page = [string]$mod.links.websiteUrl
    if (-not $page) { $page = "$($script:CfSite)/wow/addons/$($mod.slug)" }
    $authors = (@($mod.authors) | Where-Object { $_ -and $_.name } | ForEach-Object { [string]$_.name } | Select-Object -Unique) -join ', '
    return [pscustomobject]@{
        ProjectId = [long]$mod.id
        Slug      = [string]$mod.slug
        Name      = [string]$mod.name
        Authors   = $authors
        Url       = $page
    }
}

function Resolve-CfProject {
    param([string] $Slug, [long] $ProjectId)
    try { $data = Get-WidgetProject -Slug $Slug -ProjectId $ProjectId }
    catch {
        $widgetError = Get-ErrorMessage $_
        Write-Verbose "CFWidget lookup failed ($widgetError), asking the CurseForge API proxy instead."
        $proxyError = $null
        try {
            $project = Get-ProxyProject -Slug $Slug -ProjectId $ProjectId
            if ($project) { return $project }
        }
        catch {
            $proxyError = Get-ErrorMessage $_
            if ($proxyError -like '*is not a World of Warcraft add-on*') { throw }
        }
        $message = $widgetError
        if ($proxyError) { $message += " (the fallback lookup failed too: $proxyError)" }
        if ($Slug) {
            $message += " If the add-on page exists, add it by the 'Project ID' shown on its CurseForge page instead, e.g.:  $($script:ScriptName) add 1234567"
        }
        throw $message
    }
    if ($data.game -and [string]$data.game -ne 'wow') {
        throw "CurseForge project $($data.id) ($($data.title)) is not a World of Warcraft add-on."
    }
    $page = [string]$data.urls.curseforge
    if (-not $page) { $page = [string]$data.urls.project }
    $resolvedSlug = $Slug
    if ($page -match '/wow/addons/(?<slug>[^/?#]+)') { $resolvedSlug = $Matches['slug'] }
    if (-not $page) {
        if ($resolvedSlug) { $page = "$($script:CfSite)/wow/addons/$resolvedSlug" }
        else { $page = "$($script:CfSite)/projects/$($data.id)" }
    }
    $authors = (@($data.members) | Where-Object { $_ -and $_.username } | ForEach-Object { [string]$_.username } | Select-Object -Unique) -join ', '
    return [pscustomobject]@{
        ProjectId = [long]$data.id
        Slug      = $resolvedSlug
        Name      = [string]$data.title
        Authors   = $authors
        Url       = $page
    }
}

function Get-CfFilesUrl {
    param([long] $ProjectId, [int] $Page = 0, [switch] $AllFlavors)
    $url = "$($script:CfSite)/api/v1/mods/$ProjectId/files?pageIndex=$Page&pageSize=$($script:CfPageSize)&sort=dateCreated&sortDescending=true&removeAlphas=false"
    if (-not $AllFlavors) { $url += "&gameFlavorId=$($script:FlavorTypeId)" }
    return $url
}

function ConvertFrom-CfApiFile {
    param($File)
    $releaseType = 1
    try { if ($File.releaseType) { $releaseType = [int]$File.releaseType } } catch { }
    return [pscustomobject]@{
        Id           = [long]$File.id
        DisplayName  = [string]$File.displayName
        FileName     = [string]$File.fileName
        Date         = ConvertTo-UtcDate $File.dateCreated
        ReleaseType  = $releaseType
        GameVersions = ConvertTo-StringList $File.gameVersions
        Length       = [long]$File.fileLength
        IsForever    = (@($File.gameVersionTypeIds) -contains $script:FlavorTypeId)
        EarlyAccess  = [bool]$File.isEarlyAccessContent
    }
}

# Fallback when the CurseForge file list is unavailable: CFWidget only has version strings, so
# WoW Forever files are recognized by their game version (1.6x.y) or a "-forever" file name.
function ConvertFrom-WidgetFile {
    param($File)
    $releaseType = 1
    if ([string]$File.type -eq 'beta') { $releaseType = 2 }
    elseif ([string]$File.type -eq 'alpha') { $releaseType = 3 }
    $versions = ConvertTo-StringList $File.versions
    $isForever = (@($versions | Where-Object { $_ -match '^1\.6\d\.\d+' }).Count -gt 0) -or
                 ("$($File.name) $($File.display)" -match '\b(forever|camelot)\b')
    return [pscustomobject]@{
        Id           = [long]$File.id
        DisplayName  = [string]$File.display
        FileName     = [string]$File.name
        Date         = ConvertTo-UtcDate $File.uploaded_at
        ReleaseType  = $releaseType
        GameVersions = $versions
        Length       = [long]$File.filesize
        IsForever    = $isForever
        EarlyAccess  = $false
    }
}

function Select-LatestFile {
    param([object[]] $Files, [string] $Channel)
    $maxRank = $script:ChannelRank[[string]$Channel]
    if (-not $maxRank) { $maxRank = 1 }
    $candidates = @($Files | Where-Object { $_ -and $_.IsForever -and -not $_.EarlyAccess -and $_.ReleaseType -le $maxRank })
    if ($candidates.Count -eq 0) { return $null }
    return ($candidates |
        Sort-Object -Property @{ Expression = { $_.Date }; Descending = $true }, @{ Expression = { $_.Id }; Descending = $true } |
        Select-Object -First 1)
}

# Looks up the newest WoW Forever file for each target (objects with projectId + channel), in parallel.
# Returns a hashtable "projectId" -> @{ Latest; Files; Error }.
function Get-UpdateInfo {
    param([object[]] $Targets)
    $result = @{}
    $Targets = @($Targets | Where-Object { $_ })
    if ($Targets.Count -eq 0) { return $result }
    $urls = @{}
    foreach ($target in $Targets) { $urls[[string]$target.projectId] = Get-CfFilesUrl -ProjectId $target.projectId }
    $responses = Invoke-HttpGetMany -Urls @($urls.Values)
    foreach ($target in $Targets) {
        $key = [string]$target.projectId
        $info = [pscustomobject]@{ Latest = $null; Files = @(); Error = $null }
        try {
            $response = $responses[$urls[$key]]
            if ($response.Ok) {
                $files = New-Object System.Collections.Generic.List[object]
                $page = ConvertFrom-Json -InputObject $response.Content
                foreach ($file in @($page.data)) { if ($file) { $files.Add((ConvertFrom-CfApiFile $file)) } }
                $total = 0
                if ($page.pagination) { $total = [int]$page.pagination.totalCount }
                $latest = Select-LatestFile -Files $files.ToArray() -Channel $target.channel
                $index = 1
                while (-not $latest -and $total -gt ($index * $script:CfPageSize) -and $index -lt 20) {
                    $more = Invoke-HttpGetJson (Get-CfFilesUrl -ProjectId $target.projectId -Page $index)
                    foreach ($file in @($more.data)) { if ($file) { $files.Add((ConvertFrom-CfApiFile $file)) } }
                    $latest = Select-LatestFile -Files $files.ToArray() -Channel $target.channel
                    $index++
                }
                $info.Files = $files.ToArray()
                $info.Latest = $latest
                if (-not $latest -and $total -gt ($index * $script:CfPageSize)) {
                    $info.Error = "no $($target.channel) file among the newest $($index * $script:CfPageSize) WoW Forever files"
                }
            }
            else {
                $failure = Format-HttpFailure $response
                Write-Verbose "CurseForge file list failed for project $key ($failure), trying CFWidget instead."
                try {
                    $widget = Get-WidgetProject -ProjectId $target.projectId
                    $files = @(foreach ($file in @($widget.files)) { if ($file) { ConvertFrom-WidgetFile $file } })
                    $info.Files = @($files | Where-Object { $_.IsForever })
                    $info.Latest = Select-LatestFile -Files $files -Channel $target.channel
                }
                catch { $info.Error = $failure }
            }
        }
        catch { $info.Error = Get-ErrorMessage $_ }
        $result[$key] = $info
    }
    return $result
}

# Game versions an add-on supports (used to explain why there is no WoW Forever build).
function Get-AvailableGameVersions {
    param([long] $ProjectId)
    try {
        $page = Invoke-HttpGetJson (Get-CfFilesUrl -ProjectId $ProjectId -AllFlavors)
        return @(@($page.data) | ForEach-Object { @($_.gameVersions) } | Where-Object { $_ } | Select-Object -Unique |
            Sort-Object -Property @{ Expression = { try { [version]$_ } catch { [version]'0.0' } } } -Descending)
    }
    catch { return @() }
}

function Get-RequiredDependencies {
    param([long] $ProjectId)
    $data = Invoke-HttpGetJson "$($script:CfSite)/api/v1/mods/$ProjectId/dependencies?index=0&pageSize=50"
    return @(@($data.data) | Where-Object { $_ -and $_.id -and ([string]$_.type -eq 'RequiredDependency' -or [string]$_.type -eq '3') } | ForEach-Object {
            $slug = [string]$_.slug
            $url = "$($script:CfSite)/projects/$($_.id)"
            if ($slug) { $url = "$($script:CfSite)/wow/addons/$slug" }
            [pscustomobject]@{ ProjectId = [long]$_.id; Slug = $slug; Name = [string]$_.name; Authors = [string]$_.authorName; Url = $url }
        })
}

function Get-CfDownloadUrls {
    param([long] $ProjectId, $File)
    $id = [long]$File.Id
    $urls = New-Object System.Collections.Generic.List[string]
    $urls.Add("$($script:CfSite)/api/v1/mods/$ProjectId/files/$id/download")
    if ($File.FileName) {
        $part = "files/$([long][math]::Floor($id / 1000))/$($id % 1000)/$([Uri]::EscapeDataString($File.FileName))"
        $urls.Add("https://mediafilez.forgecdn.net/$part")
        $urls.Add("https://edge.forgecdn.net/$part")
    }
    return $urls.ToArray()
}

function Receive-AddonArchive {
    param([long] $ProjectId, $File, [string] $Destination)
    $failures = New-Object System.Collections.Generic.List[string]
    foreach ($url in @(Get-CfDownloadUrls -ProjectId $ProjectId -File $File)) {
        try {
            Save-HttpFile -Url $url -Path $Destination
            $length = (Get-Item -LiteralPath $Destination).Length
            if ($File.Length -gt 0 -and $length -ne $File.Length) { throw "incomplete download ($length of $($File.Length) bytes)" }
            if (-not (Test-ZipArchive $Destination)) { throw 'the file is not a valid zip archive' }
            return
        }
        catch {
            $failures.Add("$url - $(Get-ErrorMessage $_)")
            Write-Verbose "Download failed: $url - $(Get-ErrorMessage $_)"
        }
    }
    throw ("The download failed:`n      " + ($failures -join "`n      "))
}

#endregion

#region Install / uninstall -----------------------------------------------------------------------

function Initialize-Zip {
    if ($script:ZipReady) { return }
    try { Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop } catch { }
    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }
    $script:ZipReady = $true
}

function Test-ZipArchive {
    param([string] $Path)
    Initialize-Zip
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try { return ($zip.Entries.Count -gt 0) } finally { $zip.Dispose() }
    }
    catch { return $false }
}

function Get-TocBaseName {
    param([string] $FileName)
    $base = [IO.Path]::GetFileNameWithoutExtension($FileName)
    if ($base -match '^(?<name>.+?)[-_](?:Mainline|Standard|Classic|Vanilla|TBC|BCC|Wrath|WOTLKC|Cata|Mists|Camelot|Forever)$') {
        return $Matches['name']
    }
    return $base
}

function Get-TocTitle {
    param([string] $FolderPath, [string] $FolderName)
    $tocs = @(Get-ChildItem -LiteralPath $FolderPath -Filter '*.toc' -File -ErrorAction SilentlyContinue)
    $toc = $tocs | Where-Object { (Get-TocBaseName $_.Name) -eq $FolderName } | Select-Object -First 1
    if (-not $toc) { $toc = $tocs | Select-Object -First 1 }
    if (-not $toc) { return $null }
    foreach ($line in [IO.File]::ReadAllLines($toc.FullName)) {
        if ($line -match '^\s*##\s*Title\s*:\s*(?<title>.+)$') {
            $title = (($Matches['title'] -replace '\|c[0-9A-Fa-f]{8}', '') -replace '\|r', '').Trim()
            if ($title) { return $title }
        }
    }
    return $null
}

# Finds the add-on folders inside an extracted archive. Returns @{ Root; Folders }.
function Get-ExtractedAddonFolders {
    param([string] $Root, [int] $Depth = 0)
    $dirs = @(Get-ChildItem -LiteralPath $Root -Directory -Force | Where-Object { $_.Name -ne '__MACOSX' -and -not $_.Name.StartsWith('.') })
    $rootTocs = @(Get-ChildItem -LiteralPath $Root -File -Filter '*.toc' -Force)
    if ($rootTocs.Count -gt 0) {
        # Add-on files sit directly in the archive root: wrap them in a folder named after the .toc
        $name = Get-TocBaseName $rootTocs[0].Name
        $wrapRoot = "$Root.wrapped"
        $target = Join-Path $wrapRoot $name
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        foreach ($item in @(Get-ChildItem -LiteralPath $Root -Force)) { Move-Item -LiteralPath $item.FullName -Destination $target }
        return [pscustomobject]@{ Root = $wrapRoot; Folders = [string[]]@($name) }
    }
    if ($dirs.Count -eq 1 -and $Depth -lt 2) {
        # A single container folder without a .toc (e.g. "MyAddon-1.0/MyAddon/...") is unwrapped
        $innerTocs = @(Get-ChildItem -LiteralPath $dirs[0].FullName -File -Filter '*.toc' -Force)
        $innerAddons = @(Get-ChildItem -LiteralPath $dirs[0].FullName -Directory -Force |
            Where-Object { @(Get-ChildItem -LiteralPath $_.FullName -File -Filter '*.toc' -Force).Count -gt 0 })
        if ($innerTocs.Count -eq 0 -and $innerAddons.Count -gt 0) {
            return (Get-ExtractedAddonFolders -Root $dirs[0].FullName -Depth ($Depth + 1))
        }
    }
    if ($dirs.Count -eq 0) { throw 'The downloaded archive does not contain any add-on folders.' }
    return [pscustomobject]@{ Root = $Root; Folders = [string[]]@($dirs | ForEach-Object { $_.Name }) }
}

function Remove-PathSafe {
    param([string] $Path)
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        if (-not (Test-Path -LiteralPath $Path)) { return }
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -eq 4) { throw "Could not delete '$Path': $(Get-ErrorMessage $_)" }
            Start-Sleep -Milliseconds (300 * $attempt)
        }
    }
}

# Installs folders in two phases so a failure never leaves a half-installed add-on behind:
# 1) every folder is staged next to its target as "<name>.fam-new" (nothing is changed yet),
# 2) targets are swapped in; old versions wait as "<name>.fam-old" and are restored on any failure.
function Install-AddonFolders {
    param([string] $SourceRoot, [string[]] $Folders, [string] $AddOnsDir)
    $staged = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($name in $Folders) {
            $target = Get-AddonFolderPath -AddOnsDir $AddOnsDir -Name $name
            $source = [IO.Path]::Combine($SourceRoot, $name)
            $staging = "$target.fam-new"
            Remove-PathSafe $staging
            $staged.Add($staging)
            $moved = $false
            try { [IO.Directory]::Move($source, $staging); $moved = $true } catch { }   # only works on the same drive
            if (-not $moved) { Copy-Item -LiteralPath $source -Destination $staging -Recurse -Force }
        }
    }
    catch {
        $failure = Get-ErrorMessage $_
        foreach ($staging in $staged) { try { Remove-PathSafe $staging } catch { } }
        throw "Copying the add-on files failed, nothing was changed: $failure"
    }

    $swapped = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($name in $Folders) {
            $target = Get-AddonFolderPath -AddOnsDir $AddOnsDir -Name $name
            $backup = "$target.fam-old"
            Remove-PathSafe $backup
            if (Test-Path -LiteralPath $target) { [IO.Directory]::Move($target, $backup) }
            [IO.Directory]::Move("$target.fam-new", $target)
            $swapped.Add($name)
        }
    }
    catch {
        $failure = Get-ErrorMessage $_
        foreach ($name in $Folders) {
            $target = [IO.Path]::Combine($AddOnsDir, $name)
            try {
                if ($swapped.Contains($name)) { Remove-PathSafe $target }
                if ((Test-Path -LiteralPath "$target.fam-old") -and -not (Test-Path -LiteralPath $target)) { [IO.Directory]::Move("$target.fam-old", $target) }
                Remove-PathSafe "$target.fam-new"
            }
            catch { Write-Msg Warn "Could not restore '$name': $(Get-ErrorMessage $_)" }
        }
        throw "Installing the add-on folders failed and was rolled back: $failure"
    }
    foreach ($name in $Folders) {
        try { Remove-PathSafe ([IO.Path]::Combine($AddOnsDir, $name) + '.fam-old') }
        catch { Write-Verbose "Could not delete the backup of '$name': $(Get-ErrorMessage $_)" }
    }
}

# Deletes the saved settings of the given add-on folders: WTF\Account\*\SavedVariables\<folder>.lua(.bak),
# the per-character copies and the add-on's lines in the AddOns.txt files.
function Remove-AddonSettings {
    param($State, [string[]] $Folders)
    $accountRoot = Join-Path (Get-WtfPath $State) 'Account'
    if (-not $Folders -or $Folders.Count -eq 0 -or -not (Test-Path -LiteralPath $accountRoot -PathType Container)) { return 0 }
    $names = @{}
    foreach ($folder in $Folders) { $names[$folder.ToLowerInvariant()] = $true }
    $removed = 0
    $savedVariableDirs = @(Get-ChildItem -LiteralPath $accountRoot -Directory -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq 'SavedVariables' })
    foreach ($dir in $savedVariableDirs) {
        foreach ($file in @(Get-ChildItem -LiteralPath $dir.FullName -File -Force -ErrorAction SilentlyContinue)) {
            if ($file.Name -match '^(?<name>.+?)\.lua(?:\.bak)?$' -and $names.ContainsKey($Matches['name'].ToLowerInvariant())) {
                Remove-Item -LiteralPath $file.FullName -Force
                $removed++
            }
        }
    }
    foreach ($addOnsTxt in @(Get-ChildItem -LiteralPath $accountRoot -Filter 'AddOns.txt' -File -Recurse -Force -ErrorAction SilentlyContinue)) {
        $lines = [IO.File]::ReadAllLines($addOnsTxt.FullName)
        $keep = @($lines | Where-Object { -not ($_ -match '^\s*(?<name>[^:]+?)\s*:' -and $names.ContainsKey($Matches['name'].ToLowerInvariant())) })
        if ($keep.Count -ne $lines.Count) { [IO.File]::WriteAllLines($addOnsTxt.FullName, [string[]]$keep) }
    }
    return $removed
}

function Test-UpdateAvailable {
    param($Addon, [long] $LatestId, $LatestDate)
    if ($LatestId -le 0 -or $LatestId -eq [long]$Addon.fileId) { return $false }
    $installedDate = ConvertTo-UtcDate $Addon.fileDate
    $latestDate = ConvertTo-UtcDate $LatestDate
    if ($installedDate -and $latestDate) {
        if ($latestDate -lt $installedDate) { return $false }
        if ($latestDate -eq $installedDate) { return ($LatestId -gt [long]$Addon.fileId) }
    }
    return $true
}

# Remembers the result of an update check in the add-on record (used by 'list -Offline').
function Set-CheckResult {
    param($Addon, $Info)
    if (-not $Addon -or -not $Info) { return }
    $Addon.checkedAt = ConvertTo-IsoDate ([DateTime]::UtcNow)
    if ($Info.Error) { $Addon.checkResult = 'error'; return }
    if ($Info.Latest) {
        $Addon.checkResult = 'ok'
        $Addon.latest = [pscustomobject]@{
            fileId      = [long]$Info.Latest.Id
            displayName = [string]$Info.Latest.DisplayName
            version     = Get-ShortVersion $Info.Latest.DisplayName $Info.Latest.FileName
            fileDate    = ConvertTo-IsoDate $Info.Latest.Date
            releaseType = [string]$script:ReleaseTypeName[[int]$Info.Latest.ReleaseType]
        }
    }
    else {
        $Addon.checkResult = 'noforever'
        $Addon.latest = $null
    }
}

# Downloads a CurseForge file, installs its folders into Interface\AddOns and records it in the state.
function Install-AddonFile {
    param($State, $Project, $File, [string] $Channel, $Existing)
    $addOnsDir = Get-AddOnsPath $State
    if (-not (Test-Path -LiteralPath $addOnsDir -PathType Container)) { New-Item -ItemType Directory -Path $addOnsDir -Force | Out-Null }
    Initialize-Zip
    $work = Join-Path ([IO.Path]::GetTempPath()) ('fam-' + [Guid]::NewGuid().ToString('N').Substring(0, 10))
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    try {
        $zipPath = Join-Path $work 'addon.zip'
        Write-Msg Step "Downloading $($File.FileName) ($(Format-Size $File.Length))"
        Receive-AddonArchive -ProjectId $Project.ProjectId -File $File -Destination $zipPath
        $extractDir = Join-Path $work 'x'
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractDir)
        $content = Get-ExtractedAddonFolders -Root $extractDir
        $folders = [string[]]@($content.Folders)

        $owners = Get-FolderOwners -State $State -ExcludeProjectId $Project.ProjectId
        $previous = @{}
        if ($Existing) { foreach ($folder in @($Existing.folders)) { if ($folder) { $previous[$folder.ToLowerInvariant()] = $true } } }
        foreach ($folder in $folders) {
            $key = $folder.ToLowerInvariant()
            if ($owners.ContainsKey($key)) {
                Write-Msg Dim "Folder '$folder' is shared with: $(@($owners[$key]) -join ', ')"
            }
            elseif (-not $previous.ContainsKey($key) -and (Test-Path -LiteralPath (Join-Path $addOnsDir $folder))) {
                Write-Msg Dim "Replacing the existing folder '$folder' (it was not managed by this tool)"
            }
        }

        Install-AddonFolders -SourceRoot $content.Root -Folders $folders -AddOnsDir $addOnsDir

        if ($Existing) {
            $newFolders = @{}
            foreach ($folder in $folders) { $newFolders[$folder.ToLowerInvariant()] = $true }
            foreach ($old in @($Existing.folders)) {
                if (-not $old) { continue }
                $key = $old.ToLowerInvariant()
                if ($newFolders.ContainsKey($key) -or $owners.ContainsKey($key)) { continue }
                $oldPath = Get-AddonFolderPath -AddOnsDir $addOnsDir -Name $old
                if (Test-Path -LiteralPath $oldPath) {
                    Remove-PathSafe $oldPath
                    Write-Msg Dim "Removed the folder '$old' (no longer part of the add-on)"
                }
            }
        }

        $name = $Project.Name
        if (-not $name) { $name = Get-TocTitle -FolderPath (Join-Path $addOnsDir $folders[0]) -FolderName $folders[0] }
        if (-not $name) { $name = $folders[0] }
        $now = ConvertTo-IsoDate ([DateTime]::UtcNow)
        $installedAt = $now
        if ($Existing -and $Existing.installedAt) { $installedAt = $Existing.installedAt }
        $url = $Project.Url
        if (-not $url -and $Project.Slug) { $url = "$($script:CfSite)/wow/addons/$($Project.Slug)" }

        $record = [pscustomobject]@{
            projectId    = [long]$Project.ProjectId
            slug         = [string]$Project.Slug
            name         = [string]$name
            authors      = [string]$Project.Authors
            url          = [string]$url
            channel      = $Channel
            fileId       = [long]$File.Id
            fileName     = [string]$File.FileName
            displayName  = [string]$File.DisplayName
            version      = Get-ShortVersion $File.DisplayName $File.FileName
            releaseType  = [string]$script:ReleaseTypeName[[int]$File.ReleaseType]
            fileDate     = ConvertTo-IsoDate $File.Date
            gameVersions = ConvertTo-StringList $File.GameVersions
            folders      = $folders
            installedAt  = $installedAt
            updatedAt    = $now
            checkedAt    = $null
            checkResult  = ''
            latest       = $null
        }
        Set-AddonRecord -State $State -Record $record
        return $record
    }
    finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$work.wrapped" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Installs a CurseForge project (and its required dependencies) unless it is already managed.
function Add-Project {
    param($State, $Project, [string] $Channel, [hashtable] $Visited, [int] $Depth = 0)
    $key = [string]$Project.ProjectId
    if ($Visited.ContainsKey($key)) { return }
    $Visited[$key] = $true
    $label = $Project.Name
    if (-not $label) { $label = $Project.Slug }
    if (-not $label) { $label = "project $key" }

    $existing = Get-AddonById -State $State -ProjectId $Project.ProjectId
    if ($existing -and -not $script:Force) {
        if ($Depth -eq 0) {
            Write-Msg Info "$($existing.name) is already installed ($($existing.version)). Use 'update' to update it or 'add -Force' to reinstall it."
        }
        return
    }
    if ($existing -and -not $script:Channel) { $Channel = $existing.channel }   # a reinstall keeps the chosen channel

    $info = (Get-UpdateInfo -Targets @([pscustomobject]@{ projectId = $Project.ProjectId; channel = $Channel }))[$key]
    if ($info.Error) { throw "Could not read the files of '$label' from CurseForge: $($info.Error)" }
    if (-not $info.Latest) {
        if (@($info.Files).Count -gt 0) {
            throw "'$label' has WoW Forever files, but none in the '$Channel' channel. Try again with -Channel beta or -Channel alpha."
        }
        $versions = @(Get-AvailableGameVersions -ProjectId $Project.ProjectId | Select-Object -First 8)
        $hint = ''
        if ($versions.Count -gt 0) { $hint = " It is only available for game versions: $($versions -join ', ')." }
        throw "'$label' has no WoW Forever version on CurseForge, so it cannot be installed.$hint"
    }

    $record = Install-AddonFile -State $State -Project $Project -File $info.Latest -Channel $Channel -Existing $existing
    $script:InstalledCount++
    Set-CheckResult -Addon $record -Info $info
    Save-State $State
    $channelNote = ''
    if ($Channel -ne 'release') { $channelNote = ", $Channel channel" }
    Write-Styled @("  $($script:Glyph.Check) ", 'Ok', 'Installed ', 'Text', $record.name, 'Name', " $($record.version)", 'Newer',
        "  ($(@($record.folders).Count) folder(s)$channelNote)", 'Dim')
    Write-Msg Dim "Folders: $(Format-FolderList $record.folders 6)"

    if (-not $script:SkipDependencies) {
        $dependencies = @()
        try { $dependencies = @(Get-RequiredDependencies -ProjectId $Project.ProjectId) }
        catch {
            $script:HadErrors = $true
            Write-Msg Warn "Could not check the required dependencies of $($record.name): $(Get-ErrorMessage $_)"
        }
        foreach ($dependency in $dependencies) {
            if (-not $dependency) { continue }
            if (Get-AddonById -State $State -ProjectId $dependency.ProjectId) { continue }
            if ($Visited.ContainsKey([string]$dependency.ProjectId)) { continue }
            Write-Msg Info "$($record.name) requires $($dependency.Name) - installing it as well"
            try { Add-Project -State $State -Project $dependency -Channel 'release' -Visited $Visited -Depth ($Depth + 1) }
            catch {
                $script:HadErrors = $true
                Write-Msg Warn "Could not install the required dependency $($dependency.Name): $(Get-ErrorMessage $_)"
            }
        }
    }
}

# Deletes an add-on's folders (unless another managed add-on uses them) and, by default, its settings.
function Uninstall-Addon {
    param($State, $Addon)
    $result = [pscustomobject]@{ Folders = 0; Settings = 0; Shared = @() }
    $owners = Get-FolderOwners -State $State -ExcludeProjectId $Addon.projectId
    $exclusive = New-Object System.Collections.Generic.List[string]
    $shared = New-Object System.Collections.Generic.List[string]
    foreach ($folder in @($Addon.folders)) {
        if (-not $folder) { continue }
        if ($owners.ContainsKey($folder.ToLowerInvariant())) { $shared.Add($folder) } else { $exclusive.Add($folder) }
    }
    if ($State.wowPath -and (Test-Path -LiteralPath (Get-FlavorPath $State) -PathType Container)) {
        $addOnsDir = Get-AddOnsPath $State
        foreach ($folder in $exclusive) {
            $path = Get-AddonFolderPath -AddOnsDir $addOnsDir -Name $folder
            if (Test-Path -LiteralPath $path) { Remove-PathSafe $path; $result.Folders++ }
            Remove-PathSafe "$path.fam-new"
            Remove-PathSafe "$path.fam-old"
        }
        if (-not $script:KeepSettings -and $exclusive.Count -gt 0) {
            $result.Settings = Remove-AddonSettings -State $State -Folders $exclusive.ToArray()
        }
    }
    else {
        Write-Msg Warn 'The WoW Forever folder was not found, so only the list entry is removed.'
    }
    Remove-AddonRecord -State $State -ProjectId $Addon.projectId
    $result.Shared = $shared.ToArray()
    return $result
}

#endregion

#region List rendering ---------------------------------------------------------------------------

function New-Cell {
    param([string] $Text, [string] $Style)
    return [pscustomobject]@{ Text = $Text; Style = $Style }
}

# One table row: Index, Name, Installed, Latest, LatestIsNewer, Channel, Date, Status (key of StatusInfo).
function New-ListRow {
    param([int] $Index, $Addon, $Info, [int] $MissingCount)
    $latestVersion = ''
    $latestDate = $null
    $latestId = 0
    if ($Info -and -not $Info.Error -and $Info.Latest) {
        $latestVersion = Get-ShortVersion $Info.Latest.DisplayName $Info.Latest.FileName
        $latestDate = $Info.Latest.Date
        $latestId = [long]$Info.Latest.Id
    }
    elseif ((-not $Info -or $Info.Error) -and $Addon.latest) {
        $latestVersion = [string]$Addon.latest.version
        $latestDate = $Addon.latest.fileDate
        $latestId = [long]$Addon.latest.fileId
    }
    $isNewer = Test-UpdateAvailable -Addon $Addon -LatestId $latestId -LatestDate $latestDate

    $status = 'unknown'
    if ($Info) {
        if ($Info.Error) { $status = 'error' }
        elseif (-not $Info.Latest) { $status = 'noforever' }
        elseif ($isNewer) { $status = 'update' }
        else { $status = 'ok' }
    }
    elseif ($Addon.checkResult -eq 'noforever') { $status = 'noforever' }
    elseif ($Addon.checkResult -eq 'error') { $status = 'error' }
    elseif ($Addon.checkResult -eq 'ok' -and $Addon.latest) {
        if ($isNewer) { $status = 'update' } else { $status = 'ok' }
    }
    if ($MissingCount -gt 0) { $status = 'missing' }

    $date = $latestDate
    if (-not $date) { $date = $Addon.fileDate }
    return [pscustomobject]@{
        Index         = $Index
        Name          = [string]$Addon.name
        Installed     = [string]$Addon.version
        Latest        = $latestVersion
        LatestIsNewer = $isNewer
        Channel       = [string]$Addon.channel
        Date          = Format-ShortDate $date
        Status        = $status
    }
}

function Write-AddonTable {
    param([object[]] $Rows)
    # Soft = width a flexible column keeps as long as possible when the window is too narrow
    # Hide = order in which optional columns disappear on narrow windows (0 = always shown)
    $columns = @(
        [pscustomobject]@{ Header = '#';         Min = 1;  Soft = 1;  Max = 4;  Flex = $false; Right = $true;  Hide = 0 }
        [pscustomobject]@{ Header = 'Add-on';    Min = 12; Soft = 26; Max = 42; Flex = $true;  Right = $false; Hide = 0 }
        [pscustomobject]@{ Header = 'Installed'; Min = 8;  Soft = 14; Max = 26; Flex = $true;  Right = $false; Hide = 0 }
        [pscustomobject]@{ Header = 'Latest';    Min = 8;  Soft = 14; Max = 26; Flex = $true;  Right = $false; Hide = 0 }
        [pscustomobject]@{ Header = 'Channel';   Min = 7;  Soft = 7;  Max = 7;  Flex = $false; Right = $false; Hide = 2 }
        [pscustomobject]@{ Header = 'Released';  Min = 10; Soft = 10; Max = 10; Flex = $false; Right = $false; Hide = 1 }
        [pscustomobject]@{ Header = 'Status';    Min = 12; Soft = 18; Max = 18; Flex = $false; Right = $false; Hide = 0 }
    )
    $table = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Rows) {
        $status = $script:StatusInfo[[string]$row.Status]
        if (-not $status) { $status = $script:StatusInfo['unknown'] }
        $installed = [string]$row.Installed
        if (-not $installed) { $installed = '?' }
        $latest = [string]$row.Latest
        if (-not $latest) { if ($row.Status -eq 'noforever') { $latest = $script:Glyph.Dash } else { $latest = '?' } }
        $latestStyle = 'Dim'
        if ($row.LatestIsNewer) { $latestStyle = 'Newer' }
        $channelStyle = 'Release'
        if ($row.Channel -eq 'beta') { $channelStyle = 'Beta' } elseif ($row.Channel -eq 'alpha') { $channelStyle = 'Alpha' }
        $nameStyle = 'Name'
        if ($row.Status -eq 'missing' -or $row.Status -eq 'noforever') { $nameStyle = 'Text' }
        $table.Add([object[]]@(
                (New-Cell ([string]$row.Index) 'Dim'),
                (New-Cell ([string]$row.Name) $nameStyle),
                (New-Cell $installed 'Version'),
                (New-Cell $latest $latestStyle),
                (New-Cell ([string]$row.Channel) $channelStyle),
                (New-Cell ([string]$row.Date) 'Dim'),
                (New-Cell "$($status.Glyph) $($status.Label)" $status.Style)
            ))
    }

    # Column widths: natural width within [Min, Max]. On narrow windows optional columns are hidden
    # first (Released, then Channel), then flexible columns shrink towards their minimum width.
    $count = $columns.Count
    $widths = New-Object 'int[]' $count
    for ($c = 0; $c -lt $count; $c++) {
        $w = $columns[$c].Header.Length
        foreach ($cells in $table) { if ($cells[$c].Text.Length -gt $w) { $w = $cells[$c].Text.Length } }
        $widths[$c] = [Math]::Min([Math]::Max($w, $columns[$c].Min), $columns[$c].Max)
    }
    $indent = 2
    $gap = 2
    $available = (Get-ConsoleWidth) - 1
    $visible = New-Object System.Collections.Generic.List[int]
    for ($c = 0; $c -lt $count; $c++) { $visible.Add($c) }
    foreach ($priority in 1, 2) {
        $softTotal = $indent + $gap * ($visible.Count - 1)
        foreach ($c in $visible) {
            if ($columns[$c].Flex) { $softTotal += [Math]::Min($widths[$c], $columns[$c].Soft) } else { $softTotal += $widths[$c] }
        }
        if ($softTotal -le $available) { break }
        foreach ($c in @($visible | Where-Object { $columns[$_].Hide -eq $priority })) { [void]$visible.Remove($c) }
    }
    $total = $indent + $gap * ($visible.Count - 1)
    foreach ($c in $visible) { $total += $widths[$c] }
    while ($total -gt $available) {
        # shrink the flexible column that is furthest above its soft width (versions before names on ties)
        $pick = -1
        foreach ($c in $visible) {
            if (-not $columns[$c].Flex -or $widths[$c] -le $columns[$c].Min) { continue }
            if ($pick -lt 0 -or ($widths[$c] - $columns[$c].Soft) -ge ($widths[$pick] - $columns[$pick].Soft)) { $pick = $c }
        }
        if ($pick -lt 0) { break }
        $widths[$pick]--
        $total--
    }
    $last = $visible[$visible.Count - 1]

    $header = New-Object System.Collections.Generic.List[object]
    $rule = New-Object System.Collections.Generic.List[object]
    $header.Add(' ' * $indent); $header.Add('')
    $rule.Add(' ' * $indent); $rule.Add('')
    foreach ($c in $visible) {
        $text = Limit-Text $columns[$c].Header $widths[$c]
        if ($columns[$c].Right) { $text = $text.PadLeft($widths[$c]) } elseif ($c -ne $last) { $text = $text.PadRight($widths[$c]) }
        $header.Add($text); $header.Add('Head')
        $rule.Add($script:Glyph.Line * $widths[$c]); $rule.Add('Rule')
        if ($c -ne $last) { $header.Add(' ' * $gap); $header.Add(''); $rule.Add(' ' * $gap); $rule.Add('') }
    }
    Write-Styled $header.ToArray()
    Write-Styled $rule.ToArray()

    foreach ($cells in $table) {
        $parts = New-Object System.Collections.Generic.List[object]
        $parts.Add(' ' * $indent); $parts.Add('')
        foreach ($c in $visible) {
            $text = Limit-Text $cells[$c].Text $widths[$c]
            if ($columns[$c].Right) { $text = $text.PadLeft($widths[$c]) } elseif ($c -ne $last) { $text = $text.PadRight($widths[$c]) }
            $parts.Add($text); $parts.Add($cells[$c].Style)
            if ($c -ne $last) { $parts.Add(' ' * $gap); $parts.Add('') }
        }
        Write-Styled $parts.ToArray()
    }
    Write-Styled $rule.ToArray()
    Write-ListSummary -Rows $Rows
}

function Write-ListSummary {
    param([object[]] $Rows)
    $counts = @{}
    foreach ($row in $Rows) { $counts[[string]$row.Status] = 1 + [int]$counts[[string]$row.Status] }
    $total = @($Rows).Count
    $items = New-Object System.Collections.Generic.List[object]
    if ($total -eq 1) { $items.Add(@('1 add-on', 'Name')) } else { $items.Add(@("$total add-ons", 'Name')) }
    # status key, singular text, plural text
    $labels = @(
        @('ok', 'up to date', 'up to date'),
        @('update', 'update available', 'updates available'),
        @('missing', 'with missing files', 'with missing files'),
        @('noforever', 'without Forever build', 'without Forever build'),
        @('error', 'check failed', 'checks failed'),
        @('unknown', 'not checked', 'not checked')
    )
    foreach ($label in $labels) {
        $n = [int]$counts[$label[0]]
        if ($n -eq 0) { continue }
        $text = $label[2]
        if ($n -eq 1) { $text = $label[1] }
        $items.Add(@("$n $text", $script:StatusInfo[$label[0]].Style))
    }
    # Join with separators and wrap onto further lines when the window is too narrow
    $separator = "  $($script:Glyph.Middot)  "
    $maxWidth = (Get-ConsoleWidth) - 1
    $parts = New-Object System.Collections.Generic.List[object]
    $lineLength = 2
    $parts.Add('  '); $parts.Add('')
    for ($i = 0; $i -lt $items.Count; $i++) {
        $text = [string]$items[$i][0]
        if ($i -gt 0) {
            if ($lineLength + $separator.Length + $text.Length -gt $maxWidth) {
                Write-Styled $parts.ToArray()
                $parts.Clear(); $parts.Add('  '); $parts.Add('')
                $lineLength = 2
            }
            else {
                $parts.Add($separator); $parts.Add('Dim')
                $lineLength += $separator.Length
            }
        }
        $parts.Add($text); $parts.Add([string]$items[$i][1])
        $lineLength += $text.Length
    }
    Write-Styled $parts.ToArray()
    $fixable = [int]$counts['update'] + [int]$counts['missing']
    if ($fixable -gt 0) {
        Write-Styled @('  Run ', 'Dim', "$($script:ScriptName) update [name]", 'Cmd', ' to install the updates.', 'Dim')
    }
    Write-Host ''
}

#endregion

#region Commands ---------------------------------------------------------------------------------

function Invoke-SetPathCommand {
    param($State, [string[]] $Items)
    $text = (@($Items) -join ' ').Trim()
    if (-not $text) {
        $root = Find-WowInstallation
        if (-not $root) {
            throw "Could not find a World of Warcraft folder with WoW Forever automatically. Please pass it, e.g.:  $($script:ScriptName) set-path ""C:\Program Files (x86)\World of Warcraft"""
        }
        Write-Msg Info "Found World of Warcraft in $root"
    }
    else {
        $root = Resolve-WowRoot $text
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "The folder does not exist: $root" }
    }
    $flavorPath = Join-Path $root $script:FlavorDir
    if (-not (Test-Path -LiteralPath $flavorPath -PathType Container)) {
        $clients = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^_\w+_$' } | ForEach-Object { $_.Name })
        $message = "There is no WoW Forever installation ('$($script:FlavorDir)' folder) in '$root'."
        if ($clients.Count -gt 0) { $message += " Found other clients: $($clients -join ', ')." }
        if (-not $script:Force) { throw "$message Use -Force to save the folder anyway." }
        Write-Msg Warn $message
    }
    $previous = [string]$State.wowPath
    $State.wowPath = $root
    Save-State $State
    Write-Msg Ok "World of Warcraft folder set to: $root"
    Write-Msg Dim "Add-ons are installed into: $(Get-AddOnsPath $State)"
    $managed = @($State.addons).Count
    if ($previous -and $previous -ne $root -and $managed -gt 0) {
        Write-Msg Warn "$managed managed add-on(s) were installed in the previous folder ($previous)."
        Write-Msg Dim "If WoW was moved there is nothing to do. For a different installation, run 'update -Force' to install them here,"
        Write-Msg Dim "or 'remove' the ones you do not want (remove deletes same-named folders and settings in the new folder)."
    }
}

function Invoke-ConfigCommand {
    param($State)
    Write-Banner 'Configuration'
    if ($State.wowPath) {
        $flavorPath = Get-FlavorPath $State
        $addOnsPath = Get-AddOnsPath $State
        if (Test-Path -LiteralPath $State.wowPath -PathType Container) { Write-KeyValue 'WoW folder' $State.wowPath 'Text' "$($script:Glyph.Check) found" 'Ok' }
        else { Write-KeyValue 'WoW folder' $State.wowPath 'Text' "$($script:Glyph.Cross) not found" 'Error' }
        if (Test-Path -LiteralPath $flavorPath -PathType Container) { Write-KeyValue 'WoW Forever client' $flavorPath 'Text' "$($script:Glyph.Check) found" 'Ok' }
        else { Write-KeyValue 'WoW Forever client' $flavorPath 'Text' "$($script:Glyph.Cross) not found" 'Error' }
        if (Test-Path -LiteralPath $addOnsPath -PathType Container) { Write-KeyValue 'AddOns folder' $addOnsPath }
        else { Write-KeyValue 'AddOns folder' $addOnsPath 'Text' '(created on first install)' 'Dim' }
        Write-KeyValue 'Settings (WTF)' (Get-WtfPath $State)
    }
    else {
        Write-KeyValue 'WoW folder' 'not set' 'Warn' "run: $($script:ScriptName) set-path <folder>" 'Dim'
    }
    Write-KeyValue 'State file' $script:StatePath
    Write-KeyValue 'Managed add-ons' ([string]@($State.addons).Count)
    Write-KeyValue 'CurseForge flavor' "WoW Forever (game flavor $($script:FlavorTypeId))"
    Write-KeyValue 'Color mode' $script:ColorMode
    Write-KeyValue 'PowerShell' "$($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    Write-Host ''
}

function Invoke-AddCommand {
    param($State, [string[]] $Items)
    $Items = @($Items | Where-Object { $_ })
    if ($Items.Count -eq 0) { throw "Please give a CurseForge URL, e.g.:  $($script:ScriptName) add https://www.curseforge.com/wow/addons/eqol" }
    Assert-WowPath $State
    $channelName = 'release'
    if ($script:Channel) { $channelName = $script:Channel }
    Write-Banner "Installing into $(Get-AddOnsPath $State)"
    $visited = @{}
    foreach ($item in $Items) {
        Write-Styled @("  $($script:Glyph.Diamond) ", 'Brand', $item, 'Name')
        try {
            $reference = ConvertFrom-AddonReference $item
            $project = Resolve-CfProject -Slug $reference.Slug -ProjectId $reference.ProjectId
            Add-Project -State $State -Project $project -Channel $channelName -Visited $visited
        }
        catch {
            $script:HadErrors = $true
            Write-Msg Error (Get-ErrorMessage $_)
        }
        Write-Host ''
    }
    if ($script:InstalledCount -gt 0) { Write-RunningGameHint $State }
}

function Invoke-RemoveCommand {
    param($State, [string[]] $Items)
    $Items = @($Items | Where-Object { $_ })
    if ($Items.Count -eq 0) { throw "Please name the add-on to remove, e.g.:  $($script:ScriptName) remove eqol   (see 'list' for names and numbers)" }
    $targets = @(Resolve-ManagedAddons -State $State -Items $Items)
    Write-Banner "Removing $($targets.Count) add-on(s)"
    foreach ($addon in $targets) {
        Write-Styled @("  $($script:Glyph.Diamond) ", 'Brand', $addon.name, 'Name', "  $($addon.version)", 'Dim')
        Write-Msg Dim "Folders: $(Format-FolderList $addon.folders 6)"
    }
    if ($script:KeepSettings) { Write-Msg Info 'Saved settings are kept (-KeepSettings).' }
    else { Write-Msg Warn 'The saved settings of these add-ons (WTF SavedVariables) are deleted as well. Use -KeepSettings to keep them.' }
    if (-not $script:Force) {
        $answer = ''
        try { $answer = Read-Host "  Remove $($targets.Count) add-on(s)? [y/N]" }
        catch { throw 'Confirmation is not possible here. Run the command again with -Force to remove without asking.' }
        if ($answer -notmatch '^\s*(y|yes)\s*$') {
            Write-Msg Info 'Nothing was removed.'
            Write-Host ''
            return
        }
    }
    if (-not $script:KeepSettings) {
        Write-RunningGameHint $State 'WoW Forever is running: it may write the settings back when you log out. Close the game for a clean removal.'
    }
    foreach ($addon in $targets) {
        try {
            $result = Uninstall-Addon -State $State -Addon $addon
            Save-State $State
            $settingsNote = "$($result.Settings) settings file(s) deleted"
            if ($script:KeepSettings) { $settingsNote = 'settings kept' }
            Write-Msg Ok "Removed $($addon.name)  ($($result.Folders) folder(s), $settingsNote)"
            if (@($result.Shared).Count -gt 0) { Write-Msg Dim "Kept folder(s) still used by other add-ons: $($result.Shared -join ', ')" }
        }
        catch {
            $script:HadErrors = $true
            Write-Msg Error "$($addon.name): $(Get-ErrorMessage $_)"
        }
    }
    Write-Host ''
}

# Footer of 'list': kept low-key at the end so the add-on table stays in focus
function Write-GameDirectoryFooter {
    param($State)
    if (-not $State.wowPath) {
        Write-Styled @('  Game directory: ', 'Dim', 'not set', 'Warn', "  (run: $($script:ScriptName) set-path <folder>)", 'Dim')
    }
    elseif (Test-Path -LiteralPath (Get-FlavorPath $State) -PathType Container) {
        Write-Styled @('  Game directory: ', 'Dim', (Get-FlavorPath $State), 'Sub')
    }
    else {
        Write-Styled @('  Game directory: ', 'Dim', (Get-FlavorPath $State), 'Sub', '  (not found)', 'Error')
    }
    Write-Host ''
}

function Invoke-ListCommand {
    param($State)
    $addons = @(Get-SortedAddons $State)
    $subtitle = ''
    if ($script:Offline) { $subtitle = 'Offline: showing the status of the last check' }
    Write-Banner $subtitle -Credit
    if ($addons.Count -eq 0) {
        Write-Msg Info 'No add-ons are installed yet.'
        Write-Styled @('    Add one with:  ', 'Dim', "$($script:ScriptName) add https://www.curseforge.com/wow/addons/<name>", 'Cmd')
        Write-Host ''
        Write-GameDirectoryFooter $State
        return
    }
    $infos = @{}
    if (-not $script:Offline) {
        $checking = "  $($script:Glyph.Arrow) Checking $($addons.Count) add-on(s) on CurseForge..."
        if ($script:ColorMode -eq 'ansi') { Write-Styled @($checking, 'Dim') -NoNewline } else { Write-Styled @($checking, 'Dim') }
        $infos = Get-UpdateInfo -Targets $addons
        if ($script:ColorMode -eq 'ansi') { Write-Host "`r$([char]27)[2K" -NoNewline } else { Write-Host '' }
        foreach ($addon in $addons) { Set-CheckResult -Addon $addon -Info $infos[[string]$addon.projectId] }
        try { Save-State $State } catch { Write-Verbose "Could not save the check results: $(Get-ErrorMessage $_)" }
    }
    $flavorExists = $State.wowPath -and (Test-Path -LiteralPath (Get-FlavorPath $State) -PathType Container)
    $rows = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($addon in $addons) {
        $index++
        $missing = 0
        if ($flavorExists) { $missing = @(Get-MissingFolders -State $State -Addon $addon).Count }
        $rows.Add((New-ListRow -Index $index -Addon $addon -Info $infos[[string]$addon.projectId] -MissingCount $missing))
    }
    Write-AddonTable -Rows $rows.ToArray()
    Write-GameDirectoryFooter $State
    if (-not $script:Offline -and @($rows | Where-Object { $_.Status -eq 'error' }).Count -gt 0) { $script:HadErrors = $true }
}

function Invoke-UpdateCommand {
    param($State, [string[]] $Items)
    Assert-WowPath $State
    $Items = @($Items | Where-Object { $_ })
    if (@($State.addons).Count -eq 0) {
        Write-Banner
        Write-Msg Info "No add-ons are installed yet. Add one with:  $($script:ScriptName) add https://www.curseforge.com/wow/addons/<name>"
        Write-Host ''
        return
    }
    if ($Items.Count -gt 0) { $targets = @(Resolve-ManagedAddons -State $State -Items $Items) }
    else { $targets = @(Get-SortedAddons $State) }
    Write-Banner "Updating $($targets.Count) add-on(s) in $(Get-AddOnsPath $State)"
    $channelChanged = @{}
    if ($script:Channel) {
        foreach ($addon in $targets) {
            if ($addon.channel -ne $script:Channel) {
                Write-Msg Info "$($addon.name): release channel changed from $($addon.channel) to $($script:Channel)"
                $addon.channel = $script:Channel
                $channelChanged[[string]$addon.projectId] = $true
            }
        }
    }
    Write-Msg Step "Checking $($targets.Count) add-on(s) on CurseForge..."
    $infos = Get-UpdateInfo -Targets $targets
    $updated = 0; $current = 0; $skipped = 0; $failed = 0
    foreach ($addon in $targets) {
        $info = $infos[[string]$addon.projectId]
        Set-CheckResult -Addon $addon -Info $info
        if ($info.Error) {
            $failed++
            Write-Msg Error "$($addon.name): the update check failed - $($info.Error)"
            continue
        }
        if (-not $info.Latest) {
            $skipped++
            Write-Msg Warn "$($addon.name): no $($addon.channel) file for WoW Forever is available on CurseForge"
            continue
        }
        $missing = @(Get-MissingFolders -State $State -Addon $addon)
        $isNewer = Test-UpdateAvailable -Addon $addon -LatestId $info.Latest.Id -LatestDate $info.Latest.Date
        # After an explicit channel switch the channel's newest file is installed even if it is older
        $switchFile = $channelChanged.ContainsKey([string]$addon.projectId) -and [long]$info.Latest.Id -ne [long]$addon.fileId
        if (-not ($script:Force -or $isNewer -or $switchFile -or $missing.Count -gt 0)) {
            $current++
            Write-Styled @("  $($script:Glyph.Dot) ", 'Ok', $addon.name, 'Name', "  $($addon.version)", 'Dim', '  up to date', 'Ok')
            continue
        }
        try {
            if ($missing.Count -gt 0 -and -not $isNewer) { Write-Msg Info "$($addon.name): $($missing.Count) folder(s) missing - reinstalling" }
            $from = $addon.version
            $project = [pscustomobject]@{ ProjectId = $addon.projectId; Slug = $addon.slug; Name = $addon.name; Authors = $addon.authors; Url = $addon.url }
            $record = Install-AddonFile -State $State -Project $project -File $info.Latest -Channel $addon.channel -Existing $addon
            Set-CheckResult -Addon $record -Info $info
            Save-State $State
            $updated++
            if ($from -eq $record.version) {
                Write-Styled @("  $($script:Glyph.Check) ", 'Ok', 'Reinstalled ', 'Text', $record.name, 'Name', " $($record.version)", 'Version')
            }
            elseif ($switchFile -and -not $isNewer) {
                Write-Styled @("  $($script:Glyph.Check) ", 'Ok', 'Switched ', 'Text', $record.name, 'Name', "  $from", 'Version',
                    " $($script:Glyph.Right) ", 'Dim', $record.version, 'Newer', "  ($($record.channel) channel)", 'Dim')
            }
            else {
                Write-Styled @("  $($script:Glyph.Up) ", 'Update', 'Updated ', 'Text', $record.name, 'Name', "  $from", 'Version',
                    " $($script:Glyph.Right) ", 'Dim', $record.version, 'Newer')
            }
        }
        catch {
            $failed++
            Write-Msg Error "$($addon.name): $(Get-ErrorMessage $_)"
        }
    }
    Save-State $State
    if ($failed -gt 0) { $script:HadErrors = $true }
    $separator = "  $($script:Glyph.Middot)  "
    $summary = New-Object System.Collections.Generic.List[object]
    $summary.Add("  $updated updated"); if ($updated -gt 0) { $summary.Add('Update') } else { $summary.Add('Dim') }
    $summary.Add($separator); $summary.Add('Dim')
    $summary.Add("$current already up to date"); $summary.Add('Ok')
    if ($skipped -gt 0) { $summary.Add($separator); $summary.Add('Dim'); $summary.Add("$skipped without Forever build"); $summary.Add('Warn') }
    if ($failed -gt 0) { $summary.Add($separator); $summary.Add('Dim'); $summary.Add("$failed failed"); $summary.Add('Error') }
    Write-Host ''
    Write-Styled $summary.ToArray()
    if ($updated -gt 0) { Write-RunningGameHint $State }
    Write-Host ''
}

function Show-Help {
    Write-Banner 'Manage CurseForge add-ons for World of Warcraft: Forever' -Credit
    Write-Styled @('  USAGE', 'Head')
    Write-Styled @("    $($script:ScriptName) ", 'Cmd', '<command> [arguments] [options]', 'Text')
    Write-Host ''
    Write-Styled @('  COMMANDS', 'Head')
    $commands = @(
        @('set-path [folder]', 'Set the World of Warcraft folder (auto-detected when omitted)'),
        @('config', 'Show the current configuration'),
        @('add <url> [<url> ...]', 'Install add-ons from CurseForge, e.g. https://www.curseforge.com/wow/addons/eqol'),
        @('remove <addon> [...]', 'Uninstall add-ons and delete their files and saved settings'),
        @('list', 'Show installed add-ons, versions and update status'),
        @('update [<addon> ...]', 'Update all add-ons, or only the given ones'),
        @('help', 'Show this help')
    )
    foreach ($entry in $commands) { Write-Styled @(('    ' + $entry[0].PadRight(26)), 'Cmd', $entry[1], 'Text') }
    Write-Styled @('    <addon> = number shown by list, CurseForge URL, slug (eqol), project ID or part of the name', 'Dim')
    Write-Host ''
    Write-Styled @('  OPTIONS', 'Head')
    $options = @(
        @('-Channel release|beta|alpha', 'add/update: release channel to follow (default: release)'),
        @('-Force', 'add/update: reinstall even if up to date; remove: no confirmation'),
        @('-KeepSettings', 'remove: keep the saved settings (SavedVariables)'),
        @('-SkipDependencies', 'add: do not install required dependencies'),
        @('-Offline', 'list: do not contact CurseForge, show the last known status'),
        @('-NoColor', 'disable colors'),
        @('-StateFile <path>', 'use another state file')
    )
    foreach ($entry in $options) { Write-Styled @(('    ' + $entry[0].PadRight(30)), 'Key', $entry[1], 'Text') }
    Write-Host ''
    Write-Styled @('  EXAMPLES', 'Head')
    foreach ($example in @(
            'set-path "C:\Program Files (x86)\World of Warcraft"',
            'add https://www.curseforge.com/wow/addons/eqol',
            'list',
            'update',
            'update eqol -Channel beta',
            'remove eqol')) {
        Write-Styled @("    $($script:ScriptName) ", 'Dim', $example, 'Cmd')
    }
    Write-Host ''
    Write-Styled @('  State file: ', 'Dim', $script:StatePath, 'Text')
    Write-Host ''
}

#endregion

#region Main -------------------------------------------------------------------------------------

Initialize-Console
if ($StateFile) { $script:StatePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($StateFile) }
else { $script:StatePath = Get-DefaultStatePath }

$exitCode = 0
try {
    $commandName = ([string]$Command).Trim().ToLowerInvariant()
    $unknownOption = @($Arguments | Where-Object { $_ -match '^-[A-Za-z]' }) | Select-Object -First 1
    if ($unknownOption) { throw "Unknown option '$unknownOption'. Run '$($script:ScriptName) help' to see all options." }
    $mutating = @('list', 'ls', 'add', 'install', 'remove', 'rm', 'uninstall', 'delete', 'update', 'up', 'upgrade', 'set-path', 'setpath')
    if ($commandName -in $mutating) { Enter-InstanceLock }
    switch ($commandName) {
        { $_ -in @('help', '-h', '--help', '/?', '?') } { Show-Help; break }
        { $_ -in @('list', 'ls') } { Invoke-ListCommand (Read-State); break }
        { $_ -in @('add', 'install') } { Invoke-AddCommand (Read-State) $Arguments; break }
        { $_ -in @('remove', 'rm', 'uninstall', 'delete') } { Invoke-RemoveCommand (Read-State) $Arguments; break }
        { $_ -in @('update', 'up', 'upgrade') } { Invoke-UpdateCommand (Read-State) $Arguments; break }
        { $_ -in @('set-path', 'setpath') } { Invoke-SetPathCommand (Read-State) $Arguments; break }
        { $_ -in @('config', 'info') } { Invoke-ConfigCommand (Read-State); break }
        default { throw "Unknown command '$Command'. Run '$($script:ScriptName) help' to see all commands." }
    }
    if ($script:HadErrors) { $exitCode = 1 }
}
catch {
    Write-Msg Error (Get-ErrorMessage $_)
    Write-Verbose $_.ScriptStackTrace
    Write-Host ''
    $exitCode = 1
}
finally {
    Exit-InstanceLock
}
exit $exitCode

#endregion
