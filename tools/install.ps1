#Requires -Version 5.1
<#
    Task Reporter - installer
    -------------------------
    Double-click Install.bat in the folder above instead of running this by hand.

        tools\install.ps1                  normal
        tools\install.ps1 -NonInteractive  answers yes to everything

    A full record of the run is written to install.log in the project folder.
#>
[CmdletBinding()]
param([switch]$NonInteractive)

# ============================================================================
#  CONFIG - what this project needs
# ============================================================================

$App = @{
    Name          = 'Task Reporter'
    Blurb         = 'File task reports into task_reports.xlsx, from a real desktop window'
    Repo          = 'https://github.com/semyonsw/task_reporter.git'
    Issues        = 'https://github.com/semyonsw/task_reporter/issues'
    WindowsOnly   = $true

    MinPython     = '3.10'
    Venv          = '.winenv'          # the name windows\build.bat has always used
    Requirements  = $null
    PipExtra      = @('openpyxl', 'pywebview', 'pyinstaller')
    VerifyImports = @('openpyxl', 'webview', 'PyInstaller')

    MinNode       = $null
    NpmDirs       = @()
    EnvKeys       = @()

    ExtraSteps = @(
        @{ Title  = 'Checking the app runs from source'
           Action = {
               param($PyExe)
               $r = Run -Exe $PyExe -Arguments @('task-report-maker.py', '--board')
               if ($r.Code -eq 0) {
                   OK 'the app starts, reads the task board and exits cleanly'
               } elseif ($r.Output -match 'openpyxl') {
                   Stop-Install 'the app cannot load openpyxl, so it cannot touch the workbook.' @(
                       'Delete the .winenv folder in this project and run Install again.')
               } else {
                   Warn 'the app did not run cleanly from source.' 'The output is at the end of install.log. The build below may still work.'
               }
           } },

        @{ Title  = 'Building TaskReporter.exe (the double-click app)'
           Action = {
               param($PyExe)

               # A running copy holds TaskReporter.exe open, so the new one
               # could not replace it.  Say so before spending a minute building.
               $running = @(Get-Process -Name 'TaskReporter' -ErrorAction SilentlyContinue)
               if ($running.Count -gt 0) {
                   Warn ('Task Reporter is running right now (' + $running.Count + ' window(s)), so its .exe cannot be replaced.')
                   Info 'skipping the rebuild - the copy you already have keeps working.'
                   Info 'close the window and run Install again when you want the newest build.'
                   return
               }

               Info 'this is the slow step - a minute or two, and antivirus may watch it closely.'

               # Baked into the bundle so a copied or shortcutted exe can still
               # find the workbook and the task board in this folder.
               $homeFile = Join-Path $Root 'windows\project_home.txt'
               [System.IO.File]::WriteAllText($homeFile, ($Root + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))

               $spec = Join-Path $Root 'windows\TaskReporter.spec'
               if (-not (Test-Path $spec)) {
                   Warn 'windows\TaskReporter.spec is missing, so the exe cannot be built.' 'Re-clone the project. The source launcher below still works.'
                   return
               }
               $r = Run -Exe $PyExe -Arguments @('-m', 'PyInstaller', '--noconfirm', '--clean',
                        '--distpath', (Join-Path $Root 'build\dist'),
                        '--workpath', (Join-Path $Root 'build\work'), $spec)
               $built = Join-Path $Root 'build\dist\TaskReporter.exe'
               if ($r.Code -eq 0 -and (Test-Path $built)) {
                   try {
                       Copy-Item -LiteralPath $built -Destination (Join-Path $Root 'TaskReporter.exe') -Force
                       $mb = [math]::Round((Get-Item (Join-Path $Root 'TaskReporter.exe')).Length / 1MB, 1)
                       OK ('TaskReporter.exe built  (' + $mb + ' MB, no Python needed to run it)')
                   } catch {
                       Warn 'the exe was built but could not be copied into the project folder.' `
                            'It is probably still running - close the Task Reporter window and run Install again.'
                   }
               } elseif ($r.Output -match 'Access is denied|PermissionError|being used by another process') {
                   Warn 'the build could not write its files - the app is probably still running.' `
                        'Close the Task Reporter window (and check Task Manager for TaskReporter.exe), then run Install again.'
               } elseif ($r.Output -match 'VirusTotal|denied by policy|antivirus|Defender') {
                   Warn 'antivirus blocked the build.' `
                        'Allow this folder in your antivirus, or just use the source launcher - it does the same thing.'
               } else {
                   Warn 'the exe could not be built.' `
                        'Not fatal: "Start Task Reporter.bat" runs exactly the same app from source. Details at the end of install.log.'
               }
           } }
    )

    Launchers = @(
        @{ Name = 'Start Task Reporter.bat'; Body = @'
@echo off
rem Written by Install.bat. Runs the app in its own window, from source, with
rem no terminal and no browser tab. Used automatically if TaskReporter.exe
rem could not be built on this PC.
setlocal EnableExtensions
cd /d "%~dp0"

if exist "TaskReporter.exe" (
    start "" "TaskReporter.exe" %*
    exit /b 0
)

if not exist ".winenv\Scripts\pythonw.exe" (
    echo.
    echo  [ERROR] Not installed yet. Double-click Install.bat first.
    echo.
    pause
    exit /b 1
)

start "" ".winenv\Scripts\pythonw.exe" "task-report-maker.py" --app %*
'@ }
    )

    Shortcut = @{ Name = 'Task Reporter'; Target = 'TaskReporter.exe'
                  Fallback = 'Start Task Reporter.bat'; Icon = 'windows\TaskReporter.ico' }

    NextSteps = @(
        'Double-click "Task Reporter" on your Desktop  (or TaskReporter.exe here).',
        'A window opens - type what you did, press Enter, and it lands in task_reports.xlsx.',
        '',
        'The Board view keeps the running list of what still has to be done;',
        'tick the boxes and press "File Checked Tasks" to turn them into report rows.',
        '',
        'Your reports :  task_reports.xlsx',
        'Your board   :  task_board.json',
        'If it misbehaves:  .task_reporter_app.log'
    )
}

# ============================================================================
#  INSTALLER ENGINE - shared by all of the projects in this family.
#  Everything below this line is generic.  Configure the block above instead.
# ============================================================================

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$Root    = Split-Path -Parent $PSScriptRoot
$LogFile = Join-Path $Root 'install.log'

$script:StepNo   = 0
$script:Warnings = New-Object System.Collections.ArrayList
$script:Fixed    = New-Object System.Collections.ArrayList

# --------------------------------------------------------------------------
#  Output helpers
# --------------------------------------------------------------------------

function Log { param([string]$Text)
    try { Add-Content -LiteralPath $LogFile -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Text) -Encoding UTF8 } catch { }
}

function Line { param([string]$Text = '', [string]$Colour = 'Gray')
    Write-Host $Text -ForegroundColor $Colour; Log $Text
}

function Step { param([string]$Text)
    $script:StepNo++
    Write-Host ''
    Write-Host ("  [{0}] {1}" -f $script:StepNo, $Text) -ForegroundColor Cyan
    Log ("STEP {0}: {1}" -f $script:StepNo, $Text)
}

function OK { param([string]$Text)
    Write-Host '      ' -NoNewline; Write-Host 'ok' -ForegroundColor Green -NoNewline; Write-Host ("   " + $Text)
    Log ("  ok    " + $Text)
}

function Info { param([string]$Text)
    Write-Host ("           " + $Text) -ForegroundColor DarkGray; Log ("  info  " + $Text)
}

function Fixed { param([string]$Text)
    Write-Host '      ' -NoNewline; Write-Host 'fixed' -ForegroundColor Green -NoNewline; Write-Host ("  " + $Text)
    Log ("  fixed " + $Text); [void]$script:Fixed.Add($Text)
}

function Warn { param([string]$Text, [string]$Fix = '')
    Write-Host '      ' -NoNewline; Write-Host 'warn' -ForegroundColor Yellow -NoNewline; Write-Host (" " + $Text)
    if ($Fix) { Write-Host ("           -> " + $Fix) -ForegroundColor DarkYellow }
    Log ("  warn  " + $Text + "  fix: " + $Fix)
    [void]$script:Warnings.Add(@{ Text = $Text; Fix = $Fix })
}

function Stop-Install { param([string]$Text, [string[]]$Fix = @(), [string]$Kind = '')
    $err = New-Object System.Exception ($Text)
    $err.Data['fix']  = ($Fix -join "`n")
    $err.Data['kind'] = $Kind
    throw $err
}

function Box { param([string[]]$Lines, [string]$Colour = 'Cyan')
    $width = 0
    foreach ($l in $Lines) { if ($l.Length -gt $width) { $width = $l.Length } }
    $width += 4
    $rule = '  +' + ('-' * $width) + '+'
    Write-Host $rule -ForegroundColor $Colour
    foreach ($l in $Lines) {
        Write-Host '  |' -ForegroundColor $Colour -NoNewline
        Write-Host ('  ' + $l.PadRight($width - 2)) -NoNewline
        Write-Host '|' -ForegroundColor $Colour
    }
    Write-Host $rule -ForegroundColor $Colour
    foreach ($l in $Lines) { Log ("| " + $l) }
}

function Ask { param([string]$Question, [string]$Default = 'y')
    if ($NonInteractive) { return ($Default -eq 'y') }
    $suffix = '[Y/n]'
    if ($Default -ne 'y') { $suffix = '[y/N]' }
    while ($true) {
        Write-Host ('      ' + $Question + ' ' + $suffix + ' ') -ForegroundColor White -NoNewline
        $answer = Read-Host
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = $Default }
        switch ($answer.Trim().ToLower()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
        }
    }
}

# --------------------------------------------------------------------------
#  Process runner - captures everything, never throws on a non-zero exit
# --------------------------------------------------------------------------

function Run { param([string]$Exe, [string[]]$Arguments = @(), [string]$WorkDir = $Root)
    Log ("  run   " + $Exe + ' ' + ($Arguments -join ' '))
    $prevPref = $ErrorActionPreference
    $prevLoc  = Get-Location
    $code = 1603
    $text = ''
    try {
        # PowerShell's call operator quotes each array element for us, which is
        # the only way arguments like  -c "import sys; print(1)"  survive.
        $ErrorActionPreference = 'Continue'
        if ($WorkDir -and (Test-Path -LiteralPath $WorkDir)) { Set-Location -LiteralPath $WorkDir }
        $global:LASTEXITCODE = 0
        $out = (& $Exe @Arguments 2>&1 |
                    ForEach-Object {
                        if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.ToString() } else { $_ }
                    } | Out-String)
        $code = $global:LASTEXITCODE
        if ($null -eq $code) { $code = 0 }
        $text = $out
    } catch {
        $text = ('could not run "' + $Exe + '": ' + $_.Exception.Message)
        $code = 1603
    } finally {
        Set-Location $prevLoc
        $ErrorActionPreference = $prevPref
    }
    # Strip the colour escapes tools emit, so both the log and every pattern
    # match below see plain text.
    if ($text) { $text = [regex]::Replace($text, "$([char]27)\[[0-9;]*[A-Za-z]", '') }
    Log ("  exit  " + $code)
    if ($text -and $text.Trim()) { Log ("  ----- " + $text.Trim()) }
    return @{ Code = $code; Output = [string]$text }
}

function Get-Npm {
    foreach ($n in @('npm.cmd', 'npm.exe')) {
        $src = Have $n
        if ($src) { return $src }
    }
    $node = Have 'node'
    if ($node) {
        $guess = Join-Path (Split-Path -Parent $node) 'npm.cmd'
        if (Test-Path $guess) { return $guess }
    }
    $src = Have 'npm'
    if ($src -and $src -notmatch '\.ps1$') { return $src }   # npm.ps1 cannot be called like an exe
    return $null
}

function Have { param([string]$Name)
    $c = Get-Command $Name -ErrorAction SilentlyContinue
    if ($c) { return $c.Source } else { return $null }
}

# --------------------------------------------------------------------------
#  Prerequisite: Python
# --------------------------------------------------------------------------

# A Python that cannot import ssl cannot download anything, and one that cannot
# import venv or sqlite3 will fail later in a much more confusing way.  So the
# probe checks all three and prints the OpenSSL version as its proof of health.
$script:PythonProbe = 'import platform;print(platform.python_version());import ssl,sqlite3,venv;print(ssl.OPENSSL_VERSION)'
$script:BadPythons  = New-Object System.Collections.ArrayList

function Get-PythonVersion { param([string]$Exe, [string[]]$Prefix = @())
    $r = Run -Exe $Exe -Arguments ($Prefix + @('-c', 'import platform;print(platform.python_version())'))
    if ($r.Code -ne 0) { return $null }
    $m = [regex]::Match($r.Output, '(\d+)\.(\d+)\.(\d+)')
    if (-not $m.Success) { return $null }
    return [version]$m.Value
}

function Test-PythonCandidate { param([string]$Exe, [string[]]$Prefix = @())
    $r = Run -Exe $Exe -Arguments ($Prefix + @('-c', $script:PythonProbe))
    $m = [regex]::Match($r.Output, '(?m)^\s*(\d+\.\d+\.\d+)\s*$')
    if (-not $m.Success) { return $null }          # not a usable Python at all
    $info = @{ Exe = $Exe; Prefix = $Prefix; Version = [version]$m.Groups[1].Value
               Healthy = $false; Problem = '' }
    if ($r.Code -eq 0 -and $r.Output -match 'OpenSSL') { $info.Healthy = $true; return $info }
    if ($r.Output -match "_ssl|No module named 'ssl'|ssl\.py") {
        $info.Problem = 'its SSL support is broken, so pip cannot download anything'
    } elseif ($r.Output -match 'sqlite3') {
        $info.Problem = 'its sqlite3 module is broken'
    } elseif ($r.Output -match "No module named 'venv'|ensurepip") {
        $info.Problem = 'it cannot create virtual environments'
    } else {
        $info.Problem = 'it fails a basic self-check'
    }
    return $info
}

function Find-Python { param([version]$Min)
    $candidates = New-Object System.Collections.ArrayList
    # The "py" launcher knows about every Python on the box.  3.12 first: it has
    # the widest choice of ready-made packages, so it needs no C++ compiler.
    if (Have 'py') {
        foreach ($tag in @('-3.12', '-3.13', '-3.11', '-3.10', '-3')) {
            [void]$candidates.Add(@{ Exe = 'py'; Prefix = @($tag) })
        }
    }
    foreach ($name in @('python', 'python3')) {
        $src = Have $name
        if ($src) { [void]$candidates.Add(@{ Exe = $src; Prefix = @() }) }
    }
    foreach ($glob in @(
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "$env:ProgramFiles\Python3*\python.exe",
        "C:\Python3*\python.exe",
        "$env:LOCALAPPDATA\Python\pythoncore-3*\python.exe",
        "$env:USERPROFILE\miniconda3\python.exe",
        "$env:USERPROFILE\anaconda3\python.exe")) {
        foreach ($hit in (Get-ChildItem -Path $glob -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)) {
            [void]$candidates.Add(@{ Exe = $hit.FullName; Prefix = @() })
        }
    }

    $seen = @{}
    $script:BadPythons.Clear()
    foreach ($c in $candidates) {
        $key = ($c.Exe + ' ' + ($c.Prefix -join ' '))
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $info = Test-PythonCandidate -Exe $c.Exe -Prefix $c.Prefix
        if (-not $info) { continue }
        $where = $key
        if ($info.Healthy) {
            Info ('found Python ' + $info.Version + '  (' + $where + ')')
            if ($info.Version -ge $Min -and $info.Version.Major -eq 3) { return $info }
        } else {
            Info ('found Python ' + $info.Version + ' (' + $where + ') but ' + $info.Problem + ' - skipping it')
            [void]$script:BadPythons.Add(('Python ' + $info.Version + ' at "' + $where + '": ' + $info.Problem))
        }
    }
    return $null
}

function Install-Python { param([version]$Min, [string]$Series = '3.12', [switch]$Quiet)
    if (-not $Quiet) {
        if ($script:BadPythons.Count) {
            Warn ('a Python is installed, but not one this app can use.')
            foreach ($b in $script:BadPythons) { Info $b }
        } else {
            Warn ('Python ' + $Min + ' or newer was not found on this computer.')
        }
    }
    if (Have 'winget') {
        if ($Quiet -or (Ask ('Install Python ' + $Series + ' automatically now (via winget)?'))) {
            Info ('downloading and installing Python ' + $Series + ' - this takes a couple of minutes...')
            $r = Run -Exe 'winget' -Arguments @('install', '--id', ('Python.Python.' + $Series), '-e',
                        '--source', 'winget', '--accept-package-agreements',
                        '--accept-source-agreements', '--disable-interactivity')
            if ($r.Code -eq 0 -or $r.Output -match 'Successfully installed|already installed') {
                # winget only puts python on the PATH of *new* shells - refresh ours.
                $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                            [System.Environment]::GetEnvironmentVariable('Path', 'User')
                $found = Find-Python -Min $Min
                if ($found) { Fixed ('installed Python ' + $found.Version); return $found }
                Warn 'Python was installed but still cannot be found from this window.'
                Info 'close this window and double-click Install again - a fresh window will see it.'
            } else {
                Warn 'the automatic Python install did not complete.' 'Install it by hand with the link below, then run this installer again.'
            }
        }
    } else {
        Warn 'winget is not available on this Windows version, so Python cannot be installed automatically.'
    }
    if ($Quiet) { return $null }
    $fix = @(
        '1. Download Python from  https://www.python.org/downloads/windows/',
        '   (3.12 or 3.13 - avoid the newest release for a week or two: not every',
        '    package has been rebuilt for it yet.)',
        '2. On the FIRST screen of the Python installer, tick "Add python.exe to PATH".',
        '3. Finish the install, then double-click Install again.')
    if ($script:BadPythons.Count) {
        $fix += @('', 'Note: an Anaconda/Miniconda Python was found but its SSL is broken.',
                      'Either fix it (open "Anaconda Prompt" and run: conda install -y openssl)',
                      'or just install a normal Python as above - both can live side by side.')
    }
    Stop-Install 'Python is required and no usable copy could be installed automatically.' $fix
}

# --------------------------------------------------------------------------
#  Prerequisite: Node.js
# --------------------------------------------------------------------------

function Get-NodeVersion {
    $exe = Have 'node'
    if (-not $exe) { return $null }
    $r = Run -Exe $exe -Arguments @('--version')
    if ($r.Code -ne 0) { return $null }
    $m = [regex]::Match($r.Output, '(\d+)\.(\d+)\.(\d+)')
    if (-not $m.Success) { return $null }
    return [version]$m.Value
}

function Ensure-Node { param([version]$Min)
    $v = Get-NodeVersion
    if ($v -and $v -ge $Min) { OK ('Node.js ' + $v); return $true }
    if ($v) { Warn ('Node.js ' + $v + ' is older than the required ' + $Min + '.') }
    else    { Warn 'Node.js was not found on this computer.' }

    if (Have 'winget') {
        if (Ask 'Install the current Node.js LTS automatically now (via winget)?') {
            Info 'installing Node.js LTS - this takes a couple of minutes...'
            $r = Run -Exe 'winget' -Arguments @('install', '--id', 'OpenJS.NodeJS.LTS', '-e',
                        '--source', 'winget', '--accept-package-agreements',
                        '--accept-source-agreements', '--disable-interactivity')
            if ($r.Code -eq 0) {
                $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                            [System.Environment]::GetEnvironmentVariable('Path', 'User')
                $v = Get-NodeVersion
                if ($v -and $v -ge $Min) { Fixed ('installed Node.js ' + $v); return $true }
            }
            Warn 'the automatic Node.js install did not complete.'
        }
    }
    Stop-Install ('Node.js ' + $Min + ' or newer is required and could not be installed automatically.') @(
        '1. Download the LTS installer from  https://nodejs.org/en/download',
        '2. Install it with the default options.',
        '3. Close this window, then double-click Install again.')
}

# --------------------------------------------------------------------------
#  pip, with a diagnosis for every common failure
# --------------------------------------------------------------------------

function Explain-PipFailure { param([string]$Output)
    $o = $Output
    if ($o -match 'Microsoft Visual C\+\+ 14|vcvarsall|error: command .*cl\.exe') {
        return @{ Why = 'a package wanted to compile C code and this PC has no C++ compiler.'
                  Fix = @('Install the free "Build Tools for Visual Studio" (tick "Desktop development with C++"):',
                          '  https://visualstudio.microsoft.com/visual-cpp-build-tools/',
                          'Then run this installer again.') }
    }
    if ($o -match 'CERTIFICATE_VERIFY_FAILED|SSLError|SSLCertVerificationError') {
        return @{ Why = 'the HTTPS connection to pypi.org could not be verified - usually a company proxy or antivirus.'
                  Fix = @('If you are on a work network, ask IT for the proxy address and set it first:',
                          '  set HTTPS_PROXY=http://proxy.company.com:8080',
                          'Then run this installer again. (It already retried with the certificate check relaxed.)') }
    }
    if ($o -match 'ProxyError|Tunnel connection failed|407 ') {
        return @{ Why = 'a proxy server refused the connection.'
                  Fix = @('Set your proxy and try again:',
                          '  set HTTPS_PROXY=http://user:password@proxy.company.com:8080') }
    }
    if ($o -match 'Temporary failure in name resolution|Failed to establish a new connection|getaddrinfo failed|Network is unreachable|Connection reset') {
        return @{ Why = 'the download could not reach pypi.org - the internet connection dropped or is blocked.'
                  Fix = @('Check that this PC is online (open pypi.org in a browser), then run the installer again.') }
    }
    if ($o -match 'No matching distribution found|Could not find a version that satisfies') {
        $pkg = ([regex]::Match($o, 'for ([A-Za-z0-9_.\-]+)')).Groups[1].Value
        return @{ Kind = 'python-mismatch'
                  Why = ('no ready-made build of ' + $pkg + ' exists for this Python version.')
                  Fix = @('Install Python 3.12 - every package this app needs has a build for it:',
                          '  https://www.python.org/downloads/windows/',
                          'Then delete the .venv folder in this project and run the installer again.') }
    }
    if ($o -match 'No space left on device|not enough space|ENOSPC') {
        return @{ Why = 'the disk is full.'
                  Fix = @('Free up a couple of gigabytes on this drive and run the installer again.') }
    }
    if ($o -match 'Access is denied|WinError 5|Permission denied|being used by another process') {
        return @{ Why = 'a file could not be written - the app is probably still running, or antivirus locked it.'
                  Fix = @('Close the app (check the tray and Task Manager), then run the installer again.') }
    }
    if ($o -match 'is not recognized as an internal or external command|cannot find the path') {
        return @{ Why = 'the Python installation looks broken - pip is missing from it.'
                  Fix = @('Repair Python from Settings > Apps > Python > Modify > Repair,',
                          'or reinstall it from https://www.python.org/downloads/windows/') }
    }
    return @{ Why = 'pip stopped with an error (the full text is at the end of install.log).'
              Fix = @('Read the last lines of install.log in this folder - they name the package that failed.',
                      'A retry often works if the cause was a dropped connection.') }
}

function Invoke-Pip { param([string]$PyExe, [string[]]$PyPrefix, [string[]]$PipArgs, [string]$What)
    $attempts = @(
        @{ Label = ''                       ; Extra = @() },
        @{ Label = 'retrying with a longer timeout'          ; Extra = @('--timeout', '60', '--retries', '5') },
        @{ Label = 'retrying with the certificate check relaxed' ; Extra = @('--timeout', '60', '--retries', '5',
                       '--trusted-host', 'pypi.org', '--trusted-host', 'files.pythonhosted.org', '--trusted-host', 'pypi.python.org') },
        @{ Label = 'retrying with pre-built wheels only'     ; Extra = @('--only-binary', ':all:', '--timeout', '60') }
    )
    $last = $null
    for ($i = 0; $i -lt $attempts.Count; $i++) {
        $a = $attempts[$i]
        if ($a.Label) { Info $a.Label }
        $r = Run -Exe $PyExe -Arguments ($PyPrefix + @('-m', 'pip') + $PipArgs + $a.Extra + @('--disable-pip-version-check'))
        if ($r.Code -eq 0) {
            if ($i -gt 0) { Fixed ($What + ' (' + $a.Label + ')') } else { OK $What }
            return
        }
        $last = $r.Output
        # A compiler error or a genuinely missing package will not be fixed by a retry.
        if ($last -match 'No matching distribution found|Could not find a version that satisfies') { break }
        if ($i -eq 0) { Warn ($What + ' failed on the first try.') }
    }
    $diag = Explain-PipFailure -Output $last
    $kind = ''
    if ($diag.ContainsKey('Kind')) { $kind = $diag.Kind }
    Stop-Install ($What + ' failed: ' + $diag.Why) $diag.Fix -Kind $kind
}

# --------------------------------------------------------------------------
#  npm, with the same treatment
# --------------------------------------------------------------------------

function Explain-NpmFailure { param([string]$Output)
    $o = $Output
    if ($o -match 'EBADENGINE|Unsupported engine') {
        return @{ Why = 'the installed Node.js version is not supported by this project.'
                  Fix = @('Install the current Node.js LTS from https://nodejs.org/en/download and run the installer again.') }
    }
    if ($o -match 'ERESOLVE') {
        return @{ Why = 'two packages asked for conflicting versions of the same dependency.'
                  Fix = @('The installer already retried with --legacy-peer-deps.',
                          'If it still fails, delete node_modules and package-lock.json, then run the installer again.') }
    }
    if ($o -match 'ENOENT.*package-lock|can only install packages when your package.json and package-lock.json') {
        return @{ Why = 'package-lock.json does not match package.json, so the exact-versions install refused to run.'
                  Fix = @('The installer already retried with "npm install".',
                          'If it still fails, delete node_modules and package-lock.json and run the installer again.') }
    }
    if ($o -match 'EACCES|EPERM|operation not permitted') {
        return @{ Why = 'npm could not write into the project folder.'
                  Fix = @('Close any editor or terminal using this folder, then run the installer again.',
                          'If the folder is inside OneDrive, pause OneDrive syncing while installing.') }
    }
    if ($o -match 'ENOTFOUND|ETIMEDOUT|ECONNRESET|network|EAI_AGAIN') {
        return @{ Why = 'npm could not reach the package registry.'
                  Fix = @('Check that this PC is online, then run the installer again.',
                          'Behind a company proxy:  npm config set proxy http://proxy.company.com:8080') }
    }
    if ($o -match 'ENOSPC|no space left') {
        return @{ Why = 'the disk is full.'
                  Fix = @('Free up a few gigabytes on this drive and run the installer again.') }
    }
    return @{ Why = 'npm stopped with an error (the full text is at the end of install.log).'
              Fix = @('Read the last lines of install.log in this folder - npm names the failing package.') }
}

function Invoke-Npm { param([string]$Dir, [string]$What)
    $npm = Get-Npm
    if (-not $npm) { Stop-Install 'npm was not found even though Node.js is installed.' @('Reinstall Node.js from https://nodejs.org/en/download') }

    $hasLock = Test-Path (Join-Path $Dir 'package-lock.json')
    $attempts = New-Object System.Collections.ArrayList
    if ($hasLock) { [void]$attempts.Add(@{ Label = ''; Args = @('ci', '--no-audit', '--no-fund') }) }
    [void]$attempts.Add(@{ Label = 'retrying with "npm install"'; Args = @('install', '--no-audit', '--no-fund') })
    [void]$attempts.Add(@{ Label = 'retrying with --legacy-peer-deps'; Args = @('install', '--no-audit', '--no-fund', '--legacy-peer-deps') })

    $last = $null
    for ($i = 0; $i -lt $attempts.Count; $i++) {
        $a = $attempts[$i]
        if ($a.Label) { Info $a.Label }
        $r = Run -Exe $npm -Arguments $a.Args -WorkDir $Dir
        if ($r.Code -eq 0) {
            if ($i -gt 0) { Fixed ($What + ' (' + $a.Label + ')') } else { OK $What }
            return
        }
        $last = $r.Output
        if ($i -eq 0) { Warn ($What + ' failed on the first try.') }
    }
    $diag = Explain-NpmFailure -Output $last
    Stop-Install ($What + ' failed: ' + $diag.Why) $diag.Fix
}

# --------------------------------------------------------------------------
#  Virtual environment
# --------------------------------------------------------------------------

function Ensure-Venv { param($Python, [string]$VenvDir)
    $full   = Join-Path $Root $VenvDir
    $pyExe  = Join-Path $full 'Scripts\python.exe'
    if (Test-Path $pyExe) {
        # Reuse it only if it is genuinely healthy: an environment layered on a
        # conda interpreter often has no working ssl, and pip would then fail
        # with a misleading certificate error.
        $probe = Test-PythonCandidate -Exe $pyExe
        if ($probe -and $probe.Healthy) {
            OK ('existing environment reused  (' + $VenvDir + ', Python ' + $probe.Version + ')')
            return $pyExe
        }
        if ($probe) { Warn ('the existing ' + $VenvDir + ' environment is unusable: ' + $probe.Problem) }
        else { Warn ('the existing ' + $VenvDir + ' folder is broken.') }
        Info 'deleting and rebuilding it...'
        Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $full) {
            Stop-Install ('the old ' + $VenvDir + ' folder could not be deleted.') @(
                'Something is still using it. Close the app and any open terminals,',
                ('then delete this folder by hand and run the installer again:  ' + $full))
        }
        Fixed 'removed the broken environment'
    }

    Info ('creating a private Python environment in ' + $VenvDir + ' ...')
    $r = Run -Exe $Python.Exe -Arguments ($Python.Prefix + @('-m', 'venv', $full))
    if ($r.Code -ne 0) {
        if ($r.Output -match 'ensurepip|No module named venv') {
            Stop-Install 'this Python installation is missing the "venv" module.' @(
                'Repair Python: Settings > Apps > Python > Modify > Repair,',
                'or reinstall from https://www.python.org/downloads/windows/')
        }
        if ($r.Output -match 'Access is denied|Permission') {
            Stop-Install 'the environment folder could not be created - permission denied.' @(
                'Move the project somewhere under your own user folder (for example the Desktop),',
                'or right-click Install and choose "Run as administrator".')
        }
        Stop-Install 'the private Python environment could not be created.' @(
            'The reason is at the end of install.log in this folder.')
    }
    if (-not (Test-Path $pyExe)) {
        Stop-Install 'the private Python environment was created but has no python.exe in it.' @(
            'Delete the ' + $VenvDir + ' folder and run the installer again.')
    }
    OK ('private Python environment created  (' + $VenvDir + ')')
    return $pyExe
}

# --------------------------------------------------------------------------
#  Shortcuts and launchers
# --------------------------------------------------------------------------

function New-Launcher { param([string]$Name, [string]$Body)
    $path = Join-Path $Root $Name
    if ($script:VenvPy) { $Body = $Body.Replace('__PYTHON__', $script:VenvPy) }
    $text = ($Body -replace "`r`n", "`n") -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
    OK ('launcher written  (' + $Name + ')')
    return $path
}

function New-Shortcut { param([string]$Name, [string]$Target, [string]$Icon = '', [string]$Description = '', [string]$Arguments = '')
    $made = New-Object System.Collections.ArrayList
    try {
        $shell = New-Object -ComObject WScript.Shell
        $places = @($shell.SpecialFolders('Desktop'), $shell.SpecialFolders('Programs'))
        foreach ($dir in $places) {
            if (-not $dir) { continue }
            $lnk = $shell.CreateShortcut((Join-Path $dir ($Name + '.lnk')))
            $lnk.TargetPath       = $Target
            $lnk.Arguments        = $Arguments
            $lnk.WorkingDirectory = $Root
            $lnk.Description      = $Description
            if ($Icon -and (Test-Path $Icon)) { $lnk.IconLocation = $Icon }
            $lnk.Save()
            [void]$made.Add((Join-Path $dir ($Name + '.lnk')))
        }
    } catch {
        Warn ('the Desktop shortcut could not be created: ' + $_.Exception.Message) `
             ('Start the app from ' + (Split-Path -Leaf $Target) + ' in this folder instead.')
        return
    }
    foreach ($m in $made) { OK ('shortcut: ' + $m) }
}

# --------------------------------------------------------------------------
#  .env handling
# --------------------------------------------------------------------------

function Ensure-EnvKey { param([hashtable]$Spec)
    $file = Join-Path $Root $Spec.File
    $text = ''
    if (Test-Path $file) { $text = Get-Content -LiteralPath $file -Raw }

    $pattern = '(?m)^\s*' + [regex]::Escape($Spec.Key) + '\s*=\s*(.*)$'
    $m = [regex]::Match($text, $pattern)
    if ($m.Success) {
        $val = $m.Groups[1].Value.Trim().Trim('"').Trim("'")
        if ($val -and $val -notmatch '^(your|paste|changeme|xxx|<)' ) {
            OK ($Spec.Key + ' is already set in ' + $Spec.File)
            return
        }
    }

    Warn ($Spec.Key + ' is not set yet, so the app cannot talk to its API.') `
         ('Get a key here: ' + $Spec.Url)
    $value = ''
    if (-not $NonInteractive) {
        Line ''
        Line ('      Paste your ' + $Spec.Prompt + ' and press Enter.') 'White'
        Line  '      (Press Enter on its own to skip - you can fill it in later.)' 'DarkGray'
        Write-Host '      > ' -NoNewline -ForegroundColor White
        $value = (Read-Host).Trim().Trim('"')
    }
    if (-not $value) {
        Warn ($Spec.Key + ' was left empty.') `
             ('Open ' + $Spec.File + ' in Notepad later and put your key after ' + $Spec.Key + '=')
        if (-not (Test-Path $file)) {
            [System.IO.File]::WriteAllText($file, ($Spec.Key + "=`r`n"), (New-Object System.Text.UTF8Encoding($false)))
            Info ('created an empty ' + $Spec.File + ' for you')
        }
        return
    }
    if ($m.Success) { $text = [regex]::Replace($text, $pattern, ($Spec.Key + '=' + $value)) }
    else {
        if ($text -and -not $text.EndsWith("`n")) { $text += "`r`n" }
        $text += ($Spec.Key + '=' + $value + "`r`n")
    }
    [System.IO.File]::WriteAllText($file, $text, (New-Object System.Text.UTF8Encoding($false)))
    Fixed ($Spec.Key + ' saved into ' + $Spec.File)
}

# --------------------------------------------------------------------------
#  Import / smoke checks
# --------------------------------------------------------------------------

function Test-Imports { param([string]$PyExe, [string[]]$Modules)
    if (-not $Modules -or $Modules.Count -eq 0) { return }
    $bad = New-Object System.Collections.ArrayList
    foreach ($mod in $Modules) {
        $r = Run -Exe $PyExe -Arguments @('-c', ('import ' + $mod))
        if ($r.Code -ne 0) { [void]$bad.Add($mod) } 
    }
    if ($bad.Count -eq 0) {
        if ($Modules.Count -eq 1) { OK 'the required package imports cleanly' }
        else { OK ('all ' + $Modules.Count + ' required packages import cleanly') }
        return
    }
    Stop-Install ('these packages installed but will not load: ' + ($bad -join ', ')) @(
        'Delete the .venv folder in this project and run the installer again -',
        'that rebuilds the environment from scratch and fixes almost every case of this.')
}

# --------------------------------------------------------------------------
#  Main
# --------------------------------------------------------------------------

$startedAt = Get-Date
try {
    if (Test-Path $LogFile) { Move-Item -LiteralPath $LogFile -Destination ($LogFile + '.old') -Force -ErrorAction SilentlyContinue }
    Log ('installer started for ' + $App.Name)
    Log ('PowerShell ' + $PSVersionTable.PSVersion + ' on ' + [System.Environment]::OSVersion.VersionString)
    Log ('project root: ' + $Root)

    Write-Host ''
    Box -Lines (@(($App.Name + '  -  Installer'), $App.Blurb)) -Colour 'Cyan'
    Line ''
    Line ('  This window installs everything the app needs. Nothing leaves your PC') 'DarkGray'
    Line ('  except package downloads. A full record is written to install.log.') 'DarkGray'

    Step 'Checking this computer'
    if ($App.WindowsOnly -and -not ($env:OS -eq 'Windows_NT')) {
        Stop-Install 'this app only runs on Windows.' @('On Linux or macOS use install.sh instead.')
    }
    OK ('Windows ' + [System.Environment]::OSVersion.Version)
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Stop-Install 'PowerShell 5.1 or newer is required.' @('Install Windows Management Framework 5.1 from Microsoft, then try again.')
    }
    OK ('PowerShell ' + $PSVersionTable.PSVersion)
    $free = [math]::Round((Get-PSDrive -Name ($Root.Substring(0,1)) -ErrorAction SilentlyContinue).Free / 1GB, 1)
    if ($free -and $free -lt 1) { Warn ('only ' + $free + ' GB free on this drive.') 'Free up some space if the install fails part way.' }
    OK ('project folder: ' + $Root)

    $venvPy = $null
    if ($App.MinPython) {
        Step ('Looking for Python ' + $App.MinPython + ' or newer')
        $python = Find-Python -Min ([version]$App.MinPython)
        if (-not $python) { $python = Install-Python -Min ([version]$App.MinPython) }
        OK ('using Python ' + $python.Version)

        # One body of work, so it can be run a second time on a different Python.
        $installPackages = {
            param($Py)

            if ($App.Venv) {
                Step 'Preparing the private Python environment'
                $script:VenvPy = Ensure-Venv -Python $Py -VenvDir $App.Venv
            } else {
                # No environment of our own: pin down the exact interpreter, so
                # later steps cannot land on a different one than the one chosen.
                $script:VenvPy = $Py.Exe
                if ($Py.Prefix.Count) {
                    $r = Run -Exe $Py.Exe -Arguments ($Py.Prefix + @('-c', 'import sys;print(sys.executable)'))
                    $line = (($r.Output -split "`n") | Where-Object { $_ -match '\.exe\s*$' } | Select-Object -First 1)
                    if ($line) { $script:VenvPy = $line.Trim() }
                }
                OK ('using ' + $script:VenvPy)
            }

            if ($App.Requirements -or ($App.PipExtra -and $App.PipExtra.Count)) {
                Step 'Updating the Python package tools'
                $r = Run -Exe $script:VenvPy -Arguments @('-m', 'pip', 'install', '--upgrade', 'pip', 'setuptools', 'wheel', '--disable-pip-version-check')
                if ($r.Code -eq 0) { OK 'pip, setuptools and wheel are up to date' }
                else { Warn 'pip could not be updated - carrying on with the version that is there.' 'Harmless unless the next step also fails.' }
            }

            if ($App.Requirements -or ($App.PipExtra -and $App.PipExtra.Count)) {
                Step 'Installing the Python packages'
                Info 'first run downloads a few megabytes - give it a minute.'
                if ($App.Requirements) {
                    $req = Join-Path $Root $App.Requirements
                    if (-not (Test-Path $req)) {
                        Stop-Install ($App.Requirements + ' is missing from this folder.') @(
                            'This download is incomplete. Get the whole project again:',
                            ('  git clone ' + $App.Repo))
                    }
                    Invoke-Pip -PyExe $script:VenvPy -PyPrefix @() -PipArgs @('install', '-r', $req) -What ('packages from ' + $App.Requirements)
                }
                if ($App.PipExtra -and $App.PipExtra.Count) {
                    Invoke-Pip -PyExe $script:VenvPy -PyPrefix @() -PipArgs (@('install') + $App.PipExtra) -What ('packages: ' + ($App.PipExtra -join ', '))
                }
            }
        }

        try {
            & $installPackages $python
        } catch {
            $kind = ''
            if ($_.Exception.Data) { $kind = [string]$_.Exception.Data['kind'] }
            if ($kind -ne 'python-mismatch') { throw }

            # The chosen Python is too new (or too old) for one of the packages.
            # Fetching the version that does have ready-made builds is a fix the
            # installer can apply on its own, so try that before giving up.
            Warn ('Python ' + $python.Version + ' is too new for one of the packages this app needs.')
            Info 'trying to fix that by installing Python 3.12 alongside it...'
            $better = Install-Python -Min ([version]$App.MinPython) -Series '3.12' -Quiet
            if (-not $better -or $better.Version -eq $python.Version) { throw }
            Fixed ('switching the install over to Python ' + $better.Version)
            $old = Join-Path $Root $App.Venv
            if ($App.Venv -and (Test-Path $old)) { Remove-Item -LiteralPath $old -Recurse -Force -ErrorAction SilentlyContinue }
            & $installPackages $better
        }
        $venvPy = $script:VenvPy

        if ($App.VerifyImports -and $App.VerifyImports.Count) {
            Step 'Checking that every package really works'
            Test-Imports -PyExe $venvPy -Modules $App.VerifyImports
        }
    }

    if ($App.MinNode) {
        Step ('Looking for Node.js ' + $App.MinNode + ' or newer')
        Ensure-Node -Min ([version]$App.MinNode) | Out-Null

        foreach ($nd in $App.NpmDirs) {
            $dir = Join-Path $Root $nd.Path
            if (-not (Test-Path (Join-Path $dir 'package.json'))) {
                Warn ('no package.json in ' + $nd.Path + ' - skipping.') 'Re-clone the project if you expected one here.'
                continue
            }
            Step ('Installing the Node packages for ' + $nd.Label)
            Info 'this is the slow step - a few minutes on a first run.'
            Invoke-Npm -Dir $dir -What ($nd.Label + ' dependencies')
        }
    }

    if ($App.EnvKeys -and $App.EnvKeys.Count) {
        Step 'Setting up your API key'
        foreach ($spec in $App.EnvKeys) { Ensure-EnvKey -Spec $spec }
    }

    if ($App.ExtraSteps) {
        foreach ($extra in $App.ExtraSteps) {
            Step $extra.Title
            & $extra.Action $venvPy
        }
    }

    if ($App.Launchers) {
        Step 'Writing the launchers'
        foreach ($l in $App.Launchers) { New-Launcher -Name $l.Name -Body $l.Body | Out-Null }
    }

    if ($App.Shortcut) {
        Step 'Putting a shortcut on your Desktop and Start menu'
        $target = Join-Path $Root $App.Shortcut.Target
        if (-not (Test-Path $target) -and $App.Shortcut.Fallback) {
            $target = Join-Path $Root $App.Shortcut.Fallback
            Info ('pointing the shortcut at ' + $App.Shortcut.Fallback + ' instead')
        }
        if (-not (Test-Path $target)) {
            Warn ('the shortcut target is missing: ' + $target) 'Start the app from this folder instead.'
        } else {
            $icon = ''
            if ($App.Shortcut.Icon) { $icon = Join-Path $Root $App.Shortcut.Icon }
            New-Shortcut -Name $App.Shortcut.Name -Target $target -Icon $icon -Description $App.Blurb
        }
    }

    # ---------------------------------------------------------------- summary
    $secs = [int]((Get-Date) - $startedAt).TotalSeconds
    Write-Host ''
    Box -Lines (@(('INSTALL COMPLETE  -  ' + $App.Name), ('finished in ' + $secs + ' seconds'))) -Colour 'Green'

    if ($script:Fixed.Count) {
        Write-Host ''
        Line '  Problems found and fixed along the way:' 'Green'
        foreach ($f in $script:Fixed) { Line ('    - ' + $f) 'Green' }
    }
    if ($script:Warnings.Count) {
        Write-Host ''
        Line '  Warnings - the app will run, but read these:' 'Yellow'
        foreach ($w in $script:Warnings) {
            Line ('    - ' + $w.Text) 'Yellow'
            if ($w.Fix) { Line ('      ' + $w.Fix) 'DarkYellow' }
        }
    }

    Write-Host ''
    Line '  How to start it:' 'White'
    foreach ($n in $App.NextSteps) { Line ('    ' + $n) 'Gray' }
    Write-Host ''
    Line ('  Full log: ' + $LogFile) 'DarkGray'
    Write-Host ''
    exit 0

} catch {
    $msg = $_.Exception.Message
    $fix = ''
    if ($_.Exception.Data -and $_.Exception.Data['fix']) { $fix = [string]$_.Exception.Data['fix'] }
    Log ('FAILED: ' + $msg)
    if ($fix) { Log ('FIX: ' + $fix) }
    if ($_.ScriptStackTrace) { Log $_.ScriptStackTrace }

    Write-Host ''
    Box -Lines (@('INSTALL STOPPED  -  nothing is broken, it just did not finish')) -Colour 'Red'
    Write-Host ''
    Line '  What went wrong:' 'Red'
    foreach ($l in ($msg -split "`n")) { Line ('    ' + $l.Trim()) 'White' }
    if ($fix) {
        Write-Host ''
        Line '  How to fix it:' 'Yellow'
        foreach ($l in ($fix -split "`n")) { Line ('    ' + $l) 'Gray' }
    }
    Write-Host ''
    Line '  Still stuck?' 'White'
    Line ('    1. The full technical log is at  ' + $LogFile) 'Gray'
    Line  '    2. See the Troubleshooting table in INSTALL.md' 'Gray'
    if ($App.Issues) { Line ('    3. Open an issue with install.log attached:  ' + $App.Issues) 'Gray' }
    Write-Host ''
    exit 1
}
