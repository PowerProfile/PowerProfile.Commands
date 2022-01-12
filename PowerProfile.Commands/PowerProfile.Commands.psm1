#Requires -Version 5.1
#Requires -Modules @{ ModuleName='PowerProfile'; ModuleVersion='1.0.0' }

function New-PowerProfile {

<#
.SYNOPSIS
    Creates profile directories for PowerProfile.

.DESCRIPTION
    Creates profile directories, based on current system environment.
    Run this command in every terminal on any device where you would
    like to have individual settings.

.LINK
    https://www.PowerProfile.sh/
#>

    [CmdletBinding(PositionalBinding=$false,SupportsShouldProcess,ConfirmImpact='Low')]
    Param(
        [switch]${UserProfile},
        [switch]${PSHostProfile},
        [switch]${TerminalProfile},

        [switch]${AllConditionalDirectories},
        [switch]${Architecture},
        [switch]${Machine},
        [switch]${Platform},
        [ValidateSet('Core','Desktop','All')]
        [string]${PSEditions}='None',

        [switch]${ConfigFolder},
        [switch]${FunctionsFolder},
        [switch]${ModulesFolder}
    )

    if (-Not $UserProfile -and -not $PSHostProfile -and -not $TerminalProfile) {
        $UserProfile = $true
        $PSHostProfile = $true
        $TerminalProfile = $true
    }

    if ($AllConditionalDirectories) {
        if (-Not $Architecture) {
            $Architecture = $true
        }
        if (-Not $Machine) {
            $Machine = $true
        }
        if (-Not $Platform) {
            $Platform = $true
        }
        if ($PSEditions -eq 'None') {
            $PSEditions = 'All'
        }
    }

    if (-Not $ConfigFolder -and -not $FunctionsFolder -and -not $ModulesFolder) {
        $FunctionsFolder = $true
    }

    #TODO
    function New-PoProfileContentFolders {
        Param(
            [string]${Path},
            [switch]${FunctionsFolder},
            [switch]${ModulesFolder}
        )

        $subdirs = @()
        if (${FunctionsFolder}) {
            $subdirs += 'Functions'
        }
        if (${ModulesFolder}) {
            $subdirs += 'Modules'
        }

        foreach (${subdir} in ${subdirs}) {
            $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction Ignore -Confirm:$false
            return $subdirs
        }
    }

    Write-Host "    Creating profile directories in: $PROFILEHOME"

    $ProfilesList = Get-PoProfileProfilesList
    $List = @()

    if ($UserProfile) {
        $List += $ProfilesList[0]
    }
    if ($PSHostProfile) {
        $List += $ProfilesList[1]
    }
    if (
        $TerminalProfile -and
        ($null -ne $ProfilesList[2])
    ) {
        $List += $ProfilesList[2]
    }

    foreach ($ProfileName in $List) {
        $ProfileDir = [System.IO.Path]::Combine($PROFILEHOME,$ProfileName)
        if ([System.IO.Directory]::Exists($ProfileDir)) {
            $color = $PSStyle.Foreground.BrightWhite + $PSStyle.Bold
        } else {
            $color = $PSStyle.Foreground.BrightGreen + $PSStyle.Bold
        }
        $txt = Write-PoProfilePathTree -Path @($ProfileName)

        if ($PSCmdlet.ShouldProcess($txt, $ProfileDir, 'Create profile directory')) {
            $null = New-Item -ItemType Directory -Path $ProfileDir -ErrorAction Ignore -Confirm:$false
            Write-Host ([System.Environment]::NewLine + "${color}${txt}" + $PSStyle.Reset)
        }

        $PoProfileSubDirs = Get-PoProfileSubDirs -Name $ProfileName -Platform $Platform -Architecture $Architecture -Machine $Machine -PSEditions $PSEditions
        if ($PoProfileSubDirs) {
            $LastSubDir = $PoProfileSubDirs[-1] -split [IO.Path]::DirectorySeparatorChar
            $i = 0
            foreach ($ProfileSubDir in $PoProfileSubDirs) {
                $ProfileSubDirPath = [System.IO.Path]::Combine($ProfileDir,$ProfileSubDir)
                if ([System.IO.Directory]::Exists($ProfileSubDirPath)) {
                    $color = $PSStyle.Foreground.White
                } else {
                    $color = $PSStyle.Foreground.Green
                }
                $ProfileSubDirSegs = $ProfileSubDir -split [IO.Path]::DirectorySeparatorChar
                $Prequel = if ($i -eq 0) { $ProfileName } else { $PoProfileSubDirs[$i-1] -split [IO.Path]::DirectorySeparatorChar }
                $Sequel = if (($i+1) -lt $PoProfileSubDirs.Count) { $PoProfileSubDirs[$i+1] -split [IO.Path]::DirectorySeparatorChar }
                $txt = Write-PoProfilePathTree -Path @($ProfileSubDirSegs) -LastPath $LastSubDir -Prequel @($Prequel) -Sequel @($Sequel)
                if ($PSCmdlet.ShouldProcess($txt, $ProfileSubDirPath, 'Create conditional profile sub-directory')) {
                    $null = New-Item -ItemType Directory -Path $ProfileSubDirPath -ErrorAction Ignore -Confirm:$false
                    Write-Host ("${color}${txt}" + $PSStyle.Reset)
                }
                $i++
            }
        }
    }
}

function Optimize-PowerProfile {
    [CmdletBinding(SupportsShouldProcess)]
    Param()

    if ($null -eq $PROFILEHOME) {
        throw 'Missing global variable PROFILEHOME'
    }

    Get-ChildItem -Path $PROFILEHOME -Directory -Recurse | Sort-Object { $_.FullName.Split([System.IO.Path]::DirectorySeparatorChar).Count } -Descending | ForEach-Object {
        if ((Test-Path -Path $_.FullName) -and ($null -eq (Get-ChildItem -Path $_.FullName -FollowSymlink))) {
            if ($PSCmdlet.ShouldProcess($_.FullName,'delete')) {
                Remove-Item -Path $_.FullName -Confirm:$false
            }
        }
    }
}

function Write-PoProfilePathTree {
    Param(
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Prequel,

        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Sequel,

        [Parameter(Mandatory=$true,Position=0)]
        [System.Collections.ArrayList]$Path,

        [System.Collections.ArrayList]$LastPath=$Path,

        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Parent,

        [Int]$Depth=1,

        [string]$Prefix='',

        [string]$Indent='    '
    )

    if ($Prequel) {
        if ($Depth -gt 1) {
            if ($Sequel -and ($Sequel.Count -ge $Path.Count)) {
                $node = '┼──'
            } else {
                $node = '└──'
            }
        } else {
            if ($Sequel -and ($Sequel.Count -ge $Path.Count)) {
                if (($LastPath[0] -eq $Path[0]) -and ($Path.Count -eq 1)) {
                    $node = '└──'
                } else {
                    $node = '├──'
                }
            } else {
                $node = '└──'
            }
        }
    }

    if ($node) {
        if ($Depth -gt 1) {
            if ($LastPath[0] -eq $Parent[0]) {
                $Prefix = '   ' + $node
            } else {
                $Prefix = '│  ' + $node
            }
        } else {
            $Prefix = $Prefix + $node
        }
    }

    $icon = '📁'
    if ($Path[0] -match '^Profile')     {$icon = '🗂 '}
    if ($Path[0] -match '^_Arch')       {$icon = '🏢'}
    if ($Path[0] -match '^_Machine')    {$icon = '💻'}
    if ($Path[0] -match '^_Platform')   {$icon = '💾'}
    if ($Path[0] -match '^_PSEdition')  {$icon = '🐚'}
    if (-Not $Prequel -or ($Prequel[$Depth-1] -ne $Path[0])) {
        $return += "${Indent}${Prefix}$icon " + $Path[0]
    }

    if ($Path.Count -gt 1) {
        [System.Collections.ArrayList]$Parent += ,$Path[0]
        $null = $Path.RemoveAt(0)
        $sreturn = Write-PoProfilePathTree -Path $Path -Parent $Parent -Depth $($Depth+1) -Prefix $Prefix -LastPath $LastPath -Indent $Indent
        if ($sreturn) {
            $return = $return + $sreturn
        }
    }

    return $return
}

function Uninstall-PowerProfile {
    Param(
        # Also delete installed PowerProfile modules
        [switch]$DeleteModules
    )

    $ProfilePS1 = $PROFILE.CurrentUserAllHosts

    Write-PoProfileProgress -ProfileTitle "`nUninstalling PowerProfile" -ScriptCategory ''

    Reset-PoProfileState -Force

    if (
        (Get-FileHash -Path (Join-Path (Split-Path (Get-Module -Name PowerProfile).Path) 'profile.ps1')).Hash -ne
        (Get-FileHash -Path $ProfilePS1).Hash
    ) {
        Remove-Item -Force -Path $ProfilePS1 -ErrorAction Ignore -Confirm:$false

        $OrigProfilePS1 = Get-ChildItem -File -Path (Split-Path $ProfilePS1) -Filter 'profile.ps1.bak.*' | Sort-Object -Property Name -Descending | Select-Object -First 1
        if ($OrigProfilePS1) {
            Move-Item $OrigProfilePS1.FullName $ProfilePS1 -ErrorAction Ignore -Confirm:$false
            Write-PoProfileProgress -ScriptTitleType Confirmation -NoCounter -ScriptTitle @(
                "Your original $($PSStyle.FormatHyperlink((Split-Path -Leaf $ProfilePS1),'file://'+$ProfilePS1)) was restored successfully."
            )
        }
    } else {
        $bak = $ProfilePS1 + '.bak.' + (Get-Date).ToString('yyyy-MM-dd_HHmmss')
        Rename-Item $ProfilePS1 $bak

        Write-PoProfileProgress -ScriptTitleType Information -ScriptTitle @(
            $PSStyle.Italic + (Split-Path -Leaf $ProfilePS1) + $PSStyle.ItalicOff + ' backed up:'
            "The file contains changes outside of PowerProfile and was renamed to $($PSStyle.FormatHyperlink((Split-Path -Leaf $bak),'file://'+(Split-Path $bak)))"
            'for further investigation.'
        )
    }

    if ($DeleteModules) {
        $p = Join-Path (Split-Path $ProfilePS1) 'Modules' 'PowerProfile*'
        foreach ($m in (Get-ChildItem -Directory -Path $p)) {
            if (-not [bool]($m.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                Remove-Item -Recurse -Path $m.FullName -Confirm:$false
            }
        }
        Write-PoProfileProgress -ScriptTitleType Confirmation -NoCounter -ScriptTitle @(
            "All PowerProfile modules in $($PSStyle.FormatHyperlink('profile modules folder','file://'+(Split-Path $p))) were deleted."
        )
    }

    Remove-Module -Force -Name PowerProfile* -ErrorAction Ignore -Confirm:$false

    function Global:prompt {
        "PS $($executionContext.SessionState.Path.CurrentLocation)$('>' * ($nestedPromptLevel + 1)) ";
        # .Link
        # https://go.microsoft.com/fwlink/?LinkID=225750
        # .ExternalHelp System.Management.Automation.dll-help.xml
    }
}
