#Requires -Version 5.1

function Set-PSConfig {

<#
.SYNOPSIS
    Add/change/remove a PowerShell configuration item

.DESCRIPTION
    Adds an item to the PowerShell configuration

.PARAMETER Name
    Key name of the configuration item

.PARAMETER Value
    Value of the configuration item

.PARAMETER Scope
    Set the AllUsers (shared/LocalMachine) configuration instead of
    CurrentUser (per-user) configuration.

.PARAMETER Remove
    Removes the desired configuration item

.PARAMETER Registry
    Enforce to use the configuration from Windows Registry.

.PARAMETER Json
    Enforce to use the configuration from powershell.config.json.

.PARAMETER PassThru
    Returns an object with the PowerShell configuration. By default, this function does not generate any output.

.INPUTS
    PSObject

.OUTPUTS
    PSObject. Only when using -PassThru parameter.

.LINK
    https://PowerProfile.sh/
#>

    [CmdletBinding(DefaultParameterSetName='SetKey')]
    Param(
        [Parameter(Mandatory=$True,Position=0,ParameterSetName='SetKey')]
        [Parameter(Mandatory=$True,Position=0,ParameterSetName='RemoveKey')]
        [ValidateScript({
            ($_ -cmatch '\.ExecutionPolicy$') -or
            ($_ -ceq 'PSModulePath') -or
            ($_ -ceq 'ExperimentalFeatures') -or
            ($_ -ceq 'LogIdentity') -or
            ($_ -ceq 'LogLevel') -or
            ($_ -ceq 'LogChannels') -or
            ($_ -ceq 'LogKeywords') -or
            ($_ -ceq 'WindowsPowerShellCompatibilityModuleDenyList') -or
            $(throw "Unknown PowerShell configuration key '$_'. Valid keys are: *.ExecutionPolicy PSModulePath ExperimentalFeatures LogIdentity LogLevel LogChannels LogKeywords")
        })]
        [string]$Name,

        [Parameter(Mandatory=$True,Position=1,ParameterSetName='SetKey')]
        [AllowEmptyString()]
        [AllowNull()]
        $Value,

        [Parameter(ParameterSetName='SetKey')]
        [Parameter(ParameterSetName='RemoveKey')]
        [ValidateSet('CurrentUser','AllUsers')]
        [string]${Scope}='CurrentUser',

        [Parameter(Mandatory=$True,ParameterSetName='RemoveKey')]
        [switch]$Remove,

        [switch]$Registry,
        [switch]$Json,
        [switch]$PassThru
    )

    if ($Registry -and $Json) {
        if ($IsCoreCLR -or -not $IsWindows) {
            Write-Verbose 'Implicitly enforcing JSON configuration'
            $Registry = $false
        } else {
            Write-Verbose 'Implicitly enforcing Registry configuration'
            $Json = $false
        }
    }
    elseif ($IsWindows -and -not $IsCoreCLR -and -not $Registry -and -not $Json) {
        Write-Verbose 'Windows PowerShell detected, assuming Registry configuration'
        $Registry = $true
    }

    if ($Registry) {
        if (-Not $IsWindows) {
            throw 'Registry is only available on Windows machines.'
        }

        if ($Scope -eq 'AllUsers') {
            $path = 'Registry::HKEY_LOCAL_MACHINE'
        } else {
            $path = 'Registry::HKEY_CURRENT_USER'
        }
        $path += '\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell\'

        if (-Not (Test-Path $path)) {
            Write-Verbose "Creating new path $path"
            $null = New-Item -Path $path -Force -ErrorAction Stop
        }
        $null = New-ItemProperty -Path $path -Name $Name -Value $Value -Force -ErrorAction Stop

        if ($PassThru) {
            return (Get-PSConfig -Scope $Scope -Registry)
        }
    }
    else {
        if (-Not $IsCoreCLR) {
            throw 'Windows PowerShell does not use powershell.config.json configuration files.'
        }

        $psconfig = Get-PSConfig -Scope $Scope -Json

        if ($psconfig -is [System.Management.Automation.PSObject]) {
            if ($Value -eq [bool]::TrueString -or $Value -eq [bool]::FalseString) {
                $Value = [System.Convert]::ToBoolean($Value)
            }
            if ($Remove) {
                $psconfig.PSObject.Properties.Remove($Name)
            }
            elseif ($null -eq $psconfig.PSObject.Properties.Item($Name)) {
                $psconfig | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
            }
            else {
                $psconfig.$Name = $Value
            }
        }
        else {
            $psconfig = New-Object PSObject
            if (-Not $Remove) {
                Write-Verbose 'Generating new configuration object'
                $psconfig | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
            }
        }

        if ($Scope -eq 'AllUsers') {
            $path = [System.IO.Path]::Combine($PSHOME,'powershell.config.json')
        } else {
            $path = [System.IO.Path]::Combine((Split-Path $PROFILE.CurrentUserCurrentHost),'powershell.config.json')
        }

        if (($psconfig.PSObject.Properties).Count -gt 0) {
            $baseDir = Split-Path -Path $path
            if (-Not ([System.IO.Directory]::Exists($baseDir))) {
                Write-Verbose "Creating new directory $baseDir"
                $null = New-Item -Type Container -Force $baseDir -ErrorAction Stop
            }
            Write-Verbose "Serializing configuration object to $path"
            ConvertTo-Json $psconfig -Compress | Set-Content -Path $path -Encoding ASCII
        } elseif ([System.IO.File]::Exists($path)) {
            Write-Verbose "Deleting empty configuration file $path"
            Remove-Item -Path $path
        }

        if ($PassThru) {
            return $psconfig
        }
    }
}

function Get-PSConfig {

<#
.SYNOPSIS
    Get PowerShell JSON configuration

.DESCRIPTION
    Reads the PowerShell configuration

.PARAMETER Name
    Return the value of a specific configuration item only

.PARAMETER Scope
    Read the AllUsers (shared/LocalMachine) configuration instead of
    CurrentUser (per-user) configuration.

.PARAMETER Registry
    Enforce to use the configuration from Windows Registry.

.PARAMETER Json
    Enforce to use the configuration from powershell.config.json.

.INPUTS
    PSObject

.OUTPUTS
    String, when specifying $Name, otherwise PSObject.

.LINK
    https://PowerProfile.sh/
#>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0)]
        [string]$Name,

        [ValidateSet('CurrentUser','AllUsers')]
        [string]${Scope}='CurrentUser',

        [switch]$Registry,
        [switch]$Json
    )

    if ($Registry -and $Json) {
        if ($IsCoreCLR -or -not $IsWindows) {
            $Registry = $false
        } else {
            $Json = $false
        }
    }
    elseif ($IsWindows -and -not $IsCoreCLR -and -not $Registry -and -not $Json) {
        $Registry = $true
    }

    if ($Registry) {
        if (-Not $IsWindows) {
            throw 'Registry is only available on Windows machines.'
        }

        if ($Scope -eq 'AllUsers') {
            $path = 'Registry::HKEY_LOCAL_MACHINE'
        } else {
            $path = 'Registry::HKEY_CURRENT_USER'
        }
        $path += '\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell\'

        if ($Name) {
            return (Get-ItemProperty -Path $path -Name $Name -ErrorAction Ignore)
        }

        $dir = Get-Item -Path $path -ErrorAction Ignore

        if ($dir) {
            $property = Get-ItemProperty -Path $path
            $psconfig = New-Object PSObject
            foreach ($Name in $dir.Property) {
                $psconfig | Add-Member -MemberType NoteProperty -Name $Name -Value $property.$Name
            }
        }
    }
    else {
        if (-Not $IsCoreCLR) {
            throw 'Windows PowerShell does not use powershell.config.json configuration files.'
        }

        if ($Scope -eq 'AllUsers') {
            $path = [System.IO.Path]::Combine($PSHOME,'powershell.config.json')
        } else {
            $path = [System.IO.Path]::Combine((Split-Path $PROFILE.CurrentUserCurrentHost),'powershell.config.json')
        }

        if ([System.IO.File]::Exists($path)) {
            $psconfig = [System.IO.File]::ReadAllText($path) | ConvertFrom-Json -ErrorAction Ignore
        }
    }

    if ($psconfig -is [System.Management.Automation.PSObject]) {
        if ($Name) {
            return $psconfig.$Name
        } else {
            return $psconfig
        }
    }

    return $null
}

function Remove-PSConfig {

<#
.SYNOPSIS
    Remove PowerShell configuration item

.DESCRIPTION
    Removes a configuration item from powershell.config.json

.PARAMETER Name
    Key name of the configuration item

.PARAMETER Scope
    Set the AllUsers (shared/LocalMachine) configuration instead of
    CurrentUser (per-user) configuration.

.PARAMETER Registry
    Enforce to use the configuration from Windows Registry.

.PARAMETER Json
    Enforce to use the configuration from powershell.config.json.

.PARAMETER PassThru
    Returns an object with the PowerShell configuration. By default, this function does not generate any output.

.PARAMETER Clear
    Deletes the powershell.config.json file to clear the entire PowerShell configuration.

.INPUTS
    PSObject

.OUTPUTS
    PSObject. Only when using -PassThru parameter.

.LINK
    https://PowerProfile.sh/
#>

    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSObject])]
    Param(
        [Parameter(Mandatory=$True,ParameterSetName='Clear')]
        [switch]$Clear,

        [Parameter(Mandatory=$True,Position=0,ParameterSetName='Remove')]
        [string]$Name,

        [Parameter(ParameterSetName='Remove')]
        [Parameter(ParameterSetName='Clear')]
        [ValidateSet('CurrentUser','AllUsers')]
        [string]${Scope}='CurrentUser',

        [Parameter(ParameterSetName='Remove')]
        [switch]$Registry,

        [Parameter(ParameterSetName='Remove')]
        [switch]$Json,

        [Parameter(ParameterSetName='Remove')]
        [switch]$PassThru
    )

    if ($Clear) {
        If ($Registry) {
            throw 'PowerShell configuration in Registry cannot be cleared at once.'
        }
        if ($Scope -eq 'AllUsers') {
            $path = [System.IO.Path]::Combine($PSHOME,'powershell.config.json')
        } else {
            $path = [System.IO.Path]::Combine((Split-Path $PROFILE.CurrentUserCurrentHost),'powershell.config.json')
        }
        if ([System.IO.File]::Exists($path)) {
            Remove-Item -Path $path -Force -ErrorAction Ignore
        }
        return
    }

    Try {
        $psconfig = Set-PSConfig -Remove -Name $Name -Scope $Scope -Registry:$Registry -Json:$Json -PassThru:$PassThru
    }
    Catch {
        Write-Error $_.Exception.Message
        return
    }

    if ($PassThru) {
        return $psconfig
    }
}
