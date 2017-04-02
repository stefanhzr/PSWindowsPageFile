<#
    The MIT License (MIT)

    Copyright (c) 2017 Stefan H

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>

#Requires -RunAsAdministrator
#Requires -Version 2.0

function Get-PageFile {
    [CmdletBinding(DefaultParameterSetName='None')]

    param(
        # Remote computer name
        [Parameter(ParameterSetName='setRemote', Mandatory=$true)]
        [string]
        $ComputerName,

        # Credentials for remote computer
        [Parameter(ParameterSetName='setRemote', Mandatory=$true)]
        [pscredential]
        $Credential
    )

    if ($PSCmdlet.ParameterSetName -eq 'setRemote') {
        Get-WmiObject -ComputerName $ComputerName -Credential $Credential -Class Win32_PageFileSetting -EnableAllPrivileges
    } else {
        Get-WmiObject -Class Win32_PageFileSetting -EnableAllPrivileges
    }
}

function Set-PageFile {
    [CmdletBinding(DefaultParameterSetName='None')]

    param(
        # Target drive for new page file
        [Parameter(Mandatory=$true)]
        [char]
        $DriveLetter,

        # Initial page file size
        [Parameter(ParameterSetName='setSize', Mandatory=$true)]
        [Parameter(ParameterSetName='setSizeRemote', Mandatory=$true)]
        [int]
        $InitialSize,

        # Maximum page file size
        [Parameter(ParameterSetName='setSize', Mandatory=$true)]
        [Parameter(ParameterSetName='setSizeRemote', Mandatory=$true)]
        [int]
        $MaximumSize,

        # Remote computer name
        [Parameter(ParameterSetName='setRemote', Mandatory=$true)]
        [Parameter(ParameterSetName='setSizeRemote', Mandatory=$true)]
        [string]
        $ComputerName,

        # Remote computer username
        [Parameter(ParameterSetName='setRemote', Mandatory=$true)]
        [Parameter(ParameterSetName='setSizeRemote', Mandatory=$true)]
        [string]
        $UserName
    )

    begin {
        # Get credentials (remote only) and disk information
        if (@('setRemote', 'setSizeRemote') -ccontains $PSCmdlet.ParameterSetName) {
            $account = Get-Credential -UserName "$ComputerName\$UserName" -Message "Enter password for account `"$ComputerName\$UserName`" to continue."
            $diskInfo = Get-WmiObject -ComputerName "$ComputerName" -Credential $account -Class Win32_LogicalDisk -Filter "DeviceID=`'$($DriveLetter):`'"
        } else {
            $diskInfo = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID=`'$($DriveLetter):`'"
        }

        <#
            Error handling
        #>

        if (@('setSize', 'setSizeRemote') -ccontains $PSCmdlet.ParameterSetName) {
            # Check if -InitialSize is greater than -MaximumSize
            if ($MaximumSize -lt $InitialSize) {
                Write-Error '"-InitialSize" cannot be greater than "-MaximumSize"'
                break
            }

            # Check if target has enough free space
            if (($diskInfo.FreeSpace / 1MB) -lt $MaximumSize) {
                Write-Error "$([ComponentModel.Win32Exception]112) [$($DriveLetter):]"
                break
            }
        }
    }

    process {
        try {
            <#
                Automatic management
            #>

            # Get configuration
            if (@('setRemote', 'setSizeRemote') -ccontains $PSCmdlet.ParameterSetName) {
                $wmiComputerSystem = Get-WmiObject -ComputerName "$ComputerName" -Credential $account -Class Win32_ComputerSystem -EnableAllPrivileges
            } else {
                $wmiComputerSystem = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges
            }

            # Disable automatic management
            $wmiComputerSystem.AutomaticManagedPagefile = $false
            $wmiComputerSystem.Put() | Out-Null

            <#
                Current pagefile
            #>

            # Get configuration
            if (@('setRemote', 'setSizeRemote') -ccontains $PSCmdlet.ParameterSetName) {
                $wmiPageFile = Get-PageFile -ComputerName $ComputerName -Credential $account
            } else {
                $wmiPageFile = Get-PageFile
            }

            # Delete existing pagefile's
            $wmiPageFile | ForEach-Object {
                $_.Delete()
            }

            <#
                New pagefile
            #>

            # New configuration values
            $cfgPageFile = @{
                Name = "$($DriveLetter):\pagefile.sys"
                InitialSize = 0
                MaximumSize = 0
            }

            # Create new pagefile
            if (@('setRemote', 'setSizeRemote') -ccontains $PSCmdlet.ParameterSetName) {
                Set-WmiInstance -ComputerName "$ComputerName" -Credential $account -Class Win32_PageFileSetting -EnableAllPrivileges -Arguments $cfgPageFile | Out-Null
            } else {
                Set-WmiInstance -Class Win32_PageFileSetting -EnableAllPrivileges -Arguments $cfgPageFile | Out-Null
            }

            <#
                Change pagefile size
            #>

            # Check if parameters "-InitialSize" and "-MaximumSize" were specified
            if ($PSCmdlet.ParameterSetName -eq 'setSize' -or 'setSizeRemote') {
                # Get previously created pagefile
                if (@('setRemote', 'setSizeRemote') -ccontains $PSCmdlet.ParameterSetName) {
                    $wmiPageFile = Get-PageFile -ComputerName $ComputerName -Credential $account
                } else {
                    $wmiPageFile = Get-PageFile
                }

                # Change values and apply
                $wmiPageFile.InitialSize = $InitialSize
                $wmiPageFile.MaximumSize = $MaximumSize
                $wmiPageFile.Put() | Out-Null
            }
        } catch {
            throw
        }
    }
}
