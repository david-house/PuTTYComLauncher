
#Constants
$SERIAL_PORT_NAME_PATTERN = ".*\(COM\d*\).*"
$COM_PORT_PATTERN = ".*(COM\d*).*"
$COM_PORT_NUMBER_PATTERN = ".*\(COM(\d*)\).*"
$SERIAL_PORT_SETTINGS = @(
    "9600,8,n,1,N"
    "19200,8,n,1,N"
    "115200,8,n,1,N"
)
[int]$selected_setting = 2

$MARQUEE = @"

----------------------------------------------------------------------------------------------
______    _____ _______   __  _____ ________  ___  _                            _               
| ___ \  |_   _|_   _\ \ / / /  __ |  _  |  \/  | | |                          | |              
| |_/ _   _| |   | |  \ V /  | /  \| | | | .  . | | |     __ _ _   _ _ __   ___| |__   ___ _ __ 
|  __| | | | |   | |   \ /   | |   | | | | |\/| | | |    / _`` | | | | '_ \ / __| '_ \ / _ | '__|
| |  | |_| | |   | |   | |   | \__/\ \_/ | |  | | | |___| (_| | |_| | | | | (__| | | |  __| |   
\_|   \__,_\_/   \_/   \_/    \____/\___/\_|  |_/ \_____/\__,_|\__,_|_| |_|\___|_| |_|\___|_|   
----------------------------------------------------------------------------------------------
by EBRAddict
"@


# Get the list of plug and play objects with (COM##) in the name
$serial_ports = Get-WMIObject Win32_PnPEntity | Where-Object {$_.Name -match  $SERIAL_PORT_NAME_PATTERN} | Select-Object Caption


$ports = @()

foreach ($serial_port in $serial_ports) {
    # Extract the port name and port number from the caption string and create a new PSCustomObject to hold them
    $serial_port_name = $serial_port.Caption -match $COM_PORT_PATTERN | % {$Matches[1]}
    [int]$serial_port_number = $serial_port.Caption -match $COM_PORT_NUMBER_PATTERN | % {$Matches[1]}
        
        $port = New-Object PSCustomObject

        Add-Member -InputObject $port -NotepropertyName PortNumber -NotePropertyValue $serial_port_number
        Add-Member -InputObject $port -NotepropertyName "Port Name" -NotePropertyValue $serial_port_name
        Add-Member -InputObject $port -NotepropertyName  Caption -NotePropertyValue $serial_port.Caption
            
        $ports += $port

    }

# Sort the ports in numeric order (for ports > COM9 to show in proper sequence)
$sorted_ports = $ports | Sort-Object PortNumber

# Assign each port a letter starting with "A"
$i = 65
foreach ($port in $sorted_ports) {
    Add-Member -InputObject $port -NotepropertyName "Port Key" -NotePropertyValue $([char]$i)
    $i++
}


Write-Host -Object $MARQUEE -ForegroundColor Green

Write-Host ($sorted_ports | Select-Object "Port Key", "Port Name", Caption | Format-Table | Out-String)

Write-Host "Press a port key to launch PuTTY, or a space to cycle through serial port settings"
Write-Host

Write-Host  -NoNewLine  "Selected configuration: $($SERIAL_PORT_SETTINGS[$selected_setting])"

# Read the key press and find the corresponding port if available
$key = [Console]::ReadKey($true)
$selected_port = $sorted_ports | Where-Object {$_."Port Key" -eq $key.Key}

while ((($key.Key) -eq "Spacebar") -or ($selected_port -ne $null)) {

    # If user pressed a port letter, launch PuTTY with the selected COM port and configuration, then exit the loop
    if ($selected_port -ne $null) {
        Write-Host
        Write-Host
        Write-Host "You pressed key `"$($key.Key)`", launching PuTTY for `"$($selected_port.Caption)`" with serial configuration `"$($SERIAL_PORT_SETTINGS[$selected_setting])`""
        Write-Host
        & PuTTY.exe -serial $($selected_port."Port Name") -sercfg $($SERIAL_PORT_SETTINGS[$selected_setting])
        break
        
    }

    # If user pressed space then cycle through the serial port setting options
    if ($key.Key -eq "Spacebar") {
        $selected_setting = (($selected_setting + 1) % $($SERIAL_PORT_SETTINGS.Count))
        Write-Host  -NoNewLine  "`rSelected configuration: $($SERIAL_PORT_SETTINGS[$selected_setting])"
    }

    # Recycle the loop
    $key = [Console]::ReadKey($true)
    $selected_port = $sorted_ports | Where-Object {$_."Port Key" -eq $key.Key}
}

