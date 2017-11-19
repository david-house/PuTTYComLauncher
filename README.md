# PuTTYComLauncher

This is a simple PowerShell command line utility to launch PuTTY with a selected serial port configuration.

![alt text](https://github.com/david-house/PuTTYComLauncher/raw/master/images/puttycomlauncher-screenshot.png "Screenshot")

## To use

From a PowerShell prompt, enter
```
.\PuTTYComLauncher.ps1
```
## Selecting your configuration
To launch PuTTY as a serial console with the displayed configuration press a corresponding letter. To cycle the baud rate, etc., settings press space.

## Behind the scenes
This script looks for any WMI object with a name like "(COM#)". If it finds one, it adds it to the menu. The list of port configurations is an array at the top of the script block. You can add any valid arrangement and change the order to suit your needs.

Note this has only been tested with a few devices on Windows 10 with PowerShell 5.1

