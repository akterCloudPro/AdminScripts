<#
.SYNOPSIS
    Extracts the ProductCode GUID from an MSI installer without installing it.

.DESCRIPTION
    This PowerShell script uses the Windows Installer COM object to read 
    the ProductCode from the Property table of a given MSI file. 
    It works on MSI files that are not installed on the system.

.PARAMETER Path
    Full path to the MSI installer file.

.EXAMPLE
    $Path = "C:\Installers\MyApp.msi"
    # Script will output the ProductCode of MyApp.msi
#>

# Path to the MSI installer file
$Path = "C:\Path\To\Installer.msi"

# Create a Windows Installer COM object
$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer

# Open the MSI database in read-only mode (0 = read-only)
$Database = $WindowsInstaller.GetType().InvokeMember(
    "OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path, 0)
)

# Prepare a query to get the 'ProductCode' from the Property table
$View = $Database.GetType().InvokeMember(
    "OpenView", "InvokeMethod", $null, $Database, "SELECT `Value` FROM `Property` WHERE `Property`='ProductCode'"
)

# Execute the query
$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)

# Fetch the result record
$Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)

# Retrieve the ProductCode string from the record (1 = first column)
$Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
