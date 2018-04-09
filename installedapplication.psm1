# This code is based on a technet example:
# https://social.technet.microsoft.com/wiki/contents/articles/34637.powershell-list-and-export-installed-programs-local-or-remote.aspx
# Therefore I assume the copyright owner for this file is the original
# author of the article: Prashanth Jayaram

Function Get-InstalledApplication {
    #Registry Hives
    $Object =@()

    $excludeArray = ("Security Update for Windows",
            "Update for Windows",
            "Update for Microsoft .NET",
            "Security Update for Microsoft",
            "Hotfix for Windows",
            "Hotfix for Microsoft .NET Framework",
            "Hotfix for Microsoft Visual Studio 2007 Tools",
            "Microsoft Visual C++ 2010",
            "cwbin64a",
            "Hotfix")

    [long]$HIVE_HKROOT = 2147483648
    [long]$HIVE_HKCU = 2147483649
    [long]$HIVE_HKLM = 2147483650
    [long]$HIVE_HKU = 2147483651
    [long]$HIVE_HKCC = 2147483653
    [long]$HIVE_HKDD = 2147483654

    $Query = Get-WmiObject -query "Select AddressWidth, DataWidth,Architecture from Win32_Processor"
    foreach ($i in $Query) {
        if ($i.AddressWidth -eq 64) {
            $OSArch = '64-bit'
        } else {
            $OSArch = '32-bit'
        }
    }

    switch ($OSArch) {
        "64-bit" {
            $RegProv = GWMI -Namespace "root\Default" -list | where{$_.Name -eq "StdRegProv"}
            $Hive = $HIVE_HKLM
            $RegKey_64BitApps_64BitOS = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
            $RegKey_32BitApps_64BitOS = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $RegKey_32BitApps_32BitOS = "Software\Microsoft\Windows\CurrentVersion\Uninstall"

            #############################################################################

            # Get SubKey names

            $SubKeys = $RegProv.EnumKey($HIVE, $RegKey_64BitApps_64BitOS)

            # Make Sure No Error when Reading Registry

            if ($SubKeys.ReturnValue -eq 0) { # Loop through all returned subkeys
                foreach ($Name in $SubKeys.sNames) {
                    $SubKey = "$RegKey_64BitApps_64BitOS\$Name"
                    $ValueName = "DisplayName"
                    $ValuesReturned = $RegProv.GetStringValue($Hive, $SubKey, $ValueName)
                    $AppName = $ValuesReturned.sValue
                    $Version = ($RegProv.GetStringValue($Hive, $SubKey, "DisplayVersion")).sValue
                    $Publisher = ($RegProv.GetStringValue($Hive, $SubKey, "Publisher")).sValue
                    $donotwrite = $false

                    if ($AppName.length -gt "0") {
                        foreach ($exclude in $excludeArray) {
                            if ($AppName.StartsWith($exclude) -eq $TRUE) {
                                $donotwrite = $true
                                break
                            }
                        }
                        if ($donotwrite -eq $false) {
                            $Object += New-Object PSObject -Property @{
                                Application = $AppName;
                                Architecture  = "64-BIT";
                                Version = $Version;
                                Publisher = $Publisher;
                            }
                        }
                    }
                }
            }

            #############################################################################

            $SubKeys = $RegProv.EnumKey($HIVE, $RegKey_32BitApps_64BitOS)

            # Make Sure No Error when Reading Registry

            if ($SubKeys.ReturnValue -eq 0) {

                # Loop Through All Returned SubKEys

                foreach ($Name in $SubKeys.sNames) {

                    $SubKey = "$RegKey_32BitApps_64BitOS\$Name"

                    $ValueName = "DisplayName"
                    $ValuesReturned = $RegProv.GetStringValue($Hive, $SubKey, $ValueName)
                    $AppName = $ValuesReturned.sValue
                    $Version = ($RegProv.GetStringValue($Hive, $SubKey, "DisplayVersion")).sValue
                    $Publisher = ($RegProv.GetStringValue($Hive, $SubKey, "Publisher")).sValue
                    $donotwrite = $false

                    if ($AppName.length -gt "0") {
                        foreach ($exclude in $excludeArray) {
                            if ($AppName.StartsWith($exclude) -eq $TRUE) {
                                $donotwrite = $true
                                break
                            }
                        }
                        if ($donotwrite -eq $false) {
                            $Object += New-Object PSObject -Property @{
                                Application = $AppName;
                                Architecture  = "32-BIT";
                                Version = $Version;
                                Publisher = $Publisher;
                            }
                        }
                    }
                }
            }
        } #End of 64 Bit

        ######################################################################################

        ###########################################################################################

        "32-bit" {

            $RegProv = GWMI -Namespace "root\Default" -list | where {$_.Name -eq "StdRegProv"}
            $Hive = $HIVE_HKLM
            $RegKey_32BitApps_32BitOS = "Software\Microsoft\Windows\CurrentVersion\Uninstall"

            #############################################################################

            # Get SubKey names

            $SubKeys = $RegProv.EnumKey($HIVE, $RegKey_32BitApps_32BitOS)

            # Make Sure No Error when Reading Registry

            if ($SubKeys.ReturnValue -eq 0) {  # Loop Through All Returned SubKEys
                foreach ($Name in $SubKeys.sNames) {
                    $SubKey = "$RegKey_32BitApps_32BitOS\$Name"
                    $ValueName = "DisplayName"
                    $ValuesReturned = $RegProv.GetStringValue($Hive, $SubKey, $ValueName)
                    $AppName = $ValuesReturned.sValue
                    $Version = ($RegProv.GetStringValue($Hive, $SubKey, "DisplayVersion")).sValue
                    $Publisher = ($RegProv.GetStringValue($Hive, $SubKey, "Publisher")).sValue

                    if ($AppName.length -gt "0") {
                        $Object += New-Object PSObject -Property @{
                            Application = $AppName;
                            Architecture  = "32-BIT";
                            Version = $Version;
                            Publisher = $Publisher;
                        }
                    }
                }
            }
        } #End of 32 bit
    } # End of Switch

    $object
}
