# This code is based on a technet example:
# https://social.technet.microsoft.com/wiki/contents/articles/34637.powershell-list-and-export-installed-programs-local-or-remote.aspx
# Therefore I assume the copyright owner for this file is the original
# author of the article: Prashanth Jayaram

# Example 1: OutputType CSV and Output path E:\Output
# PS P:\> Get-InstalledApplication -Computername hqdbsp18 -OutputType CSV -OutFile foo.csv
# 
# Example 2: OutputType GridView
# PS P:\>Get-InstalledApplication -Computername hqdbsp18 -OutputType GridView
# 
# Example 3: OutputType Console
# PS P:\>Get-InstalledApplication -Computername hqdbsp18 -OutputType Console
# 
# Example 4: Invalid OutputType
# PS P:\> Get-InstalledApplication -Computername hqdbsp18 -OutputType Grid
# 
# Example 5: Loop through each system to pull the software list
# foreach($srv in Get-Content E:\Server.txt) {
#     Get-InstalledApplication -Computername $srv -OutputType CSV -OutFile foo.csv
# }

Function Get-InstalledApplication {
    Param(
            [Parameter(Mandatory=$false)][string[]]$Computername,
            [String[]]$OutputType,
            [string]$outfile)

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

    if ($Computername.count -eq 0) {
        $Computername += $env:computername
    }

    foreach ($EachServer in $Computername) {
        $Query = Get-WmiObject -ComputerName $Computername -query "Select AddressWidth, DataWidth,Architecture from Win32_Processor"
        foreach ($i in $Query) {
            if ($i.AddressWidth -eq 64) {
                $OSArch = '64-bit'
            } else {
                $OSArch = '32-bit'
            }
        }

        switch ($OSArch) {
            "64-bit" {
                $RegProv = GWMI -Namespace "root\Default" -list -computername $EachServer| where{$_.Name -eq "StdRegProv"}
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
                                    ServerName = $EachServer;
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
                                    ServerName = $EachServer;
                                    Version = $Version;
                                    Publisher= $Publisher;
                                }
                            }
                        }
                    }
                }
            } #End of 64 Bit

            ######################################################################################

            ###########################################################################################

            "32-bit" {

                $RegProv = GWMI -Namespace "root\Default" -list -computername $EachServer| where {$_.Name -eq "StdRegProv"}
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
                                ServerName = $EachServer;
                                Version = $Version;
                                Publisher = $Publisher;
                            }
                        }
                    }
                }
            } #End of 32 bit
        } # End of Switch
    }

    #$AppsReport

    $column1  =  @{expression = "ServerName"; width = 15; label = "Name"; alignment = "left"}
    $column2  =  @{expression = "Architecture"; width = 10; label = "32/64 Bit"; alignment = "left"}
    $column3  =  @{expression = "Application"; width = 80; label = "Application"; alignment = "left"}
    $column4  =  @{expression = "Version"; width = 15; label = "Version"; alignment = "left"}
    $column5  =  @{expression = "Publisher"; width = 30; label = "Publisher"; alignment = "left"}

    if ($outputType -eq "Console") {
        "#"*80
            "Installed Software Application Report"
            "Number of Installed Application count : $($object.count)"
            "Generated $(get-date)"
            "Generated from $(gc env:computername)"
            "#"*80
            $object |Format-Table $column1, $column2, $column3 ,$column4, $column5
    } elseif ($OutputType -eq "GridView") {
        $object|Out-GridView
    } elseif ($OutputType -eq "CSV") {
        $object |
            Sort-Object -Property Application |
            export-csv -path $outfile -NoTypeInformation
    } else {
        write-host " Invalid Output Type $OutputType"
    }
}
