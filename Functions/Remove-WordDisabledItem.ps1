
Function Remove-WordDisabledItem {

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [Alias("Path")]
        [Alias("FileName")]
        [ValidateScript({Test-Path -Path $_})]
        [String]
        $File
    )

    process {

        # Kudos to
        # https://stackoverflow.com/questions/751048/how-to-programatically-re-enable-WordDocumentuments-in-the-ms-office-list-of-disabled-fil

        $WordVersion = Get-WordVersion

        # Converts the File Name string to UTF16 Hex
        $FileNameInUTF16Hex = ""
        [System.Text.Encoding]::ASCII.GetBytes($File.ToLower()) | ForEach-Object -Process { 
            $FileNameInUTF16Hex += "{0:X2}00" -f $_
        }

        Try {
            # Tests to see if the Disabled items registry key exists
            $DisabledItemsRegistryKey = Get-Item `
                -Path "HKCU:\Software\Microsoft\Office\${WordVersion}.0\Word\Resiliency\DisabledItems\" `
                -ErrorAction SilentlyContinue
        }
        Catch {
            # Nothing yet
        }

        If ($NULL -ne $DisabledItemsRegistryKey) {

            # Cycles through all the properties and deletes it if it contains the file name.
            Foreach ($DisabledItem in $DisabledItemsRegistryKey.Property) {

                $Value = ""

                ($DisabledItemsRegistryKey | Get-ItemProperty).$DisabledItem | ForEach-Object -Process {
                    $Value += "{0:X2}" -f $_
                }

                If ($Value.Contains($FileNameInUTF16Hex)) {

                    Write-Verbose "Removing $File from the List of Disabled Items."

                    $DisabledItemsRegistryKey | Remove-ItemProperty -name $DisabledItem

                }
            }
        }
    }
}