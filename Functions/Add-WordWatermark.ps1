Function Add-WordWatermark {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("App")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $WordApplication,

        [Parameter(Mandatory=$False)]
        [ValidateSet(
            "Asap",
            "Asap2",
            "Confidential",
            "Confidential2",
            "DoNotCopy",
            "DoNotCopy2",
            "Draft",
            "Draft2",
            "Sample",
            "Sample2",
            "Urgent",
            "Urgent2"
        )]
        [String]
        $Type = "Draft"
    )

    process {

        # Can this be changed to $WordDocument.Templates so that we can call either by $WordApplication or $WordDocument?
        $WordApplication.Templates.LoadBuildingBlocks()

        # Try to get english ones
        $BuildingBlocks = 
            $WordApplication.Templates | 
            Where-Object { (($_.name -eq 'Built-In Building Blocks.dotx') -and ($_.LanguageID -eq 1033)) } | 
            Select-Object -First 1

        # Revert to Default
        If (-not ($BuildingBlocks)) {
            $BuildingBlocks = 
                $WordApplication.Templates | 
                Where-Object { (($_.name -eq 'Built-In Building Blocks.dotx')) } | 
                Select-Object -First 1
        }

        If ($BuildingBlocks) {

            $ItemIndex = switch($Type) {
                "Asap" {1}
                "Asap2" {2}
                "Confidential" {3}
                "Confidential2"{4}
                "DoNotCopy" {5}
                "DoNotCopy2" {6}
                "Draft" {7}
                "Draft2" {8}
                "Sample" {9}
                "Sample2" {10}
                "Urgent" {11}
                "Urgent2" {12}
            }

            Write-Verbose -Message "Adding Watermark"

            $Watermark = $BuildingBlocks.BuildingBlockEntries.Item($ItemIndex)

            $SectionIndex++

            $WordApplication.ActiveDocument.Sections | ForEach-Object -Process {

                $_.Headers | ForEach-Object -Process {

                    If ($SectionIndex -eq 2) {
                        [void]$Watermark.Insert($_.Range, $True)
                    }
                }
            }
        }
    }
}