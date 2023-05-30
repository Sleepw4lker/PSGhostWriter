Function Set-WordSelectionToPattern {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Pattern,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoWrap = $False
    )

    process {

        $Selection = $WordDocument.ActiveWindow.Selection

        $Selection.Find.ClearFormatting() 
        $Selection.Find.Forward = $True

        If ($NoWrap -eq $False) {
            $Selection.Find.Wrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue
        }
        Else {
            $Selection.Find.Wrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindStop
        }

        $Selection.Find.Text = $Pattern

        [void]$Selection.Find.Execute()

        return $Selection.Find.Found

    }
}