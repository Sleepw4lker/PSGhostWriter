Function Remove-WordSelection {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument
    )

    process {

        $Selection = $WordDocument.ActiveWindow.Selection

        # https://docs.microsoft.com/en-us/office/vba/api/Word.Selection.Delete
        [void]$Selection.Delete()

        # https://docs.microsoft.com/de-de/office/vba/api/word.selection.typebackspace
        $Selection.TypeBackSpace()

    }
}