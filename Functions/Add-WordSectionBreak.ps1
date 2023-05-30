Function Add-WordSectionBreak {

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

        # https://docs.microsoft.com/en-us/office/vba/api/word.selection.insertbreak
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdbreaktype
        $WordDocument.ActiveWindow.Selection.InsertBreak(
            [Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreak
        )

    }
}