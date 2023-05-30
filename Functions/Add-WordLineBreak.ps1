Function Add-WordLineBreak {

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
        $Selection.TypeParagraph()

    }
}