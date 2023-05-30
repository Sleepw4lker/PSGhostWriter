Function Set-WordStyleForSelection {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Style = $Null
    )

    process {

        $Selection = $WordDocument.ActiveWindow.Selection

        $NewStyle = $WordDocument.Styles($Style)
        
        If ($NewStyle) {
            $Selection.Range.Style = $NewStyle
        } 
    }
}