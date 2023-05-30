Function Set-WordSelectionToTopOfDocument {

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

        # https://docs.microsoft.com/en-us/office/vba/api/word.selection.homekey
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdmovementtype
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdunits

        # This method returns an integer that indicates the number of characters the selection 
        # was actually moved, or it returns 0 (zero) if the move was unsuccessful.
        # This method corresponds to functionality of the HOME key.
        [void]$WordDocument.ActiveWindow.Selection.HomeKey(
            [Microsoft.Office.Interop.Word.WdUnits]::wdStory,
            [Microsoft.Office.Interop.Word.WdMovementType]::wdMove
        )
        
    }
}