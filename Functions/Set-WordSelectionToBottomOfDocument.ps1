Function Set-WordSelectionToBottomOfDocument {

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

        # https://technet.microsoft.com/en-us/library/ee692877.aspx
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdmovementtype
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdunits

        # This method returns an integer that indicates the number of characters the selection 
        # or active end was actually moved, or it returns 0 (zero) if the move was unsuccessful. 
        # This method corresponds to functionality of the END key.
        [void]$WordDocument.ActiveWindow.Selection.EndKey(
            [Microsoft.Office.Interop.Word.WdUnits]::wdStory,
            [Microsoft.Office.Interop.Word.WdMovementType]::wdMove
        )
    }
}