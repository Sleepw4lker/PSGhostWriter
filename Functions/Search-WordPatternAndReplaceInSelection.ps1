Function Search-WordPatternAndReplaceInSelection {

    # You must pass a "Word.Selection" Object

    [cmdletbinding()]
    Param (
        [Parameter(
            Position = 0,
            Mandatory = $True, 
            ValuefromPipeline = $True
        )]
        [ValidateNotNullOrEmpty()]
        [String]
        $Pattern,

        # To Do: Parameter Validation for the Selection
        [Parameter(Mandatory=$True)]
        [Object]
        $Selection,

        # Null or Empty allowed
        [Parameter(Mandatory=$False)]
        [String]
        $ReplaceWith
    )

    process {

        # Prohibit Function failure when an empty String is passed
        If ((-not [String]::IsNullOrEmpty($Pattern)) -and (-not [String]::IsNullOrEmpty($ReplaceWith))) {

            # ToDo: Verify the below statement, it should not be necessary to do a loop at this point

            # We must assume that we have multiple occurrences, thus we
            # repeat the process as long as we find new occurrences
            do {

                # Search and Replace
                # see https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.find.execute.aspx
                # see https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdreplace.aspx
                If ($Selection.Find.Execute(
                    $Pattern,   # FindText. The text to be searched for. Use an empty string ("") to search for formatting only. 
                                # You can search for special characters by specifying appropriate character codes. For example, "^p" 
                                # corresponds to a paragraph mark and "^t" corresponds to a tab character.
                    [Microsoft.Office.Core.MsoTriState]::msoFalse, # MatchCase. True to specify that the find text be case sensitive. 
                                # Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
                    [Microsoft.Office.Core.MsoTriState]::msoFalse, # MatchWholeWord. True to have the find operation locate only entire words,
                                # not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
                    [Microsoft.Office.Core.MsoTriState]::msoFalse, # MatchWildCards. True to have the find text be a special search operator. 
                                # Corresponds to the Use wildcards check box in the Find and Replace dialog box.
                    [Microsoft.Office.Core.MsoTriState]::msoFalse, # MatchSoundslike. True to have the find operation locate words that sound 
                                # similar to the find text. Corresponds to the Sounds like check box in the Find and Replace dialog box.
                    [Microsoft.Office.Core.MsoTriState]::msoFalse, # MatchAllWordForms. True to have the find operation locate all forms 
                                # of the find text (for example, "sit" locates "sitting" and "sat"). Corresponds to the Find all word forms 
                                # check box in the Find and Replace dialog box.
                    [Microsoft.Office.Core.MsoTriState]::msoTrue, # Forward. True to search forward (toward the end of the document).
                    [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue, # Warp. Controls what happens if the search begins at 
                                # a point other than the beginning of the document and the end of the document is reached (or vice versa if 
                                # Forward is set to False). This argument also controls what happens if there is a selection or range and 
                                # the search text is not found in the selection or range. Can be one of the WdFindWrap constants.
                    [Microsoft.Office.Core.MsoTriState]::msoFalse, # Format. True to have the find operation locate formatting in 
                                # addition to, or instead of, the find text.
                    $ReplaceWith,   # ReplaceWith. The replacement text. To delete the text specified by the Find argument, use an empty string (""). 
                                # You specify special characters and advanced search criteria just as you do for the Find argument. To specify a 
                                # graphic object or other nontext item as the replacement, move the item to the Clipboard and specify "^c" for ReplaceWith.
                    [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceAll # Replace. Specifies how many replacements are to be made: one, 
                                # all, or none. Can be any WdReplace constant. 
                )) {
                    Write-Verbose -Message "Replaced Term ""$Pattern"" with ""$ReplaceWith"""
                }

            } While ($Selection.Find.Found)

        }
    }
}