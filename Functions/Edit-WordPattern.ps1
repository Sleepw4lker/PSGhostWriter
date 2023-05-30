Function Edit-WordPattern {

    # ToDo: Include Text Markers (Yellow and so on)

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
        $Underline = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Italic = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Bold = $False
    )

    process {

        Write-Verbose -Message "Editing Pattern ""$Pattern"""

        # We must search without wrapping to avoid an endless loop
        Set-WordSelectionToTopOfDocument -WordDocument $WordDocument

        Do {

            $Found = Set-WordSelectionToPattern -WordDocument $WordDocument -Pattern $Pattern -NoWrap

            If ($Found -eq $True) {

                $Selection = $WordDocument.ActiveWindow.Selection

                $Selection.Font.Italic = $Italic
                $Selection.Font.Bold = $Bold
                $Selection.Font.Underline = $Underline

            }

        } While ($Found -eq $True)
    }
}