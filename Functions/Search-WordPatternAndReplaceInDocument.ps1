Function Search-WordPatternAndReplaceInDocument {

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
        
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        # Null or Empty allowed
        [Parameter(Mandatory=$False)]
        [String]
        $ReplaceWith,

        [Parameter(Mandatory=$False)]
        [Switch]
        $IncludeHeaders = $False
    )

    process {

        $Selection = $WordDocument.ActiveWindow.Selection

        # Prohibit Function failure when an empty String is passed
        If ((-not [String]::IsNullOrEmpty($Pattern)) -and (-not [String]::IsNullOrEmpty($ReplaceWith))) {

            Search-WordPatternAndReplaceInSelection -Selection $Selection -Pattern $Pattern -ReplaceWith $ReplaceWith

            If ($IncludeHeaders -eq $True) {

                $WordDocument.Sections | ForEach-Object -Process {

                    $_.Headers | ForEach-Object -Process {
                        Search-WordPatternAndReplaceInSelection -Selection $_.Range -Pattern $Pattern -ReplaceWith $ReplaceWith
                    }

                    $_.Footers | ForEach-Object -Process {
                        Search-WordPatternAndReplaceInSelection -Selection $_.Range -Pattern $Pattern -Replacewith $ReplaceWith
                    }
                }
            }
        }
    }
}