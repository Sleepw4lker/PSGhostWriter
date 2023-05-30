Function Close-WordDocument {

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

        Write-Verbose -Message "Closing current Document"

        # Check version of Word installed and discard changes
        If ($(Get-WordVersion) -eq 14) {
            $WordDocument.Close([ref]$False)
        }
        Else {
            # Office 2013 or newer
            $WordDocument.Close($False)  
        }
    }
}