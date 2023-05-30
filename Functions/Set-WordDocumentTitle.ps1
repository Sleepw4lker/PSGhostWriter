
Function Set-WordDocumentTitle {

    [cmdletbinding()]
    Param (
        [Parameter(
            Position = 0,
            Mandatory = $True, 
            ValuefromPipeline = $True
        )]
        [ValidateNotNullOrEmpty()]
        [String]
        $Title,

        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument
    )

    process {

        Write-Verbose "Setting Document Title to $Title"

        $WordDocument.BuiltInDocumentProperties("Title") = $Title

    }
}