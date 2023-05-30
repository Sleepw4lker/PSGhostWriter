Function Update-WordDocumentFields {

    # You must pass a "Word.Application" Object      

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

        # Fields before ToC, otherwise the ToC will not honor the Lists of Figures and Tables correctly
        Write-Verbose -Message "Updating Document Fields"

        [void]$WordDocument.Fields.Update()

        Write-Verbose -Message "Updating Table(s) of Contents"

        # https://docs.microsoft.com/en-us/office/vba/api/word.tablesofcontents
        $WordDocument.TablesOfContents | ForEach-Object -Process {

            # https://docs.microsoft.com/en-us/office/vba/api/word.tableofcontents
            [void]$_.Update()
            
        }
    }
}