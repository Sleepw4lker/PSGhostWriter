Function Add-WordDocument {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$True)]
        [Alias("Path")]
        [Alias("FileName")]
        [ValidateScript({Test-Path -Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Above = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Below = $False
    )

    process {

        If ($Above) { Set-WordSelectionToTopOfDocument -WordDocument $WordDocument }
        If ($Below) { Set-WordSelectionToBottomOfDocument -WordDocument $WordDocument }

        <#
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdinformation
        # Cursor Position greater 1 means that this is not an empty Line
        If ($Selection.Information([Microsoft.Office.Interop.Word.WdInformation]::wdFirstCharacterColumnNumber) -gt 1) {
            # https://docs.microsoft.com/en-us/office/vba/api/word.selection.typeparagraph
            $Selection.TypeParagraph()
        }
        #>

        Write-Verbose -Message "Inserting $File"

        # Append the Document to the Base Document
        # https://technet.microsoft.com/en-us/library/ee692877.aspx
        # https://docs.microsoft.com/en-us/office/vba/api/word.selection.insertfile
        $WordDocument.ActiveWindow.Selection.InsertFile($File)

    }
}