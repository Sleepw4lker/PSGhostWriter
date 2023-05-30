
Function Set-WordDocumentTemplate {

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
        $File
    )

    Write-Verbose "Setting Document Styles Template to $File"

    # https://docs.microsoft.com/en-us/office/vba/api/word.document.attachedtemplate
    $WordDocument.AttachedTemplate = $File

    Write-Verbose "Copying Styles from Template"

    # The original Code Sample says to use $WordApplication.ActiveDocument.AttachedTemplate.FullName()
    # but as this may return a HTTP URL if the Files are stored on OneDrive and the Option 
    # to use Office to update Documents is selected, and the Path is exactly the same, we use $File instead

    # https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CopyStylesFromTemplate
    $WordDocument.CopyStylesFromTemplate($File)
}