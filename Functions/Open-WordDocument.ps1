Function Open-WordDocument {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordApp")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $WordApplication,

        # To-Do: Verify against allowed Extensions
        [Parameter(Mandatory=$True)]
        [Alias("Path")]
        [Alias("FileName")]
        [ValidateScript({Test-Path -Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$False)]
        [Switch]
        $ReadOnly = $False
    )

    process {

        Write-Verbose -Message "Opening Document $File. Read-Only: $($ReadOnly.IsPresent)"
        
        # Arrrrghhhh
        Start-Sleep -Seconds 1
        
        # Use [Type]::Missing for parameters that you want used with their default value.
        $DefaultValue = [Type]::Missing

        # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.documents.opennorepairdialog?view=word-pia
        $WordDocument = $WordApplication.Documents.OpenNoRepairDialog(
            $File,	        # FileName. The name of the document (paths are accepted).
            $False,	        # ConfirmConversions. True to display the Convert File dialog 
                            # box if the file isn't in Microsoft Word format.
            $ReadOnly.IsPresent,    # ReadOnly. True to open the document as read-only. 
                            # This argument doesn't override the read-only recommended setting on a saved document. 
                            # For example, if a document has been saved with read-only recommended turned on, 
                            # setting the ReadOnly argument to False will not cause the file to be opened as read/write.
            $DefaultValue,	# AddToRecentFiles. True to add the file name to the list of recently used files at the bottom of the File menu.
            $DefaultValue,	# PasswordDocument. The password for opening the document.
            $DefaultValue,	# PasswordTemplate. The password for opening the template.
            $True,	        # Revert. Controls what happens if FileName is the name of an open document. 
                            # True to discard any unsaved changes to the open document and reopen the file. False to activate the open document.
            $DefaultValue, 	# WritePasswordDocument. The password for saving changes to the document.
            $DefaultValue,	# WritePasswordTemplate. The password for saving changes to the template.
            $DefaultValue,  # Format. The file converter to be used to open the document.
                            # Can be one of the WdOpenFormat constants. The default value is wdOpenFormatAuto. 
                            # To specify an external file format, apply the OpenFormat property to a FileConverter 
                            # object to determine the value to use with this argument.
            $DefaultValue,	# Encoding. The document encoding (code page or character set) to be used by Microsoft 
                            # Word when you view the saved document. Can be any valid MsoEncoding constant. 
                            # For the list of valid MsoEncoding constants, see the Object Browser in the Visual Basic Editor.
                            #  The default value is the system code page.
            $DefaultValue,	# Visible. True if the document is opened in a visible window. The default value is True.
            $DefaultValue,	# OpenConflictDocument. Specifies whether to open the conflict file for a document with an offline conflict.
            $False          # OpenAndRepair. True to repair the document to prevent document corruption.
        )

        # You can mark the current document as already spell-checked
        # https://stackoverflow.com/questions/3822103/how-do-i-programatically-
        # turn-off-the-wavy-red-lines-in-a-microsoft-word-WordDocumentumen
        # Hope this speeds up the processing of the Document
        $WordDocument.SpellingChecked = $True

        # Returns an Microsoft.Office.Interop.Word.Document
        return $WordDocument

    }
}