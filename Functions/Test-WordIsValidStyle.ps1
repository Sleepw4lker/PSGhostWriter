Function Test-WordIsValidStyle {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Style,

        [Parameter(Mandatory=$False)]
        [ValidateSet("Table","Paragraph")]
        [String]
        $Type  = "Paragraph"
    )
    
    process {

        # This is by far the fastest Method that I found
        # $WordDocument.Styles | Where-Object { $_.NameLocal -eq $Style } would take ages in comparison
        Try { 
            $StyleObject = $WordDocument.Styles($Style) | Select-Object NameLocal,Type
        }
        Catch {
            # Not Style found, skip here and exit
            return $False
        }

        [Int]$TypeToSearchFor =  Switch ($Type) {
            "Paragraph" { [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph }
            "Table" { [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeTable }
        }

        return ($StyleObject.Type -eq $TypeToSearchFor)

    }
}