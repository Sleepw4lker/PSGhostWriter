Function Set-WordHeadersLinkedToSection {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [Int]
        $Section      
    )

    process {

        Write-Verbose "Linking all Headers to the One in Section $Section"

        $WordDocument.Sections | ForEach-Object -Process {

            $SectionIndex++

            $_.Headers | ForEach-Object -Process {

                If ($SectionIndex -gt $Section) { $_.LinktoPrevious = $True }
            }
        }
    }
}