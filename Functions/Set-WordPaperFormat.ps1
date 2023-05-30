
Function Set-WordPaperFormat {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$True)]
        [ValidateSet(
            "10x14",
            "11x17",
            "Letter",
            "LetterSmall",
            "Legal",
            "Executive",
            "A3",
            "A4",
            "A4Small",
            "A5",
            "B4",
            "B5",
            "CSheet",
            "DSheet",
            "ESheet",
            "FanfoldLegalGerman",
            "FanfoldStdGerman",
            "FanfoldUS",
            "Folio",
            "Ledger",
            "Note",
            "Quarto",
            "Statement",
            "Tabloid",
            "Envelope9",
            "Envelope10",
            "Envelope11",
            "Envelope12",
            "Envelope14",
            "EnvelopeB4",
            "EnvelopeB5",
            "EnvelopeB6",
            "EnvelopeC3",
            "EnvelopeC4",
            "EnvelopeC5",
            "EnvelopeC6",
            "EnvelopeC65",
            "EnvelopeDL",
            "EnvelopeItaly",
            "EnvelopeMonarch",
            "EnvelopePersonal")]
        [String]
        $Format
    )

    process {

        $Size = Switch ($Format) {
            "10x14" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaper10x14 }
            "11x17" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaper11x17 }
            "Letter" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLetter }
            "LetterSmall" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLetterSmall }
            "Legal" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLegal }
            "Executive" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperExecutive }
            "A3" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA3 }
            "A4" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4 }
            "A4Small" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4Small }
            "A5" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA5 }
            "B4" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperB4 }
            "B5" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperB5 }
            "CSheet" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperCSheet }
            "DSheet" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperDSheet }
            "ESheet" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperESheet }
            "FanfoldLegalGerman" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFanfoldLegalGerman }
            "FanfoldStdGerman" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFanfoldStdGerman }
            "FanfoldUS" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFanfoldUS }
            "Folio" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFolio }
            "Ledger" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLedger }
            "Note" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperNote }
            "Quarto" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperQuarto }
            "Statement" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperStatement }
            "Tabloid" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperTabloid }
            "Envelope9" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope9 }
            "Envelope10" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope10 }
            "Envelope11" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope11 }
            "Envelope12" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope12 }
            "Envelope14" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope14 }
            "EnvelopeB4" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeB4 }
            "EnvelopeB5" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeB5 }
            "EnvelopeB6" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeB6 }
            "EnvelopeC3" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC3 }
            "EnvelopeC4" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC4 }
            "EnvelopeC5" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC5 }
            "EnvelopeC6" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC6 }
            "EnvelopeC65" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC65 }
            "EnvelopeDL" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeDL }
            "EnvelopeItaly" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeItaly }
            "EnvelopeMonarch" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeMonarch }
            "EnvelopePersonal" { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopePersonal }
            default { [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4 }
        }

        Write-Verbose -Message "Setting Paper Format to ""$Format"""

        # https://docs.microsoft.com/en-us/office/vba/api/Word.sections
        $WordDocument.Sections | ForEach-Object -Process {

            # https://docs.microsoft.com/en-us/office/vba/api/word.pagesetup.papersize
            # https://gallery.technet.microsoft.com/office/Change-Default-Paper-Size-451f74f8
            $_.PageSetup.PaperSize = $Size
        }
    }
}