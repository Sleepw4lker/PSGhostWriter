Function Close-WordApplication {

    [Alias("Exit-WordApplication")]
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("App")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $WordApplication
    )

    process {

        If ($WordApplication.Application.Documents.Count -gt 0) {
            Close-WordDocument -App $WordApplication
        }

        Write-Verbose -Message "Exiting Word Application"

        $WordApplication.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApplication)

    }
}