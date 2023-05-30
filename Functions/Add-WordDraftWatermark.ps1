Function Add-WordDraftWatermark {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("App")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $WordApplication
    )

    process {

        Add-WordWatermark -App $WordApplication -Type "Draft"
        
    }
}