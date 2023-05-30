If ($PSVersionTable.PSEdition -ne "Desktop") {
    Write-Error -Message "This module is only compatible with the Desktop Edition of Windows PowerShell."
    return
}

Try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
}
Catch {
    Write-Error -Message "Microsoft Word seems not to be installed."
    return
}

$ModuleRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

. $ModuleRoot\Functions\Add-WordDocument.ps1
. $ModuleRoot\Functions\Add-WordDraftWatermark.ps1
. $ModuleRoot\Functions\Add-WordLineBreak.ps1
. $ModuleRoot\Functions\Add-WordPageBreak.ps1
. $ModuleRoot\Functions\Add-WordPictureToHeader.ps1
. $ModuleRoot\Functions\Add-WordSectionBreak.ps1
. $ModuleRoot\Functions\Close-WordDocument.ps1
. $ModuleRoot\Functions\Close-WordApplication.ps1
. $ModuleRoot\Functions\Edit-WordPattern.ps1
. $ModuleRoot\Functions\Get-WordVersion.ps1
. $ModuleRoot\Functions\New-WordApplication.ps1
. $ModuleRoot\Functions\New-WordTable.ps1
. $ModuleRoot\Functions\New-WordLine.ps1
. $ModuleRoot\Functions\Open-WordDocument.ps1
. $ModuleRoot\Functions\Remove-WordDisabledItem.ps1
. $ModuleRoot\Functions\Remove-WordSelection.ps1
. $ModuleRoot\Functions\Save-WordDocument.ps1
. $ModuleRoot\Functions\Search-WordPatternAndReplaceInDocument.ps1
. $ModuleRoot\Functions\Search-WordPatternAndReplaceInSelection.ps1
. $ModuleRoot\Functions\Set-WordDocumentTemplate.ps1
. $ModuleRoot\Functions\Set-WordDocumentTitle.ps1
. $ModuleRoot\Functions\Set-WordFootersLinkedToSection.ps1
. $ModuleRoot\Functions\Set-WordHeadersLinkedToSection.ps1
. $ModuleRoot\Functions\Set-WordPaperFormat.ps1
. $ModuleRoot\Functions\Set-WordSelectionToBottomOfDocument.ps1
. $ModuleRoot\Functions\Set-WordSelectionToPattern.ps1
. $ModuleRoot\Functions\Set-WordSelectionToTopOfDocument.ps1
. $ModuleRoot\Functions\Set-WordStyleForSelection.ps1
. $ModuleRoot\Functions\Test-WordIsValidStyle.ps1
. $ModuleRoot\Functions\Update-WordDocumentFields.ps1