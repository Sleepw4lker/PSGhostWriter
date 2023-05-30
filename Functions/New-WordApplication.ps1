Function New-WordApplication {

    # Returns a "Word.Application" Object

    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [ValidateRange(8,64)]
        [int]
        $MinimumVersion = 8
    )

    process {

        # We both check if Word is installed at all and if so, which Version
        If ((Get-WordVersion) -lt $MinimumVersion) {
            Throw "No compatible Version of Microsoft Word installed. Needing $MinimumVersion, whereas the installed Version is $WordVersion"
        }

        Write-Verbose -Message "Spawning a new Word Application Instance"

        Try {
            $WordApplication = New-Object -ComObject Word.Application
        }
        Catch {
            Throw "Unable to open Microsoft Word"
        }

        # Checking if the -Verbose Argument was given.
        # In this case, we also make the Application Window visible.
        $WordApplication.Visible = ([System.Management.Automation.ActionPreference]::SilentlyContinue -ne $VerbosePreference)

        # I hate such lousy Workarounds. But Word seems to sometimes reject RPC Calls 
        # if we directly return the Object after launching the App.
        # This must be enough for now.
        Start-Sleep -Seconds 5

        return $WordApplication
    }
}