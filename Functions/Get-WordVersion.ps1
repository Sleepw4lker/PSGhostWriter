
Function Get-WordVersion {

    [cmdletbinding()]
    param()

    process {

        # We both check if Word is installed at all and if so, which Version
        Try { 
            $WordVersion = (Get-ItemProperty HKLM:\Software\Classes\Word.Application\CurVer).'(default)'
        }
        Catch {
            Throw "Word seems not to be installed"
        }

        $WordVersion = $WordVersion.TrimStart("Word.Application.")
        [int]$WordVersion = [convert]::ToInt32($WordVersion, 10)

        return $WordVersion

    }
}