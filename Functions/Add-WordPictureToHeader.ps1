
Function Add-WordPictureToHeader {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        # https://docs.microsoft.com/en-us/dotnet/api/system.drawing.bitmap?view=netframework-4.7.2
        # GDI+ supports the following file formats: BMP, GIF, EXIF, JPG, PNG and TIFF. 

        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path -Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [Int]
        $Section,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Left = 0,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Top = 0,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Width = 300,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Height = 34
    )

    begin {
        Add-Type -AssemblyName System.Drawing
    }

    process {

        Write-Verbose "Analyzing $File"

        # Getting the Image, reading Witdth and Height, calculating a Scaling Factor
        $ImageObject = New-Object System.Drawing.Bitmap $File

        # Grabbing the relevant Data
        $OldWidth   = $ImageObject.Width
        $OldHeight  = $ImageObject.Height

        # Releasing the Handle on the Image Object
        # https://social.msdn.microsoft.com/Forums/vstudio/en-US/7aea078c-5419-4e00-bfd6-c472ab925857/how-do-i-let-properly-powershell-release-a-handle-to-a-systemdrawingbitmap-object?forum=netfxbcl
        $ImageObject.Dispose()

        Write-Verbose "Original Dimensions: $($OldWidth)x$($OldHeight)"
        Write-Verbose "Dimension Limits: $($Width)x$($Height)"

        # Determining which (Height or Width) defines the Scaling
        $ScalingFactor = [math]::Min(
            [double]($Width / $OldWidth), 
            [double]($Height / $OldHeight)
            )

        Write-Verbose "Scaling Factor is $($ScalingFactor)"

        # Rounding the Numbers
        $NewWidth   = [math]::Floor($OldWidth * $ScalingFactor)
        $NewHeight  = [math]::Floor($OldHeight * $ScalingFactor)

        Write-Verbose "New Dimensions: $($NewWidth)x$($NewHeight)"

        Write-Verbose "Inserting Picture $File at current Selection"

        $WordDocument.Sections($Section).Headers | ForEach-Object -Process {

            # https://docs.microsoft.com/en-us/office/vba/api/word.shapes.addpicture
            $Range = $_.Range

            [void]$WordDocument.Shapes.AddPicture(
                $File,
                $False,
                $True,
                $Left,
                $Top,
                $NewWidth,
                $NewHeight,
                $Range
            )
        }
    }
}