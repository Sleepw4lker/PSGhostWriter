
Function New-WordTable {

    # Create a Table, return a Table Object

    [cmdletbinding()]
    Param (
        [Parameter(
            Position = 0,
            Mandatory = $True, 
            ValuefromPipeline = $True
        )]    
        [System.Management.Automation.PSObject]$Object,

        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Doc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $WordDocument,

        [Parameter(Mandatory=$False)]
        [Alias("CaptionBelow")]
        [ValidateNotNullOrEmpty()]
        [String]
        $Caption,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $CaptionAbove,

        [Parameter(Mandatory=$False)]
        [ValidateScript({
            Test-WordIsValidStyle -WordDocument $WordDocument -Style $_ -Type Table
        })]
        [String]
        $TableStyle,

        [Parameter(Mandatory=$True)]
        [ValidateScript({
            Test-WordIsValidStyle -WordDocument $WordDocument -Style $_
        })]
        [String]
        $Style,

        [Parameter(Mandatory=$False)]
        [ValidateScript({
            Test-WordIsValidStyle -WordDocument $WordDocument -Style $_
        })]
        [String]
        $HeaderStyle,

        [Parameter(Mandatory=$False)]
        [Switch]
        $RepeatHeader,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoNewLine,

        [Parameter(Mandatory=$False)]
        [Switch]
        $StyleColumnBands,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoStyleFirstColumn,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoStyleHeadingRows,

        [Parameter(Mandatory=$False)]
        [Switch]
        $StyleLastColumn,

        [Parameter(Mandatory=$False)]
        [Switch]
        $StyleLastRow,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoStyleRowBands
    )

    begin {

        $Selection = $WordDocument.ActiveWindow.Selection

        $Columns = 1
        $Rows = 1
        $CurrentRow = 1

        [Object[]]$Properties

        # https://msdn.microsoft.com/en-us/vba/word-vba/articles/tables-add-method-word
        $Table = $WordDocument.Tables.Add(
            $Selection.Range,
            $Rows,
            $Columns,
            # https://docs.microsoft.com/en-us/office/vba/api/word.wddefaulttablebehavior
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            # https://docs.microsoft.com/en-us/office/vba/api/word.wdautofitbehavior
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitFixed
        )

    }

    process {

        Write-Verbose "Row $CurrentRow"

        # Before the first Row is written into the Table, we first have to build the Header Row
        # Thus we enumerate the Properties, create the necessary columns and fill them with the Property Names 
        If ($CurrentRow -eq 1) {

            # The first Object that enters the pipeline determines the List of Properties to display
            # It is therefores advised that you explicitly specify a list of properties when pipelining 
            # into the function by using Select-Object
            # https://stackoverflow.com/questions/51894114/powershell-is-there-a-way-to-get-proper-order-of-properties-in-select-object-so?rq=1
            $Properties = ($Object.PSObject.Properties).Name

            # If there are fewer Columns than our Data, add the Columns
            For ($i = $Table.Range.Columns.Count; $i -lt $Properties.Count; $i++) {
                [void]$Table.Columns.Add()
            }

            # Fill the Cells with the Property Names
            $CurrentColumn = 1
            $Properties | ForEach-Object -Process { 

                $CurrentCell = $Table.Cell($Currentrow, $CurrentColumn).Range
                $CurrentCell.Text = ($_ -as [System.string])
                If ($HeaderStyle) {
                    $CurrentCell.Style = $WordDocument.Styles($HeaderStyle)
                }
                $CurrentColumn++

            }

            Write-Verbose "Table has now $($Table.Range.Columns.Count) Columns)"

        }

        # The first Row that has Data is Row #1, as #1 is the Header Row
        $CurrentRow++

        # Insert a Row for the Data
        [void]$Table.Rows.Add()

        # Fill the Cells with the Data
        $CurrentColumn = 1
        $Properties | ForEach-Object -Process {

            # Objects may have a differing Property Set
            # Thus checking if the Property is present, and if not, we leave the Cell empty
            If ($_ -in ($Object.PSObject.Properties).Name) {

                $CellValue = $Object."$($_)" -as [System.string]
                Write-Verbose "Property $($_) has Value $CellValue in Row $Currentrow, Col $CurrentColumn"
                $CurrentCell = $Table.Cell($Currentrow, $CurrentColumn).Range
                $CurrentCell.Text = $CellValue
                If ($Style) {
                    $CurrentCell.Style = $WordDocument.Styles($Style)
                }

            }
            $CurrentColumn++

        }

        Write-Verbose "Table has now $($Table.Range.Rows.Count) Rows"

    }

    End {

        If ($CaptionAbove) {
            # https://msdn.microsoft.com/en-us/vba/word-vba/articles/selection-insertcaption-method-word
            $Table.Range.InsertCaption(
                # https://docs.microsoft.com/en-us/office/vba/api/word.wdcaptionlabelid
                [Microsoft.Office.Interop.Word.WdCaptionLabelID]::wdCaptionTable,
                ": $CaptionAbove",
                $False, 
                [Microsoft.Office.Interop.Word.WdCaptionPosition]::wdCaptionPositionAbove
            )
        }

        If ($Caption) {
            # https://msdn.microsoft.com/en-us/vba/word-vba/articles/selection-insertcaption-method-word
            $Table.Range.InsertCaption(
                # https://docs.microsoft.com/en-us/office/vba/api/word.wdcaptionlabelid
                [Microsoft.Office.Interop.Word.WdCaptionLabelID]::wdCaptionTable,
                ": $Caption",
                $False, 
                [Microsoft.Office.Interop.Word.WdCaptionPosition]::wdCaptionPositionBelow
            )
        }

        If ($TableStyle) {
            # https://docs.microsoft.com/en-us/office/vba/api/word.table.style
            # https://docs.microsoft.com/en-us/office/vba/api/word.tablestyle
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdstyletype
            # Paragraph, Table and other Styles are all referenced via the same Style Property
            # Thus, if applying a Table Style to a Table and then a Font Style, the first will be overridden by the first
            $Table.Style = $TableStyle
        }

        If ($RepeatHeader.IsPresent) {
            # https://docs.microsoft.com/en-us/office/vba/api/word.row.headingformat
            $Table.Rows(1).HeadingFormat = $True
        }

        $Table.ApplyStyleColumnBands = $StyleColumnBands.IsPresent
        $Table.ApplyStyleFirstColumn = (-not $NoStyleFirstColumn.IsPresent)
        $Table.ApplyStyleHeadingRows = (-not $NoStyleHeadingRows.IsPresent)
        $Table.ApplyStyleLastColumn = $StyleLastColumn.IsPresent
        $Table.ApplyStyleLastRow = $StyleLastRow.IsPresent
        $Table.ApplyStyleRowBands = (-not $NoStyleRowBands.IsPresent)

        $Table.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)

        # Set-WordTableAutoFitBehavior
        # Set-WordTableBehavior

        If (-not $NoNewLine.IsPresent) {

            # Move the Selection below the Table when finished and before returning the Object
            $Table.Select()

            $Selection.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

            # If the Table has a Caption, the Cursor will otherwise land at the beginning of the Caption
            If ($Caption) {

                # Move to End of Line
                $Selection.EndKey([Microsoft.Office.Interop.Word.wdUnits]::wdLine)
                
                # Enter Key
                $Selection.TypeParagraph()
                
            }

        }

        return $Table
    }
}