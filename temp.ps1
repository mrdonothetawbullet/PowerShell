If($PSCmdlet.ParameterSetName -eq 'Index')
            {
                $RangeValue = Convert-ToExcelRange -StartRowIndex $StartRowIndex -StartColumnIndex $StartColumnIndex -EndRowIndex $EndRowIndex -EndColumnIndex $EndColumnIndex 
                $SLTable = $WorkBookInstance.CreateTable($StartRowIndex,$StartColumnIndex,$ENDRowIndex,$ENDColumnIndex)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange - StartRow/StartColumn '{1}':'{2}' & EndRow/EndColumn '{3}':'{4}' " -f $TableStyle, $StartRowIndex,$StartColumnIndex,$ENDRowIndex,$ENDColumnIndex)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle) 

            }
