function New-ExcelInstance {
    $ExcelInstance = New-Object -ComObject Excel.Application
    return $ExcelInstance
}

function Stop-ExcelInstance {
    param (
        [Parameter(Mandatory)][Microsoft.Office.Interop.Excel.ApplicationClass]$ExcelInstance,
        [switch]$SaveBeforeQuit
    )

    if ($SaveBeforeQuit) {
        for ($i = 1; $i -le $ExcelIntance.Workbooks.Count; $i++) {
            $ExcelInstance.Workbooks[$i].Save()
        }
    }
    $ExcelInstance.Quit()
}

function Open-ExcelFile {
    # https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
    param (
        [Parameter(Mandatory)][Microsoft.Office.Interop.Excel.ApplicationClass]$ExcelInstance,
        [Parameter(Mandatory)]$ExcelFilePath,
        $UpdateLinks = $false,
        $ReadOnly = $false,
        $Format = 5,
        $Password,
        $WriteResPassword,
        $IgnoreReadOnlyRecommended
    )

    $Workbook = $ExcelInstance.Workbooks.Open(
            $ExcelFilePath, $UpdateLinks, $ReadOnly, $Format, $Password, $WriteResPassword, $IgnoreReadOnlyRecommended
        )
    return $Workbook
}

function Update-ExcelFile {
    param (
        [Parameter(Mandatory)]$Workbook
    )

    $LinkSources = $Workbook.LinkSources()
    $i = 0 
    foreach ($Link in $LinkSources) {
        Write-Progress -Activity "Excel file update" -PercentComplete ($i * 100 / $LinkSources.count) -Status "Updating links" -CurrentOperation $Link
        $Workbook.UpdateLink($Link, 1)
        $i++
    }

    Write-Progress -Activity "Excel file update" -Status "Refreshing pivot tables"
    $Workbook.RefreshAll()
}
