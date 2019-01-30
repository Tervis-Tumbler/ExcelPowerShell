function New-ExcelInstance {
    $ExcelInstance = New-Object -ComObject Excel.Application
    return $ExcelInstance
}

function Stop-ExcelInstance {
    param (
        [Parameter(Mandatory)][ref]$ExcelInstance,
        [switch]$SaveBeforeQuit
    )

    if ($SaveBeforeQuit) {
        for ($i = 1; $i -le $ExcelInstance.Value.Workbooks.Count; $i++) {
            $ExcelInstance.Value.Workbooks[$i].Save()
        }
    }
    $ExcelInstance.Value.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelInstance.Value)
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}

function Open-ExcelFile {
    # https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
    param (
        [Parameter(Mandatory)][ref]$ExcelInstance,
        [Parameter(Mandatory)]$ExcelFilePath,
        $UpdateLinks = $false,
        $ReadOnly = $false,
        $Format = 5,
        $Password,
        $WriteResPassword,
        $IgnoreReadOnlyRecommended
    )

    $Workbook = $ExcelInstance.Value.Workbooks.Open(
            $ExcelFilePath, $UpdateLinks, $ReadOnly, $Format, $Password, $WriteResPassword, $IgnoreReadOnlyRecommended
        )
    return $Workbook
}

function Update-ExcelFile {
    param (
        [Parameter(Mandatory)][ref]$Workbook
    )

    $LinkSources = $Workbook.Value.LinkSources()
    $i = 0 
    foreach ($Link in $LinkSources) {
        Write-Progress -Activity "Excel file update" -PercentComplete ($i * 100 / $LinkSources.count) -Status "Updating links" -CurrentOperation $Link
        $Workbook.Value.UpdateLink($Link, 1)
        $i++
    }

    Write-Progress -Activity "Excel file update" -Status "Refreshing pivot tables"
    $Workbook.Value.RefreshAll()
}
