function Get-DuplicateFiles {
    param(
        [string]$folder,
        [int]$min_size
    )

    [long]$minimum_size_bytes = $min_size * 1024 * 1024

    $files = Get-ChildItem $folder -File -Recurse | Where-Object { $_.Length -ge $minimum_size_bytes } | Select-Object FullName, Length, @{Name = 'MD5'; Expression = { (Get-FileHash $_.FullName -Algorithm md5).Hash } }

    # Group the files by their MD5 hash
    $groups = $files | Group-Object MD5 | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group

    # Create a new array that contains only the duplicate files
    $duplicates = foreach ($file in $groups) {
        [PSCustomObject]@{
            Name = Split-Path -Leaf $file.FullName
            Path = $file.FullName
            Size = $file.Length
            Hash = $file.MD5
        }
    }

    # Sort the duplicate files by their MD5 hash
    $duplicates = $duplicates | Sort-Object Hash

    # Write the duplicate files to an Excel file
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)
    $sheet.Cells.Item(1, 1).Value2 = "File Name"
    $sheet.Cells.Item(1, 2).Value2 = "Full Path"
    $sheet.Cells.Item(1, 3).Value2 = "Size"
    $sheet.Cells.Item(1, 4).Value2 = "MD5 Hash"
    $sheet.Range("A1:D1").Font.Bold = $true
    $row = 2

    $previousHash = ""
    $groupColor1 = 0xB0B0B0
    $useFixedColor = $false
    $first_run = $true
    
    foreach ($duplicate in $duplicates) {
        if ($previousHash -ne $duplicate.Hash) {
            If (-not $first_run) {
                $sheet.Cells.Item($row, 3).Value2 = '{0:N3} MB' -f ($totalsize / 1MB)
                $sheet.cells.item($row, 3).Borders.LineStyle = 1
                $sheet.cells.item($row, 3).Borders.Weight = 2
                $totalsize = 0
                if ($useFixedColor) {
                    $sheet.Range(("C$row")).Interior.Color = $groupColor1
                }
                $row++
            }
            $row++
            $previousHash = $duplicate.Hash
    
            # Toggle the flag to use fixed color every other group
            if ($useFixedColor) {
                $useFixedColor = $false
            }
            else {
                $useFixedColor = $true
            }
        }
        $first_run = $false
        if ($useFixedColor) {
            $sheet.Range(("A$row") + ":" + ("D$row")).Interior.Color = $groupColor1
        }
        # Write the data to the current row
        $range = $sheet.Range(("A$row") + ":" + ("D$row"))
        $range.Borders.LineStyle = 1
        $range.Borders.Weight = 2
        $sheet.Cells.Item($row, 1).Value2 = $duplicate.Name
        $sheet.Cells.Item($row, 2).Value2 = $duplicate.Path
        $sheet.Cells.Item($row, 3).Value2 = '{0:N3} MB' -f ($duplicate.Size / 1MB)
        $totalsize = $totalsize + $duplicate.Size
        $sheet.Cells.Item($row, 4).Value2 = $duplicate.Hash
        $row++
        $current = [array]::IndexOf($duplicates, $duplicate)
        if ($current -eq ($duplicates.Length - 1)) {
            $sheet.Cells.Item($row, 3).Value2 = '{0:N3} MB' -f ($totalsize / 1MB)
            $sheet.cells.item($row, 3).Borders.LineStyle = 1
            $sheet.cells.item($row, 3).Borders.Weight = 2
        }
    }


    # Auto-fit the columns and rows
    $range = $sheet.Range("A1:D$row")
    $range.EntireColumn.AutoFit() | Out-Null
    $range.EntireRow.AutoFit() | Out-Null
    $excel.Visible = $true
    
}

$folder = Read-Host -Prompt "What is the Parent folder to scan all it's contents?"
$size_min = Read-Host -Prompt "In MB, what is the minimun size for files larger to be recorded?"
Get-DuplicateFiles -folder $folder -min_size $size_min