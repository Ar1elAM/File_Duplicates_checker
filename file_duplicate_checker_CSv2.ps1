function Get-DuplicateFiles {
    param(
        [string]$folder,
        [int]$min_size,
        [string]$saveas
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
            Size = '{0:N3} MB' -f ($file.Length / 1MB)
            Hash = $file.MD5
        }
    }

    # Sort the duplicate files by their MD5 hash
    $duplicates = $duplicates | Sort-Object Hash

    # Write the duplicate files to a CSV file
    $duplicates | Export-Csv -Path $saveas -NoTypeInformation -Encoding UTF8

}

$folder = Read-Host -Prompt "What is the Parent folder to scan all it's contents?"
$size_min = Read-Host -Prompt "In MB, what is the minimun size for files larger to be recorded?"
$saveas = Read-Host -Prompt "Where do you want to save the file `"Duplicates.csv`"?`nEx.: FILE\LOCATION\ONLY"
if (-not $saveas.Substring(($saveas.Length - 1)) -eq "\") {
    $saveas = $saveas + "\"
}
$saveas = $saveas + "\DuplicateFiles.csv"
Get-DuplicateFiles -folder $folder -min_size $size_min -saveas $saveas
