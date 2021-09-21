# Param - Required
param([string]$Excel="")

If ($Excel -eq "") {
	Write-Output "Failed"
	throw "No excel argument"
}

$ExcelFilePath = Split-Path -Path $excel
$ExcelName = Split-Path -Path $excel -Leaf
$ExcelTempDir = $ExcelFilePath + '\' + $ExcelName + "_temp"
$ExcelFileSaved = $ExcelFilePath + "\" + $ExcelName + "_unprotected.xlsx"

If (Test-Path $ExcelTempDir){
	Remove-Item $ExcelTempDir
}

If (Test-Path $ExcelFileSaved){
	Remove-Item $ExcelFileSaved
}

Add-Type -A System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::ExtractToDirectory($Excel, $ExcelTempDir)

$InputDir = $ExcelTempDir + "\xl\worksheets\"

$sheetXMLs = Get-ChildItem $InputDir -filter *.xml

foreach ($Input in $sheetXMLs) {
	
	# Load the existing document
	$Doc = [xml](Get-Content $Input.FullName)

# Remove all tag with this name
$DeleteNames = "sheetProtection"

($Doc.worksheet.ChildNodes |Where-Object { $DeleteNames -contains $_.Name }) | ForEach-Object {
	# Remove each node from its parent
	[void]$_.ParentNode.RemoveChild($_)
}

	# Save the modified document
	$Doc.Save($Input.FullName)
}

# $Input = $ExcelTempDir + "\xl\worksheets\sheet1.xml"
# $Output = $ExcelTempDir + "\xl\worksheets\sheet1.xml"
[System.IO.Compression.ZipFile]::CreateFromDirectory($ExcelTempDir, $ExcelFileSaved) ;

Remove-Item $ExcelTempDir -Force -Recurse

Write-Output "Success"

# [Environment]::Exit(200)


