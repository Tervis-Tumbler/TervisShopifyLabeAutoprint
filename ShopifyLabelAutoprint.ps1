<#
Note:
Works best when used with Sumatra PDF as the default PDF viewer.
Create a scheduled task with command - 
powershell -WindowStyle Hidden -File "C:\ShopifyLabelAutoprint\ShopifyLabelAutoprint.ps1"
#>

# $LabelPrinter = "Zebra ZP 500 (ZPL)"
# $PacklistPrinter = "Brother HL-6180DW series Printer"
# $PathToMonitor = "C:\Users\ospreystore\Downloads"

# $RootPath = "C:\ShopifyLabelAutoprint"
$InProgressPath = "$RootPath\InProgress"
$ArchivePath = "$RootPath\Archive"

#Testing
$PathToMonitor = "C:\users\hperez\OneDrive - Tervis\Downloads"
$RootPath = "D:\ShopifyLabelAutoprint"
$LabelPrinter = "Adobe PDF"
$PacklistPrinter = "Print to Evernote"

# Setup
New-Item -Path $InProgressPath -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path $ArchivePath -ItemType Directory -ErrorAction SilentlyContinue

# File monitor loop
while ($true) {
    $LabelZipFiles = Get-ChildItem -Path $PathToMonitor -Filter Documents_*.zip
    if (-not $LabelZipFiles) { Start-Sleep 1; continue }

    foreach ($ZipFile in $LabelZipFiles) {
        # Archive download
        $ArchiveZip = Move-Item -Path $ZipFile.FullName -Destination $ArchivePath -PassThru -Force

        # Cleanup and unzip download
        Get-ChildItem -Path $InProgressPath | ForEach-Object { $_ | Remove-Item }
        Expand-Archive -Path $ArchiveZip.FullName -DestinationPath $InProgressPath
        
        # Print shipping labels
        $ZPLFiles = Get-ChildItem -Path $InProgressPath -Filter *.zplii
        Get-CimInstance -Class Win32_Printer -Filter "Name='$LabelPrinter'" | 
            Invoke-CimMethod -MethodName SetDefaultPrinter - | 
            Out-Null
        foreach ($ZPLFile in $ZPLFiles) {
            $TempLabelPath = "$InProgressPath\$($ZPLFile.BaseName).txt"
            "`${" | Out-File -Encoding utf8 -FilePath $TempLabelPath -Force
            Get-Content -Raw -Path $ZPLFile.FullName | Add-Content -Path $TempLabelPath
            "}$" | Add-Content -Path $TempLabelPath        
            Start-Process -FilePath $TempLabelPath -Verb Print -Wait
        }

        # Print packlists
        $PDFFiles = Get-ChildItem -Path $InProgressPath -Filter *.pdf
        Get-CimInstance -Class Win32_Printer -Filter "Name='$PacklistPrinter'" | 
            Invoke-CimMethod -MethodName SetDefaultPrinter | 
            Out-Null
        foreach ($PDFFile in $PDFFiles) {
            Start-Process -FilePath $PDFFile.FullName -Verb Print -Wait
        }
    }
}
