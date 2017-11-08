# Use the -baseDir Param to define at which level the directory will be recreated in the destination
# gci 'C:\admin\HR_Migration\Bewerbungen 2017' -Recurse -Filter eva*.docx | .\ConvertToPDF.ps1 -docType docx -destDir C:\temp -baseDir C:\admin\HR_Migration
# C:\_Themis_365\ConvertOfficeDocsToPDF-master>  gci 'z:\Migration\1 MIG\BGM (Archiv)' -Filter *.do* -Recurse  | .\ConvertToPDF.ps1 -docType doc -destDir C:\_Themis_365\ConvertOfficeDocsToPDF-master\target -baseDir 'z:\Migration\1 MIG\BGM(Archiv)'
[cmdletbinding()]
param(
[Parameter(Mandatory=$false, 
Position=0, 
ParameterSetName="LiteralPath", 
ValueFromPipeline, 
HelpMessage="Literal path to one or more locations.")][string[]] $LiteralPath,
[Parameter()][ValidateSet('doc', 'xls', 'ppt')][string] $docType,
[Parameter(Mandatory=$true)][string] $destDir,
[Parameter(Mandatory=$true)][string] $baseDir,
[Parameter(Mandatory=$false)][string] $logFile
)
Begin {
    $baseDir = $baseDir.TrimEnd('\')
    switch ($docType) {
        "xls" { $officeApp = New-Object -comobject excel.application }
        "doc" { $officeApp = New-Object -comobject word.application
                $officeApp.Visible = $false }
        "ppt" { $officeApp =  New-Object -comobject powerpoint.application }
        Default { $officeApp = "unknown filetype" }
    }
    Write-Host "Begin converting $($officeApp)"
}

Process {
    $xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]
    #PowerPOint does not user visible, use Presentation.Open("name.pptx", null, null, $false) the third si the WithWindow Parameter
    #$officeApp.visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    
    function getAndEnsureNewFullPath ($file) {
        $newPath = $file.DirectoryName.Replace($baseDir, $destDir)
        #$newPath2 = $newPath.Replace("\d", "\d2")
        Write-Host -f cyan "New Path2 $($newPath2)"
        Write-Host -f cyan "New Path $($newPath) - $($destDir) - $($baseDir) $($file.DirectoryName)"
        #if(!(Test-Path $newPath2)) {
        #    New-Item $newPath2 -ItemType Directory -Force | Out-Null
        #}

        if(!(Test-Path $newPath)) {
            New-Item $newPath -ItemType Directory -Force | Out-Null
            
        }
        if($file.Extension.EndsWith("x"))
        {
        Join-Path $newPath -ChildPath ($file.BaseName + ".pdf")
        }
        else {
            Join-Path $newPath -ChildPath ($file.BaseName + "." + $docType + ".pdf")
        }
    }

    if($_.Extension -match $docType) {
        #Convert to String otherwise you get PSObject, abd that does not work with officedoc.SaveAs (!) it just hangs
        #[string]$filepath = Join-Path $destDir -ChildPath ($_.BaseName + '.pdf')
        
        [string]$filepath = getAndEnsureNewFullPath $_

        $file = $_.FullName
        Write-Host -f yellow "converting $($_.FullName), target: $($filepath)"
        
        switch -Wildcard ($docType) {
            "ppt*" {
                try {
                    $ppopt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
                    $officeDoc = $officeApp.Presentations.Open($file, $null, $null, $false)
                    $officeDoc.SaveAs($filepath, $ppopt)
                    $officeDoc.Close()
                    Add-Content -Path $logFile "Sucessfully converted $($file)"
                }
                catch [Exception]
                {
                    $msg = $file + "-> " + $_.Exception.GetType().FullName + "`n" + $_.Exception.Message
                    Add-Content -Path $logFile $msg
                    $officeDoc.Close()
                }
            }
            "doc*" { 
                try {
                    $wordopt = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF
                    $officeDoc = $officeApp.Documents.Open($file)
                    $officeDoc.SaveAs($filepath, $wordopt)
                    $officeDoc.Close()
                    Add-Content -Path $logFile "Sucessfully converted $($file)"
                }
                catch [Exception]
                {
                    $msg = $file + "-> " + $_.Exception.GetType().FullName + "`n" + $_.Exception.Message
                    Add-Content -Path $logFile $msg
                    $officeDoc.Close()
                }
            }
            "xls*" { 
                try {
                    Write-Host -f cyan "$($file) - $($filepath)"
                    #$filepath2 = $filepath.Replace("\d", "\d2")
                    #Copy-Item -Path $file -Destination ($filepath2.Replace(".pdf", ".xlsx"))
                    $officeDoc =  $officeApp.Workbooks.Open($file, 3)
                    $officeDoc.Saved = $true
                    $officeDoc.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)
                    $officeDoc.Close()
                    Add-Content -Path $logFile "Sucessfully converted $($file)"
                }
                catch [Exception]
                {
                    $msg = $file + "-> " + $_.Exception.GetType().FullName + "`n" + $_.Exception.Message
                    Add-Content -Path $logFile $msg
                    $officeDoc.Close()
                }
            }
            Default {  }
        }
    }
}

End {
    $officeApp.Quit()
    $officeApp = $null
}
