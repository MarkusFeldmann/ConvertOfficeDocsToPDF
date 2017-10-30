#  gci C:\SourceDocs |  .\ConvertToPDF.ps1 -docType docx -destDir c:\DestinationDocs
[cmdletbinding()]
param(
[Parameter(Mandatory=$false, 
Position=0, 
ParameterSetName="LiteralPath", 
ValueFromPipeline, 
HelpMessage="Literal path to one or more locations.")][string[]] $LiteralPath,
[Parameter()][ValidateSet('docx', 'xlsx', 'pptx')][string] $docType,
[Parameter(Mandatory=$true)][string] $destDir
)
Begin {
    switch ($docType) {
        "xlsx" { $officeApp = New-Object -comobject excel.application }
        "docx" { $officeApp = New-Object -comobject word.application }
        "pptx" { $officeApp =  New-Object -comobject powerpoint.application }
        Default { $officeApp = "unknown filetype" }
    }
    Write-Host "Begin converting $($officeApp)"
}

Process {
    $xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]
    #PowerPOint does not user visible, use Presentation.Open("name.pptx", null, null, $false) the third si the WithWindow Parameter
    #$officeApp.visible = [Microsoft.Office.Core.MsoTriState]::msoFalse    

    if($_.Extension -match $docType) {
        #Convert to String otherwise you get PSObject, abd that does not work with officedoc.SaveAs (!) it just hangs
        [string]$filepath = Join-Path $destDir -ChildPath ($_.BaseName + '.pdf')
        $file = $_.FullName
        Write-Host -f yellow "converting $($_.FullName), target: $($filepath)"
        
        switch ($docType) {
            "pptx" {
                $ppopt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
                $officeDoc = $officeApp.Presentations.Open($file, $null, $null, $false)
                $officeDoc.SaveAs($filepath, $ppopt)
                $officeDoc.Close()
            }
            "docx" { 
                $wordopt = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF
                $officeDoc = $officeApp.Documents.Open($file)
                $officeDoc.SaveAs($filepath, $wordopt)
                $officeDoc.Close()
            }
            "xlsx" { 
                $officeDoc =  $officeApp.Workbooks.Open($file, 3)
                $officeDoc.Saved = $true
                $officeDoc.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)
                $officeDoc.Close()
            }
            Default {  }
        }
    }
}

End {
    $officeApp.Quit()
    $officeApp = $null
}