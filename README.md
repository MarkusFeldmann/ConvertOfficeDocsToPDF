ConvertToPDF.ps1
Converts all documents of a certain type in a source directory to a pdf of the same name in the destination directory

Usage:
gci C:\SourceDocs |  .\ConvertToPDF.ps1 -docType docx -destDir c:\DestinationDocs

Remarks:
Works with docx, xlsx and pptx

---------------------------------------
GetFileHash.ps1
Get the file hash for a given file and appends them optionally to a logging file

Usage:
gci "C:\DestinationDocs\" -Recurse | .\GetFileHash.ps1 -logFilePath 'C:\SourceDocs\logfile.txt'

