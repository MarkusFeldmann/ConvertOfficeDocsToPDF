# gci "C:\DestinationDocs" -recurse |? { !$_.PSIsContainer } | .\getHashFromFile.ps1

[cmdletbinding()]
param(
[Parameter(Mandatory=$false, 
Position=0, 
ParameterSetName="LiteralPath", 
ValueFromPipeline, 
HelpMessage="Literal path to one or more locations.")][string[]] $LiteralPath,
[Parameter(Mandatory=$false)][string] $logFilePath
)
   
Process {
    $fm = [System.IO.FileMode]::Open
    $md5 = [System.Security.Cryptography.MD5]::Create()
    function getHash ($f)
    {
        $s = New-Object System.IO.FileStream($f.FullName, $fm)
        $b = $md5.ComputeHash($s)
        $hash = [System.BitConverter]::ToString($b).Replace("-", "")
        Write-Host -f cyan "$($hash) for $($f.FullName)"
        $s.Close()
        $hash
    }

    if(!$_.PSIsContainer) { 
        $hash = getHash $_ 
        if(![string]::IsNullOrEmpty($logFilePath))
        {
            $fileinfo = [String]::Format("{0};{1}", $hash, $_.FullName)
            Add-Content -Path $logFilePath -Value $fileinfo
        }
    }
}