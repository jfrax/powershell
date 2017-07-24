clear

Add-Type -AssemblyName 'System.Windows.Forms'

$resultMissing = "";

foreach($c in [Windows.Forms.Clipboard]::GetText() -Split "\r\n") {

    $found = $FALSE;

    foreach($file in [IO.Directory]::EnumerateFiles("F:\GFI Exports") | Split-Path -leaf) {
        $str = [io.path]::GetFileNameWithoutExtension($file);
        
        if($str.Trim() -eq $c.Trim()) {
            $found = $TRUE;
            break;
        }

    }
    
    if(!$found)
    {
        $resultMissing += $c + "`r`n";
    }

    
}


[Windows.Forms.Clipboard]::SetText( $resultMissing )