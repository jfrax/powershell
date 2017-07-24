clear

Add-Type -AssemblyName 'System.Windows.Forms'

$result = "";


foreach($file in [IO.Directory]::EnumerateFiles("F:\GFI Exports") | Split-Path -leaf) {
    $str = [io.path]::GetFileNameWithoutExtension($file)
    $result += "('" + $str + "'),"
}


[Windows.Forms.Clipboard]::SetText( $result )