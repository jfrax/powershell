Clear-Host;
Add-Type -AssemblyName 'System.Windows.Forms'

$result = ""

foreach($c in [Windows.Forms.Clipboard]::GetText() -Split "\r\n") {
    $result += "'" + $c.Trim() + "',"
}

[Windows.Forms.Clipboard]::SetText( $result )