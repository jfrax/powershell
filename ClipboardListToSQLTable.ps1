clear
Add-Type -AssemblyName 'System.Windows.Forms'

$result = ""
$counter = 0;
foreach($c in [Windows.Forms.Clipboard]::GetText() -Split "\r\n") {
    $counter++;

    if ($counter -eq 999) {

        $result += [System.Environment]::NewLine;
        $result += [System.Environment]::NewLine;

        $counter = 0;
    }

    $result += "('" + $c + "'),"
}

[Windows.Forms.Clipboard]::SetText( $result )


