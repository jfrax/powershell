$un = 'sburkhalter';

$u = get-aduser -identity $un –properties thumbnailphoto;

[System.Io.File]::WriteAllBytes("C:\Users\jfraxedas\OneDrive - RTI Surgical\powershell\$un.jpg", $u.Thumbnailphoto)