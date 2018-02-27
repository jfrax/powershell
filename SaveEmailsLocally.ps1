$DestinationPath = "D:\MailArchive\";

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

 $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]

 $outlook = new-object -comobject outlook.application

 $namespace = $outlook.GetNameSpace("MAPI")

 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)

#Removes invalid Characters for file names from a string input and outputs the clean string 
#Similar to VBA CleanString() Method 
#Currently set to replace all illegal characters with a hyphen (-) 
Function Remove-InvalidFileNameChars { 
 
    param( 
        [Parameter(Mandatory=$true, Position=0)] 
        [String]$Name 
    ) 
 
    return [RegEx]::Replace($Name, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '-') 
} 

foreach ($email in $folder.Items) { 
     
    
    [string]$subject = $email.Subject
    [string]$sentOn = $email.SentOn 
    
    $fileName = Remove-InvalidFileNameChars -Name ($sentOn + "-" + $subject + ".eml");
    
    $dest = $DestinationPath + $fileName;
    
    $email.SaveAs($dest, $olSaveType::olMSG);
    
} 