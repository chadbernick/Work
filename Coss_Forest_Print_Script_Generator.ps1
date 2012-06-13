#Region "Comment Block"
      # Adapted from the drive mapping script.
#EndRegion


#Region "Assemblies"
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#EndRegion

#Region "###################  Functions  #######################"
# --------------------------------------------------
Function Msg-Box($String){
[Windows.Forms.MessageBox]::Show($String)
}

Function Get-Members($GroupName, $GroupCompareCollection){
      # Echo to screen what was passed
      Write-Host "Passed Group is $GroupName" -ForegroundColor Green
      Write-Host "Passed Compare is $GroupCompareCollection" -ForegroundColor Green

      # Create a new collectioin to convert the input data
      $tmpGroupCol = New-Object System.Collections.ArrayList
      # Data to return
      $ReturnUsers = New-Object System.Collections.ArrayList
      
      # Convert the string to a collection
      If($GroupCompareCollection -ne $null){
            ForEach($iGroup in $GroupCompareCollection){
            # Add to Collection
            [Void]$tmpGroupCol.Add($iGroup)
            }
      }
      If($GroupCompareCollection.Contains($GroupName)){
            Write-Host "Group is already in Collection - Exiting Function" -ForegroundColor Green
            $UserCol = $ReturnUsers
            Return ,$UserCol
      }
      Else{
            [Void]$tmpGroupCol.Add($GroupName)
      }
      
      # Get a list of groups contained within the passed group
      $Mem_DN = $GroupName.DN
      $AllObjs = Get-QADGroupMember -SizeLimit 0 -Identity $Mem_DN | Select LogonName, samAccountName, DN, member
      
      ForEach($Obj in $AllObjs){
            # Check to see if the Member field is not null
            $Check = $Obj.Member
            # If not null, we have a group
            
            If($Check -ne $null){
                  $NestedUsers = Get-Members $Obj $tmpGroupCol
                  # Add the users to the $ReturnUsers Collection

                  ForEach($User in $NestedUsers){
                        [Void]$ReturnUsers.Add($User)
                  }
            }
      }     
      
      # From original group - get users and add to the $ReturnUsers Collection
      
      ForEach($LocalUser in $AllObjs){
            # Check if the member field is null
            $Check = $Obj.Member
            
            # If null, we have a user
            If($Check -eq $null){
                  $lName = $LocalUser.LogonName
                  [Void]$ReturnUsers.Add($lName)
            }
            
      }     
      
      # Return a collection/list of users
            $UserCol = $ReturnUsers
            Return ,$UserCol
      
}

Function Create-LogonScript($UserList, $PrnPort, $MapPath){
      # Cycle each user - create the initial logon script
      
      ForEach($iUser in $UserList){
            $lScript = "$WorkSpaceDir\$iUser"+ "_PrnMaps.kix"
            $Mapping = "ADDPRINTERCONNECTION($Q"+ $MapPath+ "$Q)"
            $Mapping | Out-File $lScript -Append -Encoding ASCII
      }

}

Function Create-WorkSpace{
# Create Directories for storing logon scripts
$bTempDir = Test-Path($TempDir)
$bWorkSpaceDir = Test-Path($WorkSpaceDir)

      If($bWorkSpaceDir -eq $false){
            If($bTempDir -eq $false){
                  mkdir $TempDir
                  mkdir $WorkSpaceDir
            }
            Else{
                  mkdir $WorkSpaceDir
            }
      }
}

Function Zip-Files($SourceDir, $ZipFileName){


      if (test-path $ZipFileName) { 
            # Delete the current zip and replace
            Remove-Item $ZipFileName
            #echo "Zip file already exists at $ZipFileName" 
        return 
      } 

      set-content $ZipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18)) 
      (dir $ZipFileName).IsReadOnly = $false 

      $ZipFile = (new-object -com shell.application).NameSpace($ZipFileName) 

      $ZipFile.CopyHere($SourceDir)
}

Function Return-ArchiveFile($NumOfDaysToKeep){
      # Find out where we are in archiving
      $bArchive = Test-Path(".\Archives.dat")
      
      $MyPath = Get-Location
      
      If($bArchive -eq $true){
            $Archive_Num = GC ".\Archives.dat"
            If([int]$Archive_Num -ge [int]$NumOfDaysToKeep){
                  # Reset the Archive number and return file name Archive_1.zip
                  $Archive_Num = "1"
                  $Archive_Num | Out-File ".\Archives.dat"
                  Return "$MyPath\Archive_1.zip"
            }
            Else{
                  # Return the archive file number
                  $Archive_Num = [int]$Archive_Num+ 1
                  $Archive_Num | Out-File ".\Archives.dat"
                  $aFile = "$MyPath\Archive_"+ $Archive_Num+ ".zip"
                  Return $aFile
            }
      }
      Else{
            # File doesn't exist so let's create it
            "1" | Out-File ".\Archives.dat"
            Return "$MyPath\Archive_1.zip"
      }
}

Function Write-Log($DataToWrite, $LogFile){
      $ErrTime = Get-Date -Format "[MMddyy]@HH:mm"
      "$ErrTime     $DataToWrite" | Out-File $LogFile -Append
}

Function Return-Scripts{
      # new array for the script names
      $col_Scripts = New-Object System.Collections.ArrayList
      $HoldScripts = dir "$WorkSpaceDir\*.kix"
      
      If($HoldScripts -ne $null){
            ForEach($sFile in $HoldScripts){
                  $iScr = $sFile.Name
                  $col_Scripts.Add($iScr)
            }
            $col_Scripts | Out-File "$TempDir\List.txt"
            $true | Out-File "$TempDir\bList"
      }
      Else{
            $false | Out-File "$TempDir\bList"
      }
}

Function ReWrite-Script($Data, $ScriptToEdit){

      # Insert Header
      $Filler = "                                                  "
      $Filler | Out-File $ScriptToEdit 
      "Rem *************************************************" | Out-File $ScriptToEdit -Append
      "Rem * Logon Script Generated by:                    *" | Out-File $ScriptToEdit -Append
      "Rem * Cross Forest Logon Script Generation Engine   *" | Out-File $ScriptToEdit -Append
      "Rem * Concept:  Chad Bernick                        *" | Out-File $ScriptToEdit -Append
      "Rem * Written by:  Michael Heath                    *" | Out-File $ScriptToEdit -Append
      "Rem * Technical Assist:  Scott Reese                *" | Out-File $ScriptToEdit -Append
      "Rem *************************************************" | Out-File $ScriptToEdit -Append
      $Filler | Out-File $ScriptToEdit -Append
      $Data | Out-File $ScriptToEdit -Append
            }

Function Compare-Scripts($NewScript, $CurScript){
      # Make sure the remote file exist
      $sExist = Test-Path($CurScript)
      
      If($sExist -eq $true){
            # compare 2 scripts... if they are different, replace the current with the new
            $A = GC $NewScript
            $B = GC $CurScript
            
            $cResult = Compare-Object $A $B
            
            If($cResult -eq $null){
                  # Files are the same
            }
            Else{
                  # Files are different
                  # Delete the current script and replace with the new script
                  Remove-Item -Path $CurScript
                  Copy-Item -Path $NewScript -Destination $CurScript
            }
      }
      Else{
            # Remote File doesn't exist so let's copy it over
            Copy-Item -Path $NewScript -Destination $CurScript
      }
}

Function Compare-Directories($DirA, $DirB){
# Compares the contents of two directories.  $DirA is always the directory that is affected.

$ListDirA = Dir $DirA -Name
$ListDirB = Dir $DirB -Name

$Difference = Compare-Object $ListDirA $ListDirB

      If($Difference -ne $null){

            ForEach($Change in $Difference){
      $M = $Change.SideIndicator
      
      If($M -eq "=>"){
            # Change is on the right - DirB - move files from DirB to DirA
            $File1 = "$DirB\"+ $Change.InputObject
            $File2 = "$DirA\"+ $Change.InputObject
            Copy-Item -Path $File1 -Destination $File2
      }
      Else{
            # Change is on the left - DirA - Delete files from DirA
            $File2 = "$DirA\"+ $Change.InputObject
            Remove-Item $File2
      }
}
      }
}
  
Function Mail-SvrTeam($eAttachment){
      #  Set password and username
      $DecP = Decrypt-String $EncP "I am the one"           
      $secpasswd = ConvertTo-SecureString $DecP -AsPlainText -Force
      $mycreds = New-Object System.Management.Automation.PSCredential (“fsbp\svcpowershelltask”, $secpasswd)

      $emailFrom = "serverteam@fsdp.com"
      #$emailTo = "heathmichael@bfdp.com"
      $subject = "Printer Mappings Complete"
      $body = "Printer Mappings Have Completed."+ $cr
      $smtpServer = "relay.fsdp.com"
      
      If($emailTo -ne $null){
            # Send Notification to Server Team 
            If(Test-Path($eAttachment)){
                  Send-MailMessage -To $emailTo -From $emailFrom -Body $body -Subject $subject -Attachments $eAttachment -SmtpServer $smtpServer -Credential $mycreds 
            }
            Else{
                  Send-MailMessage -To $emailTo -From $emailFrom -Body $body -Subject $subject -SmtpServer $smtpServer -Credential $mycreds 
            }
      }
}

Function Mail-MissingUNC($eReport){
      #  Set password and username
      $DecP = Decrypt-String $EncP "I am the one"           
      $secpasswd = ConvertTo-SecureString $DecP -AsPlainText -Force
      $mycreds = New-Object System.Management.Automation.PSCredential (“fsbp\svcpowershelltask”, $secpasswd)

      $emailFrom = "serverteam@fsdp.com"
      #$emailTo = "ServerTeam@fsdp.com"
      #$emailTo = "heathmichael@bfdp.com"
      $subject = "Missing UNC Path in Group"
      $body = $eReport+ $cr
      $smtpServer = "relay.fsdp.com"
      
      If($emailTo -ne $null){
            Send-MailMessage -To $emailTo -From $emailFrom -Body $body -Subject $subject -SmtpServer $smtpServer -Credential $mycreds
      }
}
#EndRegion

#Region "Encrypt/Decrypt Functions and Assemblies"
# Load the Security assembly to use with this script 
#################
[Reflection.Assembly]::LoadWithPartialName("System.Security")

#################
# This function is to Encrypt A String.
# $string is the string to encrypt, $passphrase is a second security "password" that has to be passed to decrypt.
# $salt is used during the generation of the crypto password to prevent password guessing.
# $init is used to compute the crypto hash -- a checksum of the encryption
#################
function Encrypt-String($String, $Passphrase, $salt="SaltCrypto", $init="IV_Password", [switch]$arrayOutput){
      # Create a COM Object for RijndaelManaged Cryptography
      $r = new-Object System.Security.Cryptography.RijndaelManaged
      # Convert the Passphrase to UTF8 Bytes
      $pass = [Text.Encoding]::UTF8.GetBytes($Passphrase)
      # Convert the Salt to UTF Bytes
      $salt = [Text.Encoding]::UTF8.GetBytes($salt)

      # Create the Encryption Key using the passphrase, salt and SHA1 algorithm at 256 bits
      $r.Key = (new-Object Security.Cryptography.PasswordDeriveBytes $pass, $salt, "SHA1", 5).GetBytes(32) #256/8
      # Create the Intersecting Vector Cryptology Hash with the init
      $r.IV = (new-Object Security.Cryptography.SHA1Managed).ComputeHash( [Text.Encoding]::UTF8.GetBytes($init) )[0..15]
      
      # Starts the New Encryption using the Key and IV   
      $c = $r.CreateEncryptor()
      # Creates a MemoryStream to do the encryption in
      $ms = new-Object IO.MemoryStream
      # Creates the new Cryptology Stream --> Outputs to $MS or Memory Stream
      $cs = new-Object Security.Cryptography.CryptoStream $ms,$c,"Write"
      # Starts the new Cryptology Stream
      $sw = new-Object IO.StreamWriter $cs
      # Writes the string in the Cryptology Stream
      $sw.Write($String)
      # Stops the stream writer
      $sw.Close()
      # Stops the Cryptology Stream
      $cs.Close()
      # Stops writing to Memory
      $ms.Close()
      # Clears the IV and HASH from memory to prevent memory read attacks
      $r.Clear()
      # Takes the MemoryStream and puts it to an array
      [byte[]]$result = $ms.ToArray()
      # Converts the array from Base 64 to a string and returns
      return [Convert]::ToBase64String($result)
}

function Decrypt-String($Encrypted, $Passphrase, $salt="SaltCrypto", $init="IV_Password"){
      # If the value in the Encrypted is a string, convert it to Base64
      if($Encrypted -is [string]){
            $Encrypted = [Convert]::FromBase64String($Encrypted)
      }

      # Create a COM Object for RijndaelManaged Cryptography
      $r = new-Object System.Security.Cryptography.RijndaelManaged
      # Convert the Passphrase to UTF8 Bytes
      $pass = [Text.Encoding]::UTF8.GetBytes($Passphrase)
      # Convert the Salt to UTF Bytes
      $salt = [Text.Encoding]::UTF8.GetBytes($salt)

      # Create the Encryption Key using the passphrase, salt and SHA1 algorithm at 256 bits
      $r.Key = (new-Object Security.Cryptography.PasswordDeriveBytes $pass, $salt, "SHA1", 5).GetBytes(32) #256/8
      # Create the Intersecting Vector Cryptology Hash with the init
      $r.IV = (new-Object Security.Cryptography.SHA1Managed).ComputeHash( [Text.Encoding]::UTF8.GetBytes($init) )[0..15]


      # Create a new Decryptor
      $d = $r.CreateDecryptor()
      # Create a New memory stream with the encrypted value.
      $ms = new-Object IO.MemoryStream @(,$Encrypted)
      # Read the new memory stream and read it in the cryptology stream
      $cs = new-Object Security.Cryptography.CryptoStream $ms,$d,"Read"
      # Read the new decrypted stream
      $sr = new-Object IO.StreamReader $cs
      # Return from the function the stream
      Write-Output $sr.ReadToEnd()
      # Stops the stream      
      $sr.Close()
      # Stops the crypology stream
      $cs.Close()
      # Stops the memory stream
      $ms.Close()
      # Clears the RijndaelManaged Cryptology IV and Key
      $r.Clear()
}

#EndRegion

#Region "Variables & Declarations"
  $INIFIle = ".\PSettings.ini"
  $BakINI = ".\PSettings.bak"
  
#Region "Settings File"  
# Settings File
If(Test-Path($INIFile)){

      $SetFile = GC "$INIFile"
      "#" | Out-File "$BakINI"

# Check for blank lines in file and remove them
      ForEach($Line in $SetFile){
      
            If($Line -ne ""){
                  $Line | Out-File "$BakINI" -Append
            }
      }

      $SetFile = GC "$BakINI"

# Extract all variables from Settings file
      ForEach($Setting in $SetFile){
            $i = $Setting.Substring(0,1)
            
            IF($i -ne "#"){
                  # Select only the first 4 Letters
                  # For the ecryption, we cannot send the whole string to lowercase - so reserve a string for it.
                  $pEncrypt = $Setting                
                  $Setting = $Setting.ToLower()
                  $Ltrs = $Setting.Substring(0,4)
                  
                  Switch($Ltrs)
                  {

                        "temp"{
                              $TempDir = $Setting.Replace("tempdir=", "")
                        }
                  
                       "work"{
                              $WorkSpace = $Setting.Replace("workspace=", "")
                        }
                        
                        "nets"{
                              $NetScriptDir = $Setting.Replace("netscriptdir=", "")
                        }
                        
                        "ou_g"{
                              $OU_GUID = $Setting.Replace("ou_guid=", "")
                        }
                        
                        "arch"{
                              $Archive_Days = $Setting.Replace("archive_days=", "")
                        }
                        
                        "doma"{
                              $DC = $Setting.Replace("domaincontroller=", "")
                        }
                        
                        "back"{
                              $BackupDir = $Setting.Replace("backupdir=", "")
                        }
                        
                        "encp"{
                              $EncP = $pEncrypt.Replace("Encp=", "")
                        }
                        "emai"{
                              $emailTo = $Setting.Replace("email=", "")
                        }
                  
                        Default{
                              # Something isn't right..  I don't know what to do but I will at least log it to a file
                              Write-Log "Extra Data Was Found in Settings. The Error is The Data - see next" ".\lsce.log"
                              Write-Log $Setting ".\lsce.log"
                        }
                  }                 
                  
            }
      }
}

#Region "Settings Values"
# If any values are null, set the defaults
If($TempDir -eq $null){
      $TempDir = "C:\Temp"
}

If($NetScriptDir -eq $null){
      $NetScriptDir = "prnScripts"
}

If($Archive_Days -eq $null){
      $Archive_Days = "1"
}

If($DC -eq $null){
      $DC = "USINDMVGCDOM001"
}

If($BackupDir -eq $null){
      $BackupDir = Get-Location
}

If($OU_GUID -eq $null){
      $Drv_OU = "LDAP://<GUID=3f903994-1cab-43df-b8f0-fad5ed87cf58>"
}

If($WorkSpace -eq $null){
      $WorkSpace = "PrnSpace"
}

If($EncP -eq $null){
      # Password not set for email - email will not function
      Write-Log "The Encrypted Password for the Email Function is Not Set - Email will not be sent." ".\lsce.log"
}

If($emailTo -eq ""){
      # Make sure a blank is made null - Email will not function
      $emailTo = $null
}

If($emailTo -eq $null){
      # Email Send Address is not set.  Email will not function
      Write-Log "No Email Address was specified - Email will not be sent." ".\lsce.log"
}
#EndRegion

#Region "Format the values"
      $WorkSpaceDir = "$TempDir\$WorkSpace"
      $NetScriptDir = "\\$DC\netlogon\$NetScriptDir"

      # check that our NetScriptDir exists and if not, create it
      $bNetScriptDir = Test-Path($NetScriptDir)
      If($bNetScriptDir -eq $false){
      mkdir $NetScriptDir
}
      
      # Check our Backup directory exists if not, set the default
      $bBackup = Test-Path($BackupDir)
      
      If($bBackup -eq $false){
            $BackupDir = Get-Location
            Write-Log "Backup Directory does not exist. Manually create it. Defaulting to $BackupDir" ".\lsce.log"
      }
      
      $Archive_Days = [int]$Archive_Days * 1  #  Change this multiplier value based on how many times a day script will run.
#EndRegion

#EndRegion

$eAttach = ".\PrinterMappingErrors.txt"
$cr = [char]10
$Q = [char]34

#EndRegion

#Region "Main"
      Create-WorkSpace
      # Set the Guid for DP_Groups\Printer_Mappings in DP.net.
      $Drv_OU = [System.DirectoryServices.DirectoryEntry] "LDAP://<GUID=3f903994-1cab-43df-b8f0-fad5ed87cf58>"
      $Drv_Group = $Drv_OU.distinguishedName
      $Col_Drive_Groups = Get-QADGroup -SizeLimit 0 -SearchRoot $Drv_Group
      
      
      # Clean the Temp and Temp\Logon directories
      del "$TempDir\*.*"
      del "$WorkSpaceDir\*.*"
      
      sleep 10
ForEach($Group in $Col_Drive_Groups){
      # Set FileName for exporting to text
      $FileName = $Group.Name
      
      # Get the Group Distinguished Name
      $G_DN = $Group.DN
      
      # Get the Drive Letter and UNC Path for mapping the drive
      $UncPath = $Group.Description
      
      If($UncPath -eq $null){
            # If there is nothing in the UNC Path, we need to alert the server team
            $eData = "$FileName is missing a unc path for printer mapping."
            Mail-MissingUNC $eData
      }
      ElseIf($UncPath -eq ""){
            # If there is nothing in the UNC Path, we need to alert the server team
            $eData = "$FileName is missing a unc path for printer mapping."
            Mail-MissingUNC $eData  
      }
            $OnlyUsers = Get-Members $Group ""
      
            # Sort users in alphabetical order
            $OnlyUsers.Sort()
            $OnlyUsers = $OnlyUsers | Select -Unique
            
            # Export the Collection of unique users to a text file
            Create-LogonScript $OnlyUsers "lpt2" $UncPath
            
      }

      # Clean up logon scripts
      Return-Scripts # Creates the temp File for the Scripts

      $ListOfScripts = GC "$TempDir\List.txt"
      
      # Compare our new scripts to those on the server - if they are the same, don't change them
      ForEach($nScript in $ListOfScripts){
            $RemoteScript = "$NetScriptDir\$nScript"
            $LocalScript = "$WorkSpaceDir\$nScript"
            Compare-Scripts $LocalScript $RemoteScript
      }
      
      # Compare Directories and remove dead scripts from server
      Compare-Directories $NetScriptDir $WorkSpaceDir

      # Zip and copy new bat script files 
      $ArchiveFile = Return-ArchiveFile $Archive_Days
      Zip-Files $WorkSpaceDir $ArchiveFile
      
      # Mail the server team if there are any drive duplicates
      
            Mail-SvrTeam $eAttach


      # Delete the bak file   
      If(Test-Path($BakINI)){
            Remove-Item $BakINI
      }
#EndRegion






Michael Heath
System Administrator
Firestone Diversified Products, LLC
Direct: 317.428.2623  Cell: 317.503.9902

