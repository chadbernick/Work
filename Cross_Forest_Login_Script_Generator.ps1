cls

#Region "Comment Block"
      '@
            Script .....:     Logon Script Creation Engine
            Purpose     ....: Dynamically create logon scripts for users
                                    based on group membership.
                                    
            Prerequisites:    
                                    1.  Place/Create groups in a single OU
                                    2.  Group name must contain the Drive letter as the first character
                                    3.  Only the UNC Path must be written in the Description Field of the Group
                                    4.  OU GUID Must be utilized in the script
            
      '     
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

Function SampleGet-Members($GroupName, $GroupList){
      $retUserCol = New-Object System.Collections.ArrayList
      
      Write-Host "You are in Get-Members" -ForegroundColor Green
      [Void]$retUserCol.Add("NothingHere")
      [Void]$retUserCol.Add("NothingHereEither")
      [Void]$retUserCol.Add("NorIsThereAnythingHere")
      
      write-host $retUserCol -ForegroundColor Red
      $GetMembers = $retUserCol
      Write-Host "GetMembers:  $GetMembers" -ForegroundColor Yellow
      Return $GetMembers
}

#EndRegion

 #Region "Variables & Declarations"
# Set up temp collections
$Collection = New-Object System.Collections.ArrayList

# Collections used in Function Get-NestedGroups
$Col_Members = New-Object System.Collections.ArrayList
$Col_Groups = New-Object System.Collections.ArrayList
$Col_LoopCheck = New-Object System.Collections.ArrayList

# Used in place of $ResultantSet_of_GroupMembers
$Collection_Results = New-Object System.Collections.ArrayList

#EndRegion

#Region "Main"
      # Set the Guid for DP_Groups\Drive_Mappings in DP.net.
      $Drv_OU = [System.DirectoryServices.DirectoryEntry] "LDAP://<GUID=67ebc32b-9934-42a9-ac9c-1be4c81c86f2>"
      $Drv_Group = $Drv_OU.distinguishedName
      $Col_Drive_Groups = Get-QADGroup -SizeLimit 0 -SearchRoot $Drv_Group

ForEach($Group in $Col_Drive_Groups){
      # Set FileName for exporting to CSV
      $FileName = $Group.Name
      $G_DN = $Group.DN
      
      
      $OnlyUsers = Get-Members $Group ""
      
      # Sort users in alphabetical order
      $OnlyUsers.Sort()
      
      # Export the Collection of unique users to a text file
      $OnlyUsers | Select -Unique | Out-File "C:\temp\$FileName.txt"
      Write-Host "Write to File" -ForegroundColor Yellow
      }

#EndRegion


