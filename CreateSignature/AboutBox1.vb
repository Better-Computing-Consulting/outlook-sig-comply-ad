Public NotInheritable Class ProgInfo

   Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Dim info As String = ""
      info &= "Usage: " & vbCrLf & vbCrLf
      info &= "Enter User ID, click Check." & vbCrLf & vbCrLf
      info &= "The program will prefill form and preview signature based on the user's " & vbCrLf & "information in Active Directory." & vbCrLf & vbCrLf
      info &= "You may edit the fields directly and click Preview to update the signature." & vbCrLf & vbCrLf
      info &= "The preferred way for editing the signature is by changing the user profile in " & vbCrLf & "Active directory and clicking on Check again." & vbCrLf & vbCrLf
      info &= "The fields in the form link to user's properties in Active Directory " & vbCrLf & "as follows:" & vbCrLf & vbCrLf
      info &= "Form Field".PadRight(20, " ") & "Tab".PadRight(21, " ") & "Property".PadRight(60, " ") & vbCrLf
      info &= "".PadRight(60, "=") & vbCrLf
      info &= "Name".PadRight(20) & "General".PadRight(21, " ") & "First Name +" & vbCrLf
      info &= "".PadRight(20) & "".PadRight(21, " ") & "Initials +" & vbCrLf
      info &= "".PadRight(20) & "".PadRight(21, " ") & "Last Name" & vbCrLf
      info &= "Title".PadRight(20) & "Organization".PadRight(21, " ") & "Title" & vbCrLf
      info &= "Position".PadRight(20) & "General".PadRight(21, " ") & "Description" & vbCrLf
      info &= "Direct Line".PadRight(20) & "General".PadRight(21, " ") & "Telephone Number" & vbCrLf
      info &= "Cell Phone".PadRight(20) & "Telephones".PadRight(21, " ") & "Mobile" & vbCrLf
      info &= "Direct Fax".PadRight(20) & "Telephones".PadRight(21, " ") & "Fax" & vbCrLf
      info &= "Email Address".PadRight(20) & "General".PadRight(21, " ") & "E-mail" & vbCrLf
      info &= "Office".PadRight(20) & "Profile".PadRight(21, " ") & "Home folder" & vbCrLf & vbCrLf
      info &= "If a San Francisco user shows the Los Angeles information it means that his" & vbCrLf & "home directory is not in a San Fransico server. "
      info &= "The information in the signature" & vbCrLf & "can be corrected by selecting the San Francisco office in the drop down box," & vbCrLf & "but his home directory should be "
      info &= "moved to a San Francisco Server." & vbCrLf & vbCrLf
      info &= "Click Save to create an htm file of the signature (as displayed in the screen) in:" & vbCrLf & vbCrLf
      info &= "\\fpas\Apps\UserSignatures\<USERID>\firmsignature.htm" & vbCrLf & vbCrLf & vbCrLf
      info &= "To apply signature to a User's Outlook:" & vbCrLf & vbCrLf
      info &= "1) Add user to Active Directory group ""Apply Outlook Firm Signature""" & vbCrLf & vbCrLf
      info &= "2) Have the user logout and log back into the computer." & vbCrLf & vbCrLf
      Me.Label1.Text = info
   End Sub
End Class
