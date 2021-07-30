Imports System.IO
Imports System.DirectoryServices
Public Class Form1
#Region " Variables"
   Dim htmText As String = ""
   Dim tmpHtmText = ""
   Dim userID As String = ""
   Dim userFName As String = ""
   Dim userInitials As String = ""
   Dim userLName As String = ""
   Dim userTitle As String = ""
   Dim userPosition As String = ""
   Dim userDirPhone As String = ""
   Dim userCellPhone As String = ""
   Dim userDirFax As String = ""
   Dim userEmail As String = ""
   Dim userHome As String = ""
   Dim LA_address1 As String = "1100 Glendon Avenue&nbsp;|&nbsp;14th Floor"
   Dim LA_address2 As String = "Los Angeles, CA 90024.3503"
   Dim LA_phone As String = "310.500.3500"
   Dim LA_fax As String = "310.500.3501"
   Dim SF_address1 As String = "199 Fremont Street | 20th Floor"
   Dim SF_address2 As String = "San Francisco, CA 94105.2255"
   Dim SF_phone As String = "415.489.7700"
   Dim SF_fax As String = "415.489.7701"
   Dim tempfile = ""
   Dim app_path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\", "")
#End Region
   Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
      Dim precode As String = "<tr><td nowrap><div><font face=""Berlin Sans FB"" size=2>"
      Dim postcode As String = "</font></div></td></tr>"
      Dim s As String = ""
      If Me.txtTitle.Text <> "" Then
         s = htmText.Replace("enter-name-here", Me.txtName.Text & ",")
      Else
         s = htmText.Replace("enter-name-here", Me.txtName.Text)
      End If
      s = s.Replace("enter-email-here", Me.txtEmail.Text)
      s = s.Replace("enter-title-here", Me.txtTitle.Text)
      s = s.Replace("enter-position-here", Me.txtPosition.Text)
      If Trim(Me.txtDirPhone.Text) <> "" Then
         s = s.Replace("enter-dirphone-here", precode & "dir:&nbsp;&nbsp;" & Trim(Me.txtDirPhone.Text) & postcode)
      Else
         s = s.Replace("enter-dirphone-here", "")
      End If
      If Trim(Me.txtCell.Text) <> "" Then
         s = s.Replace("enter-cellphone-here", precode & "cell:&nbsp;&nbsp;" & Trim(Me.txtCell.Text) & postcode)
      Else
         s = s.Replace("enter-cellphone-here", "")
      End If
      If trim(Me.txtDirectFax.Text) <> "" Then
         s = s.Replace("enter-dirfax-here", precode & "dir fax:&nbsp;&nbsp;" & Trim(Me.txtDirectFax.Text) & postcode)
      Else
         s = s.Replace("enter-dirfax-here", "")
      End If
      If InStr(Me.boxOffice.SelectedItem, "Los Angeles") Then
         s = s.Replace("enter-address1-here", LA_address1)
         s = s.Replace("enter-address2-here", LA_address2)
         s = s.Replace("enter-mainphone-here", LA_phone)
         s = s.Replace("enter-mainfax-here", LA_fax)
      Else
         s = s.Replace("enter-address1-here", SF_address1)
         s = s.Replace("enter-address2-here", SF_address2)
         s = s.Replace("enter-mainphone-here", SF_phone)
         s = s.Replace("enter-mainfax-here", SF_fax)
      End If
      tempfile = app_path & "\" & userID & "-tmp-" & String.Format("{0:yyMMddHHmmss}", Now) & ".htm"
      Dim sw As StreamWriter = File.CreateText(tempfile)
      tmpHtmText = s
      sw.Write(tmpHtmText)
      sw.Flush()
      sw.Close()
      Me.WebBrowser1.Navigate(New Uri(tempfile))
   End Sub
   Function SetHTMText() As String
      Dim s As String = ""
      Dim objStreamReader As StreamReader
      Dim strLine As String
      objStreamReader = New StreamReader(app_path & "\template.htm")
      strLine = objStreamReader.ReadLine
      Do While Not strLine Is Nothing
         s &= strLine & vbCrLf
         strLine = objStreamReader.ReadLine
      Loop
      objStreamReader.Close()
      Return s
   End Function
   Private Sub btnCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheck.Click
      Dim sName As String = Trim(Me.txtUserID.Text)
      If sName = "" Then
         MsgBox("Cannot check a blank User ID")
         Exit Sub
      End If
      Dim entry As New DirectoryEntry("LDAP://lysdom")
      Dim seach As New DirectorySearcher(entry)
      seach.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & sName & "))"
      Try
         If seach.FindAll.Count = 0 Then
            MsgBox(sName & " User ID not found")
            Exit Sub
         Else
            For Each r As SearchResult In seach.FindAll
               userID = UCase(GetProperty(r, "sAMAccountName"))
               userEmail = LCase(GetProperty(r, "mail"))
               userFName = Trim(GetProperty(r, "givenName"))
               userInitials = Trim(GetProperty(r, "initials"))
               userLName = Trim(GetProperty(r, "sn"))
               userTitle = Trim(GetProperty(r, "title"))
               userPosition = Trim(GetProperty(r, "description"))
               userDirPhone = Trim(GetProperty(r, "telephoneNumber"))
               userDirFax = Trim(GetProperty(r, "facsimileTelephoneNumber"))
               userCellPhone = Trim(GetProperty(r, "mobile"))
               userHome = LCase(GetProperty(r, "homeDirectory"))
            Next
            Me.txtUserID.Text = Me.userID
            If Me.userInitials <> "" Then
               If Me.userInitials.EndsWith(".") Then
                  Me.txtName.Text = Me.userFName & " " & Me.userInitials & " " & Me.userLName
               Else
                  Me.txtName.Text = Me.userFName & " " & Me.userInitials & ". " & Me.userLName
               End If
            Else
               Me.txtName.Text = Me.userFName & " " & Me.userLName
            End If
            Me.txtEmail.Text = Me.userEmail
            Me.txtDirPhone.Text = Me.userDirPhone
            Me.txtCell.Text = Me.userCellPhone
            Me.txtDirectFax.Text = Me.userDirFax
            Me.txtTitle.Text = Me.userTitle
            Me.txtPosition.Text = Me.userPosition
            Me.btnSave.Enabled = True
            Me.btnPreview.Enabled = True
            Me.boxOffice.Enabled = True
            If InStr(userHome, "\\sf") Then
               Me.boxOffice.SelectedItem = "San Francisco"
            Else
               Me.boxOffice.SelectedItem = "Los Angeles"
            End If
            Me.btnPreview_Click(sender, e)
         End If
      Catch ex As Exception
         MsgBox(ex.Message)
      End Try
   End Sub
   Function GetProperty(ByVal r As SearchResult, ByVal PropertyName As String) As String
      Dim s As String = ""
      If r.Properties.Contains(PropertyName) Then
         s = r.Properties.Item(PropertyName)(0)
      End If
      Return s
   End Function
   Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      htmText = SetHTMText()
      Me.boxOffice.Enabled = False
      Me.btnPreview.Enabled = False
      Me.btnSave.Enabled = False
   End Sub
   Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
      Dim dest_path As String = "\\fpas\Apps\UserSignatures\" & userID
      If Not Directory.Exists(dest_path) Then
         Directory.CreateDirectory(dest_path)
      End If
      Dim sr As StreamWriter
      sr = New StreamWriter(dest_path & "\firmsignature.htm")
      sr.Write(tmpHtmText)
      sr.Close()
      File.Copy(app_path & "\lysr_logo.jpg", dest_path & "\lysr_logo.jpg", True)
      MsgBox("Signature file created in " & dest_path & "\firmsignature.htm")
   End Sub
   Private Sub txtUserID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserID.TextChanged
      Me.btnCheck.Enabled = True
   End Sub
   Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
      Dim h As New ProgInfo
      h.ShowDialog()
   End Sub
   Private Sub WebBrowser1_Navigated(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated
      File.Delete(tempfile)
   End Sub

End Class
