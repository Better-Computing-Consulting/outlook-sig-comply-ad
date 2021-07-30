   Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
      Dim s2 As String = ""
      If Me.txtTitle.Text <> "" Then
         s2 = htmText.Replace("enter-name-here", Me.txtName.Text & ",")
      Else
         s2 = htmText.Replace("enter-name-here", Me.txtName.Text)
      End If
      Dim s3 As String = s2.Replace("enter-email-here", Me.txtEmail.Text)
      Dim s4 As String = s3.Replace("enter-title-here", Me.txtTitle.Text)
      Dim s5 As String = ""
      If Me.txtDirPhone.Text <> "" Then
         s5 = s4.Replace("enter-dirphone-here", "dir:&nbsp;&nbsp;" & Me.txtDirPhone.Text)
      Else
         s5 = s4.Replace("enter-dirphone-here", "")
      End If
      Dim s6 As String = ""
      If Me.txtCell.Text <> "" Then
         s6 = s5.Replace("enter-cellphone-here", "cell:&nbsp;&nbsp;" & Me.txtCell.Text)
      Else
         s6 = s5.Replace("enter-cellphone-here", "")
      End If
      Dim s7 As String = ""
      If Me.txtDirectFax.Text <> "" Then
         s7 = s6.Replace("enter-dirfax-here", "dir fax:&nbsp;&nbsp;" & Me.txtDirectFax.Text)
      Else
         s7 = s6.Replace("enter-dirfax-here", "")
      End If
      Dim s8 As String = ""
      Dim s9 As String = ""
      Dim s10 As String = ""
      Dim s11 As String = ""
      If Me.boxOffice.Text = "Los Angeles" Then
         s8 = s7.Replace("enter-address1-here", LA_address1)
         s9 = s8.Replace("enter-address2-here", LA_address2)
         s10 = s9.Replace("enter-mainphone-here", LA_phone)
         s11 = s10.Replace("enter-mainfax-here", LA_fax)
      ElseIf Me.boxOffice.Text = "San Francisco" Then
         s8 = s7.Replace("enter-address1-here", SF_address1)
         s9 = s8.Replace("enter-address2-here", SF_address2)
         s10 = s9.Replace("enter-mainphone-here", SF_phone)
         s11 = s10.Replace("enter-mainfax-here", SF_fax)
      Else
         MsgBox("Please select a valid office from the drop-down menu")
         Exit Sub
      End If
      Dim path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
      path = path.Replace("file:\", "")
      tempfile = path & "\" & userID & "-tmp-" & String.Format("{0:yyMMddHHmmss}", Now) & ".htm"
      Dim sw As StreamWriter = File.CreateText(tempfile)
      sw.Write(s11)
      sw.Flush()
      sw.Close()
      Me.WebBrowser1.Navigate(New Uri(tempfile))
   End Sub