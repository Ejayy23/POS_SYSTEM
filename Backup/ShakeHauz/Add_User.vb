Imports System.Data.OleDb
Imports System.IO

Public Class Add_User
    Dim cnn As OleDb.OleDbConnection = mycon()
    Dim imgName As String
    Dim daImage As OleDbDataAdapter
    Dim dsImage As DataSet
    Dim sql As String
    Dim result As Integer
    Dim cmd As New OleDb.OleDbCommand
    Private Sub imgsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgsave.Click
        Try
            Dim dlgImage As FileDialog = New OpenFileDialog()

            dlgImage.Filter = "Image File (*.jpg;*.bmp;*.gif;*.png)|*.jpg;*.bmp;*.gif"

            If dlgImage.ShowDialog() = DialogResult.OK Then
                imgName = dlgImage.FileName

                Dim newimg As New Bitmap(imgName)

                imgsave.SizeMode = PictureBoxSizeMode.StretchImage
                imgsave.Image = DirectCast(newimg, Image)
            End If
            dlgImage = Nothing
        Catch ae As System.ArgumentException
            imgName = " "
            MessageBox.Show(ae.Message.ToString())
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString())
        End Try
    End Sub

    Public Sub clearfields()
        For Each crt As Control In GroupBox1.Controls
            If crt.GetType Is GetType(TextBox) Then
                crt.Text = Nothing
            End If
            cbogender.Text = ""
            imgsave.Image = Nothing
        Next
    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        clearfields()
    End Sub

    Private Sub btnsave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        'new code
        cnn.Open()

        Dim arrImage() As Byte
        Dim strImage As String
        Dim myMs As New IO.MemoryStream
        '
        If Not IsNothing(imgsave.Image) Then
            Me.imgsave.Image.Save(myMs, Me.imgsave.Image.RawFormat)
            arrImage = myMs.GetBuffer
            strImage = "?"
        Else
            arrImage = Nothing
            strImage = "NULL"
        End If

        Dim cmd As New OleDb.OleDbCommand
        cmd.Connection = cnn
        cmd.CommandText = "INSERT INTO  tblusers(username,userpassword,userfname,userlname, usermname, usergender, useremail, useraddress,userimage) VALUES ('" & txtusername.Text & "','" & txtpassword.Text & "','" & txtfname.Text & "', '" & txtlname.Text & "','" & txtmname.Text & "', '" & cbogender.Text & "', '" & txtemail.Text & "', '" & txtaddress.Text & "'," & strImage & ")"
        ' cmd.CommandText = "INSERT INTO tblstudent(stdid, [name], photo) VALUES(" & Me.txtID.Text & ",'" & _
        'Me.txtName.Text & "'," & strImage & ")"

        If strImage = "?" Then
            cmd.Parameters.Add(strImage, OleDb.OleDbType.Binary).Value = arrImage
        End If
        '
        MsgBox("Data save successfully!")

        cmd.ExecuteNonQuery()
        cnn.Close()

        'fill user in combobox to refresh
        main.filluser()


        '--------------------------------------------------
        '    Try
        '        cnn.Open()
        '        sql = "INSERT INTO  tblusers(username,userpassword,userfname,userlname, usermname, usergender, useremail, useraddress,userimage) VALUES ('" & txtusername.Text & "','" & txtpassword.Text & "','" & txtfname.Text & "', '" & txtlname.Text & "','" & txtmname.Text & "', '" & cbogender.Text & "', '" & txtemail.Text & "', '" & txtaddress.Text & "'," & " @Img)"

        '        If imgName <> "" Then
        '            Dim fs As FileStream

        '            fs = New FileStream(imgName, FileMode.Open, FileAccess.Read)

        '            Dim picByte As Byte() = New Byte(fs.Length - 1) {}

        '            fs.Read(picByte, 0, System.Convert.ToInt32(fs.Length))

        '            fs.Close()
        '            Dim imgParam As New OleDbParameter()

        '            imgParam.OleDbType = OleDbType.Binary
        '            imgParam.ParameterName = "Img"
        '            imgParam.Value = picByte



        '            cmd.Parameters.Add(imgParam)

        '            With cmd
        '                .CommandText = sql
        '                .Connection = cnn
        '            End With

        '            result = cmd.ExecuteNonQuery
        '            If result > 0 Then
        '                MsgBox("User Saved")
        '                clearfields()
        '                main.filluser()
        '                cnn.Close()
        '            Else
        '                MsgBox("Access Denied!")
        '            End If
        '        End If

        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    Finally
        '        cnn.Close()
        '    End Try
    End Sub
End Class