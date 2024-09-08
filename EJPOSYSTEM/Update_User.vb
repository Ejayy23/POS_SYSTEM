Imports System.Data.OleDb
Imports System.IO

Public Class Update_User

    Dim cnn As OleDb.OleDbConnection = mycon()
    Dim imgName As String
    Dim daImage As OleDbDataAdapter
    Dim dsImage As DataSet
    Dim sql As String
    Dim result As Integer
    Dim cmd As New OleDb.OleDbCommand
    Dim ds As New DataSet
    Dim tables As DataTableCollection
    Dim da As New OleDb.OleDbDataAdapter

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        'NEW CODE 
        cnn.Open()

        Dim arrImage() As Byte
        Dim strImage As String
        Dim myMs As New IO.MemoryStream
        '
        If Not IsNothing(imgupdate.Image) Then
            Me.imgupdate.Image.Save(myMs, Me.imgupdate.Image.RawFormat)
            arrImage = myMs.GetBuffer
            strImage = "?"
        Else
            arrImage = Nothing
            strImage = "NULL"
        End If

        Dim cmd As New OleDb.OleDbCommand
        cmd.Connection = cnn
        cmd.CommandText = "UPDATE tblusers SET userfname='" & txtfname.Text & "',userlname= '" & txtlname.Text & "', " & _
        " usermname='" & txtmname.Text & "', useraddress='" & txtaddress.Text & "', " & _
        " useremail='" & txtemail.Text & "', username= '" & txtusername.Text & "' ," & _
        " userpassword='" & txtpassword.Text & "', usergender= '" & cbogender.Text & "'," & _
        " userimage = " & strImage & " WHERE userID = " & cboID.Text

        ' cmd.CommandText = "INSERT INTO tblstudent(stdid, [name], photo) VALUES(" & Me.txtID.Text & ",'" & _
        'Me.txtName.Text & "'," & strImage & ")"

        If strImage = "?" Then
            cmd.Parameters.Add(strImage, OleDb.OleDbType.Binary).Value = arrImage
        End If
        '
        MsgBox("User updated successfully!")

        cmd.ExecuteNonQuery()
        cnn.Close()
        'fill user in combobox to refresh
        main.filluser()



        '-----------------------------------------------------------
        'Try
        '    Dim result As DialogResult = MessageBox.Show("Confirm update?", "", MessageBoxButtons.YesNo)

        '    If result = Windows.Forms.DialogResult.Yes Then
        '        cnn.Open()

        'sql = "UPDATE tblusers SET userfname='" & txtfname.Text & "',userlname= '" & txtlname.Text & "', " & _
        '" usermname='" & txtmname.Text & "', useraddress='" & txtaddress.Text & "', " & _
        '" useremail='" & txtemail.Text & "', username= '" & txtusername.Text & "' ," & _
        '" userpassword='" & txtpassword.Text & "', usergender= '" & cbogender.Text & "'," & _
        '" userimage = " & " @Img " & " WHERE userID = " & cboID.Text

        '        ' sql = "INSERT INTO  tblusers(username,userpassword,userfname,userlname, usermname, usergender, useremail, useraddress,userimage) VALUES ('" & txtusername.Text & "','" & txtpassword.Text & "','" & txtfname.Text & "', '" & txtlname.Text & "','" & txtmname.Text & "', '" & cbogender.Text & "', '" & txtemail.Text & "', '" & txtaddress.Text & "'," & " @Img)"
        '        'get image from database

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
        '                MsgBox("User Updated")
        '                'clearfields()
        '                cnn.Close()
        '            Else
        '                MsgBox("Access Denied!")
        '            End If
        '        End If
        '    Else
        '        MsgBox("Access Denied!")
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally
        '    cnn.Close()
        'End Try
    End Sub

    Private Sub imgsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgupdate.Click
        Try
            Dim dlgImage As FileDialog = New OpenFileDialog()

            dlgImage.Filter = "Image File (*.jpg;*.bmp;*.gif;*.png)|*.jpg;*.bmp;*.gif"

            If dlgImage.ShowDialog() = DialogResult.OK Then
                imgName = dlgImage.FileName

                Dim newimg As New Bitmap(imgName)

                imgupdate.SizeMode = PictureBoxSizeMode.StretchImage
                imgupdate.Image = DirectCast(newimg, Image)
            End If
            dlgImage = Nothing
        Catch ae As System.ArgumentException
            imgName = " "
            MessageBox.Show(ae.Message.ToString())
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString())
        End Try
    End Sub

    Public Sub filluser()

        'fill user ID in a combobox
        ds = New DataSet
        tables = ds.Tables
        da = New OleDb.OleDbDataAdapter("SELECT userID from tblusers", mycon)
        da.Fill(ds, "tblusers")
        Dim view1 As New DataView(tables(0))
        With cboID
            .DataSource = ds.Tables("tblusers")
            .DisplayMember = "userID"
            .ValueMember = "userID"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With

    End Sub

    Private Sub Update_User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        filluser()
    End Sub

    Private Sub cboID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboID.SelectedIndexChanged
        Dim sql As String
        Dim cmd As New OleDb.OleDbCommand
        cmd.Parameters.AddWithValue("@userID", 3)
        Dim dr As OleDb.OleDbDataReader
        Try
            sql = "SELECT * FROM tblusers where userID =" & cboID.Text & ""

            cnn.Open()
            With cmd
                .CommandText = sql
                .Connection = cnn
            End With
            dr = cmd.ExecuteReader
            While dr.Read()


                txtfname.Text = dr("userfname")
                txtlname.Text = dr("userlname")
                txtmname.Text = dr("usermname")
                txtfullname.Text = txtlname.Text + ", " + txtfname.Text + " " + txtmname.Text
                txtemail.Text = dr("useremail")
                txtaddress.Text = dr("useraddress")
                txtusername.Text = dr("username")
                txtpassword.Text = dr("userpassword")
                cbogender.Text = dr("usergender")

                Dim data As Byte() = DirectCast(dr("userimage"), Byte())
                Dim ms As New MemoryStream(data)
                imgupdate.Image = Image.FromStream(ms)
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()


        End Try
    End Sub
End Class