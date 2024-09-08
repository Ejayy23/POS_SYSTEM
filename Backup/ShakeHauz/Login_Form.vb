Public Class Login_Form
    Dim cnn As OleDb.OleDbConnection = mycon()
    Dim cmd As New OleDb.OleDbCommand
    Dim da As New OleDb.OleDbDataAdapter
    Dim sql As String

    Private Sub btnlogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnlogin.Click
        Dim dt As New DataTable
        sql = "SELECT * FROM tblusers WHERE username ='" & txtuser.Text & "' AND userpassword = '" & txtpassword.Text & "'"
        Try
            cnn.Open()
            With cmd
                .Connection = cnn
                .CommandText = Sql
            End With
            'FILLING THE DATA IN A SPICIFIC TABLE OF THE DATABASE
            da.SelectCommand = cmd
            dt = New DataTable
            da.Fill(dt)

            'DECLARING AN INTEGER TO SET THE MAXROWS OF THE TABLE
            Dim maxrow As Integer = dt.Rows.Count

            'CHECKING IF THE DATA IS EXIST IN THE ROW OF THE TABLE
            If maxrow > 0 Then
                'MsgBox("Welcome " & dt.Rows(0).Item(1))

                main.tlslogin.Text = "Logout"

                main.TabControl1.Visible = True

                If txtuser.Text = "Admin" Or txtuser.Text = "admin" Then
                    main.GroupBox5.Enabled = True
                Else
                    main.GroupBox5.Enabled = False
                End If

                Me.Hide()

                txtuser.Text = ""
                txtpassword.Text = ""
            Else
                MsgBox("Invalid User Account. Please try again.")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        cnn.Close()
    End Sub
End Class