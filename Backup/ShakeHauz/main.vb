Imports System.IO
'Imports System.Drawing
Public Class main

    Dim cnn As OleDb.OleDbConnection = mycon()
    Dim mySQLCommand As New OleDb.OleDbCommand
    Dim mySQLStrg As String
    Dim ds As DataSet
    Dim da As New OleDb.OleDbDataAdapter
    Dim tables As DataTableCollection
    Dim sum As Double


    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        salesrecord()
        filluser()

    End Sub


    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        Dim g As Graphics
        Dim sText As String
        Dim iX As Integer
        Dim iY As Integer
        Dim sizeText As SizeF
        Dim ctlTab As TabControl

        ctlTab = CType(sender, TabControl)
        g = e.Graphics
        sText = ctlTab.TabPages(e.Index).Text
        sizeText = g.MeasureString(sText, ctlTab.Font)
        iX = e.Bounds.Left + 6
        iY = e.Bounds.Top + (e.Bounds.Height - sizeText.Height) / 2
        g.DrawString(sText, ctlTab.Font, Brushes.Black, iX, iY)
    End Sub

    Private Sub TextBox7_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        If sum > Val(TextBox7.Text) Then
            TextBox8.Text = "0"
        Else
            TextBox8.Text = Val(TextBox7.Text) - sum
        End If
    End Sub

    Private Sub main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Login_Form.txtuser.Text = "Administrator" Or Login_Form.txtuser.Text = "administrator" Then
            GroupBox5.Enabled = True
        Else
            GroupBox5.Enabled = False
        End If
        TabControl1.Visible = False
        Timer1.Start()
        'load list of shake
        fillshake()
        'load list of foods
        fillfoods()
        'load list of user ID
        filluser()
        clearfields()
        'load sales record
        salesrecord()
    End Sub

    Public Sub fillshake()

        'fill list of shakes in a combobox
        ds = New DataSet
        tables = ds.Tables
        da = New OleDb.OleDbDataAdapter("SELECT itemname from tblitems where itemtype = 'shake'", mycon)
        da.Fill(ds, "tblitems")
        Dim view1 As New DataView(tables(0))
        With ComboBox1
            .DataSource = ds.Tables("tblitems")
            .DisplayMember = "itemname"
            .ValueMember = "itemname"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub

    Public Sub fillfoods()
        'fill list of foods in a combobox

        ds = New DataSet
        tables = ds.Tables
        da = New OleDb.OleDbDataAdapter("SELECT itemname from tblitems where itemtype = 'foods'", mycon)
        da.Fill(ds, "tblitems")
        Dim view1 As New DataView(tables(0))
        With ComboBox2
            .DataSource = ds.Tables("tblitems")
            .DisplayMember = "itemname"
            .ValueMember = "itemname"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
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

    Public Sub clearfields()

        ComboBox1.Text = "SELECT FRUIT SHAKE"
        ComboBox2.Text = "SELECT FAST FOOD"

        For Each crt As Control In GroupBox3.Controls
            If crt.GetType Is GetType(TextBox) Then
                crt.Text = Nothing
            End If
        Next
        Dim Rbtn As RadioButton
        For Each Rbtn In GroupBox6.Controls
            If TypeOf Rbtn Is RadioButton Then Rbtn.Checked = False
        Next
        For Each Rbtn In GroupBox7.Controls
            If TypeOf Rbtn Is RadioButton Then Rbtn.Checked = False
        Next
    End Sub

    Public Sub clearfields2()

        For Each crt As Control In GroupBox2.Controls
            If crt.GetType Is GetType(TextBox) Then
                crt.Text = Nothing
            End If
            ListBox2.Items.Clear()
        Next
    End Sub

    Public Sub fillfruitprice()
        Dim sql As String
        Dim cmd As New OleDb.OleDbCommand
        Dim dr As OleDb.OleDbDataReader
        Try
            sql = "SELECT * FROM tblitems where itemname = '" & ComboBox1.Text & "' "

            cnn.Open()
            With cmd
                .CommandText = sql
                .Connection = cnn
            End With
            dr = cmd.ExecuteReader
            While dr.Read()

                If rdbmed.Checked = True Then
                    TextBox3.Text = dr("itemmediumprice")
                ElseIf rdblarge.Checked = True Then
                    TextBox3.Text = dr("itemlargeprice")
                ElseIf rdbreg.Checked = True Then
                    TextBox3.Text = dr("itemregularprice")
                End If

            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()
        End Try
    End Sub

    Public Sub fillfastfoodprice()
        Dim sql As String
        Dim cmd As New OleDb.OleDbCommand
        Dim dr As OleDb.OleDbDataReader
        Try
            sql = "SELECT * FROM tblitems where itemname = '" & ComboBox2.Text & "'"

            cnn.Open()
            With cmd
                .CommandText = sql
                .Connection = cnn
            End With
            dr = cmd.ExecuteReader
            While dr.Read()
                If rdb2med.Checked = True Then
                    TextBox4.Text = dr("itemmediumprice")
                ElseIf rdb3large.Checked = True Then
                    TextBox4.Text = dr("itemlargeprice")
                ElseIf rdb1reg.Checked = True Then
                    TextBox4.Text = dr("itemregularprice")
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        For x As Integer = 0 To ListBox2.Items.Count - 1
            sum += Val(ListBox2.Items.Item(x).ToString)
        Next
        txttotal.Text = sum.ToString
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Add_User.Show()
    End Sub
   
    Private Sub btnloaduserinfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnloaduserinfo.Click

        Dim sql As String
        Dim cmd As New OleDb.OleDbCommand
        cmd.Parameters.AddWithValue("@userID", 3)
        Dim dr As OleDb.OleDbDataReader
        Try
            sql = "SELECT * FROM tblusers where userID =" & cboID.Text & ""
            If Not cnn.State = ConnectionState.Open Then
                cnn.Open()
            End If

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
                imgretrieve.Image = Image.FromStream(ms)
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()


        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Update_User.Show()

        Show()

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim total1 As Double
        Dim total2 As Double

        If ComboBox1.Text = "SELECT FRUIT SHAKE" And ComboBox2.Text = "SELECT FAST FOOD" Then


        ElseIf ComboBox1.Text = "SELECT FRUIT SHAKE" Then
            total1 = Val(TextBox4.Text) * Val(TextBox6.Text)
            ListBox2.Items.Add(total1)
            items.Text &= ComboBox2.Text + " (x" + TextBox6.Text + "), " + vbCrLf
        ElseIf ComboBox2.Text = "SELECT FAST FOOD" Then
            total2 = Val(TextBox3.Text) * Val(TextBox5.Text)
            ListBox2.Items.Add(total2)
            items.Text &= ComboBox1.Text + " (x" + TextBox5.Text + "), " + vbCrLf
        Else
            total2 = Val(TextBox3.Text) * Val(TextBox5.Text)
            total1 = Val(TextBox4.Text) * Val(TextBox6.Text)
            ListBox2.Items.Add(total1)
            ListBox2.Items.Add(total2)
            items.Text &= ComboBox1.Text + " (x" + TextBox5.Text + "), " + vbCrLf & ComboBox2.Text + " (x" + TextBox6.Text + "), " + vbCrLf
        End If





        'Dim Newline As String
        'Newline = System.Environment.NewLine
        'If ComboBox1.Text = "SELECT FRUIT SHAKE" Then

        'Else
        '    Dim fruit As String = ComboBox1.Text
        '    total1 = Val(TextBox3.Text) * Val(TextBox5.Text)
        '    ListBox2.Items.Add(total1)
        '    '------------USING LISTBOX
        '    'ListBox1.Items.Add(fruit + " (x" + TextBox5.Text + ")")
        '    '------------USING TEXTBOX
        '    items.Text = items.Text + Newline & ComboBox1.Text + " (x" + TextBox5.Text + "), "

        'End If

        'If ComboBox2.Text = "SELECT FAST FOOD" Then

        'Else

        '    Dim foods As String = ComboBox2.Text
        '    total1 = Val(TextBox4.Text) * Val(TextBox6.Text)
        '    ListBox2.Items.Add(total1)
        '    '-------------USING LISTBOX
        '    'ListBox1.Items.Add(foods + " (x" + TextBox6.Text + ")")
        '    '-------------USING TEXTBOX
        '    items.Text = items.Text + Newline & ComboBox2.Text + " (x" + TextBox6.Text + "), "

        'End If

        clearfields()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
       
        clearfields()
    End Sub

    Private Sub tncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tncancel.Click
        clearfields2()
    End Sub

    Private Sub rdbreg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbreg.CheckedChanged
        fillfruitprice()
    End Sub

    Private Sub rdbmed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbmed.CheckedChanged
        fillfruitprice()
    End Sub

    Private Sub rdblarge_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdblarge.CheckedChanged
        fillfruitprice()
    End Sub

    Private Sub rdb1reg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdb1reg.CheckedChanged
        fillfastfoodprice()
    End Sub

    Private Sub rdb2med_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdb2med.CheckedChanged
        fillfastfoodprice()
    End Sub

    Private Sub rdb3large_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdb3large.CheckedChanged
        fillfastfoodprice()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim sql As String
        Dim result As Integer
        Dim cmd As New OleDb.OleDbCommand

        Try
            Dim Dresult As DialogResult = MessageBox.Show("Are you sure you want to delete this user?", "Shakehauz Cashiering System", MessageBoxButtons.YesNo)
            If Dresult = Windows.Forms.DialogResult.Yes Then
                sql = "DELETE * FROM tblusers  WHERE userID=" & cboID.Text
                cnn.Open()
                With cmd
                    .CommandText = sql
                    .Connection = cnn
                End With

                result = cmd.ExecuteNonQuery
                If result > 0 Then
                    MsgBox("User Deleted")
                    filluser()
                    btnloaduserinfo.PerformClick()
                    cnn.Close()

                Else
                    MsgBox("NO RECORD HASS BEEN DELTED!")
                End If
            Else

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()


        End Try
    End Sub

  

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lbltime.Text = TimeOfDay
        lbldate.Text = Date.Today
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        recordbydate()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        salesrecord()
    End Sub

    Private Sub btnreport_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnreport.Click
        message = "all reports"
        frmReports.Show()
    End Sub

    Private Sub btndailyreport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndailyreport1.Click
        message = "daily reports"
        frmReports.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        Dim cmd As New OleDb.OleDbCommand
        Dim result As Integer

        Try
            cnn.Open()
            sql = "INSERT INTO  tblsales(customer_name, customer_address, customer_odered_items, customer_total_payment, customer_cash_ammount, customer_cash_changed,customer_date_ordered,customer_time_ordered ) VALUES ('" & txtcustomername.Text & "', '" & txtcustomeraddress.Text & "','" & items.Text & ", " & "', '" & txttotal.Text & "', '" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & lbldate.Text & "', '" & lbltime.Text & "')"

           
            With cmd
                .CommandText = sql
                .Connection = cnn
            End With

          
            result = cmd.ExecuteNonQuery
            'clear textboxes

            If result > 0 Then
                MsgBox("Saved")
                sum = 0
                clearfields2()
                'connection close
                cnn.Close()
                
            Else
                MsgBox("Access Denied!")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()
        End Try
    End Sub

 
    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tlslogin.Click

        If tlslogin.Text = "Logout" Then
            MsgBox("Logged Out!")
            TabControl1.Visible = False
            tlslogin.Text = "Login"

        Else
            Login_Form.Show()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Me.Close()
        Login_Form.Close()

    End Sub

 
    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        DateTimePicker1.Text = Today
        salesrecord()
    End Sub
End Class