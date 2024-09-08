Module Loadsalesrecord
    Dim da As New OleDb.OleDbDataAdapter
    Dim dt As DataTable
    Dim sql As String
    Dim cnn As OleDb.OleDbConnection = mycon()
    Dim cmd As New OleDb.OleDbCommand

    Public Sub salesrecord()
        dt = New DataTable
        Try
            sql = "SELECT customer_receipt as [Receipt No], customer_name as [Name],customer_odered_items as [Ordered Items],customer_total_payment as [Total Payment],customer_cash_ammount as [Cash Amount],customer_cash_changed as [Cash Change],customer_date_ordered as [Date],customer_time_ordered as [Time] FROM tblsales"
            cnn.Open()
            With cmd
                .CommandText = sql
                .Connection = cnn
            End With

            da.SelectCommand = cmd
            da.Fill(dt)
            main.DataGridView1.DataSource = dt

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()
        End Try
    End Sub

    Public Sub recordbydate()
        dt = New DataTable
        Try
            sql = "SELECT customer_receipt as [Receipt No], customer_name as [Name],customer_odered_items as [Ordered Items],customer_total_payment as [Total Payment],customer_cash_ammount as [Cash Amount],customer_cash_changed as [Cash Change],customer_date_ordered as [Date],customer_time_ordered as [Time] FROM tblsales WHERE customer_date_ordered = '" & main.DateTimePicker1.Text & "'"
            cnn.Open()
            With cmd
                .CommandText = sql
                .Connection = cnn
            End With

            da.SelectCommand = cmd
            da.Fill(dt)
            main.DataGridView1.DataSource = dt

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cnn.Close()
        End Try
    End Sub
End Module
