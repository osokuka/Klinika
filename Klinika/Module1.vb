Imports System.Data
Imports System.Data.OleDb
Module Module1
    Public Function filldatagrid()
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim tables As DataTableCollection
        Dim source1 As New BindingSource
        ds = New DataSet
        tables = ds.Tables
        Dim con As OleDbConnection = New OleDbConnection(Form1.cn)
        da = New OleDbDataAdapter("SELECT Pacientet.Emri, Pacientet.Mbiemri, Pacientet.Emriprindit AS [Emri Prindit], Pacientet.Ditlindja, Pacientet.Adresa, Pacientet.NrTelefonit, Pacientet.Email, Pacientet.Nrpersonal, Pacientet.kartashendetsore, Pacientet.Kartafam, Pacientet.Datargjistrimit, Pacientet.RegjistruarNga, ID , gjinia FROM Pacientet;", con) 'Change items to your database name
        da.Fill(ds, "items") 'Change items to your database name
        Dim view As New DataView(tables(0))
        source1.DataSource = view
        Form1.DataGridView1.DataSource = view
        da.Dispose()
        tables.Clear()
        Return Nothing

    End Function

    Public Function fillmjekun()
        'im cn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=d:\northwind.mdb")
        Dim da As New OleDbDataAdapter()
        Dim dt As New DataTable()
        Dim con As New OleDbConnection(Form1.cn)
        Try
            con.Open()
            da.SelectCommand = New OleDbCommand("select EmriMjekut &' '& Mbiemri as Mjeku, ID from Mjeket", con)
            da.Fill(dt)

            Form1.ComboBox2.DataSource = dt
            Form1.ComboBox2.DisplayMember = "Mjeku"
            Form1.ComboBox2.ValueMember = "ID"

            ' or this way


        Catch ex As Exception

        End Try
        con.Close()

        Return Nothing
        Form1.ComboBox2.Text = ""
    End Function

    Public Function fillkalendarin()
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim tables As DataTableCollection
        Dim source1 As New BindingSource
        ds = New DataSet
        tables = ds.Tables
        Dim con As OleDbConnection = New OleDbConnection(Form1.cn)
        da = New OleDbDataAdapter("TRANSFORM First(Pacientet.Emri & ' '& Pacientet.Emriprindit & ' ' & Pacientet.Mbiemri) AS Pacienti SELECT Format(takimet.koha,'hh:nn:ss') AS Koha FROM (SELECT takimet.* FROM takimet  WHERE (((takimet.Data) between dateadd('d',0, Date()) and dateadd('d',8, Date())))) AS T LEFT JOIN Pacientet ON T.pacienti = Pacientet.ID GROUP BY Format(T.koha,'hh:nn:ss') PIVOT Format([Data],'MM-DD-YY');", con) 'Change items to your database name
        da.Fill(ds, "items") 'Change items to your database name
        Dim view As New DataView(tables(0))
        source1.DataSource = view
        Form1.DataGridView2.DataSource = view
        da.Dispose()
        tables.Clear()
        Return Nothing

    End Function
    Public Function cellscolors()
        'Try
        '    If Form1.DataGridView2.Rows.Count > 0 Then
        '        For i As Integer = 0 To Form1.DataGridView2.Rows.Count - 1
        '            For j As Integer = 0 To Form1.DataGridView2.Columns.Count - 1
        '                Dim CellChange As String = Form1.DataGridView2.Rows(i).Cells(j).Value.ToString().Trim()
        '                If CellChange.Contains(Nothing) = False Then

        '                    With Form1.DataGridView2
        '                        .Rows(i).Cells(j).Style.BackColor = Color.Green
        '                    End With

        '                End If
        '            Next
        '        Next
        '    End If
        'Catch e As Exception
        '    MessageBox.Show(e.ToString())
        'End Try
    End Function
End Module
