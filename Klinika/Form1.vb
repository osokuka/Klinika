Imports System.Data
Imports System.Data.OleDb




Public Class Form1
    Public cn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\038972\Desktop\Klinika\Klinika.mdb"
    Public ID As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Forma kryesore"
        filldatagrid()
        fillmjekun()
        fillkalendarin()
        cellscolors()



        'Dim da As OleDbDataAdapter
        'Dim ds As DataSet
        'Dim tables As DataTableCollection
        'Dim source1 As New BindingSource
        'ds = New DataSet
        'tables = ds.Tables
        'Dim con As OleDbConnection = New OleDbConnection(cn)
        'da = New OleDbDataAdapter("SELECT Pacientet.Emri, Pacientet.Mbiemri, Pacientet.Emriprindit AS [Emri Prindit], Pacientet.Ditlindja, Pacientet.Adresa, Pacientet.NrTelefonit, Pacientet.Email, Pacientet.Nrpersonal, Pacientet.kartashendetsore, Pacientet.Kartafam, Pacientet.Datargjistrimit, Pacientet.RegjistruarNga, ID FROM Pacientet;", con) 'Change items to your database name
        'da.Fill(ds, "items") 'Change items to your database name
        'Dim view As New DataView(tables(0))
        'source1.DataSource = view
        'DataGridView1.DataSource = view
        'da.Dispose()
        'tables.Clear()


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        Dim i As Integer
        Me.TextBox1.Text = DataGridView1.Item(0, i).Value
        Me.TextBox2.Text = DataGridView1.Item(1, i).Value
        Me.TextBox4.Text = DataGridView1.Item(2, i).Value
        Me.DateTimePicker1.Value = DataGridView1.Item(3, i).Value
        Me.TextBox5.Text = DataGridView1.Item(5, i).Value
        Me.TextBox6.Text = DataGridView1.Item(4, i).Value
        Me.TextBox7.Text = DataGridView1.Item(7, i).Value
        Me.TextBox8.Text = DataGridView1.Item(6, i).Value
        Me.DateTimePicker2.Value = DataGridView1.Item(10, i).Value
        ID = DataGridView1.Item(12, i).Value
        Me.ComboBox1.Text = DataGridView1.Item(13, i).Value

    End Sub



    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If Me.TextBox1.Text = "" Then

            Dim con As OleDbConnection = New OleDbConnection(cn)
            Dim cmd As OleDbCommand = New OleDbCommand("Select Pacientet.Emri, Pacientet.Mbiemri, Pacientet.Emriprindit AS [Emri Prindit], Pacientet.Ditlindja, Pacientet.Adresa, Pacientet.NrTelefonit, Pacientet.Email, Pacientet.Nrpersonal, Pacientet.kartashendetsore, Pacientet.Kartafam, Pacientet.Datargjistrimit, Pacientet.RegjistruarNga , ID, gjinia FROM Pacientet WHERE Pacientet.Emri &' '& Pacientet.Mbiemri &' '& Pacientet.nrpersonal Like '%" & Me.TextBox3.Text & "%' ", con)
            ' or Where JobNo='" & SearchJob & "'  
            Try
                con.Open()
                Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
                Dim myDataSet As DataSet = New DataSet()
                myDA.Fill(myDataSet, "MyTable")
                DataGridView1.DataSource = myDataSet.Tables("MyTable").DefaultView


            Catch ex As Exception

            End Try
        Else

            fshesa()
        End If

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub


    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        'Cursors.Hand()
        PictureBox4.Image.Dispose()

        PictureBox4.Image = My.Resources.if_button_20_61451
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave
        PictureBox4.Image.Dispose()

        PictureBox4.Image = My.Resources.if_button_25_61456

    End Sub
    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover
        PictureBox2.Image.Dispose()

        PictureBox2.Image = My.Resources.if_button_21_61452
    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave
        PictureBox2.Image.Dispose()

        PictureBox2.Image = My.Resources.if_button_24_61455
    End Sub
    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs) Handles PictureBox3.MouseHover
        PictureBox3.Image.Dispose()

        PictureBox3.Image = My.Resources.if_button_22_61453
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox3.MouseLeave
        PictureBox3.Image.Dispose()

        PictureBox3.Image = My.Resources.if_button_23_61454
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        PictureBox4.Image.Dispose()
        PictureBox4.Image = My.Resources.if_button_7_61438
        'insert code
        'Me.TextBox1.Text = ""
        'Me.TextBox2.Text = ""
        'Me.TextBox4.Text = ""
        'Me.DateTimePicker1.Value = Now()
        'Me.TextBox5.Text = ""
        'Me.TextBox6.Text = ""
        'Me.TextBox7.Text = ""
        'Me.TextBox8.Text = ""
        'Me.DateTimePicker2.Value = Now()




        Using myconnection As New OleDbConnection(cn)
            myconnection.Open()
            Dim sqlQry As String = "INSERT INTO Pacientet ( Emri, Mbiemri, Emriprindit, Ditlindja, Adresa, NrTelefonit, Email, Nrpersonal, gjinia) values ( @emri, @Mbiemri, @Emriprindit, @Ditlindja, @Adresa, @NrTelefonit, @Email, @Nrpersonal, @gjinia) ;
"

            Try
                Using cmd As New OleDbCommand(sqlQry, myconnection)
                    cmd.Parameters.AddWithValue("@Emri", Me.TextBox1.Text.ToString())
                    cmd.Parameters.AddWithValue("@Mbiemri", Me.TextBox2.Text.ToString())
                    cmd.Parameters.AddWithValue("@Emriprindit", Me.TextBox4.Text.ToString())
                    cmd.Parameters.AddWithValue("@Ditlindja", Me.DateTimePicker1.Value)
                    cmd.Parameters.AddWithValue("@Adresa", Me.TextBox6.Text.ToString())
                    cmd.Parameters.AddWithValue("@NrTelefonit", Me.TextBox5.Text.ToString())
                    cmd.Parameters.AddWithValue("@Email", Me.TextBox8.Text.ToString())
                    cmd.Parameters.AddWithValue("@Nrpersonal", Me.TextBox7.Text.ToString())
                    cmd.Parameters.AddWithValue("@gjinia", Me.ComboBox1.Text)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                MsgBox("Pacienti nuk eshte regjistruar, pacienti munde te jete regjistruar me heret " & ex.ToString())
            End Try

            ' 
        End Using

        filldatagrid()

        fshesa()


    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        PictureBox2.Image.Dispose()
        PictureBox2.Image = My.Resources.if_button_8_61439
        'delete code    
        If MsgBox("A deshironi te fshini pacientin " & Me.TextBox1.Text & " " & Me.TextBox2.Text, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then


            Using myconnection As New OleDbConnection(cn)
                myconnection.Open()
                Dim sqlQry As String = "Delete * From Pacientet where ID= @ID ;
"

                Try
                    Using cmd As New OleDbCommand(sqlQry, myconnection)
                        cmd.Parameters.AddWithValue("@ID", ID)

                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    MsgBox(ex.ToString())
                End Try


            End Using

            filldatagrid()
        Else
            Exit Sub
        End If
        Me.TextBox3.Focus()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        PictureBox3.Image.Dispose()
        PictureBox3.Image = My.Resources.if_button_9_61440

        'Update code
        Using myconnection As New OleDbConnection(cn)
            myconnection.Open()
            Dim sqlQry As String = "Update Pacientet  Set Emri = @emri, Mbiemri =@Mbiemri, Emriprindit=@Emriprindit, Ditlindja =@Ditlindja, Adresa=@Adresa, NrTelefonit =@NrTelefonit, Email=@Email, Nrpersonal=@Nrpersonal, Gjinia=@gjinia Where ID =@id;
"

            Try
                Using cmd As New OleDbCommand(sqlQry, myconnection)
                    cmd.Parameters.AddWithValue("@Emri", Me.TextBox1.Text.ToString())
                    cmd.Parameters.AddWithValue("@Mbiemri", Me.TextBox2.Text.ToString())
                    cmd.Parameters.AddWithValue("@Emriprindit", Me.TextBox4.Text.ToString())
                    cmd.Parameters.AddWithValue("@Ditlindja", Me.DateTimePicker1.Value)
                    cmd.Parameters.AddWithValue("@Adresa", Me.TextBox6.Text.ToString())
                    cmd.Parameters.AddWithValue("@NrTelefonit", Me.TextBox5.Text.ToString())
                    cmd.Parameters.AddWithValue("@Email", Me.TextBox8.Text.ToString())
                    cmd.Parameters.AddWithValue("@Nrpersonal", Me.TextBox7.Text.ToString())
                    cmd.Parameters.AddWithValue("@gjinia", Me.ComboBox1.Text)
                    cmd.Parameters.AddWithValue("@id", ID)
                    cmd.ExecuteNonQuery()
                    MsgBox("Informatat per pacentin jan azhuruar.")
                End Using
            Catch ex As Exception

            End Try


        End Using

        filldatagrid()

        ID = Nothing
        fshesa()


    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox4.Text = ""
        Me.DateTimePicker1.Value = Now()
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.DateTimePicker2.Value = Now()

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        PictureBox5.Image.Dispose()
        PictureBox5.Image = My.Resources.if_button_1_61432
        Dim i As Integer
        Me.TextBox1.Text = DataGridView1.Item(0, i).Value
        Me.TextBox2.Text = DataGridView1.Item(1, i).Value
        Me.TextBox4.Text = DataGridView1.Item(2, i).Value
        Me.DateTimePicker1.Value = DataGridView1.Item(3, i).Value
        Me.TextBox5.Text = DataGridView1.Item(5, i).Value
        Me.TextBox6.Text = DataGridView1.Item(4, i).Value
        Me.TextBox7.Text = DataGridView1.Item(7, i).Value
        Me.TextBox8.Text = DataGridView1.Item(6, i).Value
        Me.DateTimePicker2.Value = DataGridView1.Item(10, i).Value
        ID = DataGridView1.Item(12, i).Value
        Me.ComboBox1.Text = DataGridView1.Item(13, i).Value
    End Sub

    Private Sub PictureBox5_MouseHover(sender As Object, e As EventArgs) Handles PictureBox5.MouseHover
        PictureBox5.Image.Dispose()
        PictureBox5.Image = My.Resources.if_button_16_61447
    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox5.MouseLeave
        PictureBox5.Image.Dispose()
        PictureBox5.Image = My.Resources.if_button_31_61462
    End Sub
    Public Function fshesa()
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox4.Text = ""
        Me.DateTimePicker1.Value = Now()
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.DateTimePicker2.Value = Now()
        Me.ComboBox1.Text = ""
        Me.TextBox3.Text = ""
        Return Nothing
        Me.TextBox3.Focus()
    End Function

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub RegjistroMjeketToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegjistroMjeketToolStripMenuItem.Click
        Form2.Show()

    End Sub
End Class

