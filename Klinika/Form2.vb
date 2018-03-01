Imports System.Data
Imports System.Data.OleDb

Public Class Form2
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Using myconnection As New OleDbConnection(Form1.cn)
            myconnection.Open()
            Dim sqlQry As String = "INSERT INTO Mjeket ( Emrimjekut, Mbiemri, kredencialet, titulli) values ( @emri, @Mbiemri, @kredencialet, @titulli) ;"

            Try
                Using cmd As New OleDbCommand(sqlQry, myconnection)
                    cmd.Parameters.AddWithValue("@Emri", Me.TextBox1.Text.ToString())
                    cmd.Parameters.AddWithValue("@Mbiemri", Me.TextBox2.Text.ToString())
                    cmd.Parameters.AddWithValue("@kredencialet", Me.TextBox3.Text.ToString())
                    cmd.Parameters.AddWithValue("@titulli", Me.TextBox4.Text)

                    cmd.ExecuteNonQuery()
                    Me.TextBox1.Text = ""
                    Me.TextBox2.Text = ""
                    Me.TextBox3.Text = ""
                    Me.TextBox4.Text = ""
                End Using
            Catch ex As Exception
                MsgBox("Mjeku nuk eshte regjistruar, Mjeku munde te jete regjistruar me heret " & ex.ToString())
            End Try

            ' 
        End Using

        'filldatagrid()

        Form1.fshesa()

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

    Private Sub Form2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        Form1.Show()
        Module1.fillmjekun()

    End Sub
End Class