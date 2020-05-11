Public Class Form1
    Dim sql As String
    Sub panggil()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_Kamar", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_Kamar")
        DataGridView1.DataSource = DS.Tables("tb_Kamar")
        DataGridView1.Enabled = True
    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sql
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggil()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sql = "insert into tb_kamar (Kode_Kamar, Nama_Kamar, Fasilitas, Fungsi, Tarif, penanggung_jawab) values ('" & TextBox1.Text & "','" & TextBox2.Text & "' , '" & TextBox3.Text & "' , '" & TextBox4.Text & "' , '" & TextBox5.Text & "' , '" & TextBox6.Text & "')"
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")

        Call panggil()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
        TextBox5.Text = DataGridView1.Item(4, i).Value
        TextBox6.Text = DataGridView1.Item(5, i).Value
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        sql = "Update tb_Kamar SET Nama_kamar = '" & TextBox2.Text & "', Fasilitas = '" & TextBox3.Text & "' , Fungsi = '" & TextBox4.Text & "' , Tarif = '" & TextBox5.Text & "' , Penaggung_Jawab = '" & TextBox6.Text & "' , Kode_kamar = '" & TextBox1.Text & " "
        Call jalan()
        MsgBox("Data Berhasil Terubah")
        Call panggil()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("Select * From tb_Kamar Where nama_kamar like '%" & TextBox7.Text & "%'", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_Kamar")
        DataGridView1.DataSource = DS.Tables("tb_Kamar")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sql = "Delete FROM tb_Kamar where Kode_kamar = '" & TextBox1.Text & "'"
        Call jalan()
        MsgBox("Data berhasil Di hapus")
        Call panggil()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        End
    End Sub
End Class
