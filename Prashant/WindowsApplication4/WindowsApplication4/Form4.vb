Imports System.Data.OleDb
Imports System.Console

Public Class Form4
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    ' Dim icount As Integer
    Dim str As String
    Dim head As New System.Data.DataTable
    ' Public no As Integer = 0
    Public headadapter As New OleDbDataAdapter
    Public headcommand As New OleDbCommandBuilder
    Dim da As New OleDbDataAdapter
    Dim dt As New DataTable

    'SEARCH
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form4 where code='" & TextBox1.Text & "'"
        cmd = New OleDbCommand(str, cn)
        dr = cmd.ExecuteReader
        While dr.Read
            TextBox2.Text = dr(1)
            TextBox3.Text = dr(2)
            TextBox4.Text = dr(3)
            TextBox5.Text = dr(4)
            TextBox6.Text = dr(5)
            TextBox7.Text = dr(6)
            TextBox8.Text = dr(7)
            TextBox9.Text = dr(8)
        End While
        dr.Close()
        cn.Close()
    End Sub
    'CLEAR
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
    End Sub
    'INSERT
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "insert into form4(code,m_name,c_name,batch_no,mfg_date,exp_date,file,cost_price,sell_price) values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "'," & TextBox4.Text & ",'" & TextBox5.Text & "','" & TextBox6.Text & "'," & TextBox7.Text & "," & TextBox8.Text & "," & TextBox9.Text & ")"
        cmd = New OleDbCommand(str, cn)
        cmd.ExecuteNonQuery()
        MsgBox("Data Inserted Successfully....")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        cn.Close()
    End Sub
    'UPDATE
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "update form4 set m_name='" & TextBox2.Text & "',c_name='" & TextBox3.Text & "',batch_no=" & TextBox4.Text & ",mfg_date='" & TextBox5.Text & "',exp_date='" & TextBox6.Text & "',file=" & TextBox7.Text & ",cost_price=" & TextBox8.Text & ",sell_price=" & TextBox9.Text & " where code='" & TextBox1.Text & "'"
        cmd = New OleDbCommand(str, cn)
        cmd.ExecuteNonQuery()
        MsgBox("Data Updated Successfully....")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
    End Sub
    'DELETE
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "delete from form4 where code='" & TextBox1.Text & "'"
        cmd = New OleDbCommand(str, cn)
        cmd.ExecuteNonQuery()
        MsgBox("Data Deleted Successfully....")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        cn.Close()
    End Sub
    'back
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Form3.Show()
        Me.Hide()

    End Sub
    'logout
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Form1.Show()
        Me.Hide()
    End Sub
End Class