Imports System.Data.OleDb
Imports System.Console

Public Class Form5
    Inherits System.Windows.Forms.Form
    Dim cn As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim str As String
    Dim dt As New DataTable
    Dim da As OleDbDataAdapter
    Dim dr As OleDbDataReader
    'Function1
    Public Sub show1()
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form4"
        cmd = New OleDbCommand(str, cn)
        dr = cmd.ExecuteReader
        While dr.Read
            ComboBox1.Items.Add(dr(1))
        End While
        dr.Close()
        cn.Close()
        ComboBox1.Text = ""
    End Sub
    'Form5
    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form4"
        cmd = New OleDbCommand(str, cn)
        dt = New DataTable
        da = New OleDbDataAdapter(cmd)
        da.Fill(dt)
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
        End If
        show1()
    End Sub
    'Exit
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Form3.Show()

    End Sub
    'Display
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form4 where m_name='" & ComboBox1.Text & "'"
        cmd = New OleDbCommand(str, cn)
        dr = cmd.ExecuteReader
        While dr.Read
            TextBox1.Text = dr(4)
            TextBox2.Text = dr(5)
            TextBox4.Text = dr(8)
            TextBox3.Text = dr(6)
        End While
        dr.Close()
        cn.Close()
    End Sub
    'Clear
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ComboBox1.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
    End Sub
    'Logout
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Hide()
        Form1.Show()
    End Sub
End Class