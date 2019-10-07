Imports System.Data.OleDb
Imports System.Console

Public Class Form6
    Inherits System.Windows.Forms.Form
    Dim cn As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim str As String
    Dim dt As New DataTable
    Dim da As OleDbDataAdapter
    Dim dr As OleDbDataReader
    Dim a As Double = 0
    Dim b As Double = 0
    Dim c As Double = 0
    Dim d As Double = 0
    Dim total As Double = 0
 
    'Function_1
    Public Sub show11()
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
    'Function_3
    Sub disp()
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form6"
        cmd = New OleDbCommand(str, cn)
        dt = New DataTable
        da = New OleDbDataAdapter(cmd)
        da.Fill(dt)
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
        End If
    End Sub
    'Form_6
    Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox2.Text = ""
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox4.Text = ""
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form4"
        cmd = New OleDbCommand(str, cn)
        dt = New DataTable
        da = New OleDbDataAdapter(cmd)
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
        End If
        Show()
    End Sub
    'Form6
    Private Sub Form6_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        show11()
    End Sub
    'Print
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        PrintDialog1.Document = PrintDocument1
        PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
        PrintDialog1.AllowSomePages = True
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
            PrintDocument1.Print()
        End If
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "delete from form6"
        cmd = New OleDbCommand(str, cn)
        cmd.ExecuteNonQuery()
        MsgBox("Print Successfully")
        DataGridView1.DataSource = ""
        Label10.Text = 0
    End Sub
    'Combo Box
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "select * from form4 where m_name='" & ComboBox1.SelectedItem & "'"
        cmd = New OleDbCommand(str, cn)
        dr = cmd.ExecuteReader
        While dr.Read
            TextBox1.Text = dr(0)
            TextBox2.Text = dr(8)
            TextBox3.Text = dr(4)
            TextBox4.Text = dr(6)
            TextBox5.Text = dr(5)
        End While
        dr.Close()
        cn.Close()
    End Sub
    'Quantity
    Private Sub TextBox7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus
        a = TextBox2.Text
        b = TextBox4.Text
        c = a / b
        d = TextBox6.Text
        TextBox7.Text = c * d
    End Sub
    'Add
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        str = "insert into form6 values('" & ComboBox1.Text & "','" & TextBox1.Text & "'," & TextBox2.Text & ",'" & TextBox3.Text & "'," & TextBox4.Text & ",'" & TextBox5.Text & "'," & TextBox6.Text & "," & TextBox7.Text & ")"
        cmd = New OleDbCommand(str, cn)
        cmd.ExecuteNonQuery()
        MsgBox("Data Inserted Successfully....")
        total = total + TextBox7.Text
        ComboBox1.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        Label10.Text = total
        cn.Close()
        disp()
    End Sub
    'Logout
    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Hide()
        Form1.Show()
    End Sub
    'back
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Form3.Show()

    End Sub

End Class