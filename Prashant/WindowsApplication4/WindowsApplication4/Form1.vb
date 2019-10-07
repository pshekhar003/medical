Imports System.Data.OleDb
Imports System.Console

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim cn As New OleDbConnection
    Dim cmd As OleDbCommand
    Dim str As String
    Dim dr As OleDbDataReader
    Dim un As String = "abc"
    Dim ps As String = " "
    Dim un1, ps1 As String

    'NEW_USER
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Form2.Show()
        Me.Hide()
    End Sub
    'LOGIN
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        un1 = TextBox1.Text
        ps1 = TextBox2.Text
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\project.accdb")
        cn.Open()
        str = "select * from form3 where user_name='" & TextBox1.Text & "'"
        cmd = New OleDbCommand(str, cn)
        dr = cmd.ExecuteReader
        While dr.Read
            un = dr(0)
            ps = dr(1)
        End While
        If un = "admin" And ps = ps1 Then

            Form3.Show()
            Me.Hide()
            Form3.Button7.Show()
            TextBox1.Text = ""
            TextBox2.Text = ""

        ElseIf un = un1 And ps = ps1 Then
            Form3.Show()
            Me.Hide()
            TextBox1.Text = ""
            TextBox2.Text = ""
        Else
            MsgBox("Invalid Username Or Password....")
        End If
    End Sub
    'RESET
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
    'EXIT
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
End Class
