Imports System.Data.OleDb
Imports System.Console

Class Form2
    Inherits System.Windows.Forms.Form
    Dim cn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim str As String
    Dim un As String = "a"
    Dim ps As String = "b"
    Dim ps1 As String = "c"
   
    'CREATE
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\project.accdb")
        cn.Open()
        un = TextBox1.Text
        ps = TextBox2.Text
        ps1 = TextBox3.Text
        If un = "" Or ps = "" Or ps1 = "" Then
            MsgBox("Invalid Process....")
        ElseIf ps = ps1 Then
            Try
                str = "insert into form3(user_name,password1) values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                cmd = New OleDbCommand(str, cn)
                cmd.ExecuteNonQuery()
                MsgBox("User_Id Successfully Created....")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
            Catch ex As Exception
                MsgBox("Invalid Process....")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
            End Try
        Else
            MsgBox("Both Passwords Are Not The Same....")
        End If
        cn.Close()
    End Sub
    'RESET
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    'EXIT
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    'LOGIN
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form1.Show()
        Me.Hide()
    End Sub
End Class