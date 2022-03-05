Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Public Class Login_page
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Login_ID As String
        Dim Pwd As String
        Dim EmployeeName As String
        Login_ID = Login.Text
        Pwd = Password.Text
        If Login.Text = "" Then
            Message.Text = "EMPLOYEE ID IS MANDATORY FIELD!"
            Message.ForeColor = System.Drawing.Color.Red
            Message.BackColor = System.Drawing.Color.White
            Exit Sub
        End If

        If Password.Text = "" Then
            Message.Text = "PASSWORD IS MANDATORY FIELD!"
            Message.ForeColor = System.Drawing.Color.Red
            Message.BackColor = System.Drawing.Color.White
            Exit Sub
        End If

        Dim connetionString_acc As String
        Dim sqlCnn_acc As SqlConnection
        Dim sqlCmd_acc As SqlCommand
        Dim sql_acc As String
        connetionString_acc = ConfigurationManager.ConnectionStrings("Data_Entry").ConnectionString
        sql_acc = "Select COUNT(*) From ER_Account where Employee_ID = '" & Login_ID & "' and Password = '" & Pwd & "'"
        sqlCnn_acc = New SqlConnection(connetionString_acc)
        sqlCnn_acc.Open()
        sqlCmd_acc = New SqlCommand(sql_acc, sqlCnn_acc)
        Dim sqlReader_acc As SqlDataReader = sqlCmd_acc.ExecuteReader()
        While sqlReader_acc.Read()

            If sqlReader_acc.Item(0) > 0 Then
                Dim connetionString_access As String
                Dim sqlCnn_access As SqlConnection
                Dim sqlCmd_access As SqlCommand
                Dim sql_access As String
                connetionString_access = ConfigurationManager.ConnectionStrings("Data_Entry").ConnectionString
                sql_access = "Select ER_Employee_DB.Employee_Name, ER_Account.Access From ER_Account Cross Join ER_Employee_DB where ER_Account.Employee_ID = ER_Employee_DB.Employee_ID and ER_Account.Employee_ID = '" & Login_ID & "' and ER_Account.Password = '" & Pwd & "'"
                sqlCnn_access = New SqlConnection(connetionString_access)
                sqlCnn_access.Open()
                sqlCmd_access = New SqlCommand(sql_access, sqlCnn_access)
                Dim sqlReader_access As SqlDataReader = sqlCmd_access.ExecuteReader()
                While sqlReader_access.Read()
                    EmployeeName = sqlReader_access.GetValue(0)
                    Session("Employee_Name") = EmployeeName
                    Session("Login_ID") = UCase(Login_ID).ToString

                    If sqlReader_access.GetValue(1) = "Admin" Then
                        Response.Redirect("ExcelExport.aspx")
                        Exit While
                    ElseIf sqlReader_access.GetValue(1) = "User" Then
                        Response.Redirect("NotAuthorized.aspx")
                        Exit While
                    ElseIf sqlReader_access.GetValue(1) = "NON PAYROLL" Then
                        Response.Redirect("NotAuthorized.aspx")
                        Exit While
                    ElseIf sqlReader_access.GetValue(1) = "TimeAttendance" Then
                        Response.Redirect("NotAuthorized.aspx")
                        Exit While
                    End If

                End While
                sqlReader_access.Close()
                sqlCmd_access.Dispose()
                sqlCnn_access.Close()

            ElseIf sqlReader_acc.Item(0) <= 0 Then

                Message.Text = "Employee ID and Password Invalid, Please Try Again!"
                Message.ForeColor = System.Drawing.Color.Red
                Message.BackColor = System.Drawing.Color.White
                Message.Font.Bold = True
            End If

        End While
        sqlReader_acc.Close()
        sqlCmd_acc.Dispose()
        sqlCnn_acc.Close()
    End Sub
End Class