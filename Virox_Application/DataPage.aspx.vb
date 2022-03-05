Imports System.Data.SqlClient
Imports System.Web.Script.Serialization

Public Class DataPage
    Inherits System.Web.UI.Page

    Private Shared PageIndex As String = 1
    Private Shared PageSize As String = 5000

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ''If Not IsPostBack Then
        BindListBoxes()
        ''End If
    End Sub

    Private Sub BindListBoxes()
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
        Dim con As New SqlConnection(strConnString)
        Dim str As String
        Dim com As SqlCommand
        Dim sqlda As SqlDataAdapter
        Dim ds As DataSet

        con.Open()


        'Department Bind
        str = "SELECT DISTINCT Department FROM VA_RULE_DEPARTMENT_VERTICAL where Employee_ID = '" & Session("Login_ID") & "' ORDER BY Department"
        com = New SqlCommand(str, con)
        sqlda = New SqlDataAdapter(com)
        ds = New DataSet
        sqlda.Fill(ds, "Department")
        lstDepartment.DataValueField = "Department"
        lstDepartment.DataTextField = "Department"
        lstDepartment.DataSource = ds
        lstDepartment.DataBind()

        'EmployeeType Bind
        str = "SELECT DISTINCT Employee_Type FROM ER_Employee_DB ORDER BY Employee_Type"

        com = New SqlCommand(str, con)
        sqlda = New SqlDataAdapter(com)
        ds = New DataSet
        sqlda.Fill(ds, "Employee_Type")
        lstEmployeeType.DataValueField = "Employee_Type"
        lstEmployeeType.DataTextField = "Employee_Type"
        lstEmployeeType.DataSource = ds
        lstEmployeeType.DataBind()

        'Status Bind
        str = "SELECT DISTINCT Status FROM ER_Employee_DB ORDER BY Status"
        com = New SqlCommand(str, con)
        sqlda = New SqlDataAdapter(com)
        ds = New DataSet
        sqlda.Fill(ds, "Status")
        lstStatus.DataValueField = "Status"
        lstStatus.DataTextField = "Status"
        lstStatus.DataSource = ds
        lstStatus.DataBind()

        'Query Bind
        str = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'ER_Employee_DB' AND COLUMN_NAME != 'Repository_ID'"
        com = New SqlCommand(str, con)
        sqlda = New SqlDataAdapter(com)
        ds = New DataSet
        sqlda.Fill(ds, "COLUMN_NAME")
        lstQuery.DataValueField = "COLUMN_NAME"
        lstQuery.DataTextField = "COLUMN_NAME"
        lstQuery.DataSource = ds
        lstQuery.DataBind()

        For Each li As ListItem In lstQuery.Items
            If li.Value = "Employee_ID" Or li.Value = "Employee_Name" Or li.Value = "Department" Then
                li.Selected = True
            End If
        Next

        con.Close()

    End Sub

    <System.Web.Services.WebMethod()>
    Public Shared Function DataListDetails(querydata As String) As DataResult
        Dim resultstring As String = ""
        Dim datadetails As New DataResult

        Dim s As JavaScriptSerializer = New JavaScriptSerializer()
        Dim searchdata As DataSearch = s.Deserialize(Of DataSearch)(querydata)

        Try
            Dim SearchTerm As String = ""

            If searchdata.selecteddepartment <> "" Then
                If SearchTerm = "" Then
                    SearchTerm += " Department IN (" + searchdata.selecteddepartment + ") "
                Else
                    SearchTerm += " AND Department IN (" + searchdata.selecteddepartment + ") "
                End If
            End If
            If searchdata.selectedemployeetype <> "" Then
                If SearchTerm = "" Then
                    SearchTerm += " Employee_Type IN (" + searchdata.selectedemployeetype + ") "
                Else
                    SearchTerm += " AND Employee_Type IN (" + searchdata.selectedemployeetype + ") "
                End If
            End If
            If searchdata.selectedstatus <> "" Then
                If SearchTerm = "" Then
                    SearchTerm += " Status IN (" + searchdata.selectedstatus + ") "
                Else
                    SearchTerm += " AND Status IN (" + searchdata.selectedstatus + ") "
                End If
            End If

            Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
            Dim con As New SqlConnection(strConnString)
            Dim query As String = "[GetEmployeeDetails_Pager]"
            Dim cmd As New SqlCommand(query)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@ColumnsFilter", searchdata.selectedcolumns)
            cmd.Parameters.AddWithValue("@SearchTerm", SearchTerm)
            cmd.Parameters.AddWithValue("@PageIndex", PageIndex)
            cmd.Parameters.AddWithValue("@PageSize", PageSize)

            Dim sqlda As SqlDataAdapter
            cmd.Connection = con
            sqlda = New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            sqlda.Fill(ds, "DataList")

            Dim DataString As String = ""
            DataString = Newtonsoft.Json.JsonConvert.SerializeObject(ds.Tables(0))

            Dim ColumnString As String = ""
            ColumnString = Newtonsoft.Json.JsonConvert.SerializeObject(ds.Tables(1))

            datadetails.data = DataString
            datadetails.columns = ColumnString

        Catch ex As Exception
            Return datadetails
        End Try
        Return datadetails
    End Function

End Class