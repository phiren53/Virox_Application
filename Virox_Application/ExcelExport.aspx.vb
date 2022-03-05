Imports System.Data
Imports System.Web.Services
Imports System.Configuration
Imports System.Data.SqlClient
Imports ClosedXML.Excel
Imports System.IO
Imports System.Web.Script.Services
Imports Newtonsoft.Json
Imports System.Reflection

<ScriptService()>
Public Class ExcelExport
    Inherits System.Web.UI.Page

    Private Shared PageIndex As String = 1
    Private Shared PageSize As String = 5000
    Private Shared WhereCondition As String = ""

    Private Shared table As Table

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try
            Dim Login_ID As String
            Dim Pwd As String
            Dim EmployeeName As String


            If HttpContext.Current.Session("Login_ID") = "" Then
                Response.Redirect("NotAuthorized.aspx")
            End If

            Dim connetionString_acc As String
            Dim sqlCnn_acc As SqlConnection
            Dim sqlCmd_acc As SqlCommand
            Dim sql_acc As String
            connetionString_acc = ConfigurationManager.ConnectionStrings("Data_Entry").ConnectionString
            sql_acc = "Select COUNT(*) From ER_Account where Employee_ID = '" & HttpContext.Current.Session("Login_ID") & "' and Access = 'Admin'"
            sqlCnn_acc = New SqlConnection(connetionString_acc)
            sqlCnn_acc.Open()
            sqlCmd_acc = New SqlCommand(sql_acc, sqlCnn_acc)
            Dim sqlReader_acc As SqlDataReader = sqlCmd_acc.ExecuteReader()
            While sqlReader_acc.Read()

                If sqlReader_acc.GetValue(0) > 0 Then
                    '         Dim connetionString_access As String
                    '         Dim sqlCnn_access As SqlConnection
                    '         Dim sqlCmd_access As SqlCommand
                    '         Dim sql_access As String
                    '         connetionString_access = ConfigurationManager.ConnectionStrings("Data_Entry").ConnectionString
                    '         sql_access = "Select ER_Employee_DB.Employee_Name, ER_Account.Access From ER_Account Cross Join ER_Employee_DB where ER_Account.Employee_ID = ER_Employee_DB.Employee_ID and ER_Account.Employee_ID = '" & HttpContext.Current.Session("Login_ID") & "'"
                    '         sqlCnn_access = New SqlConnection(connetionString_access)
                    '         sqlCnn_access.Open()
                    '         sqlCmd_access = New SqlCommand(sql_access, sqlCnn_access)
                    '         Dim sqlReader_access As SqlDataReader = sqlCmd_access.ExecuteReader()
                    '         While sqlReader_access.Read()
                    '             EmployeeName = sqlReader_access.GetValue(0)
                    '             Session("Employee_Name") = EmployeeName
                    '             Session("Login_ID") = UCase(Login_ID).ToString

                    '             If sqlReader_access.GetValue(1) = "Admin" Then
                    '                 Response.Redirect("ExcelExport.aspx")
                    '                 Exit While
                    '             ElseIf sqlReader_access.GetValue(1) = "User" Then
                    '                 Response.Redirect("NotAuthorized.aspx")
                    '                 Exit While
                    '             ElseIf sqlReader_access.GetValue(1) = "NON PAYROLL" Then
                    '                 Response.Redirect("NotAuthorized.aspx")
                    '                 Exit While
                    '             ElseIf sqlReader_access.GetValue(1) = "TimeAttendance" Then
                    '                 Response.Redirect("NotAuthorized.aspx")
                    '                 Exit While
                    '             End If

                    '         End While
                    '         sqlReader_access.Close()
                    '         sqlCmd_access.Dispose()
                    '         sqlCnn_access.Close()

                ElseIf sqlReader_acc.GetValue(0) <= 0 Then

                    'Message.Text = "Employee ID and Password Invalid, Please Try Again!"
                    'Message.ForeColor = System.Drawing.Color.Red
                    'Message.BackColor = System.Drawing.Color.White
                    'Message.Font.Bold = True
                    Response.Redirect("NotAuthorized.aspx")
                End If

            End While
            sqlReader_acc.Close()
            sqlCmd_acc.Dispose()
            sqlCnn_acc.Close()

            Dim connetionString_ORP_CAL As String
            Dim sqlCnn_ORP_CAL As SqlConnection
            Dim sqlCmd_ORP_CAL As SqlCommand
            Dim sql_ORP_CAL As String
            connetionString_ORP_CAL = ConfigurationManager.ConnectionStrings("Data_Entry").ConnectionString
            sql_ORP_CAL = "Select COUNT(*) from PY_PAYROLL_ACCESS_RIGHT  WITH (NOLOCK)  where EMPLOYEE_ID = '" & HttpContext.Current.Session("Login_ID") & "'"
            sqlCnn_ORP_CAL = New SqlConnection(connetionString_ORP_CAL)
            sqlCnn_ORP_CAL.Open()
            sqlCmd_ORP_CAL = New SqlCommand(sql_ORP_CAL, sqlCnn_ORP_CAL)
            Dim sqlReader_ORP_CAL As SqlDataReader = sqlCmd_ORP_CAL.ExecuteReader()

            While sqlReader_ORP_CAL.Read()
                If sqlReader_ORP_CAL.GetValue(0) < 1 Then
                    Message.Text = "You Are Not Authorize Access Employee Profile - " & HttpContext.Current.Session("Login_ID")
                    Message.ForeColor = System.Drawing.Color.Red
                    Message.BackColor = System.Drawing.Color.White
                    Exit Sub
                ElseIf sqlReader_ORP_CAL.GetValue(0) > 0 Then

                    If Not IsPostBack Then
                        BindListBoxes()
                        'BindGrid()

                        table = New Table()
                    End If

                    Message.Text = "Please query and click Search button!"
                    Message.ForeColor = System.Drawing.Color.DarkBlue
                    Message.BackColor = System.Drawing.Color.White

                End If
            End While
            sqlCnn_ORP_CAL.Close()
            sqlCmd_ORP_CAL.Dispose()
            sqlReader_ORP_CAL.Close()
        Catch ex As Exception

        End Try




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
        '        str = "SELECT COLUMN_NAME, PrefixCOLUMN_NAME FROM(
        'SELECT DISTINCT COLUMN_NAME, CONCAT(TABLE_NAME,'.', COLUMN_NAME) PrefixCOLUMN_NAME, ROW_NUMBER() OVER(PARTITION BY COLUMN_NAME ORDER BY COLUMN_NAME DESC) OrderID FROM INFORMATION_SCHEMA.COLUMNS 
        'WHERE TABLE_NAME IN ( N'ER_Employee_DB', N'ER_Payroll') AND COLUMN_NAME != 'Repository_ID') ColumnList
        'WHERE OrderID = '1'
        'ORDER BY COLUMN_NAME"

        str = "SELECT COLUMN_NAME, PrefixCOLUMN_NAME, DisplayName,ColumnSequence FROM(
                SELECT DISTINCT COLUMN_NAME, CONCAT(TABLE_NAME,'.', COLUMN_NAME) PrefixCOLUMN_NAME, DisplayName,ColumnSequence, ROW_NUMBER() OVER(PARTITION BY COLUMN_NAME ORDER BY COLUMN_NAME DESC) OrderID FROM INFORMATION_SCHEMA.COLUMNS 
                LEFT JOIN ColumnMapping CM ON COLUMN_NAME = CM.TableColumnName
                WHERE TABLE_NAME IN ( N'ER_Employee_DB', N'ER_Payroll') AND COLUMN_NAME != 'Repository_ID' AND COLUMN_NAME != 'Passport_For_Sticker'
                AND COLUMN_NAME != 'EMP_QUERY'
                AND COLUMN_NAME != 'Perfect_Att'
                AND COLUMN_NAME != 'TABUNG_HAJI'
                AND COLUMN_NAME != 'Amanah_Saham'
                AND COLUMN_NAME != 'ZAKAT'
                AND COLUMN_NAME != 'Payroll_ID'
                AND COLUMN_NAME != 'Deduction_for_Spouse'
                AND COLUMN_NAME != 'Deduction_for_Individual'
                AND COLUMN_NAME != 'Employee_EPF'
                AND COLUMN_NAME != 'Employer_EPF'
                AND COLUMN_NAME != 'Wife_Code'
                AND COLUMN_NAME != 'Record_Type'
                AND COLUMN_NAME != 'Employer_No'
                AND COLUMN_NAME != 'Total_Pre_Salary'
                AND COLUMN_NAME != 'Total_Pre_EPF'
                AND COLUMN_NAME != 'Repository_ID'
                AND COLUMN_NAME != 'Previous_AR'
                AND COLUMN_NAME != 'Previous_AR_EPF'
                AND COLUMN_NAME != 'Previous_PCB'
                AND COLUMN_NAME != 'OTHOUR'
                AND COLUMN_NAME != 'START_DATE'
                AND COLUMN_NAME != 'END_DATE'
                AND COLUMN_NAME != 'Month_Diff'
                AND COLUMN_NAME != 'Amanah_Saham'
                AND COLUMN_NAME != 'backupNo2'
                AND COLUMN_NAME != 'Repository_ID'
                AND COLUMN_NAME != 'City'
                AND COLUMN_NAME != 'State'
                AND COLUMN_NAME != 'Zip_Code'
                AND COLUMN_NAME != 'Photo'
                AND COLUMN_NAME != 'Spouse_Name'
                AND COLUMN_NAME != 'Spouse_Relationship'
                AND COLUMN_NAME != 'Sponse_Career'
                AND COLUMN_NAME != 'Secondary_Phone_Number'
                AND COLUMN_NAME != 'Address_No1'
                AND COLUMN_NAME != 'Address_No2'
                AND COLUMN_NAME != 'Resident_Address_No1'
                AND COLUMN_NAME != 'Resident_Address_No2'

                AND COLUMN_NAME != 'Above_18_Full_Time_Certificate'
                AND COLUMN_NAME != 'Above_18_Full_Time_Degree'
                AND COLUMN_NAME != 'Account_Classification'
                
                AND COLUMN_NAME != 'Advanced_Salary'
                AND COLUMN_NAME != 'Advanced_Salary_Amount'
                AND COLUMN_NAME != 'Bank Account'
                AND COLUMN_NAME != 'CHARGEAREA'
                AND COLUMN_NAME != 'City'
                AND COLUMN_NAME != 'Country'
                AND COLUMN_NAME != 'Country_Code'
                AND COLUMN_NAME != 'Disabled_Child'
                AND COLUMN_NAME != 'Disabled_Child_Study'
                AND COLUMN_NAME != 'DIVISION'
                AND COLUMN_NAME != 'Education_Level'
                AND COLUMN_NAME != 'Education_Remark'
                AND COLUMN_NAME != 'EMPLOYEE_STATUS_V3'
                AND COLUMN_NAME != 'EMPLOYEE_TYPE_2'
                AND COLUMN_NAME != 'Employee Category'
                AND COLUMN_NAME != 'Employer_No'
                AND COLUMN_NAME != 'EPF_19Percent'
                AND COLUMN_NAME != 'EPF_Category'
                AND COLUMN_NAME != 'Leave Category'
                AND COLUMN_NAME != 'Locker_Number'
                AND COLUMN_NAME != 'Hostel_Address'
                AND COLUMN_NAME != 'Hostel_Code'
                AND COLUMN_NAME != 'Insurance_No'
                AND COLUMN_NAME != 'New_Category'
                AND COLUMN_NAME != 'Resident_Address'
                
                AND COLUMN_NAME != 'Resident_Phone'
                AND COLUMN_NAME != 'Passport_For_Sticker'
                AND COLUMN_NAME != 'PAYMENT_METHOD'
                AND COLUMN_NAME != 'PAYMENT_STATUS'
                AND COLUMN_NAME != 'Permit_Number'
                AND COLUMN_NAME != 'Photo'
                AND COLUMN_NAME != 'Receiving_Member_Bank'
                AND COLUMN_NAME != 'Record_Type'
                AND COLUMN_NAME != 'Rentas_Status'
                AND COLUMN_NAME != 'Rentas_Transaction_Code'
                AND COLUMN_NAME != 'Retire Age'
                AND COLUMN_NAME != 'Secondary_Phone_Number'
                AND COLUMN_NAME != 'Shoes_Issue_Date'
                AND COLUMN_NAME != 'Shoes_Size'
                AND COLUMN_NAME != 'Smock_Issue_Date'
                AND COLUMN_NAME != 'Smock_Size'
                AND COLUMN_NAME != 'SOCSO_No'
                AND COLUMN_NAME != 'Categoryid'
                AND COLUMN_NAME != 'backupNo1'
                AND COLUMN_NAME != 'LM_CAT'
                AND COLUMN_NAME != 'State'
                AND COLUMN_NAME != 'Under_18_Age'
                AND COLUMN_NAME != 'University_Name'
                AND COLUMN_NAME != 'Wife_Code'
                AND COLUMN_NAME != 'WORK_PERMIT'
                AND COLUMN_NAME != 'Zip_Code'
                AND COLUMN_NAME != 'Employee_Group_1'
                AND COLUMN_NAME != 'ID_INDICATOR'
                AND COLUMN_NAME != 'OT_LEVEL'
                AND COLUMN_NAME != 'OT_RESTRICTION'
                AND COLUMN_NAME != 'PAYOUT_TYPE'
                AND COLUMN_NAME != 'Sec_Half'
                AND COLUMN_NAME != 'Sec_Half_Sal'
                AND COLUMN_NAME != 'first_half'
                AND COLUMN_NAME != 'first_half_Sal'
AND COLUMN_NAME != 'IG'
AND COLUMN_NAME != 'PREFIX'
AND COLUMN_NAME != 'Total_for_Children'
AND COLUMN_NAME != 'Mode'
AND COLUMN_NAME != 'WIFE_WORKING'
AND COLUMN_NAME != 'TRANSPORT'
AND COLUMN_NAME != 'SPIKPA'
AND COLUMN_NAME != 'SPPA'
AND COLUMN_NAME != 'Agent'
AND COLUMN_NAME != 'EMPLOYEE_L_TYPE'
AND COLUMN_NAME != 'Employee_Type_Finance'
AND COLUMN_NAME != 'Bonus_Amount'
AND COLUMN_NAME != 'Token'
AND COLUMN_NAME != 'Insurance_Agent'
AND COLUMN_NAME != 'Overtime_Rate'
AND COLUMN_NAME != 'RETIRE_AGE'

                ) ColumnList
                WHERE OrderID = '1' ORDER BY COLUMN_NAME"

        com = New SqlCommand(str, con)
        sqlda = New SqlDataAdapter(com)
        ds = New DataSet
        sqlda.Fill(ds, "COLUMN_NAME")
        lstQuery.DataValueField = "PrefixCOLUMN_NAME"
        lstQuery.DataTextField = "DisplayName"
        lstQuery.DataSource = ds
        lstQuery.DataBind()

        Session("Columns") = ds.Tables(0)

        If LoadOptions.Rows.Count > 0 Then
            Dim Data As String = LoadOptions.Rows(0)(1).ToString
            hdnSelectQuery.Value = Data.Split("|")(1).ToString()

            Dim Cols = hdnSelectQuery.Value.Split(",")
            For Each obj As String In Cols
                If obj IsNot Nothing Then
                    For Each li As ListItem In lstQuery.Items
                        If li.Value = obj Then
                            li.Selected = True
                        End If
                    Next
                End If
            Next

            Dim selecteddepartment As String = Data.Split("|")(0).ToString()
            Dim Departments = selecteddepartment.Split(",")
            For Each obj1 As String In Departments
                If obj1 IsNot Nothing Then
                    For Each li1 As ListItem In lstDepartment.Items
                        If li1.Value = obj1.Replace("'", "") Then
                            li1.Selected = True
                        End If
                    Next
                End If
            Next

            BindGrid()
        Else
            For Each li As ListItem In lstQuery.Items
                If li.Value = "ER_Employee_DB.Employee_ID" Or li.Value = "ER_Employee_DB.Employee_Name" Or li.Value = "ER_Employee_DB.Department" Then
                    li.Selected = True
                End If
            Next
            hdnSelectQuery.Value = ",ER_Employee_DB.Employee_Name,ER_Employee_DB.Employee_ID,ER_Employee_DB.Department"
        End If

        con.Close()

    End Sub
    Private Function LoadOptions() As DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
        Dim con As New SqlConnection(strConnString)
        Dim str As String = "SELECT TOP 1 * FROM ColumnMapping"
        Dim com As SqlCommand
        Dim sqlda As SqlDataAdapter
        com = New SqlCommand(str, con)
        sqlda = New SqlDataAdapter(com)
        Dim DT As DataTable = New DataTable
        sqlda.Fill(DT)

        Return DT
    End Function
    Private Sub BindGrid()
        Try
            ViewState.Remove("SortDirection")
            ViewState.Remove("SortExpression")

            Dim dt As DataTable = DataFetch().Tables(0)
            dt = SetColumnNameAndOrder(dt, False)
            gvEmployee.DataSource = dt

            gvEmployee.DataBind()

            ViewState("tables") = dt

            hdnSelectQuery.Value = Session("_Qry")
            hdnadded_options.Value = Session("added_options")
            hdncurrent_values.Value = Session("current_values")
            hdnremoved_options.Value = Session("removed_options")
            hdnselected_options.Value = Session("selected_options")
        Catch ex As Exception

        Finally

        End Try
    End Sub

    Private Function SetColumnNameAndOrder(dataTable As DataTable, changeName As Boolean) As DataTable
        Try
#Region "To Set column name As per mapping table."
            Dim ColumnTable As DataTable = CType(Session("Columns"), DataTable)


            Dim ColumnIndexTable As DataTable = ColumnTable.Clone()

            For i As Integer = 0 To dataTable.Columns().Count() - 1

                Dim dr As DataRow = ColumnTable.Select("COLUMN_NAME = '" + dataTable.Columns(i).ColumnName + "'").FirstOrDefault()
                If Not dr Is Nothing AndAlso dr.Item("DisplayName").ToString() IsNot "" Then
                    If changeName Then
                        dataTable.Columns(i).ColumnName = dr.Item("DisplayName").ToString()
                    End If
                    ColumnIndexTable.Rows.Add(dr.ItemArray)
                End If

            Next

            Dim ColumnDataView As DataView = ColumnIndexTable.DefaultView()
            ColumnDataView.Sort = "ColumnSequence ASC"
            ColumnIndexTable = ColumnDataView.ToTable()

            For Each row As DataRow In ColumnIndexTable.Rows
                If changeName Then
                    dataTable.Columns(row.Item("DisplayName").ToString()).SetOrdinal(ColumnIndexTable.Rows.IndexOf(row) + 2)
                Else
                    dataTable.Columns(row.Item("COLUMN_NAME").ToString()).SetOrdinal(ColumnIndexTable.Rows.IndexOf(row) + 2)
                End If

            Next

            Return dataTable

#End Region
        Catch ex As Exception

        End Try
    End Function

    Private Function DataFetch() As DataSet
        Try
            Dim selectedcolumns As String = ""
            Dim selectedcolumns1 As String = ""
            For Each item As ListItem In lstQuery.Items
                If item.Selected Then
                    selectedcolumns1 += "," + item.Value
                End If
            Next

            selectedcolumns = hdnSelectQuery.Value.ToString()
            Session("_Qry") = selectedcolumns
            Session("_Qry") = hdnSelectQuery.Value
            Session("added_options") = hdnadded_options.Value
            Session("current_values") = hdncurrent_values.Value
            Session("removed_options") = hdnremoved_options.Value
            Session("selected_options") = hdnselected_options.Value

            Dim selecteddepartment As String = ""
            For Each item As ListItem In lstDepartment.Items
                If item.Selected Then
                    If selecteddepartment = "" Then
                        selecteddepartment += "'" + item.Value + "'"
                    Else
                        selecteddepartment += ",'" + item.Value + "'"
                    End If
                End If
            Next

            Dim selectedemployeetype As String = ""
            For Each item As ListItem In lstEmployeeType.Items
                If item.Selected Then
                    If selectedemployeetype = "" Then
                        selectedemployeetype += "'" + item.Value + "'"
                    Else
                        selectedemployeetype += ",'" + item.Value + "'"
                    End If
                End If
            Next

            Dim selectedstatus As String = ""
            For Each item As ListItem In lstStatus.Items
                If item.Selected Then
                    If selectedstatus = "" Then
                        selectedstatus += "'" + item.Value + "'"
                    Else
                        selectedstatus += ",'" + item.Value + "'"
                    End If
                End If
            Next

            Dim SearchTerm As String = ""

            If selecteddepartment <> "" Then
                If SearchTerm = "" Then
                    SearchTerm += " Department IN (" + selecteddepartment + ") "
                Else
                    SearchTerm += " AND Department IN (" + selecteddepartment + ") "
                End If
            End If
            If selectedemployeetype <> "" Then
                If SearchTerm = "" Then
                    SearchTerm += " Employee_Type IN (" + selectedemployeetype + ") "
                Else
                    SearchTerm += " AND Employee_Type IN (" + selectedemployeetype + ") "
                End If
            End If
            If selectedstatus <> "" Then
                If SearchTerm = "" Then
                    SearchTerm += " Status IN (" + selectedstatus + ") "
                Else
                    SearchTerm += " AND Status IN (" + selectedstatus + ") "
                End If
            End If

            Dim ModifiedQueryCondition As String = ""
            Dim ConditionList() As String
            If hdnFiletData.Value.Contains("AND") Then
                ConditionList = hdnFiletData.Value.Split(New String() {"AND"}, StringSplitOptions.None)
            Else
                ModifiedQueryCondition += QueryModify(hdnFiletData.Value)
            End If

            'If ConditionList IsNot Nothing Then
            '    For i As Integer = 0 To ConditionList.Length - 1
            '        If ModifiedQueryCondition = "" Then
            '            ModifiedQueryCondition += QueryModify(ConditionList(i))
            '        Else
            '            ModifiedQueryCondition += " AND " + QueryModify(ConditionList(i))
            '        End If
            '    Next
            '    SearchTerm += ModifiedQueryCondition
            'Else
            '    SearchTerm += ModifiedQueryCondition
            'End If

            If ConditionList IsNot Nothing Then
                For i As Integer = 0 To ConditionList.Length - 1
                    If ModifiedQueryCondition = "" Then
                        ModifiedQueryCondition += QueryModify(ConditionList(i))
                    Else
                        ModifiedQueryCondition += " AND " + QueryModify(ConditionList(i))
                    End If
                Next
                SearchTerm += ModifiedQueryCondition
            Else
                If SearchTerm = "" Then
                    SearchTerm += ModifiedQueryCondition
                Else
                    If Not ModifiedQueryCondition = "" Then
                        SearchTerm += " AND " + ModifiedQueryCondition
                    End If
                End If
            End If

            Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
            Dim con As New SqlConnection(strConnString)
            Dim query As String = "[GetEmployeeDetails_Pager]"
            Dim cmd As New SqlCommand(query)
            If SearchTerm = "" Then
                SearchTerm = "Department IN ('LOGISTICS')"
            End If

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@ColumnsFilter", selectedcolumns)
            cmd.Parameters.AddWithValue("@SearchTerm", SearchTerm)
            cmd.Parameters.AddWithValue("@PageIndex", PageIndex)
            cmd.Parameters.AddWithValue("@PageSize", PageSize)

            Dim sqlda As SqlDataAdapter
            cmd.Connection = con
            If (con.State <> ConnectionState.Open) Then
                con.Open()
            End If

            sqlda = New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            sqlda.Fill(ds, "DataList")

#Region "Static Order related code commented..."

            'If (ds.Tables(0).Columns.Contains("Employee_ID")) Then
            '    ds.Tables(0).Columns("Employee_ID").SetOrdinal(2)
            'End If
            'If (ds.Tables(0).Columns.Contains("Employee_Name")) Then
            '    ds.Tables(0).Columns("Employee_Name").SetOrdinal(3)
            'End If
#End Region

            'If (ds.Tables(0).Columns.Contains("Department")) Then
            '    ds.Tables(0).Columns("Department").SetOrdinal(4)
            'End If

            For i As Integer = 0 To ds.Tables(0).Rows.Count() - 1
                For j As Integer = 0 To ds.Tables(0).Columns.Count() - 1
                    If ds.Tables(0).Columns(j).DataType.Name = "DateTime" Then
                        If Convert.ToDateTime(ds.Tables(0).Rows(i)(j)).ToString("dd/MM/yyyy") = "01/01/0001" OrElse Convert.ToDateTime(ds.Tables(0).Rows(i)(j)).ToString("dd/MM/yyyy") = "01/01/1900" Then
                            ds.Tables(0).Rows(i)(j) = DBNull.Value
                        End If
                    End If

                Next
            Next
            Return ds

        Catch ex As Exception

        Finally

        End Try
    End Function

    Private Function QueryModify(ByVal QueryData As String) As String
        Dim Condition As String = ""
        Dim FinalColumn As String = ""
        Dim FinalData As String = ""
        Dim FormatedResult As String = ""
        Dim FinalResult As String = ""

        If QueryData.Contains("=") Then
            Condition = "="
            FinalColumn = QueryData.Split("=")(0)
            FinalData = "'" + QueryData.Split("=")(1) + "'"
            FormatedResult = FinalColumn + Condition + FinalData
        ElseIf QueryData.Contains("<") Then
            Condition = "<"
            FinalColumn = QueryData.Split("<")(0)
            FinalData = "'" + QueryData.Split("<")(1) + "'"
            FormatedResult = FinalColumn + Condition + FinalData
        ElseIf QueryData.Contains(">") Then
            Condition = ">"
            FinalColumn = QueryData.Split(">")(0)
            FinalData = "'" + QueryData.Split(">")(1) + "'"
            FormatedResult = FinalColumn + Condition + FinalData
        ElseIf QueryData.Contains("<=") Then
            Condition = "<="
            FinalColumn = QueryData.Split("<=")(0)
            FinalData = "'" + QueryData.Split("<=")(1) + "'"
            FormatedResult = FinalColumn + Condition + FinalData
        ElseIf QueryData.Contains(">=") Then
            Condition = ">="
            FinalColumn = QueryData.Split(">=")(0)
            FinalData = "'" + QueryData.Split(">=")(1) + "'"
            FormatedResult = FinalColumn + Condition + FinalData
        Else
            FormatedResult = QueryData
        End If

        If FinalResult = "" Then
            FinalResult += FormatedResult
        Else
            FinalResult += " AND " + FormatedResult
        End If

        Return FinalResult
    End Function

    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        Dim selecteddepartment As String = ""
        For Each item As ListItem In lstDepartment.Items
            If item.Selected Then
                If selecteddepartment = "" Then
                    selecteddepartment += "'" + item.Value + "'"
                Else
                    selecteddepartment += ",'" + item.Value + "'"
                End If
            End If
        Next

        If selecteddepartment = "" Then
            Message.Text = "Please select Department before proceed with click on Search button"
        Else
            BindGrid()
        End If

    End Sub



    Protected Sub btnExcel_Click(sender As Object, e As EventArgs)
        Dim ds As New DataSet
        ds = DataFetch()

#Region "To Set column name As per mapping table."

        Dim excelData As DataTable = SetColumnNameAndOrder(ds.Tables(0), True)

        'Dim ColumnTable As DataTable = CType(Session("Columns"), DataTable)


        'Dim ColumnIndexTable As DataTable = ColumnTable.Clone()

        'For i As Integer = 0 To ds.Tables(0).Columns().Count() - 1

        '    Dim dr As DataRow = ColumnTable.Select("COLUMN_NAME = '" + ds.Tables(0).Columns(i).ColumnName + "'").FirstOrDefault()
        '    If Not dr Is Nothing AndAlso dr.Item("DisplayName").ToString() IsNot "" Then
        '        ds.Tables(0).Columns(i).ColumnName = dr.Item("DisplayName").ToString()
        '        ColumnIndexTable.Rows.Add(dr.ItemArray)
        '    End If

        'Next

        'Dim ColumnDataView As DataView = ColumnIndexTable.DefaultView()
        'ColumnDataView.Sort = "ColumnSequence ASC"
        'ColumnIndexTable = ColumnDataView.ToTable()

        'For Each row As DataRow In ColumnIndexTable.Rows
        '    ds.Tables(0).Columns(row.Item("DisplayName").ToString()).SetOrdinal(ColumnIndexTable.Rows.IndexOf(row) + 2)
        'Next
#End Region


        Dim wb As New XLWorkbook()
        Dim wsDetailedData = wb.Worksheets.Add(excelData, "EmployeeDetails")
        wsDetailedData.Columns().AdjustToContents()

        Response.Clear()
        Response.Buffer = True
        Response.Charset = ""
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        Response.AddHeader("content-disposition", "attachment;filename=ExcelFile.xlsx")
        Using MyMemoryStream As New MemoryStream()
            wb.SaveAs(MyMemoryStream)
            MyMemoryStream.WriteTo(Response.OutputStream)
            Response.Flush()
            Response.End()
        End Using
    End Sub

    Protected Sub lnkPrevious_Click(sender As Object, e As EventArgs)
        PageIndex = Integer.Parse(PageIndex - 1)
        BindGrid()
        If PageIndex = 1 Then
            'lnkPrevious.Enabled = False
            'lnkNext.Enabled = True
        End If
    End Sub

    Protected Sub lnkNext_Click(sender As Object, e As EventArgs)
        PageIndex = Integer.Parse(PageIndex + 1)
        BindGrid()
        If PageIndex = 1 Then
            'lnkPrevious.Enabled = True
            'lnkNext.Enabled = False
        End If
    End Sub

    Protected Sub gvEmployee_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim row As New GridViewRow(0, -1, DataControlRowType.DataRow, DataControlRowState.Normal)
        'For i As Integer = 2 To gvEmployee.HeaderRow.Cells.Count - 1
        '    Dim cell As New TableCell()
        '    If (i > 2) Then
        '        Dim txtsearch As New TextBox()
        '        Dim columnname As String = CType(gvEmployee.HeaderRow.Cells(i).Controls(0), LinkButton).Text.Replace("&#9650;", "").Replace("&#9660;", "").Trim()
        '        Dim searchcolumnname As String = CType(gvEmployee.HeaderRow.Cells(i).Controls(0), LinkButton).Attributes("searchcolumnname").Replace("&#9650;", "").Replace("&#9660;", "").Trim()
        '        txtsearch.Attributes("Placeholder") = columnname
        '        Dim ColumnPrefix As String = ""
        '        For Each item As ListItem In lstQuery.Items
        '            If item.Selected Then
        '                If (item.Text = searchcolumnname) Then
        '                    ColumnPrefix = item.Value
        '                    Exit For
        '                End If
        '            End If
        '        Next
        '        txtsearch.ID = "txt-" + gvEmployee.HeaderRow.Cells(i).Text
        '        txtsearch.CssClass = "textbox"
        '        txtsearch.Attributes.Add("runat", "server")
        '        txtsearch.Attributes.Add("backcolumn", ColumnPrefix)
        '        cell.Controls.Add(txtsearch)
        '    End If

        '    row.Controls.Add(cell)
        'Next

        'Dim tb As Table = gvEmployee.Controls(0)
        'tb.Rows.AddAt(1, row)
        'gvEmployee.HeaderRow.TableSection = TableRowSection.TableHeader

        Try
            Dim row As New GridViewRow(0, -1, DataControlRowType.DataRow, DataControlRowState.Normal)
            For i As Integer = 2 To gvEmployee.HeaderRow.Cells.Count - 1
                Dim cell As New TableCell()
                If (i > 2) Then
                    Dim txtsearch As New TextBox()
                    Dim columnname As String = CType(gvEmployee.HeaderRow.Cells(i).Controls(0), LinkButton).Text.Replace("&#9650;", "").Replace("&#9660;", "").Trim()
                    Dim searchcolumnname As String = CType(gvEmployee.HeaderRow.Cells(i).Controls(0), LinkButton).Attributes("searchcolumnname").Replace("&#9650;", "").Replace("&#9660;", "").Trim()
                    txtsearch.Attributes("Placeholder") = columnname
                    Dim ColumnPrefix As String = ""
                    For Each item As ListItem In lstQuery.Items
                        If item.Selected Then
                            If (item.Text = searchcolumnname) Then
                                ColumnPrefix = item.Value
                                Exit For
                            End If
                        End If
                    Next
                    txtsearch.ID = "txt-" + gvEmployee.HeaderRow.Cells(i).Text
                    txtsearch.CssClass = "textbox"
                    txtsearch.Attributes.Add("runat", "server")
                    txtsearch.Attributes.Add("backcolumn", ColumnPrefix)
                    cell.Controls.Add(txtsearch)
                End If

                row.Controls.Add(cell)
            Next

            Dim tb As Table = gvEmployee.Controls(0)
            tb.Rows.AddAt(1, row)
            gvEmployee.HeaderRow.TableSection = TableRowSection.TableHeader

        Catch ex As Exception

        Finally

        End Try
    End Sub

    Protected Sub btnEdit_Click(sender As Object, e As EventArgs)

    End Sub

    Protected Sub gvEmployee_RowDataBound(sender As Object, e As GridViewRowEventArgs)

        If gvEmployee.HeaderRow IsNot Nothing Then
            gvEmployee.HeaderRow.Cells(1).Visible = False
            gvEmployee.HeaderRow.Cells(2).Visible = False
        End If

        For Each row As GridViewRow In gvEmployee.Rows
            row.Cells(1).Visible = False
            row.Cells(2).Visible = False
        Next

        'If e.Row.RowType = DataControlRowType.Header Then
        '    For i As Integer = 3 To e.Row.Cells.Count - 1
        '        If e.Row.Cells(i) IsNot Nothing Then
        '            Dim chk As CheckBox = New CheckBox()
        '            chk.Text = e.Row.Cells(i).Text
        '            chk.ID = "chk" + i.ToString()
        '            chk.CssClass = "headerCheckbox"
        '            chk.Attributes.Add("onclick", "onchangetick('" + chk.ID + "','headerCheckbox')")
        '            e.Row.Cells(i).Controls.Add(chk)
        '        End If
        '    Next
        'End If
    End Sub

    '<System.Web.Services.WebMethod()>
    'Public Shared Function BindEditDropdown(ByVal datatype As String) As List(Of ListItem)
    '    Dim datalist As New List(Of ListItem)()
    '    Try
    '        Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
    '        Dim con As New SqlConnection(strConnString)
    '        Dim str As String
    '        Dim com As SqlCommand
    '        Dim sqlda As SqlDataAdapter
    '        Dim ds As DataSet

    '        con.Open()
    '        If datatype = "Department" Then
    '            str = "SELECT DISTINCT Department FROM ER_Employee_DB ORDER BY Department"
    '        ElseIf datatype = "Employee_Type" Then
    '            str = "SELECT DISTINCT Employee_Type FROM ER_Employee_DB ORDER BY Employee_Type"
    '        ElseIf datatype = "Status" Then
    '            str = "SELECT DISTINCT Status FROM ER_Employee_DB ORDER BY Status"
    '        End If
    '        com = New SqlCommand(str, con)
    '        sqlda = New SqlDataAdapter(com)
    '        ds = New DataSet
    '        sqlda.Fill(ds, "Status")

    '        con.Close()

    '        For i As Integer = 0 To ds.Tables(0).Rows().Count() - 1
    '            If datatype = "Department" Then
    '                datalist.Add(New ListItem(ds.Tables(0).Rows(i)("Department").ToString(), ds.Tables(0).Rows(i)("Department").ToString()))
    '            ElseIf datatype = "Employee_Type" Then
    '                datalist.Add(New ListItem(ds.Tables(0).Rows(i)("Employee_Type").ToString(), ds.Tables(0).Rows(i)("Employee_Type").ToString()))
    '            ElseIf datatype = "Status" Then
    '                datalist.Add(New ListItem(ds.Tables(0).Rows(i)("Status").ToString(), ds.Tables(0).Rows(i)("Status").ToString()))
    '            End If
    '        Next

    '        con.Close()
    '    Catch ex As Exception
    '        Return datalist
    '    End Try
    '    Return datalist
    'End Function

    '<System.Web.Services.WebMethod()>
    'Public Shared Function UpdateData(ByVal datalist As List(Of ER_Employee_DB)) As String
    '    Try
    '        Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
    '        Dim con As New SqlConnection(strConnString)
    '        Dim com As New SqlCommand
    '        con.Open()
    '        com.Connection = con


    '        For i As Integer = 0 To datalist.Count() - 1
    '            Dim Repository_ID As String = datalist(i).Repository_ID

    '            Dim info() As PropertyInfo = datalist(i).GetType().GetProperties()

    '            Dim columns As String = ""
    '            Dim columnname As String = ""
    '            For colIndex As Integer = 0 To info.Count() - 1
    '                If info(colIndex).GetValue(datalist(i), Nothing) IsNot Nothing AndAlso info(colIndex).Name <> "Repository_ID" Then
    '                    If columns = "" Then
    '                        columns += info(colIndex).Name + " = '" + info(colIndex).GetValue(datalist(i), Nothing) + "'"
    '                        columnname = info(colIndex).Name
    '                    Else
    '                        columns += "," + info(colIndex).Name + " = '" + info(colIndex).GetValue(datalist(i), Nothing) + "'"
    '                        columnname = info(colIndex).Name
    '                    End If
    '                End If
    '            Next

    '            Dim sqlda As SqlDataAdapter = New SqlDataAdapter
    '            Dim cmd As SqlCommand = New SqlCommand
    '            cmd.Connection = con
    '            cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'ER_Employee_DB' AND COLUMN_NAME = '" + columnname + "'"
    '            sqlda = New SqlDataAdapter(cmd)
    '            Dim ds As New DataSet
    '            sqlda.Fill(ds)
    '            If ds.Tables(0).Rows.Count > 0 Then
    '                com.CommandText = "UPDATE ER_Employee_DB SET " + columns + " Where [Repository_ID] = '" & Repository_ID & "'"
    '            Else
    '                cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'ER_Payroll' AND COLUMN_NAME = '" + columnname + "'"
    '                sqlda = New SqlDataAdapter(cmd)
    '                ds = New DataSet
    '                sqlda.Fill(ds)
    '                If ds.Tables(0).Rows.Count > 0 Then
    '                    com.CommandText = "UPDATE ER_Payroll SET " + columns + " Where [Repository_ID] = '" & Repository_ID & "'"
    '                End If
    '            End If


    '            com.ExecuteNonQuery()
    '        Next

    '        con.Close()
    '    Catch ex As Exception
    '        Return "Data updation failed"
    '    End Try
    '    Return "Data updated successfully"
    'End Function

    Protected Sub gvEmployee_Sorting(sender As Object, e As GridViewSortEventArgs)
        Dim SortDir As String = String.Empty
        Dim dt As DataTable = New DataTable()
        dt = TryCast(ViewState("tables"), DataTable)

        If True Then

            If dir() = SortDirection.Ascending Then
                dir = SortDirection.Descending
                SortDir = "Desc"
            Else
                dir = SortDirection.Ascending
                SortDir = "Asc"
            End If

            Dim sortedView As DataView = New DataView(dt)
            sortedView.Sort = e.SortExpression & " " & SortDir
            ViewState("SortExpression") = e.SortExpression
            gvEmployee.DataSource = sortedView
            gvEmployee.DataBind()
        End If
    End Sub

    Public Property dir As SortDirection
        Get

            If ViewState("SortDirection") Is Nothing Then
                ViewState("SortDirection") = SortDirection.Ascending
            End If

            Return CType(ViewState("SortDirection"), SortDirection)
        End Get
        Set(ByVal value As SortDirection)
            ViewState("SortDirection") = value
        End Set
    End Property

    Protected Sub gvEmployee_RowCreated(sender As Object, e As GridViewRowEventArgs)
        'Dim ColumnTable As DataTable = CType(Session("Columns"), DataTable)

        'If e.Row.RowType = DataControlRowType.Header Then
        '    For Each tc As TableCell In e.Row.Cells
        '        If tc.HasControls() Then
        '            Dim lnk As LinkButton = CType(tc.Controls(0), LinkButton)
        '            lnk.CssClass = "headerlink"
        '            Dim sortDir As String = If(ViewState("SortDirection") IsNot Nothing, ViewState("SortDirection").ToString(), "Ascending")
        '            Dim sortBy As String = If(ViewState("SortExpression") IsNot Nothing, ViewState("SortExpression").ToString(), "---")

        '            If lnk IsNot Nothing AndAlso sortBy = lnk.CommandArgument Then
        '                Dim sortArrow As String = If(sortDir = "Ascending", " &#9650;", " &#9660;")
        '                lnk.Text += sortArrow
        '            End If

        '            lnk.Attributes("searchcolumnname") = lnk.Text.ToString()
        '            Dim row As DataRow = ColumnTable.Select("COLUMN_NAME = '" + lnk.Text + "'").FirstOrDefault()
        '            If Not row Is Nothing AndAlso row.Item("DisplayName").ToString() IsNot "" Then
        '                lnk.Text = row.Item("DisplayName").ToString()
        '            End If

        '        End If
        '    Next
        'End If

        Try
            Dim ColumnTable As DataTable = CType(Session("Columns"), DataTable)

            If e.Row.RowType = DataControlRowType.Header Then
                For Each tc As TableCell In e.Row.Cells
                    If tc.HasControls() Then
                        Dim lnk As LinkButton = CType(tc.Controls(0), LinkButton)
                        lnk.CssClass = "headerlink"
                        Dim sortDir As String = If(ViewState("SortDirection") IsNot Nothing, ViewState("SortDirection").ToString(), "Ascending")
                        Dim sortBy As String = If(ViewState("SortExpression") IsNot Nothing, ViewState("SortExpression").ToString(), "---")

                        If lnk IsNot Nothing AndAlso sortBy = lnk.CommandArgument Then
                            Dim sortArrow As String = If(sortDir = "Ascending", " &#9650;", " &#9660;")
                            lnk.Text += sortArrow
                        End If

                        lnk.Attributes("searchcolumnname") = lnk.Text.ToString()
                        Dim row As DataRow = ColumnTable.Select("COLUMN_NAME = '" + lnk.Text + "'").FirstOrDefault()
                        If Not row Is Nothing AndAlso row.Item("DisplayName").ToString() IsNot "" Then
                            lnk.Text = row.Item("DisplayName").ToString()
                        End If

                    End If
                Next
            End If
        Catch ex As Exception

        Finally

        End Try
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString
        Dim con As New SqlConnection(strConnString)

        Dim selecteddepartment As String = ""
        For Each item As ListItem In lstDepartment.Items
            If item.Selected Then
                If selecteddepartment = "" Then
                    selecteddepartment += "'" + item.Value + "'"
                Else
                    selecteddepartment += ",'" + item.Value + "'"
                End If
            End If
        Next

        Try
            Dim Id As Integer = 0
            If LoadOptions.Rows.Count > 0 Then
                Id = Convert.ToInt32(LoadOptions.Rows(0)(0).ToString())
            End If
            Dim query As String = "[SP_InsertColumns]"
            Dim cmd As New SqlCommand(query)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Columns", selecteddepartment & "|" & hdnSelectQuery.Value)
            cmd.Parameters.AddWithValue("@ID", Id)
            cmd.Connection = con
            cmd.Connection.Open()
            cmd.ExecuteNonQuery()

            Session("_Qry") = hdnSelectQuery.Value
            BindGrid()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub
End Class