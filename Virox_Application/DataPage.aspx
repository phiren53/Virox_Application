<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DataPage.aspx.vb" Inherits="Virox_Application.DataPage" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Data Page</title>
    <link rel="stylesheet" href="Contents/Css/bootstrap.min.css">
    <link href="Contents/Css/bootstrap-multiselect.css" rel="stylesheet" type="text/css" />
    <link href="https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css" rel="stylesheet" type="text/css" />
    <link href="Contents/Css/custom.css" rel="stylesheet" type="text/css" />
    <script src="Contents/Scripts/jquery.min.js"></script>
    <script src="https://cdn.datatables.net/1.10.20/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/fixedcolumns/3.3.0/js/dataTables.fixedColumns.min.js"></script>
    <script src="Contents/Scripts/bootstrap.min.js"></script>
    <script src="Contents/Scripts/bootstrap-multiselect.js" type="text/javascript"></script>

    <script type="text/javascript">
        $(function () {
            $('[id*=lstDepartment]').multiselect({
                includeSelectAllOption: true
            });
            $('[id*=lstEmployeeType]').multiselect({
                includeSelectAllOption: true
            });
            $('[id*=lstStatus]').multiselect({
                includeSelectAllOption: true
            });
            $('[id*=lstQuery]').multiselect({
                includeSelectAllOption: true
            });
        });

        function searchData() {
            var selectedcolumns = "";
            var QuerylistBox = document.getElementById("<%= lstQuery.ClientID%>");
            for (var i = 0; i < QuerylistBox.options.length; i++) {
                if (QuerylistBox.options[i].selected) {
                    if (selectedcolumns == "") {
                        selectedcolumns += QuerylistBox.options[i].value;
                    }
                    else {
                        selectedcolumns += "," + QuerylistBox.options[i].value;
                    }
                }
            }
            var selecteddepartment = "";
            var DepartmentlistBox = document.getElementById("<%= lstDepartment.ClientID%>");
            for (var i = 0; i < DepartmentlistBox.options.length; i++) {
                if (DepartmentlistBox.options[i].selected) {
                    if (selecteddepartment == "") {
                        selecteddepartment += "'" + DepartmentlistBox.options[i].value + "'";
                    }
                    else {
                        selecteddepartment += ",'" + DepartmentlistBox.options[i].value + "'";
                    }
                }
            }
            var selectedemployeetype = "";
            var EmployeeTypelistBox = document.getElementById("<%= lstEmployeeType.ClientID%>");
            for (var i = 0; i < EmployeeTypelistBox.options.length; i++) {
                if (EmployeeTypelistBox.options[i].selected) {
                    if (selectedemployeetype == "") {
                        selectedemployeetype += "'" + EmployeeTypelistBox.options[i].value + "'";
                    }
                    else {
                        selectedemployeetype += ",'" + EmployeeTypelistBox.options[i].value + "'";
                    }
                }
            }
            var selectedstatus = "";
            var StatuslistBox = document.getElementById("<%= lstStatus.ClientID%>");
            for (var i = 0; i < StatuslistBox.options.length; i++) {
                if (StatuslistBox.options[i].selected) {
                    if (selectedstatus == "") {
                        selectedstatus += "'" + StatuslistBox.options[i].value + "'";
                    }
                    else {
                        selectedstatus += ",'" + StatuslistBox.options[i].value + "'";
                    }
                }
            }

            var array = {};

            array.selectedcolumns = selectedcolumns;
            array.selecteddepartment = selecteddepartment;
            array.selectedemployeetype = selectedemployeetype;
            array.selectedstatus = selectedstatus;
            var querydata = JSON.stringify(array);
            console.log(JSON.stringify(array));
            $.ajax({
                type: "POST",
                url: "DataPage.aspx/DataListDetails",
                data: "{'querydata' : '" + JSON.stringify(array) + "'}",
                //data: {querydata : JSON.stringify(array)},
                async: false,
                cache: false,
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var dataresult = [];
                    var columnobjects = JSON.parse(response.d.columns);
                    var dataObj = JSON.parse(response.d.data);
                    dataObj.forEach(function (obj) {
                        if (obj != "") {
                            var singleobjects = [];
                            Object.keys(obj).forEach(function (key) {
                                if (key != "RowNumber" && key != "Repository_ID") {
                                    singleobjects.push(obj[key]);
                                }
                            });
                            dataresult.push(singleobjects);
                        }
                    });
                    console.log("column :", JSON.parse(response.d.columns));
                    if ($.fn.DataTable.isDataTable('#datatablelist')) {
                        $('#datatablelist').DataTable().destroy();
                    }
                    $('#datatablelist').empty();
                    var table = $('#datatablelist').DataTable({
                        data: dataresult,
                        columns: columnobjects,
                        searching: false,
                        ordering: true,
                        info: false,
                        fixedHeader: true,
                        scrollX: true,
                        scrollY: "550px",
                        scrollCollapse: true,
                        paging: false
                    });

                    // Setup - add a text input to each footer cell
                    $('#datatablelist thead tr').clone(true).appendTo('#datatablelist thead');
                    $('#datatablelist thead tr:eq(1) th').each(function (i) {
                        var title = $(this).text();
                        $(this).html('<input type="text" placeholder="Search ' + title + '" />');

                        $('input', this).on('keyup change', function () {
                            if (table.column(i).search() !== this.value) {
                                table
                                    .column(i)
                                    .search(this.value)
                                    .draw();
                            }
                        });
                    });
                },
                failure: function (response) {
                    alert(response.d);
                },
                error: function (response) {
                    alert(response.d);
                }
            });
            return false;
        }

        $(document).ready(function () {
            
        });
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <div class="container">
            <div class="row">
                <div class="col-md-12">
                    <table class="table table-bordered">
                        <tbody>
                            <tr>
                                <td>Department:
                            <asp:ListBox ID="lstDepartment" runat="server" SelectionMode="Multiple"></asp:ListBox></td>
                                <td>Employee Type:
                            <asp:ListBox ID="lstEmployeeType" runat="server" SelectionMode="Multiple"></asp:ListBox></td>
                                <td>Status:
                            <asp:ListBox ID="lstStatus" runat="server" SelectionMode="Multiple"></asp:ListBox></td>
                                <td>Query:
                            <asp:ListBox ID="lstQuery" runat="server" SelectionMode="Multiple"></asp:ListBox></td>
                                <td>
                                    <asp:Button ID="btnSearch" runat="server" Text="Search" CssClass="btn btn-info" OnClientClick="return searchData()" />
                                    <asp:Button ID="btnEdit" runat="server" Text="Edit" CssClass="btn btn-info" />
                                    <asp:Button ID="btnExcel" runat="server" Text="Excel" CssClass="btn btn-info" />
                                    <asp:HiddenField ID="hdnFiletData" runat="server" />
                                    <asp:HiddenField ID="hdnColumnData" runat="server" />
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <div class="table-content">
                        <table id="datatablelist" class="display table table-striped table-bordered" cellspacing="0" width="100%"></table>
                    </div>
                </div>
            </div>
            <br />
            <!-- Modal Popup -->
            <div id="MyPopup" class="modal fade" role="dialog">
                <div class="modal-dialog">
                    <!-- Modal content-->
                    <div class="modal-content">
                        <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal">
                                &times;</button>
                            <h4 class="modal-title">Edit Records</h4>
                        </div>
                        <div class="modal-body">
                            <div style="width: 100%;">
                                <div id="tab" class="btn-group btn-group-justified" data-toggle="buttons">
                                    <a href="#prices" class="btn btn-default active" data-toggle="tab">
                                        <input type="radio" name="secitem" value="inputtextbox" checked="checked" />Input Textbox
                                    </a>
                                    <a href="#features" class="btn btn-default" data-toggle="tab">
                                        <input type="radio" name="secitem" value="inputdate" />Input Date
                                    </a>
                                    <a href="#requests" class="btn btn-default" data-toggle="tab">
                                        <input type="radio" name="secitem" value="inputdropdown" />Input Dropdown
                                    </a>
                                </div>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="prices">
                                        <div class="row">
                                            <div class="col-md-3 text-right" style="line-height: 34px;">
                                                Input Textbox :
                                            </div>
                                            <div class="col-md-4">
                                                <input id="textfield" type="text" class="form-control" />
                                            </div>
                                        </div>
                                    </div>
                                    <div class="tab-pane" id="features">
                                        <div class="row">
                                            <div class="col-md-3 text-right" style="line-height: 34px;">
                                                Input Date :
                                            </div>
                                            <div class="col-md-4">
                                                <input id="datefield" type="date" class="form-control" />
                                            </div>
                                        </div>
                                    </div>
                                    <div class="tab-pane" id="requests">
                                        <div class="row">
                                            <div class="col-md-3 text-right" style="line-height: 34px;">
                                                Input Dropbox :
                                            </div>
                                            <div class="col-md-4">
                                                <select id="dropfield" class="form-control">
                                                    <option></option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div id="EditTableHead"></div>
                                <table id="header-fixed"></table>
                                <div id="tablecontent">
                                </div>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <asp:Button ID="btnUpdate" runat="server" Text="Update" CssClass="btn btn-info" />
                            <button type="button" class="btn btn-danger" data-dismiss="modal">
                                Close</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <!-- Modal Popup -->
    </form>
</body>
</html>
