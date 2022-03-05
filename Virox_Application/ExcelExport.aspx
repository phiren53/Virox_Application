<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ExcelExport.aspx.vb" Inherits="Virox_Application.ExcelExport" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excel Export</title>
    <link rel="stylesheet" href="Contents/Css/bootstrap.min.css">
    <link href="Contents/Css/bootstrap-multiselect.css" rel="stylesheet" type="text/css" />
    <link href="Contents/Css/custom.css" rel="stylesheet" type="text/css" />
    <script src="Contents/Scripts/jquery.min.js"></script>
    <script src="Contents/Scripts/bootstrap.min.js"></script>
    <script src="Contents/Scripts/bootstrap-multiselect.js" type="text/javascript"></script>
    <style>
        #header-fixed {
            position: fixed;
            top: 0px;
            display: none;
            background-color: white;
        }
         th {
            position: sticky !important;
            top: 0;
            box-shadow: 0 2px 2px -1px rgba(0, 0, 0, 0.4);
        }

        th, td {
            padding: 0.25rem;
        }
    </style>
    <script language="javascript">

        var recent_options = selected_options = added_options = removed_options = [];
        var current_values, index;
        function LoadData() {
            var value = $("#hdnSelectQuery").val();
            if ($("#hdnselected_options").val() != "") {
                var _options = $("#hdnselected_options").val().split(",");
                for (var i = 0; i < _options.length; i++) {
                    if (_options[i] > -1)
                        //selected_options.push(_options[i])
                        selected_options.push($('#<%=lstQuery.ClientID %>').find("option[value='" + _options[i] + "']").index());
                }
            }

            console.log("m", selected_options);
            if ($("#hdnremoved_options").val() != "") {
                var _options1 = $("#hdnremoved_options").val().split(",");
                for (var j = 0; i < _options1.length; j++) {
                    removed_options.push(_options1[j])
                }
            }

            var _values = $("#hdnadded_options").val();
            if ($("#hdnadded_options").val() != "") {
                var _options = $("#hdnadded_options").val().split(",");
                for (var i = 0; i < _options.length; i++) {
                    //alert(_options[i]);
                    if (_options[i] != "")
                        recent_options.push(_options[i])
                }
            }
            console.log(recent_options);

        //current_values = $("#<%=lstQuery.ClientID %>").val();

            $("#<%=lstQuery.ClientID %>").change(function (e) {

                //DOIt();

                var _TotalCount = 0;
                var _TotalSelected = 0;

                current_values = $(this).val();
                //recent_options = current_values ? current_values : [];

                $($("#<%=lstQuery.ClientID %>").children("option")
                    .get()
                    .reverse())
                    .each(function () {
                        if ($(this).prop('selected')) {
                            //Do work here for only elements that are selected
                            _TotalSelected += 1;
                        }
                        else {
                            var _value = $(this).val();

                            if (value.indexOf(_value) != -1) {
                                value = value.replace("," + $(this).val(), '');
                            }


                        }

                    });

                <%--$('#<%=lstQuery.ClientID %>').children("option").each(function () {
                    if ($(this).prop('selected')) {
                        //Do work here for only elements that are selected
                        _TotalSelected += 1;
                    }
                    else {
                        var _value = $(this).val();

                        if (value.indexOf(_value) != -1) {
                            console.log("Value", _value);
                            //console.log("arr", selected_options);
                            value = value.replace("," + $(this).val(), '');
                            console.log("here", value);
                        }


                    }


                });--%>


                $('#<%=lstQuery.ClientID %>').children("option").each(function () {
                    _TotalCount += 1;
                });

                if (_TotalCount == _TotalSelected) {
                    value = "";
                    $('#<%=lstQuery.ClientID %>').children("option").each(function () {
                        if ($(this).prop('selected')) {
                            added_options = current_values.filter(function (x) { return recent_options.indexOf(x) < 0 });
                            if (added_options.length > 0) {
                                selected_options.push($('#<%=lstQuery.ClientID %>').find("option[value='" + this.value + "']").index());
                                var s = value.indexOf(e.currentTarget.options[selected_options[selected_options.length - 1]].value);
                                if (value.indexOf(e.currentTarget.options[selected_options[selected_options.length - 1]].value) == -1) {
                                    value += "," + e.currentTarget.options[selected_options[selected_options.length - 1]].value
                                }

                            }
                        }
                    });

                } else {
                    if (_TotalSelected == 0) {
                        value = "";
                    }
                    if (current_values && current_values.length > 0) {
                        added_options = current_values.filter(function (x) { return recent_options.indexOf(x) < 0 });
                        if (added_options.length > 0) {
                            selected_options.push($(this).find("option[value='" + added_options[0] + "']").index());
                            var s = value.indexOf(e.currentTarget.options[selected_options[selected_options.length - 1]].value);
                            if (value.indexOf(e.currentTarget.options[selected_options[selected_options.length - 1]].value) == -1) {
                                value += "," + e.currentTarget.options[selected_options[selected_options.length - 1]].value
                            }

                        }
                        else {
                            removed_options = recent_options.filter(function (x) { return current_values.indexOf(x) < 0 });
                            if (removed_options.length > 0) {
                                index = selected_options.indexOf($(this).find("option[value='" + removed_options[0] + "']").index());
                                if (index > -1) {
                                    var d = e.currentTarget.options[selected_options[index]];
                                    if (typeof d != 'undefined') {

                                        value = value.replace("," + e.currentTarget.options[selected_options[index]].value, '');
                                        selected_options.splice(index, 1);
                                    }
                                }

                                //console.log("last selected option is : " + selected_options[selected_options.length - 1]);
                                if (value == "")
                                    value += "," + e.currentTarget.options[selected_options[selected_options.length - 1]].value
                                else
                                    if (typeof e.currentTarget.options[selected_options[selected_options.length - 1]] === 'undefined ')
                                        value += "," + e.currentTarget.options[selected_options[selected_options.length - 1]].value

                            }
                        }
                    }
                }


                recent_options = current_values ? current_values : [];

                /*console.log(recent_options);*/
                $("#hdnSelectQuery").val(value);
                $("#hdncurrent_values").val(current_values);
                $("#hdnselected_options").val(selected_options);
                $("#hdnadded_options").val(recent_options);
                $("#hdnremoved_options").val(removed_options);

            });

            $('#<%=gvEmployee.ClientID%>').css('position', 'absolute');
            $('#<%=gvEmployee.ClientID%>').css('top', '0');
        }

        $(document).ready(function () {
            LoadData();
        });


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

        function onchangetick(id, classname) {
            var ckName = document.getElementsByClassName(classname);
            var currentchk = document.getElementById("gvEmployee_" + id);
            if (currentchk.checked) {
                $('.headerCheckbox input').attr('disabled', 'disabled');
                var currentchk = document.getElementById("gvEmployee_" + id);
                currentchk.disabled = false;
            }
            else {
                for (var i = 0; i < ckName.length; i++) {
                    $('.headerCheckbox input').removeAttr('disabled');
                }
            }
        }

        $(".headerCheckbox input").click(function () {
            $(".headerCheckbox input").each(function () {
                var ckName = document.getElementsByClassName(ckType.name);
                if (this.checked) {
                    for (var i = 0; i < ckName.length; i++) {
                        ckName[i].disabled = true;
                    }
                    this.disabled = false;
                }
                else {
                    for (var i = 0; i < ckName.length; i++) {
                        ckName[i].disabled = false;
                    }
                }
            });
        });

        function ShowPopup() {
            $("#MyPopup").modal("show");
        }

        function CheckFilter() {
            var filter = "";
            var grid = document.getElementById('<%=gvEmployee.ClientID %>');
            //Get the count of columns.
            var columnCount = grid.rows[0].cells.length;
            for (var i = 1; i < columnCount; i++) {
                //var headerfiltertext = $(grid.rows[1].cells[i]).find('input[type=text]')[0];
                var headerfiltertext = $(grid.rows[1].cells[i]).find("input[type='text']")[0];
                var id = $(headerfiltertext).attr('id');
                var _Value = $(id).val();

                if (headerfiltertext.value != "") {
                    var querytype = "";
                    if (headerfiltertext.value.includes("<") || headerfiltertext.value.includes(">") || headerfiltertext.value.includes("=")) {
                        querytype = headerfiltertext.value;
                    }
                    else {
                        querytype = " LIKE '%" + headerfiltertext.value + "%'";
                    }
                    if (filter == "") {
                        filter += $(headerfiltertext).attr("backcolumn") + " " + querytype;
                    }
                    else {
                        filter += " AND " + $(headerfiltertext).attr("backcolumn") + " " + querytype;
                    }

                }
            }
            document.getElementById('<%=hdnFiletData.ClientID %>').value = filter;
        }

        function EditData() {

            document.getElementById('tablecontent').innerHTML = "";
            var editTable = document.createElement("TABLE");
            editTable.className = "table table-bordered";
            editTable.id = "dataTable";

            var headerdata = $("#GHead").find("table");
            var grid = document.getElementById('<%=gvEmployee.ClientID %>');
            var columnCount = grid.rows[0].cells.length;
            if (grid.rows[0].cells[1].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Employee_ID" && grid.rows[0].cells[2].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Employee_Name" && grid.rows[0].cells[3].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Department") {
                var checkpresent = "No";
                for (var i = 1; i < columnCount; i++) {
                    var headercheckbox = $($(headerdata)[0].rows[0].cells[i]).find('input[type=checkbox]')[0];
                    if (headercheckbox.checked) {
                        checkpresent = "Yes";
                        break;
                    }
                }
                if (checkpresent == "No") {
                    alert("Please select one column from grid to edit data");
                    return false;
                }
                else {
                    //Add the header row.
                    var row = editTable.insertRow(-1);
                    for (var i = 1; i < columnCount; i++) {
                        if (grid.rows[0].cells[i].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Employee_ID" || grid.rows[0].cells[i].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Employee_Name" || grid.rows[0].cells[i].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Department") {
                            var headerCell = document.createElement("TH");
                            headerCell.innerHTML = grid.rows[0].cells[i].innerText.replace(' ▼', '').replace(' ▲', '');
                            headerCell.setAttribute("editcolumnname", grid.rows[0].cells[i].getElementsByTagName('a')[0].getAttribute('searchcolumnname'));
                            headerCell.width = "25%";
                            row.appendChild(headerCell);
                        }
                    }
                    for (var i = 1; i < columnCount; i++) {
                        var headercheckbox = $($(headerdata)[0].rows[0].cells[i]).find('input[type=checkbox]')[0];
                        if (headercheckbox.checked) {
                            var headerCell = document.createElement("TH");
                            headerCell.innerHTML = grid.rows[0].cells[i].innerText.replace(' ▼', '').replace(' ▲', '');
                            headerCell.setAttribute("editcolumnname", grid.rows[0].cells[i].getElementsByTagName('a')[0].getAttribute('searchcolumnname'));
                            row.appendChild(headerCell);

                            $.ajax({
                                type: "POST",
                                url: "ExcelExport.aspx/BindEditDropdown",
                                data: JSON.stringify({ 'datatype': grid.rows[0].cells[i].innerText.replace(' ▼', '').replace(' ▲', '') }),
                                contentType: "application/json; charset=utf-8",
                                dataType: "json",
                                success: function (response) {
                                    $('#dropfield').empty();
                                    $.each(response.d, function (index, value) {
                                        $('#dropfield').append('<option value="' + value.Value + '">' + value.Text + '</option>');
                                    });
                                },
                                failure: function (response) {
                                    alert("Unable to bind");
                                }
                            });
                        }
                    }
                    var id = 1;
                    for (var i = 2; i < grid.rows.length; i++) {
                        var checkbox = $(grid.rows[i].cells[0]).find('input[type=checkbox]')[0];
                        if (checkbox.checked) {
                            row = editTable.insertRow(-1);
                            var hiddenid = $(grid.rows[i].cells[0]).find('input[type=hidden]')[0];
                            for (var j = 1; j < grid.rows[i].cells.length; j++) {
                                if (grid.rows[0].cells[j].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Employee_ID" || grid.rows[0].cells[j].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Employee_Name" || grid.rows[0].cells[j].getElementsByTagName('a')[0].getAttribute('searchcolumnname').replace(' ▼', '').replace(' ▲', '') == "Department") {
                                    var cell = row.insertCell(-1);

                                    var span = document.createElement("span");
                                    span.innerText = grid.rows[i].cells[j].innerText;
                                    cell.appendChild(span);
                                }
                            }
                            for (var j = 1; j < grid.rows[i].cells.length; j++) {
                                var headercheckbox = $($(headerdata)[0].rows[0].cells[j]).find('input[type=checkbox]')[0];
                                if (headercheckbox.checked) {
                                    var cell = row.insertCell(-1);

                                    var span = document.createElement("span");
                                    span.innerText = grid.rows[i].cells[j].innerText;
                                    cell.appendChild(span);
                                    cell.width = "25%";

                                    var input = document.createElement("input");
                                    input.type = "hidden";
                                    input.name = "hdn~" + id.toString();
                                    input.id = "hdn~" + id.toString();
                                    input.value = hiddenid.value;
                                    input.className = "form-control";
                                    cell.appendChild(input);
                                }
                            }

                            id = id + 1;
                        }
                    }
                    document.getElementById('tablecontent').appendChild(editTable);

                    //FixedHeader
                    document.getElementById('EditTableHead').innerHTML = "";
                    var editFixedHeaderTable = document.createElement("TABLE");
                    editFixedHeaderTable.className = "table table-bordered";
                    editFixedHeaderTable.id = "dataFixedHeaderTable";
                    $(editFixedHeaderTable).css('position', 'fixed');
                    $(editFixedHeaderTable).css('background-color', '#fff');
                    $(editFixedHeaderTable).css('z-index', '999');

                    var fixedrow = editFixedHeaderTable.insertRow(-1);
                    for (var x = 0; x < $("#dataTable")[0].rows[0].cells.length; x++) {
                        var headerCell = document.createElement("TH");
                        headerCell.innerHTML = $("#dataTable")[0].rows[0].cells[x].innerText.replace(' ▼', '').replace(' ▲', '');
                        headerCell.setAttribute("editcolumnname", $("#dataTable")[0].rows[0].cells[x].getAttribute('editcolumnname').replace(' ▼', '').replace(' ▲', ''));
                        headerCell.width = "25%";
                        fixedrow.appendChild(headerCell);
                    }

                    document.getElementById('tablecontent').insertAdjacentElement('afterbegin', editFixedHeaderTable);

                    $("#MyPopup").modal("show");
                    setTimeout(function () {
                        $(editFixedHeaderTable).css('width', $("#dataTable")[0].scrollWidth + 1);
                    }, 500);
                    return false;
                }
            }
            else {
                alert("Employee_ID, Employee_Name & Department columns is mendatory to edit data. Please select this columns from query list & search the data");
                return false;
            }
        }

        function UpdateData() {
            var array = {};
            var datalist = [];
            var newdata = "";

            var seldata = $('input[name=secitem]:checked', '#tab').val();
            if (seldata == "inputtextbox") {
                newdata = document.getElementById('textfield').value;
            } else if (seldata == "inputdate") {
                newdata = document.getElementById('datefield').value;
            }
            else if (seldata == "inputdropdown") {
                newdata = document.getElementById('dropfield').value;
            }

            var table = document.getElementById('dataTable');
            for (var x = 1; x < table.rows.length; x++) {
                var objects = {};
                var hiddenid = document.getElementById('hdn~' + x.toString());
                objects["Repository_ID"] = hiddenid.value;
                for (var j = 0; j < table.rows[x].cells.length; j++) {
                    var txtbox = document.getElementById('hedtxt~' + (j + 1).toString());
                    if (newdata != "") {
                        objects[table.rows[0].cells[3].getAttribute('editcolumnname')] = newdata;
                    }
                }

                datalist.push(objects);
            }
            //console.log(JSON.stringify(datalist));
            array.datalist = datalist;
            $.ajax({
                type: "POST",
                url: "ExcelExport.aspx/UpdateData",
                data: JSON.stringify(array),
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    alert(response.d);
                    if (response.d == "Data updated successfully") {
                        $("#MyPopup").modal("hide");
                        document.getElementById('<%= btnSearch.ClientID %>').click();
                    }
                },
                failure: function (response) {
                    alert("Unable to update");
                }
            });
            return false;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <div class="container">
            <div class="row">
                <div class="col-md-12 text-center">
                    <h3>Employee Master Report Writer</h3>
                    <br />
                    <asp:Label ID="Message" runat="server" Text=""></asp:Label>
                </div>
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
                                    <asp:Button ID="btnSearch" runat="server" Text="Search" CssClass="btn btn-info" OnClick="btnSearch_Click" OnClientClick="CheckFilter()" Width="40%" />
                                    <asp:Button ID="btnExcel" runat="server" Text="Excel" CssClass="btn btn-info" OnClick="btnExcel_Click" Width="40%" />
                                    <asp:Button ID="btnSave" runat="server" Text="Save Columns" CssClass="btn btn-info" Width="40%" />
                                    <asp:HiddenField ID="hdnFiletData" runat="server" />
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <div class="table-content">
                        <div id="GHead"></div>
                        <asp:GridView ID="gvEmployee" runat="server" ShowHeaderWhenEmpty="True" EmptyDataText="No records Found" AllowSorting="true" OnSorting="gvEmployee_Sorting" CssClass="table table-bordered" OnDataBound="gvEmployee_DataBound" OnRowDataBound="gvEmployee_RowDataBound" OnRowCreated="gvEmployee_RowCreated">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox runat="server" />
                                        <asp:HiddenField ID="hdnPKID" runat="server" Value='<%# Eval("Repository_ID") %>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
            </div>
            <br />
            <%--  <div class="Pager">
                <asp:LinkButton ID="lnkPrevious" runat="server" OnClick="lnkPrevious_Click"><< Previous Page</asp:LinkButton>
                ||  
                <asp:LinkButton ID="lnkNext" runat="server" OnClick="lnkNext_Click">Next Page >></asp:LinkButton>
            </div>--%>

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
                            <asp:Button ID="btnUpdate" runat="server" Text="Update" CssClass="btn btn-info" OnClientClick="return UpdateData()" />
                            <button type="button" class="btn btn-danger" data-dismiss="modal">
                                Close</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <asp:HiddenField ID="hdnSelectQuery" runat="server" />
        <asp:HiddenField ID="hdnselected_options" runat="server" />
        <asp:HiddenField ID="hdnadded_options" runat="server" />
        <asp:HiddenField ID="hdnremoved_options" runat="server" />
        <asp:HiddenField ID="hdncurrent_values" runat="server" />
        <!-- Modal Popup -->
    </form>
</body>

</html>
