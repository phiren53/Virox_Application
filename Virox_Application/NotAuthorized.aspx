<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NotAuthorized.aspx.vb" Inherits="Virox_Application.NotAuthorized" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
                <br />
<asp:Table runat="server" Width="70%" BorderStyle="Double" BorderColor="Black" Font-Names="calibri" CellPadding ="4">
      <asp:TableHeaderRow runat="server" >
            <asp:TableHeaderCell>YOU ARE NOT AUTHORIZED TO ACCESS THIS FUNCTION. PLEASE CONTACT ADMINISTRATOR OR YOUR SESSION HAS BEEN EXPIRED </asp:TableHeaderCell>
        </asp:TableHeaderRow> 
        <asp:TableHeaderRow runat="server" >
            <asp:TableHeaderCell>
                <asp:Button ID="Button1" runat="server" Text="HOME PAGE" />
            </asp:TableHeaderCell>
        </asp:TableHeaderRow> 
      </asp:Table> 
        </div>
    </form>
</body>
</html>
