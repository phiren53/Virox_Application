<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login_page.aspx.vb" Inherits="Virox_Application.Login_page" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Table ID="Table1" runat="server" borderstyle="Solid" HorizontalAlign="Right"  Width="30%" Height="30" style="font-family: Verdana; text-align: left;" cellpadding="4" Font-Size ="Small" >
                           <asp:TableRow ID="TableRow1" runat="server" Width="30%" HorizontalAlign ="Center" Font-Size ="10pt" VerticalAlign ="Middle">
                            <asp:TableCell ColumnSpan ="2"><asp:Label ID="Message" runat="server" Text="REPORT WRITER LOGIN" Font-Bold="true" BackColor ="lightgray" Width="100%" Height="40"></asp:Label></asp:TableCell>
            </asp:TableRow>                                 
            <asp:TableRow ID="TableRow2" runat="server" Width="30%">
                <asp:TableCell>EMPLOYEE ID</asp:TableCell>
                 <asp:TableCell><asp:TextBox ID="Login" runat="server" MaxLength ="8" Width="80%" Height="25"></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ErrorMessage="Required Field - Employee ID" ControlToValidate="Login" Text="*"></asp:RequiredFieldValidator></asp:TableCell>                                 
            </asp:TableRow>
             <asp:TableRow ID="TableRow10" runat="server" Width="30%">
                <asp:TableCell>PASSWORD</asp:TableCell>
                <asp:TableCell><asp:TextBox ID="Password" runat="server" TextMode="Password" Width="80%" Height="25"></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="Required Field - Password" ControlToValidate="Password" Text="*"></asp:RequiredFieldValidator>
                </asp:TableCell>
            </asp:TableRow>         
           
              <asp:TableRow ID="TableRow9" runat="server"  Width="30%">
                <asp:TableCell>
                                 
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="TableRow7" runat="server"  Width="30%">
                <asp:TableCell>
                    <asp:Button ID="Button1" runat="server" Text="Sign In" Font-Names="Verdana" Width="50%" Height="30" /></asp:TableCell>
            </asp:TableRow>

              <asp:TableRow ID="TableRow3" runat="server" Width="30%" HorizontalAlign ="Center" Font-Size ="10pt">
            <asp:TableCell ColumnSpan ="2" VerticalAlign ="Middle" ><asp:Label ID="Label1" runat="server" Text="VIROX SOFTWARE HOUSE SDN BHD" Font-Bold="true" BackColor ="lightgray" Width="100%" Height="30"></asp:Label></asp:TableCell>
            </asp:TableRow>    
        </asp:Table>

   
       

        </div>
    </form>
</body>
</html>
