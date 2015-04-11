<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Plot_Graph.aspx.cs" Inherits="Projections_Capstone_Spring15.Plot_Graph" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <title>Plot Graph</title>
    <link href="Styles/main.css" rel="stylesheet" />
    <link href="Styles/jquery-ui.css" rel="stylesheet" />
</head>
<body>
    <form id="frmGraphs" runat="server">
        <div style="text-align:center">
            <span class="pageHeader">Plot your graphs</span>
        </div>
        <%--Accordian code here--%>
        <div id="accordion">
            <h3>Browse TEPC file</h3>
            <div>
                <asp:Button runat="server" ID="btnTEPCBrowse" Text="Browse" OnClick="btnTEPCBrowse_Click" />
                <asp:Label runat="server" ID="lblTEPCPath" Visible="false"></asp:Label>
            </div>
            <h3>Browse RAM TLD file</h3>
            <div>
                <asp:Button runat="server" ID="btnRAMBrowse" Text="Browse" />
                <asp:Label runat="server" ID="lblRAMPath" Visible="false"></asp:Label>
                <br />
                <div style="float: left; width:50%">
                    <p>
                        Start Date:
                    <asp:TextBox runat="server" ID="datepickerStart" />
                    </p>
                </div>
                <div style="float: left">
                    <p>
                        End Date:
                    <asp:TextBox runat="server" ID="datepickerEnd" />
                    </p>
                </div>
                </div>
        </div>
        <%--Accordian code end--%>
        <asp:Button runat="server" ID="btnPlot" text="Plot" CssClass="ui-widget button btnPlot"/>
    </form>
    <script src="Scripts/jquery.js"></script>
    <%--    <script src="//code.jquery.com/jquery-1.11.2.min.js"></script>--%>
    <script src="Scripts/jquery-ui.js"></script>
    <script>
        $(document).ready(function () {
            $("#accordion").accordion();
        });
        $(function () {
            $('[id^=datepicker]').datepicker();
        });


    </script>
</body>
</html>
