﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Plot_Graph.aspx.cs" Inherits="Projections_Capstone_Spring15.Plot_Graph" %>

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
        <div style="text-align: center">
            <span class="pageHeader">Plot your graphs</span>
        </div>
        <%--Accordian code here--%>
        <div id="accordion">
            <h3>Browse TEPC file</h3>
            <div>
                <asp:FileUpload ID="btnTEPCBrowse" runat="server" />
                <asp:Button ID="btnUploadTEPC" runat="server" OnClick="Button1_Click" Text="Upload" OnClientClick="javascript:$('#imgTEPCLoading').show();" />
                <asp:LinkButton runat="server" Text="Download average doses values" ID="lnkDownloadAvgTEPC" OnClick="lnkDownloadAvgTEPC_Click"></asp:LinkButton>
                <div>
                     <img src="Styles/images/ajax-loader.gif" id="imgTEPCLoading" runat="server" style="display: none" />
                    <asp:Label ID="lblErrorDescription" runat="server" Text="" ForeColor="Red"></asp:Label>
                </div>
            </div>
            <h3>Browse RAM TLD file</h3>
            <div>
                <asp:FileUpload ID="btnRAMBrowse" runat="server" />
                <asp:Button ID="btnUploadRAM_TLD" runat="server" OnClick="btnUploadRAM_TLD_Click" Text="Upload" OnClientClick="javascript:$('#imgRAMLoading').show();" />
                 <img src="Styles/images/ajax-loader.gif" id="imgRAMLoading" runat="server" style="display: none" />
                <div>
                    <asp:Label ID="lblErrorDescription_RAM_TLD" runat="server" ForeColor="Red"></asp:Label>
                </div>
                <br />
                <div style="float: left; width: 50%">
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
        <asp:Button runat="server" ID="btnPlot" Text="Plot" CssClass="ui-widget button btnPlot" />
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

    </form>
</body>
</html>
