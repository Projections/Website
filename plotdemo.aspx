<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="plotdemo.aspx.cs" Inherits="Projections_Capstone_Spring15.plotdemo" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <script src="Scripts/jquery.js"></script>
    <script src="http://code.highcharts.com/highcharts.js"></script>
<script src="http://code.highcharts.com/highcharts-more.js"></script>
<script src="http://code.highcharts.com/modules/exporting.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:literal ID="ltrPlot" runat="server"></asp:literal>
    </div>
        </form>
     <script src="Scripts/jquery-ui.js"></script>
        <script src="Scripts/Highcharts-4.0.1/js/highcharts.js"></script>
</body>
</html>
