<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="BaocaoHCM.aspx.cs" Inherits="XuatExcelClosedXML.Excels.BaocaoHCM" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="FeaturedContent" runat="server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">

    <asp:FileUpload ID="fileUpload" runat="server" />

    <asp:Button Text="Xuất báo cáo" ID="btnExport" OnClick="btnExport_Click" runat="server" />
</asp:Content>
