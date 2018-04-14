<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Employee.aspx.cs" Inherits="WebApplication1.Excels.Employee" %>

<%@ Register Assembly="ASP.Web.UI.PopupControl"
    Namespace="ASP.Web.UI.PopupControl"
    TagPrefix="ASPP" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="FeaturedContent" runat="server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">
    <div class="row">
        <div class="col-md-12">
            <div class="form-inline">
                Chi nhánh:
    <asp:DropDownList OnSelectedIndexChanged="ddlChiNhanh_SelectedIndexChanged" AutoPostBack="true" runat="server" ID="ddlChiNhanh"></asp:DropDownList>
                Phòng ban:
    <asp:DropDownList runat="server" ID="ddlPhongBan"></asp:DropDownList>
                <asp:Button runat="server" ID="btnChange" CssClass="btn btn-primary" Text="Xem dữ liệu" OnClick="btnChange_Click" />
                <!-- Example single danger button -->
                <div class="btn-group">
                    <button type="button" class="btn btn-danger dropdown-toggle" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                        Xuất Excel
                    </button>
                    <div class="dropdown-menu">
                        <asp:Button runat="server" CssClass="dropdown-item" ID="btnExport" Text="Xuất đơn" OnClick="btnExport_Click" />
                        <asp:Button runat="server" CssClass="dropdown-item" ID="Button1" Text="Xuất 1" OnClick="Button1_Click" />
                        <asp:Button Enabled="false" runat="server" CssClass="dropdown-item" ID="btnExport2" Text="Xuất nhóm" OnClick="btnExport2_Click" />
                        <button class="dropdown-item" type="button" id="btnOpenModal" onclick="$('#modalAdvanted').modal('show');">Xuất nâng cao</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <asp:UpdatePanel runat="server" ID="pnlUpdate">
        <ContentTemplate>
            <div class="row" style="margin-top: 20px;">
                <div class="col-md-12 col-xs-12">
                    <div class="table-responsive">
                        <div style="background-color: #ffffff">
                            <table class="table table-striped table-bordered">
                                <thead>
                                    <th>Mã số</th>
                                    <th>Họ tên</th>
                                    <th>Phòng ban</th>
                                    <th>Ngày sinh</th>
                                    <th>Địa chỉ</th>
                                    <th>Điện thoại</th>
                                </thead>
                                <tbody>
                                    <%foreach (WebApplication1.Models.Entites.NhanVien item in employees)
                                      {
                                    %>
                                    <tr>
                                        <td><%=item.MaNV %></td>
                                        <td><%=item.TenNV %></td>
                                        <td><%=item.MaPB %></td>
                                        <td><%=item.NgaySinh %></td>
                                        <td><%=item.DiaChi %></td>
                                        <td><%=item.DienThoai %></td>
                                    </tr>
                                    <%
                                      } %>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

        </ContentTemplate>
    </asp:UpdatePanel>
    <div class="modal fade" id="modalAdvanted" tabindex="-1" role="dialog">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Modal title</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                </div>
                <div class="modal-body">
                    <asp:CheckBoxList ID="ckblColumns" runat="server"></asp:CheckBoxList>
                </div>
                <div class="modal-footer">
                    <asp:Button runat="server" CssClass="btn btn-primary" ID="btnAdvanted" Text="Xuất" OnClick="btnAdvanted_Click" />
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                </div>
            </div>
        </div>
    </div>

</asp:Content>
