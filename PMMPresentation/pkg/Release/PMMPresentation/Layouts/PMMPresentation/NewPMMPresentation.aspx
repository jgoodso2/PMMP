<%@ Assembly Name="PMMPresentation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=4817109300b3adf3" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="NewPMMPresentation.aspx.cs" Inherits="PMMPresentation.Layouts.PMMPresentation.NewPMMPresentation" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <SharePoint:CssRegistration ID="CssRegistration1" runat="server" Name="<% $SPUrl:~SiteCollection/_layouts/PMMPresentation/PMMPresentation.css %>"></SharePoint:CssRegistration>
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <asp:MultiView runat="server" ID="mvMain" ActiveViewIndex="0">
    <asp:View runat="server" ID="viewDefault">
        <asp:Panel ID="pnlSubmit" runat="server" Visible="false">
	        <script type="text/javascript" language="javascript">
	            document.onload = submitForm();
	            function submitForm() {
	                SP.UI.ModalDialog.commonModalDialogClose(1, 'Submitted');
	                return false;
	            }
	        </script>
        </asp:Panel>
        <asp:Panel ID="pnlClose" runat="server" Visible="false">
	      <script type="text/javascript" language="javascript">
	          document.onload = closeForm();
	          function closeForm() {
	              SP.UI.ModalDialog.commonModalDialogClose(0, 'Canceled');
	              return false;
	          }
	      </script>
	    </asp:Panel>
        <div class="nsr-container">
            <div class="nsr-row">
                <div class="nsr-row-column">Document Name:</div>
                <div class="nsr-row-column nsr-row-column-control"><asp:TextBox runat="server" ID="txtDocumentName" Width="295px"></asp:TextBox>&nbsp;.xlsx</div>
            </div>
            <div class="nsr-row">
                <div class="nsr-row-column">Comment:</div>
                <div class="nsr-row-column nsr-row-column-control"><asp:TextBox runat="server" ID="txtComment" Width="395px" TextMode="MultiLine" Rows="4" style="OVERFLOW:auto; RESIZE:none;"></asp:TextBox></div>
            </div>
            <div class="nsr-buttonsrow">
                <asp:Button runat="server" ID="btnCreate" Text="Create" Width="120px" OnClick="btnCreate_Click" />&nbsp;&nbsp;&nbsp;
                <asp:Button runat="server" ID="btnCancel" Text="Cancel" Width="120px" OnClick="btnCancel_Click" />
            </div>
        </div>
    </asp:View>
    <asp:View runat="server" ID="viewError">
        <asp:Literal runat="server" ID="lblErrorMessage"></asp:Literal>
    </asp:View>
    </asp:MultiView>
</asp:Content>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    New PMM Presentation
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
    New PMM Presentation
</asp:Content>
