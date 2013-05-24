<%@ Assembly Name="PMMPresentation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=4817109300b3adf3" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Settings.aspx.cs" Inherits="PMMPresentation.Layouts.PMMPresentation.Settings" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <SharePoint:CssRegistration ID="CssRegistration1" runat="server" Name="<% $SPUrl:~SiteCollection/_layouts/PMMPresentation/PMMPresentation.css %>"></SharePoint:CssRegistration>
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
<asp:MultiView runat="server" ID="mvMain" ActiveViewIndex="0">
<asp:View runat="server" ID="viewDefault">
    <asp:Panel ID="pnlClose" runat="server" Visible="false">
	    <script type="text/javascript" language="javascript">
	        document.onload = closeForm();
	        function closeForm() {
	            SP.UI.ModalDialog.commonModalDialogClose(0, 'Close');
	            return false;
	        }
	    </script>
	</asp:Panel>
    <div class="srs-container">
     <div><asp:Label ID="lblLoggedInUser" runat="server"></asp:Label></div>
                <div><asp:Label ID="lblAppPoolUser" runat="server"></asp:Label></div>
        <div class="srs-row">
            <div class="srs-row-column srs-row-column-label">
                <div><b>Service URL</b></div>
                <div></div>
            </div>
            <div class="srs-row-column srs-row-column-control">
               
                <asp:TextBox runat="server" ID="txtServiceURL" Width="300px"></asp:TextBox>
                <asp:Button runat="server" ID="btnLoadProjects" Text="Load" Width="100px" OnClick="btnLoad_Click" CausesValidation="false" />
                <span class="ms-formvalidation">
                    <asp:RequiredFieldValidator runat="server" ID="txtServiceURLValidator" ControlToValidate="txtServiceURL" Display="Dynamic" ErrorMessage="You must specify a value for this required field."></asp:RequiredFieldValidator>
                </span>
            </div>    
        </div>
        <div class="srs-row">
            <div class="srs-row-column srs-row-column-label">
                <div><b>Project</b></div>
                <div></div>
            </div>
            <div class="srs-row-column srs-row-column-control">
                <asp:DropDownList runat="server" ID="ddlProject" Width="300px"></asp:DropDownList>
                <span class="ms-formvalidation">
                    <asp:RequiredFieldValidator runat="server" ID="ddlProjectValidator" ControlToValidate="ddlProject" Display="Dynamic" ErrorMessage="You must specify a value for this required field."></asp:RequiredFieldValidator>
                </span>
            </div>    
        </div>
        <div class="srs-row">
            <div class="srs-row-column srs-row-column-label">
                <div><b>Presentation Template</b></div>
                <div style="padding-top:5px;">
                    To see the current template click&nbsp;<asp:HyperLink runat="server" ID="lnkTemplate" Text="here"></asp:HyperLink>
                </div>
            </div>
            <div class="srs-row-column srs-row-column-control">
                <div><asp:FileUpload runat="server" ID="TemplateFileUpload" /></div>
            </div>
        </div>
        <div class="srs-buttonsrow">
            <asp:Button runat="server" ID="btnOK" Text="OK" Width="120px" OnClick="btnOK_Click" />&nbsp;&nbsp;&nbsp;
            <asp:Button runat="server" ID="btnCancel" Text="Cancel" Width="120px" OnClick="btnCancel_Click" CausesValidation="false" />
        </div>
    </div>
</asp:View>
<asp:View runat="server" ID="viewError">
    <asp:Literal runat="server" ID="lblErrorMessage"></asp:Literal>
</asp:View>
</asp:MultiView>
</asp:Content>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    Settings
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
    Settings
</asp:Content>
