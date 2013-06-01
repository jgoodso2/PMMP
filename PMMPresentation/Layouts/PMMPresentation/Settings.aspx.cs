using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Repository;
using PMMPresentation.Support;
using System.Linq;
using System.Web.UI.WebControls;
using Configuration = PMMPPresentation.Configuration;
using Constants = PMMP.Constants;
using System.Security.Principal;
using System.Net;
namespace PMMPresentation.Layouts.PMMPresentation
{
    public partial class Settings : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                //This gets the app pool account
//                SPSecurity.RunWithElevatedPrivileges(() =>
//{
    var obj = System.Security.Principal.WindowsIdentity.GetCurrent();
    lblLoggedInUser.Text = "The Logged in User =" + SPContext.Current.Web.CurrentUser.LoginName;
    lblAppPoolUser.Text = "The App Pool User =" + System.Security.Principal.WindowsIdentity.GetCurrent().Name;
    if (!Page.IsPostBack)
    {
       
            // Now impersonating
            // Access resources using the identity of the authenticated user
            this.txtServiceURL.Text = Configuration.ServiceURL;
            this.lnkTemplate.NavigateUrl = String.Format("{0}/{1}", this.Web.Url, Constants.TEMPLATE_FILE_LOCATION);

            if (!String.IsNullOrEmpty(this.txtServiceURL.Text))
            {
                DataRepository.ClearImpersonation();

//                CredentialCache.DefaultNetworkCredentials

                if (DataRepository.P14Login(this.txtServiceURL.Text))
                {
                    var projectList = DataRepository.ReadProjectsList();
                    this.ddlProject.DataSource = projectList.Tables["Project"];
                    this.ddlProject.DataTextField = "PROJ_NAME";
                    this.ddlProject.DataValueField = "PROJ_UID";
                    this.ddlProject.DataBind();

                    int index = 0;
                    foreach (ListItem item in this.ddlProject.Items)
                    {
                        if (item.Value == Configuration.ProjectUID)
                            break;

                        index++;
                    }

                    if (index >= this.ddlProject.Items.Count)
                        index = 0;

                    this.ddlProject.SelectedIndex = index;
                }
            }
    }
//});
            }
            catch (Exception ex)
            {
                this.ShowErrorMessage(ex);
            }

        }

        protected void btnLoad_Click(object sender, EventArgs e)
        {
            SPSecurity.RunWithElevatedPrivileges(() =>
{

    try
    {
        if (!String.IsNullOrEmpty(this.txtServiceURL.Text))
        {
                // Access resources using the identity of the authenticated user
                DataRepository.ClearImpersonation();

                if (DataRepository.P14Login(this.txtServiceURL.Text))
                {
                    var obj = System.Security.Principal.WindowsIdentity.GetCurrent();
                    var projectList = DataRepository.ReadProjectsList();
                    ddlProject.DataSource = projectList.Tables["Project"];
                    ddlProject.DataTextField = "PROJ_NAME";
                    ddlProject.DataValueField = "PROJ_UID";
                    ddlProject.DataBind();
                }
        }
    }
    catch (Exception ex)
    {
        this.ShowErrorMessage(ex);
    }
});
        }

        protected void btnOK_Click(object sender, EventArgs e)
        {
            SPSecurity.RunWithElevatedPrivileges(() =>
{
    try
    {
            // Access resources using the identity of the authenticated user
            if (!String.IsNullOrEmpty(this.txtServiceURL.Text) && !String.IsNullOrEmpty(this.ddlProject.SelectedValue))
            {
                var updateList = Configuration.ProjectUID != this.ddlProject.SelectedValue;
                Configuration.ServiceURL = this.txtServiceURL.Text;
                Configuration.ProjectUID = this.ddlProject.SelectedValue;

                if (this.TemplateFileUpload.HasFile)
                    this.Web.Files.Add(Constants.TEMPLATE_FILE_LOCATION, this.TemplateFileUpload.FileBytes, true);

                this.pnlClose.Visible = true;
            }
    }
    catch (Exception ex)
    {
        this.ShowErrorMessage(ex);
    }
});
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            this.pnlClose.Visible = true;
        }

        private void ShowErrorMessage(Exception ex)
        {
            this.lblErrorMessage.Text = String.Format("An exception has ocurred. Message:{0}. Stack Trace:{1}", ex.Message, ex.StackTrace);
            this.mvMain.ActiveViewIndex = 1;
        }
    }
}
