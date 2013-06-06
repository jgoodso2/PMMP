﻿using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;

using System.Linq;

using Configuration = PMMPPresentation.Configuration;
using Constants = PMMP.Constants;
using WCFHelpers;
using System.Security.Principal;
using Repository;

namespace PMMPresentation.Layouts.PMMPresentation
{
    /// <summary>
    /// 
    /// </summary>
    public partial class NewPMMPresentation : LayoutsPageBase
    {
        private Guid ListId
        {
            get { return new Guid(this.Request.QueryString["List"]); }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
                this.txtDocumentName.Text = String.Format("PMMPresentation_{0}_{1}", this.Web.Title, DateTime.Now.ToString("dMMMyyyy"));
        }

        #region Events

        protected void btnCreate_Click(object sender, EventArgs e)
        {
            SPSecurity.RunWithElevatedPrivileges(() =>
            {

                if (!String.IsNullOrEmpty(Configuration.ServiceURL) && !String.IsNullOrEmpty(Configuration.ProjectUID))
                {


                    if (DataRepository.P14Login(((string)SPContext.Current.Web.Properties[Constants.PROPERTY_NAME_DB_SERVICE_URL])))
                    // Start impersonating
                    {
                        var docName = this.txtDocumentName.Text + ".pptx";
                        var docLib = this.Web.Lists[this.ListId];
                        var templateFile = this.Web.GetFile(Constants.TEMPLATE_FILE_LOCATION);
                        PMMP.PMPDocument document = new PMMP.PMPDocument();
                        var stream = document.CreateDocument("Presentation", templateFile.OpenBinary(), Configuration.ProjectUID);




                        SPListItem item = docLib.RootFolder.Files.Add(docName, stream, true).Item;

                        var pmmContentType = (from SPContentType ct in docLib.ContentTypes
                                              where ct.Name.ToLower() == Constants.CT_PMM_NAME.ToLower()
                                              select ct).FirstOrDefault();

                        if (pmmContentType != null)
                        {
                            item[SPBuiltInFieldId.ContentTypeId] = pmmContentType.Id;
                            item[Constants.FieldId_Comments] = txtComment.Text;
                            item.Update();
                        }
                    }
                }
                else
                {
                    this.ShowErrorMessage("An error has ocurred trying to get the configuration parameters");
                }
            });
            this.pnlSubmit.Visible = true;
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            var docLib = this.Web.Lists[this.ListId];
            this.pnlClose.Visible = true;
        }

        #endregion

        #region Private methods

        private void ShowErrorMessage(Exception ex)
        {
            this.lblErrorMessage.Text = String.Format("An exception has ocurred. Message:{0}. Stack Trace:{1}", ex.Message, ex.StackTrace);
            this.mvMain.ActiveViewIndex = 1;
        }

        private void ShowErrorMessage(string message)
        {
            this.lblErrorMessage.Text = message;
            this.mvMain.ActiveViewIndex = 1;
        }

        #endregion
    }
}
