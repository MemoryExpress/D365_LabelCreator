namespace D365_LabelCreator
{
    using System;
    using System.ComponentModel.Composition;
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Forms;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Tables;
    using Microsoft.Dynamics.AX.Metadata.Service;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.BaseTypes;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Menus;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Security;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Reports;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Views;


    /// <summary>
    /// TODO: Say a few words about what your AddIn is going to do
    /// </summary>
    [Export(typeof(IDesignerMenu))]
    // TODO: This addin will show when user right clicks on a form root node or table root node. 
    // If you need to specify any other element, change this AutomationNodeType value.
    // You can specify multiple DesignerMenuExportMetadata attributes to meet your needs
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ITable))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IBaseField))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IView))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IViewBaseField))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IFormDesign))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IFormControl))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ISecurityPrivilege))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ISecurityDuty))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ISecurityRole))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IEdtBase))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IMenuItem))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IReportDataSetField))]
    public class AddLabelAddin : DesignerMenuBase
    {
        #region Member variables
        private const string addinName = "MMX_AddLabelAddin";
        private LabelHelper helper = new LabelHelper();
        private string labelPrefix;
        #endregion

        #region Properties
        /// <summary>
        /// Caption for the menu item. This is what users would see in the menu.
        /// </summary>
        public override string Caption
        {
            get
            {
                return AddinResources.AddLabelsCaption;
            }
        }


        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return AddLabelAddin.addinName;
            }
        }
        #endregion

        #region Callbacks
        /// <summary>
        /// /Handle the click event in the OnClick method.  Here we will test the selected object, get the object’s model and label file, and create the label.
        /// </summary>
        /// <param name="e">The context of the VS tools and metadata</param>
        public override void OnClick(AddinDesignerEventArgs e)
        {
            try
            {
                // Get the metamodel service
                IMetaModelService metaModelService = helper.MetaModelService;
                
                // Is the selected element a table?
                if (e.SelectedElement is ITable)
                {
                    ITable table = e.SelectedElement as ITable;
                    helper.setModelAndLabelFile(metaModelService.GetTableModelInfo(table.Name));

                    labelPrefix = table.Name;
                }
                else if (e.SelectedElement is IBaseField)
                {
                    IBaseField baseField = e.SelectedElement as IBaseField;
                    var table = baseField.Table;
                    helper.setModelAndLabelFile(metaModelService.GetTableModelInfo(table.Name));

                    labelPrefix = String.Format("{0}_{1}", baseField.Table.Name, baseField.Name);
                }
                if (e.SelectedElement is IView)
                {
                    IView view = e.SelectedElement as IView;
                    helper.setModelAndLabelFile(metaModelService.GetTableModelInfo(view.Name));

                    labelPrefix = view.Name;
                }
                else if (e.SelectedElement is IViewBaseField)
                {
                    IViewBaseField baseField = e.SelectedElement as IViewBaseField;
                    var view = baseField.View;
                    helper.setModelAndLabelFile(metaModelService.GetViewModelInfo(view.Name));

                    labelPrefix = String.Format("{0}_{1}", baseField.View.Name, baseField.Name);
                }
                else if (e.SelectedElement is IFormDesign)
                {
                    IFormDesign formDesign = e.SelectedElement as IFormDesign;
                    var form = formDesign.Form;
                    helper.setModelAndLabelFile(metaModelService.GetFormModelInfo(form.Name));

                    labelPrefix = formDesign.Form.Name;  
                }
                else if (e.SelectedElement is IFormControl)
                {
                    IFormControl formControl = e.SelectedElement as IFormControl;
                    IRootElement rootElement = formControl.RootElement as IRootElement;
                    if (rootElement is IFormExtension)
                    {
                        helper.setModelAndLabelFile(metaModelService.GetFormExtensionModelInfo(rootElement.Name));
                        labelPrefix = String.Format("{0}_{1}", rootElement.Name, formControl.Name);
                        labelPrefix = labelPrefix.Replace(".","_");
                    }
                    else
                    {
                        helper.setModelAndLabelFile(metaModelService.GetFormModelInfo(rootElement.Name));
                        labelPrefix = String.Format("{0}_{1}", rootElement.Name, formControl.Name);
                    }
                }
                else if (e.SelectedElement is ISecurityPrivilege)
                {
                    ISecurityPrivilege securityObject = e.SelectedElement as ISecurityPrivilege;
                    helper.setModelAndLabelFile(metaModelService.GetSecurityPrivilegeModelInfo(securityObject.Name));

                    labelPrefix = String.Format("{0}_{1}", Tags.PrivilegeTag, securityObject.Name);
                }
                else if (e.SelectedElement is ISecurityDuty)
                {
                    ISecurityDuty securityObject = e.SelectedElement as ISecurityDuty;
                    helper.setModelAndLabelFile(metaModelService.GetSecurityDutyModelInfo(securityObject.Name));

                    labelPrefix = String.Format("{0}_{1}", Tags.DutyTag, securityObject.Name);
                }
                else if (e.SelectedElement is ISecurityRole)
                {
                    ISecurityRole securityObject = e.SelectedElement as ISecurityRole;
                    helper.setModelAndLabelFile(metaModelService.GetSecurityRoleModelInfo(securityObject.Name));

                    labelPrefix = String.Format("{0}_{1}", Tags.RoleTag, securityObject.Name);
                }
                else if (e.SelectedElement is IEdtBase)
                {
                    IEdtBase edt = e.SelectedElement as IEdtBase;
                    helper.setModelAndLabelFile(metaModelService.GetExtendedDataTypeModelInfo(edt.Name));

                    labelPrefix = String.Format("{0}_{1}", Tags.EDTTag, edt.Name);           
                }
                else if (e.SelectedElement is IMenuItem)
                {
                    IMenuItem AxMenuItem = e.SelectedElement as IMenuItem;
                    labelPrefix = String.Format("{0}_{1}", Tags.MenuItemTag, AxMenuItem.Name);

                    if (AxMenuItem is IMenuItemAction)
                    {
                        helper.setModelAndLabelFile(metaModelService.GetMenuItemActionModelInfo(AxMenuItem.Name));
                    }
                    else if (AxMenuItem is IMenuItemDisplay)
                    {
                        helper.setModelAndLabelFile(metaModelService.GetMenuItemDisplayModelInfo(AxMenuItem.Name));
                        
                    }
                    else if (AxMenuItem is IMenuItemOutput)
                    {
                        helper.setModelAndLabelFile(metaModelService.GetMenuItemOutputModelInfo(AxMenuItem.Name));
                    }

                }
                else if (e.SelectedElement is IReportDataSetField)
                {
                    IReportDataSetField dataField = e.SelectedElement as IReportDataSetField;

                    helper.setModelAndLabelFile(metaModelService.GetReportModelInfo(dataField.DataSet.Report.Name));

                    labelPrefix = String.Format("{0}_{1}_{2}", dataField.DataSet.Report.Name, dataField.DataSet.Name, dataField.Name);
                }
                helper.createPropertyLabels(e.SelectedElement, labelPrefix);
            }
            catch (Exception ex)
            {
                Microsoft.Dynamics.Framework.Tools.MetaModel.Core.CoreUtility.HandleExceptionWithErrorMessage(ex);

            }
        }
        #endregion
    }
}

