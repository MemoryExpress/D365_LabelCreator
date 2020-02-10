namespace D365_LabelCreator
{
    using System;
    using System.ComponentModel.Composition;
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Forms;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Tables;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;
    using Microsoft.Dynamics.AX.Metadata.MetaModel;
    using Microsoft.Dynamics.AX.Metadata.Service;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.BaseTypes;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Menus;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Security;

    /// <summary>
    /// TODO: Say a few words about what your AddIn is going to do
    /// </summary>
    [Export(typeof(IDesignerMenu))]
    // TODO: This addin will show when user right clicks on a form root node or table root node. 
    // If you need to specify any other element, change this AutomationNodeType value.
    // You can specify multiple DesignerMenuExportMetadata attributes to meet your needs
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ITable))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IBaseField))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IFormDesign))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IFormControl))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ISecurityPrivilege))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ISecurityDuty))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ISecurityRole))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IEdtBase))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IMenuItem))]

    public class AddLabelAddin : DesignerMenuBase
    {
        #region Member variables
        private const string addinName = "MMX_AddLabelAddin";
        private LabelHelper helper = new LabelHelper();
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
                ModelInfoCollection modelInfoCollection = null;

                IMetaModelService metaModelService = null;
                // Get the metamodel provider
                IMetaModelProviders metaModelProvider = ServiceLocator.GetService(typeof(IMetaModelProviders)) as IMetaModelProviders;

                if (metaModelProvider != null)
                {
                    // Get the metamodel service
                    metaModelService = metaModelProvider.CurrentMetaModelService;
                }
                
                // Is the selected element a table?
                if (e.SelectedElement is ITable)
                {
                    ITable table = e.SelectedElement as ITable;

                    modelInfoCollection = metaModelService.GetTableModelInfo(table.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(table, table.Name, labelFile);
                }
                else if (e.SelectedElement is IBaseField)
                {
                    IBaseField baseField = e.SelectedElement as IBaseField;
                    var table = baseField.Table;
                    var labelPrefix = String.Format("{0}_{1}", baseField.Table.Name, baseField.Name);

                    modelInfoCollection = metaModelService.GetTableModelInfo(table.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(baseField, labelPrefix, labelFile);
                }
                else if (e.SelectedElement is IFormDesign)
                {
                    IFormDesign formDesign = e.SelectedElement as IFormDesign;
                    var form = formDesign.Form;

                    modelInfoCollection = metaModelService.GetFormModelInfo(form.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(formDesign, formDesign.Form.Name, labelFile);
                }
                else if (e.SelectedElement is IFormControl)
                {
                    IFormControl formControl = e.SelectedElement as IFormControl;
                    var form = formControl.RootElement as IForm;

                    modelInfoCollection = metaModelService.GetFormModelInfo(form.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    var labelPrefix = String.Format("{0}_{1}", form.Name, formControl.Name);
                    helper.createPropertyLabels(formControl, labelPrefix, labelFile);
                }
                else if (e.SelectedElement is ISecurityPrivilege)
                {
                    ISecurityPrivilege securityObject = e.SelectedElement as ISecurityPrivilege;
                    var labelPrefix = String.Format("{0}_{1}", Tags.PrivilegeTag, securityObject.Name);

                    modelInfoCollection = metaModelService.GetSecurityPrivilegeModelInfo(securityObject.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(securityObject, labelPrefix, labelFile);
                }
                else if (e.SelectedElement is ISecurityDuty)
                {
                    ISecurityDuty securityObject = e.SelectedElement as ISecurityDuty;
                    var labelPrefix = String.Format("{0}_{1}", Tags.DutyTag, securityObject.Name);

                    modelInfoCollection = metaModelService.GetSecurityDutyModelInfo(securityObject.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(securityObject, labelPrefix, labelFile);
                }
                else if (e.SelectedElement is ISecurityRole)
                {
                    ISecurityRole securityObject = e.SelectedElement as ISecurityRole;
                    var labelPrefix = String.Format("{0}_{1}", Tags.RoleTag, securityObject.Name);

                    modelInfoCollection = metaModelService.GetSecurityRoleModelInfo(securityObject.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(securityObject, labelPrefix, labelFile);
                }
                else if (e.SelectedElement is IEdtBase)
                {
                    IEdtBase edt = e.SelectedElement as IEdtBase;
                    var labelPrefix = String.Format("{0}_{1}", Tags.EDTTag, edt.Name);

                    modelInfoCollection = metaModelService.GetExtendedDataTypeModelInfo(edt.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(edt, labelPrefix, labelFile);
                }
                else if (e.SelectedElement is IMenuItem)
                {
                    IMenuItem AxMenuItem = e.SelectedElement as IMenuItem;
                    var labelPrefix = String.Format("{0}_{1}", Tags.MenuItemTag, AxMenuItem.Name);
                    if (AxMenuItem is IMenuItemAction)
                    {
                        modelInfoCollection = metaModelService.GetMenuItemActionModelInfo(AxMenuItem.Name);
                    }
                    else if (AxMenuItem is IMenuItemDisplay)
                    {
                        modelInfoCollection = metaModelService.GetMenuItemDisplayModelInfo(AxMenuItem.Name);
                    }
                    else if (AxMenuItem is IMenuItemOutput)
                    {
                        modelInfoCollection = metaModelService.GetMenuItemOutputModelInfo(AxMenuItem.Name);
                    }

                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(AxMenuItem, labelPrefix, labelFile);
                }
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);

            }
        }
        #endregion
    }
}

