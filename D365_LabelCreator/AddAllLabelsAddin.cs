namespace D365_LabelCreator
{
    using System;
    using System.Linq;
    using System.ComponentModel.Composition;
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Forms;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.Tables;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;
    using Microsoft.Dynamics.AX.Metadata.MetaModel;
    using Microsoft.Dynamics.AX.Metadata.Service;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Automation.BaseTypes;
    using System.Collections;

    /// <summary>
    /// This Addin Will Generate labels for The root element+children elements
    /// </summary>
    [Export(typeof(IDesignerMenu))]
    // This addin will show when user right clicks on a design node.
    // If you need to specify any other element, change this AutomationNodeType value.
    // You can specify multiple DesignerMenuExportMetadata attributes to meet your needs
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(ITable))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IFormDesign))]
    [DesignerMenuExportMetadata(AutomationNodeType = typeof(IBaseEnum))]
    public class DesignerContextMenuAddIn : DesignerMenuBase
    {
        #region Member variables
        private const string addinName = "MMX_AddAllLabelsAddin";
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
                return AddinResources.AddAllLabelsCaption;
            }
        }


        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return DesignerContextMenuAddIn.addinName;
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
            // Dont display Already exists message
            helper.PromptOnDuplicate = false;
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

                    // Loop through each BaseField. Similar logic coulde be added for FieldGroups Ect.
                    foreach(IBaseField baseField in table.BaseFields)
                    {
                        var labelPrefix = String.Format("{0}_{1}", baseField.Table.Name, baseField.Name);
                        helper.createPropertyLabels(baseField, labelPrefix, labelFile);
                    }
                }
                // Is this a Form Design Element?
                else if (e.SelectedElement is IFormDesign)
                {
                    IFormDesign formDesign = e.SelectedElement as IFormDesign;
                    var form = formDesign.Form;

                    modelInfoCollection = metaModelService.GetFormModelInfo(form.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(formDesign, formDesign.Form.Name, labelFile);

                    // Loop through all children FormControls.
                    this.crawlDesignControls(formDesign.VisualChildren, labelFile,form.Name);
                }
                else if (e.SelectedElement is IBaseEnum)
                {
                    IBaseEnum axEnum = e.SelectedElement as IBaseEnum;
                    var labelPrefix = String.Format("{0}_{1}", Tags.EnumTag, axEnum.Name);

                    modelInfoCollection = metaModelService.GetEnumModelInfo(axEnum.Name);
                    AxLabelFile labelFile = helper.GetLabelFile(metaModelProvider, metaModelService, modelInfoCollection);

                    helper.createPropertyLabels(axEnum, labelPrefix, labelFile);
                    // Was unable to locate a way of finding the parent enum of a enumValue. No model could be found without it.
                    // Loop through all values
                    foreach (IBaseEnumValue enumValue in axEnum.BaseEnumValues)
                    {
                        var enumValueLabelPrefix = String.Format("{0}_{1}_{2}", Tags.EnumTag, axEnum.Name, enumValue.Name);
                        helper.createPropertyLabels(enumValue, enumValueLabelPrefix, labelFile);
                    }
                }
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);
            }
        }
        #endregion

        /// <summary>
        /// recursively loop through form controls creating labels 
        /// </summary>
        /// <param name="controls">Instance of <c>IEnumerable</c> containing FormContorls</param>
        /// <param name="labelFile">Instance of <c>AXLabelFile</c> to insert into</param>
        /// <param name="formName"><c>FormName</c> to add as a prefix</param>
        public void crawlDesignControls(IEnumerable controls, AxLabelFile labelFile,string formName)
        {
            foreach(IFormControl control in controls.OfType<IFormControl>())
            {
                var labelPrefix = String.Format("{0}_{1}", formName, control.Name);
                helper.createPropertyLabels(control, labelPrefix, labelFile);
                crawlDesignControls(control.VisualChildren, labelFile, formName);
            }
        }


        
    }
}

