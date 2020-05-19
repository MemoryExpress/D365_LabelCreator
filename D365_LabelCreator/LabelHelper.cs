using EnvDTE;
using EnvDTE80;
using Microsoft.Dynamics.AX.Metadata.MetaModel;
using Microsoft.Dynamics.AX.Metadata.Service;
using Microsoft.Dynamics.Framework.Tools.Extensibility;
using Microsoft.Dynamics.Framework.Tools.Labels;
using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Dynamics.AX.Server.Core.Service;

namespace D365_LabelCreator
{
    /// <summary>
    /// Helper class containing all the utilities to create labels
    /// </summary>
    class LabelHelper
    {
        ModelInfoCollection modelInfoCollection = null;
        IMetaModelService metaModelService = null;
        // Label file to write into
        AxLabelFile currentLabelFile;
        // Get the metamodel provider
        IMetaModelProviders metaModelProvider = getMetaModelService();
       


        bool promptOnDuplicate = true;

        //Factory and controller for adding and saving labels
        LabelControllerFactory factory = new LabelControllerFactory();
        LabelEditorController labelController;

        public static IMetaModelProviders getMetaModelService()
        {
            IMetaModelProviders modelProviders;
            try
            {
                modelProviders = ServiceLocator.GetService<IMetaModelProviders>();
            }
            catch
            {
                modelProviders = new ExtensibilityService();
                ServiceLocator.RegisterService<IMetaModelProviders>(modelProviders);
            }

            if (modelProviders == null)
            {
                modelProviders = new ExtensibilityService();
                ServiceLocator.RegisterService<IMetaModelProviders>(modelProviders);
            }

            return modelProviders;
        }

        public LabelHelper()
        {
            if (metaModelProvider != null)
            {
                // Get the metamodel service
                MetaModelService = metaModelProvider.CurrentMetaModelService;
            }
            else
            {
                CoreUtility.DisplayError("Unable to obtain MetaModelProvider");
            }
        }

        /// <summary>
        /// MetaModelService to interact with metadata
        /// </summary>
        public IMetaModelService MetaModelService
        {
            get
            {
                return metaModelService;
            }

            set
            {
                metaModelService = value;
            }
        }

        /// <summary>
        /// If true, Display Warning on already existing labelId's. True by default.
        /// </summary>
        public bool PromptOnDuplicate
        {
            get
            {
                return promptOnDuplicate;
            }

            set
            {
                promptOnDuplicate = value;
            }
        }

        /// <summary>
        /// Set the model info and retreive it's label file.
        /// </summary>
        /// <param name="modelInfo">Instance of <c>ModelInfoCollection</c> to retreive label file for</param>
        public void setModelAndLabelFile(ModelInfoCollection modelInfo, AxLabelFile labelFile = null)
        {
            modelInfoCollection = modelInfo;
            if (labelFile == null)
            {
                currentLabelFile = this.GetLabelFile();
            }
            else
            {
                currentLabelFile = labelFile;
            }
        }

        /// <summary>
        /// Check for existing label, if none exists create ie
        /// </summary>
        /// <param name="labelText"><c>String</c> Value for the label</param>
        /// <param name="elementName"><c>String</c> value thats is the prefix and main portion of the labelID</param>
        /// <param name="propertyName"><c>String</c> value that is the PropertyName which becomes the Suffix</param>
        /// <param name="labelFile">Instance of <c>AXLabelFile</c> to insert into</param>
        /// <returns></returns>
        private String FindOrCreateLabel(String labelText, String elementName, String propertyName,bool delaySave = false)
        {
            string newLabelId = String.Empty;

            if(labelController == null)
            {
                labelController = factory.GetOrCreateLabelController(currentLabelFile, LabelHelper.Context);
            }
            // Don't bother if the string is empty
            if (String.IsNullOrEmpty(labelText) == false)
            {
                // Construct a label id that will be unique
                string labelId = $"{elementName}_{propertyName}";

                // Make sure the label doesn't exist.
                // What will you do if it does?
                if (labelController.Exists(labelId) == false)
                {
                    labelController.Insert(labelId, labelText, String.Empty);
                    if (!delaySave)
                    {
                        labelController.Save();
                    }
                    // Construct a label reference to go into the label property
                    newLabelId = $"@{currentLabelFile.LabelFileId}:{labelId}";
                }
                else if(promptOnDuplicate)
                {
                    CoreUtility.DisplayWarning($"Label: {labelId} already exists");
                }
            }

            return newLabelId;
        }

        #region ControlProperties
        /// <summary>
        /// Check all properties and create labels for them if needed
        /// </summary>
        /// <param name="ob">Instance of an <c>Object</c> to check for preoprties</param>
        /// <param name="labelPrefix"><c>String</c> Prefix to add to the label</param>
        /// <param name="labelFile">Instance of <c>AXLabelFile</c> to insert into</param>
        public void createPropertyLabels(Object ob, String labelPrefix)
        {
            if(currentLabelFile == null)
            {
                CoreUtility.DisplayError("Labelfile not found");
                return;
            }

            labelController = factory.GetOrCreateLabelController(currentLabelFile, LabelHelper.Context);

            var type = ob.GetType();

            var LabelProperty = getLabelProperty(type);
            if (LabelProperty != null)
            {
                var label = LabelProperty.GetValue(ob).ToString();

                if (this.IsValidLabelId(label) == false)
                {
                    var labelid = this.FindOrCreateLabel(label, labelPrefix, Tags.LabelTag,true);
                    LabelProperty.SetValue(ob, labelid != string.Empty ? labelid : label);
                }
            }

            var helpTextProperty = getHelpTextProperty(type);
            if (helpTextProperty != null)
            {
                var label = helpTextProperty.GetValue(ob).ToString();

                if (this.IsValidLabelId(label) == false)
                {
                    var labelid = this.FindOrCreateLabel(label, labelPrefix, Tags.HelpTextTag,true);
                    helpTextProperty.SetValue(ob, labelid != string.Empty ? labelid : label);
                }
            }

            var captionProperty = getCaptionProperty(type);
            if (captionProperty != null)
            {
                var label = captionProperty.GetValue(ob).ToString();

                if (this.IsValidLabelId(label) == false)
                {
                    var labelid = this.FindOrCreateLabel(label, labelPrefix, Tags.CaptionTag, true);
                    captionProperty.SetValue(ob, labelid != string.Empty ? labelid : label);
                }
            }

            var textProperty = getTextProperty(type);
            if (textProperty != null)
            {
                var label = textProperty.GetValue(ob).ToString();

                if (this.IsValidLabelId(label) == false)
                {
                    var labelid = this.FindOrCreateLabel(label, labelPrefix, Tags.TextTag, true);
                    textProperty.SetValue(ob, labelid != string.Empty ? labelid : label);
                }
            }

            var DescriptionProperty = getDescriptionProperty(type);
            if (DescriptionProperty != null)
            {
                var label = DescriptionProperty.GetValue(ob).ToString();

                if (this.IsValidLabelId(label) == false)
                {
                    var labelid = this.FindOrCreateLabel(label, labelPrefix, Tags.DescriptionTag, true);
                    DescriptionProperty.SetValue(ob, labelid != string.Empty ? labelid : label);
                }
            }

            var developerDocumentationProperty = getDeveloperDocumentationProperty(type);
            if (developerDocumentationProperty != null)
            {
                var label = developerDocumentationProperty.GetValue(ob).ToString();

                if (this.IsValidLabelId(label) == false)
                {
                    var labelid = this.FindOrCreateLabel(label, labelPrefix, Tags.DeveloperDocumentationTag, true);
                    developerDocumentationProperty.SetValue(ob, labelid != string.Empty ? labelid : label);
                }
            }

            labelController.Save();
        }


        #endregion

        #region Property Getters
        private System.Reflection.PropertyInfo getLabelProperty(Type type)
        {
            return type.GetProperty("Label", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
        }
        private System.Reflection.PropertyInfo getHelpTextProperty(Type type)
        {
            return type.GetProperty("HelpText", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
        }
        private System.Reflection.PropertyInfo getCaptionProperty(Type type)
        {
            return type.GetProperty("Caption", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
        }
        private System.Reflection.PropertyInfo getTextProperty(Type type)
        {
            return type.GetProperty("Text", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
        }
        private System.Reflection.PropertyInfo getDescriptionProperty(Type type)
        {
            return type.GetProperty("Description", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
        }
        private System.Reflection.PropertyInfo getDeveloperDocumentationProperty(Type type)
        {
            return type.GetProperty("DeveloperDocumentation", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
        }

        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="metaModelProviders">Instance of <c>IMetaModelProvider</c></param>
        /// <param name="metaModelService">Instance of <c>IMetaModelService</c></param>
        /// <param name="modelInfoCollection">Instance of <c>ModelInfoCollection</c></param>
        /// <returns>Instance of <c>AXLabelFile</c></returns>
        public AxLabelFile GetLabelFile()
        {
            // Choose the first model in the collection
            ModelInfo modelInfo = ((System.Collections.ObjectModel.Collection<ModelInfo>)modelInfoCollection)[0];

            // Construct a ModelLoadInfo
            ModelLoadInfo modelLoadInfo = new ModelLoadInfo
            {
                Id = modelInfo.Id,
                Layer = modelInfo.Layer,
            };

            // Get the list of label files from that model
            IList<String> labelFileNames = metaModelProvider.CurrentMetadataProvider.LabelFiles.ListObjectsForModel(modelInfo.Name);

            // Choose the first
            // What happens if there is no label file?
            AxLabelFile labelFile = MetaModelService.GetLabelFile(labelFileNames[0], modelLoadInfo);

            return labelFile;
        }

        /// <summary>
        /// Regex to determine if its a new Label
        /// </summary>
        private static Regex newLabelMatcher =
            new Regex("\\A(?<AtSign>\\@)(?<LabelFileId>[a-zA-Z]\\w*):(?<LabelId>[a-zA-Z]\\w*)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>
        /// Regex to determine if its a legacy label
        /// </summary>
        private static Regex legacyLabelMatcher =
            new Regex("\\A(?<LabelId>(?<AtSign>\\@)(?<LabelFileId>[a-zA-Z][a-zA-Z][a-zA-Z])\\d+)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        //These expressions will be used by the following methods.
        /// <summary>
        /// Determine if this is a valid label
        /// </summary>
        /// <param name="labelId"><c>String</c> Labelid to validate<</param>
        /// <returns>True if valid, otherwise false</returns>
        private Boolean IsValidLabelId(String labelId)
        {
            bool result = false;

            if (String.IsNullOrEmpty(labelId) == false)
            {
                result = LabelHelper.IsLegacyLabelId(labelId);
                if (result == false)
                {
                    result = (LabelHelper.newLabelMatcher.Matches(labelId).Count > 0);
                }
            }

            return result;
        }

        /// <summary>
        /// Determine if this is a valid Legacy label
        /// </summary>
        /// <param name="labelId"><c>String</c> Labelid to validate<</param>
        /// <returns>True if valid, otherwise false</returns>
        private static Boolean IsLegacyLabelId(String labelId)
        {
            bool result = false;

            if (String.IsNullOrEmpty(labelId) == false)
            {
                result = (LabelHelper.legacyLabelMatcher.Matches(labelId).Count > 0);
            }

            return result;
        }

        static DTE2 MyDte => CoreUtility.ServiceProvider.GetService(typeof(DTE)) as DTE2;
        static VSApplicationContext Context => new VSApplicationContext(MyDte.DTE);
    }
}
