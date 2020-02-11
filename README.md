# MEMX_D365_LabelCreator
Visual Studio Add-in for easy creation of Labels

*Thanks and Credit to Stoneridge Software and especially Mark Nelson for the begining code that started this*
*it hardly resembles where it started but heres the link to their post regarding automated Label creation for 365.*
https://stoneridgesoftware.com/learn-to-use-a-label-creator-add-in-extension-in-dynamics-365-for-finance-operations/

**Install**
References to some library may need to be added.
*DTES can be found here*
C:\Program Files (x86)\Common Files\Microsoft Shared\MSEnv\PublicAssemblies\
*Framwork Tools can be found here*
C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\Extensions\
*MetaData and Automation Objects here*
C:\AOSService\PackagesLocalDirectory\bin\

Once References are Added, Build the solution and copy the .DLL out of the Bin folder to here.
C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\Extensions\5hutepyf.xp2\AddinExtensions\.

**Usage**

Right click eligable Elements in the D365 Designer.
Two Options Currently exist
*Create Labels*
This will create labels for a single Elements
*Create All Labels*
This will create labels for the root element and all children. Currently works with 
	*Tables (Table +Fields, no Field Groups)*
	*FormsDesign (Form design plus all child FormControls)*
	*Enum (Enum and all values. Create labels for Enum Values independantly dosn't currently exist*
	
**Modification**
Currently the solution is designed for Simple Label formating.
The Tags.cs file contains Constants for all the Tags we use. These can be adapted to fit your naming convention if differnt.
The LabelHelper takes a simple label string that forms the first part of the labelid, then a property string which results in the end of the labelid.

so in the case of Enums label property.
Enum_ElementName_Label. (Enum_ElementName) <-Prefix using a String formated Tag+ElementName, (_Label) being the property Type.

in the Enum code block, just change the string formating to accomplish a different prefix.
