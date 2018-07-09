using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using Microsoft.SharePoint.Client;

namespace K2Field.K2NE.SPContentBroker.Helpers
{
    /// <summary>
    /// The SP Content broker uses a property for every field/column on a list/library. Some properties are added manually to 
    /// provide addditional information. Some properties are also static inside sharepoint, or added for special methods. 
    /// 
    /// To make the code a little cleaner in understanding which property is what, we use these extension methods to make our code more readale.
    /// There is usage of metadata service elements, which are discovered and added in the DiscoverServiceObject method and there are properties that we can
    /// just check on their name.
    /// </summary>
    internal static class PropertyExtensions
    {
        /// <summary>
        /// SMO property is read only
        /// </summary>
        public static bool IsReadOnly(this Property property)
        {
            return property.MetaData.GetServiceElement<Boolean>(Constants.InternalProperties.ReadOnly);
        }

        /// <summary>
        /// SMO property is internal
        /// </summary>
        public static bool IsInternal(this Property property)
        {
            return property.MetaData.GetServiceElement<Boolean>(Constants.InternalProperties.Internal);
        }

        /// <summary>
        /// SMO property is required
        /// </summary>
        public static bool IsRequired(this Property property)
        {
            return property.MetaData.GetServiceElement<Boolean>(Constants.InternalProperties.Required);
        }

        /// <summary>
        /// SMO property is an additional property of SharePoint file (FileName of File Link property)
        /// </summary>
        public static bool IsFileName(this Property property)
        {
            return (string.Compare(property.Name, Constants.SOProperties.FileName) == 0);
        }

        /// <summary>
        /// SMO property is an additional property of SharePoint file (FileName of File Link property)
        /// </summary>
        public static bool IsDocSetName(this Property property)
        {
            return (string.Compare(property.Name, Constants.SOProperties.DocSetName) == 0);
        }

        /// <summary>
        /// SMO property is an additional property of SharePoint file (FileName of File Link property)
        /// </summary>
        public static bool IsDocSetContentType(this Property property)
        {
            return (string.Compare(property.Name, Constants.SOProperties.ContentType) == 0);
        }

        /// <summary>
        /// Helps to determine if the property is the specific File property in SharePoint.
        /// </summary>
        /// <param name="property"></param>
        /// <returns>True if the property is a file, false if it's not.</returns>
        public static bool IsFile(this Property property)
        {
            if (property.IsInternal())
            {
                return false;
            }
            
            FieldType fieldType = property.MetaData.GetServiceElement<FieldType>(Constants.InternalProperties.FieldTypeKind);
            return (fieldType == FieldType.File);
        }

        /// <summary>
        /// Check LinkToItem proeprty
        /// </summary>
        public static bool IsLinkToItem(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.LinkToItem) == 0;
        }

        /// <summary>
        /// SMO property is a folder name
        /// </summary>
        public static bool IsFolderName(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.FolderName) == 0;
        }

        /// <summary>
        /// SMO property is a recursive indicator
        /// </summary>
        public static bool IsRecursivelyName(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.Recursively) == 0;
        }

        /// <summary>
        /// SMO property is an indicator for overwrite exesting document (for upload document)
        /// </summary>
        public static bool IsOverwriteExistingDocument(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.OverwriteExistingDocument) == 0;
        }

        /// <summary>
        /// SMO property is a NewFileName (for rename document)
        /// </summary>
        public static bool IsNewFileName(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.NewFileName) == 0;
        }

        /// <summary>
        /// SMO property is a DestinationFolder
        /// </summary>
        public static bool IsDestinationFolder(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.DestinationFolder) == 0;
        }

        /// <summary>
        /// SMO property is a DestinationURL (for copy/move methods)
        /// </summary>
        public static bool IsDestinationURL(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.DestinationURL) == 0;
        }

        /// <summary>
        /// SMO property is a DestinationListLibrary  (for copy/move methods)
        /// </summary>
        public static bool IsDestinationListLibrary(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.DestinationListLibrary) == 0;
        }

        /// <summary>
        /// SMO property is a DestinationLibrary  (for copy/move methods)
        /// </summary>
        public static bool IsDestinationLibrary(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.DestinationLibrary) == 0;
        }

        /// <summary>
        /// SMO property is a Permission (for permission methods)
        /// </summary>
        public static bool IsPermissionColumn(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.Permission) == 0;
        }

        /// <summary>
        /// SMO property is a UserOrGroup (for permission methods)
        /// </summary>
        public static bool IsUserOrGroup(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.UserOrGroup) == 0;
        }

        /// <summary>
        /// SMO property is a UserLogins (for permission methods)
        /// </summary>
        public static bool IsUserLogins(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.UserLogins) == 0;
        }

        /// <summary>
        /// SMO property is a GroupLogins (for permission methods)
        /// </summary>
        public static bool IsGroupLogins(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.GroupLogins) == 0;
        }

        /// <summary>
        /// SMO property is for CheckInComments
        /// </summary>
        public static bool IsCheckInComments(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.CheckInComment) == 0;
        }

        /// <summary>
        /// SMO property is for RetainCheckout
        /// </summary>
        public static bool IsRetainCheckout(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.RetainCheckOut) == 0;
        }

        /// <summary>
        /// SMO property is for UseCheckedInVersion
        /// </summary>
        public static bool IsUseCheckedInVersion(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.UseCheckedInVersion) == 0;
        }

        /// <summary>
        /// SMO property is for UseCheckedOutVersion
        /// </summary>
        public static bool IsUseCheckedOutVersion(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.UseCheckedOutVersion) == 0;
        }

    
        /// <summary>
        /// SMO property is for DocSetName
        /// </summary>
        public static bool IsNewDocumentSetName(this Property property)
        {
            return string.Compare(property.Name, Constants.SOProperties.DocSetNewName) == 0;
        }
    }
}
