using K2Field.K2NE.SPContentBroker;
using K2Field.K2NE.SPContentBroker.Helpers;
using Microsoft.SharePoint.Client;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Globalization;
using System.Xml;
using Microsoft.SharePoint.Client.Taxonomy;

namespace K2Field.K2NE.SPContentBroker.ServiceObjects
{
    public partial class SPServiceObject : ServiceObjectBase
    {
        public SPServiceObject(SPContentBroker broker) : base(broker)
        {
        }

        public override List<ServiceObject> DescribeServiceObjects()
        {
            List<ServiceObject> SOs = new List<ServiceObject>();
            using (ClientContext context = InitializeContext(base.SiteURL))
            {
                Web spWeb = context.Web;
                ListCollection lists = spWeb.Lists;
                context.Load(lists);
                context.ExecuteQuery();
                //Making this dummy call to load the Micorsoft.Sharepoint.Client.Taxanomy assembly (https://blogs.msdn.microsoft.com/boodablog/2014/07/04/taxonomy-fields-return-as-dictionaries-using-the-client-objcet-model-in-sharepoint-2013/) 
                TaxonomyItem dummy = new TaxonomyItem(context, null);
                foreach (List list in lists)
                {

                    if (list.Hidden == false || (list.Hidden && base.IncludeHiddenLists))
                    {
                        if (list.BaseType == BaseType.DocumentLibrary)
                        {
                            context.Load(list.ContentTypes);
                            context.ExecuteQuery();
                        }

                        ServiceObject so = Helper.CreateServiceObject(list.Title, list.Title, list.Description);

                        so.MetaData.DisplayName = list.Title;
                        so.MetaData.Description = list.Description;
                        if (list.BaseType == BaseType.DocumentLibrary)
                        {
                            so.MetaData.AddServiceElement(Constants.InternalProperties.ServiceFolder, "Document Libraries");
                        }
                        else
                        {
                            so.MetaData.AddServiceElement(Constants.InternalProperties.ServiceFolder, "List Items");
                        }

                        so.MetaData.AddServiceElement(Constants.InternalProperties.ListTitle, list.Title);
                        so.MetaData.AddServiceElement(Constants.InternalProperties.IsFolderEnabled, list.EnableFolderCreation);
                        so.MetaData.AddServiceElement(Constants.InternalProperties.ListBaseType, list.BaseType);
                        if (list.BaseType == BaseType.DocumentLibrary && GetDocumentSetContentType(list.ContentTypes) != null)
                        {
                            so.MetaData.AddServiceElement(Constants.InternalProperties.IsListDocumentSet, true);
                        }
                        else
                        {
                            so.MetaData.AddServiceElement(Constants.InternalProperties.IsListDocumentSet, false);
                        }
                     

                        FieldCollection fields = list.Fields;
                        context.Load(fields);
                        context.ExecuteQuery();
                        foreach (Field f in fields)
                        {
                            if (f.Hidden == false || (f.Hidden == true && base.IncludeHiddenFields))
                            {
                                // We'll use InternalName as the system name and Title as the display name.
                                // See http://blogs.perficient.com/microsoft/2009/04/static-name-vs-internal-name-vs-display-name-in-sharepoint/ for some background

                                // Some fields have no title, so then we just take the internalname.
                                string fieldTitle = f.Title;
                                if (string.IsNullOrEmpty(fieldTitle))
                                {
                                    fieldTitle = f.InternalName;
                                }

                                // Because the field title can be duplicate, we see if it already exists.
                                // If it does, we change the displayname of both existing and newly found property to something unique.
                                // This behaviour can also be forced by the Show Internal Names option.

                                Property existingProp = GetExistingProperty(so, fieldTitle);
                                string displayName = fieldTitle;
                                if (ShowInternalNames)
                                {
                                    displayName = string.Format("{0} ({1})", fieldTitle, f.InternalName);
                                }

                                if (existingProp != null)
                                {
                                    existingProp.MetaData.DisplayName = string.Format("{0} ({1})",
                                        existingProp.MetaData.GetServiceElement<string>(Constants.InternalProperties.Title),
                                        existingProp.MetaData.GetServiceElement<string>(Constants.InternalProperties.InternalName));
                                    displayName = string.Format("{0} ({1})", fieldTitle, f.InternalName);
                                }

                                AddFieldProperty(so, f);
                                FieldType _fieldtype; //We will find the Fieldtype from the MapHelper class (To get the correctoutput for field type Calculated)

                                SoType soType = MapHelper.SPTypeField(f, out _fieldtype);
                                Property prop = Helper.CreateSpecificProperty(f.InternalName, displayName, f.Description, soType);

                                prop.MetaData.AddServiceElement(Constants.InternalProperties.Hidden, f.Hidden);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.Title, fieldTitle);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.InternalName, f.InternalName);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.Id, f.Id);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.ReadOnly, AssignReadonly(f));
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.Required, f.Required);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.FieldTypeKind, _fieldtype);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.SPFieldType, f.TypeAsString);
                                prop.MetaData.AddServiceElement(Constants.InternalProperties.Internal, false);
                                so.Properties.Add(prop);
                            }
                        }

                        AddInputServiceObjectPropertie(so);

                        AddServiceObjectMethods(so);

                        SOs.Add(so);
                    }
                }

                ServiceObject ctSo = Helper.CreateServiceObject("ContentTypes", "Content Types", "Manage the collection of content types, which enables consistent handling of content across sites.");

                ctSo.MetaData.DisplayName = "Content Types";
                ctSo.MetaData.Description = "Manage the collection of content types, which enables consistent handling of content across sites.";
                ctSo.MetaData.AddServiceElement(Constants.InternalProperties.ServiceFolder, "Management");

                ctSo.Properties.Add(Helper.CreateSpecificProperty(Constants.SOProperties.ContentTypeName, "Name", "Name", SoType.Text));
                //so.Properties.Add(Helper.CreateSpecificProperty("SiteURL", "Site URL", "Site URL", SoType.Text));
                ctSo.Properties.Add(Helper.CreateSpecificProperty(Constants.SOProperties.ContentTypeGroup, "Group", "Group", SoType.Text));
                ctSo.Properties.Add(Helper.CreateSpecificProperty(Constants.SOProperties.ContentTypeReadOnly, "ReadOnly", "ReadOnly", SoType.YesNo));
                ctSo.Properties.Add(Helper.CreateSpecificProperty(Constants.SOProperties.ContentTypeHidden, "Hidden", "Hidden", SoType.Text));
                ctSo.Properties.Add(Helper.CreateSpecificProperty(Constants.SOProperties.ContentTypeCount, "Count", "Count", SoType.Number));
                ctSo.Properties.Add(Helper.CreateSpecificProperty(Constants.SOProperties.ContentTypeID, "Content Type ID", "Content Type ID", SoType.Text));

                AddContentTypeServiceObjectMethods(ctSo);

                SOs.Add(ctSo);

                return SOs;
            }
        }

        private void AddContentTypeServiceObjectMethods(ServiceObject ctSo)
        {
            AddGetContentTypeByNameMethod(ctSo);
            AddGetContentTypeByIdMethod(ctSo);
            AddGetContentTypesMethod(ctSo);
            AddGetContentTypesByParentMethod(ctSo);
        }

        public override void Execute()
        {
            switch (ServiceBroker.Service.ServiceObjects[0].Methods[0].Name)
            {
                case Constants.Methods.GetItemById:
                    ExecuteGetItemById();
                    break;
                case Constants.Methods.CreateItem:
                    ExecuteCreateItem();
                    break;
                case Constants.Methods.UpdateItemById:
                    ExecuteUpdateItemById();
                    break;
                case Constants.Methods.DeleteItemById:
                    ExecuteDeleteItemById();
                    break;
                case Constants.Methods.GetItems:
                    ExecuteGetItems();
                    break;
                case Constants.Methods.GetItemByTitle:
                    ExecuteGetItemByTitle();
                    break;
                case Constants.Methods.GetItemByName:
                    ExecuteGetItemByName();
                    break;
                case Constants.Methods.CreateDocument:
                    ExecuteCreateDocument();
                    break;
                case Constants.Methods.DeleteDocumentById:
                    ExecuteDeleteDocumentById();
                    break;
                case Constants.Methods.GetDocumentById:
                    ExecuteGetDocumentById();
                    break;
                case Constants.Methods.GetDocuments:
                    ExecuteGetDocuments();
                    break;
                case Constants.Methods.RenameDocumentById:
                    ExecuteRenameDocumentById();
                    break;
                case Constants.Methods.CopyDocumentByName:
                    ExecuteCopyDocumentByName();
                    break;
                case Constants.Methods.MoveDocumentByName:
                    ExecuteMoveDocumentByName();
                    break;
                case Constants.Methods.CreateFolder:
                    ExecuteCreateFolder();
                    break;
                case Constants.Methods.DeleteFolder:
                    ExecuteDeleteFolder();
                    break;
                case Constants.Methods.RenameFolder:
                    ExecuteRenameFolder();
                    break;
                case Constants.Methods.BreakItemInheritanceById:
                    ExecuteBreakItemInheritanceById();
                    break;
                case Constants.Methods.ResetItemInheritanceById:
                    ExecuteResetItemInheritanceById();
                    break;
                case Constants.Methods.BreakFolderInheritanceByName:
                    ExecuteBreakFolderInheritanceByName();
                    break;
                case Constants.Methods.ResetFolderInheritanceByName:
                    ExecuteResetFolderInheritanceByName();
                    break;
                case Constants.Methods.AddItemPermissionById:
                    ExecuteAddItemPermissionById();
                    break;
                case Constants.Methods.RemoveItemPermissionById:
                    ExecuteRemoveItemPermissionById();
                    break;
                case Constants.Methods.AddFolderPermissionByName:
                    ExecuteAddFolderPermissionByName();
                    break;
                case Constants.Methods.RemoveFolderPermissionByName:
                    ExecuteRemoveFolderPermissionByName();
                    break;
                case Constants.Methods.GetItemPermissionById:
                    ExecuteGetItemPermissionById();
                    break;
                case Constants.Methods.GetFolderPermissionByName:
                    ExecuteGetFolderPermissionByName();
                    break;
                case Constants.Methods.MoveFolder:
                    ExecuteMoveFolder();
                    break;
                case Constants.Methods.CheckInDocumentByName:
                    ExecuteCheckInDocumentByName();
                    break;
                case Constants.Methods.CheckInDocumentById:
                    ExecuteCheckInDocumentById();
                    break;
                case Constants.Methods.CheckOutDocumentByName:
                    ExecuteCheckOutDocumentByName();
                    break;
                case Constants.Methods.CheckOutDocumentById:
                    ExecuteCheckOutDocumentById();
                    break;
                case Constants.Methods.CreateDocumentSet:
                    ExecuteCreateDocSetByName();
                    break;
                case Constants.Methods.UpdateDocumentSetByName:
                    ExecuteUpdateDocSetByName();
                    break;
                case Constants.Methods.GetDocumentSetByName:
                    ExecuteGetDocSetByName();
                    break;
                case Constants.Methods.GetDocumentSets:
                    ExecuteGetDocumentSets();
                    break;
                case Constants.Methods.RenameDocumentSetByName:
                    ExecuteRenameDocSet();
                    break;
                case Constants.Methods.DeleteDocumentSetByName:
                    ExecuteDeleteDocSet();
                    break;

                case Constants.Methods.GetContentTypeByName:
                    ExecuteGetContentTypeByName();
                    break;
                case Constants.Methods.GetContentTypeById:
                    ExecuteGetContentTypeById();
                    break;

                case Constants.Methods.GetContentTypes:
                    ExecuteGetContentTypes();
                    break;

                case Constants.Methods.GetContentTypesByParent:
                    ExecuteGetContentTypesByParent();
                    break;
            }
        }

        private void AddInputServiceObjectPropertie(ServiceObject so)
        {
            //add recursively property
            Property recursivelyNameProperty = NewProperty(Constants.SOProperties.Recursively, Constants.SOProperties.Recursively_DisplayName, true, SoType.YesNo, "Recursively do this operation");
            so.Properties.Add(recursivelyNameProperty);
            //add folder name property
            Property folderNameProperty = NewProperty(Constants.SOProperties.FolderName, Constants.SOProperties.FolderName_DisplayName, true, SoType.Text, "The foldername to apply this operation to");
            so.Properties.Add(folderNameProperty);

            //add linkToItem property
            Property linkToItemProperty = NewProperty(Constants.SOProperties.LinkToItem, Constants.SOProperties.LinkToItem_DisplayName, true, SoType.Text, "Link To Item");
            so.Properties.Add(linkToItemProperty);

            //add userLogins property
            Property userLoginsProperty = NewProperty(Constants.SOProperties.UserLogins, Constants.SOProperties.UserLogins_DisplayName, true, SoType.Text, "User logins separated by semicolon");
            so.Properties.Add(userLoginsProperty);

            //add groupLogins property
            Property groupLoginsProperty = NewProperty(Constants.SOProperties.GroupLogins, Constants.SOProperties.GroupLogins_DisplayName, true, SoType.Text, "Group logins separated by semicolon");
            so.Properties.Add(groupLoginsProperty);

            //add permission property
            Property permissionProperty = NewProperty(Constants.SOProperties.Permission, Constants.SOProperties.Permission_DisplayName, true, SoType.Text, "Role Definition Name");
            so.Properties.Add(permissionProperty);

            //add userOrGroup property
            Property userOrGroupProperty = NewProperty(Constants.SOProperties.UserOrGroup, Constants.SOProperties.UserOrGroup_DisplayName, true, SoType.Text, "User or group login");
            so.Properties.Add(userOrGroupProperty);

            //add DestinationURL property
            Property destinationURLProperty = NewProperty(Constants.SOProperties.DestinationURL, Constants.SOProperties.DestinationURL_DisplayName, true, SoType.Text, "The destination URL");
            so.Properties.Add(destinationURLProperty);

            //add DestinationLibrary property
            Property destinationLibraryProperty = NewProperty(Constants.SOProperties.DestinationLibrary, Constants.SOProperties.DestinationLibrary_DisplayName, true, SoType.Text, "The system name of the list");
            so.Properties.Add(destinationLibraryProperty);

            //add DestinationListLibrary property
            Property destinationListLibraryProperty = NewProperty(Constants.SOProperties.DestinationListLibrary, Constants.SOProperties.DestinationListLibrary_DisplayName, true, SoType.Text, "The system name of the list/library");
            so.Properties.Add(destinationListLibraryProperty);

            //add DestinationFolder property
            Property destinationFolderProperty = NewProperty(Constants.SOProperties.DestinationFolder, Constants.SOProperties.DestinationFolder_DisplayName, true, SoType.Text, "The foldername to apply this operation to");
            so.Properties.Add(destinationFolderProperty);

            if (so.IsDocumentLibrary())
            {
                //add overwriteExistingDocument property
                Property overwriteExistingDocument = NewProperty(Constants.SOProperties.OverwriteExistingDocument,
                    Constants.SOProperties.OverwriteExistingDocument_DisplayName,
                    true, SoType.YesNo, "Overwrite Existing Document");
                so.Properties.Add(overwriteExistingDocument);
                //add newFileName property
                Property newFileNameProperty = NewProperty(Constants.SOProperties.NewFileName, Constants.SOProperties.NewFileName_DisplayName, true, SoType.Text, "New File Name");
                so.Properties.Add(newFileNameProperty);

                //properties used in Check In Functionality
                Property checkInComments = NewProperty(Constants.SOProperties.CheckInComment, Constants.SOProperties.CheckInComment_DisplayName, true, SoType.Memo, "Check In Comments");
                so.Properties.Add(checkInComments);

                Property retainCheckout = NewProperty(Constants.SOProperties.RetainCheckOut, Constants.SOProperties.RetainCheckOut_DisplayName, true, SoType.YesNo, "Retain Checkout");
                so.Properties.Add(retainCheckout);

                Property useCheckedInVersion = NewProperty(Constants.SOProperties.UseCheckedInVersion, Constants.SOProperties.UseCheckedInVersion_DisplayName, true, SoType.YesNo, "Used Checked In version");
                so.Properties.Add(useCheckedInVersion);

                //property use in CheckOut functionality
                Property useCheckedOutVersion = NewProperty(Constants.SOProperties.UseCheckedOutVersion, Constants.SOProperties.UseCheckedOutVersion_DisplayName, true, SoType.YesNo, "Used Checked Out version");
                so.Properties.Add(useCheckedOutVersion);

                if(so.IsDocumentSetLibrary())
                {
                    //property use in CheckOut functionality
                    Property docSetName = NewProperty(Constants.SOProperties.DocSetName, Constants.SOProperties.DocSetName_DisplayName, true, SoType.Text, "Document Set Name");
                    so.Properties.Add(docSetName);

                    Property docSetNewName = NewProperty(Constants.SOProperties.DocSetNewName, Constants.SOProperties.DocSetNewName_DisplayName, true, SoType.Text, "New Document Set Name");
                    so.Properties.Add(docSetNewName);
                }
            }
        }

        private void AddServiceObjectMethods(ServiceObject so)
        {
            //add methods
            AddGetItemByIdMethod(so);
            AddUpdateItemByIdMethod(so);
            AddBreakItemInheritanceByIdMethod(so);
            AddResetItemInheritanceByIdMethod(so);
            AddAddItemPermissionByIdMethod(so);
            AddRemoveItemPermissionByIdMethod(so);
            AddGetItemPermissionByIdMethod(so);
            AddGetItemsMethod(so);

            if (!so.IsDocumentLibrary())
            {
                AddCreateItemMethod(so);
                AddDeleteItemByIdMethod(so);
                AddGetItemByTitleMethod(so);
               
            }
            else
            {
                AddGetItemByNameMethod(so);
                AddCreateDocumentMethod(so);
                AddDeleteDocumentByIdMethod(so);
                AddGetDocumentByIdMethod(so);
                AddGetDocumentsMethod(so);
                AddRenameDocumentByIdMethod(so);
                AddCopyDocumentByNameMethod(so);
                AddMoveDocumentByNameMethod(so);
                AddCheckInDocumentByNameMethod(so);
                AddCheckInDocumentByIdMethod(so);
                AddCheckOutDocumentByIdMethod(so);
                AddCheckOutDocumentByNameMethod(so);

                if (so.IsDocumentSetLibrary())
                {
                    AddCreateDocumentSetByNameMethod(so);
                    AddUpdateDocumentSetByNameMethod(so);
                    AddGetDocSetMethod(so);
                    AddGetDocumentSetsMethod(so);
                    AddRenameDocumentSetMethod(so);
                    AddDeleteDocumentSetMethod(so);
                }

            }

            if (so.IsSPFolderEnabled())
            {
                AddCreateFolderMethod(so);
                AddDeleteFolderMethod(so);
                AddRenameFolderMethod(so);
                AddBreakFolderInheritanceByNameMethod(so);
                AddResetFolderInheritanceByNameMethod(so);
                AddAddFolderPermissionByNameMethod(so);
                AddRemoveFolderPermissionByNameMethod(so);
                AddGetFolderPermissionByIdMethod(so);
                AddMoveFolderMethod(so);
            }
        }

        #region Property methods

        //This method is added so that we can make property false for fieldtype Lookup and User as in that case only one property with ID should be editable.
        private bool AssignReadonly(Field f)
        {
            switch (f.FieldTypeKind)
            {
                case FieldType.Lookup:
                case FieldType.User:
                    return true;
                default:
                    return f.ReadOnlyField;
            }
        }

        private void AddFieldProperty(ServiceObject so, Field fd)
        {

            switch (fd.FieldTypeKind)
            {
                case FieldType.Lookup:
                case FieldType.User:
                    so.Properties.Add(CreateFieldProperty(Constants.InternalProperties.Suffix_ID, SoType.Text, fd, fd.ReadOnlyField));
                    so.Properties.Add(CreateFieldProperty(Constants.InternalProperties.Suffix_Value, SoType.Text, fd, true));
                    break;
                case FieldType.File:

                    Property fileNameProp = NewProperty(Constants.SOProperties.FileName, Constants.SOProperties.FileName_DisplayName, true, SoType.Text,
                        Constants.SOProperties.FileName_DisplayName);
                    
                    so.Properties.Add(fileNameProp);
                    break;
                case FieldType.Invalid:
                    if (string.Compare(fd.TypeAsString, Constants.SharePointProperties.TaxanomyFieldType) == 0 || string.Compare(fd.TypeAsString, Constants.SharePointProperties.TaxanomyFieldTypeMulti) == 0)
                    {
                        so.Properties.Add(CreateFieldProperty(Constants.InternalProperties.Suffix_Value, SoType.Text, fd, true));
                    }
                    break;
            }
        }

        private Property CreateFieldProperty(string nameSuffix, SoType soType, Field fd, bool readOnly)
        {
            //Modifying the field Title with the nameSuffix. Resolution of TFS Bug # 3265
            string title = string.Format("{0} {1}", fd.Title, nameSuffix);
            //Modifying the field display name with internal name and nameSuffix. Resolution of TFS Bug # 3265
            string displayName = string.Format("{0} ({1})", title, fd.InternalName);
            string internalName = string.Concat(fd.InternalName, nameSuffix);
     
            Property prop = Helper.CreateSpecificProperty(internalName, displayName, fd.Description, soType);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Hidden, fd.Hidden);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Title, title);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.InternalName, fd.InternalName);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Id, fd.Id);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.ReadOnly, readOnly);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Required, fd.Required);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.FieldTypeKind, fd.FieldTypeKind);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, fd.TypeAsString);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
            return prop;
        }

        private Property NewProperty(string internalName, string DisplayName, bool internalProp, SoType soType, string description)
        {
            Property prop = Helper.CreateSpecificProperty(internalName, DisplayName, description, soType);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.InternalName, internalName);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Title, DisplayName);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, internalProp);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.ReadOnly, false);
            return prop;
        }

        private Property GetExistingProperty(ServiceObject so, string title)
        {
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.Title), title, true) == 0)
                {
                    return prop;
                }
            }
            return null;
        }

        #endregion

        private string GetSiteURL()
        {
            string siteURL = base.SiteURL;
            if (base.IsDynamicSiteURL)
            {
                siteURL = base.GetStringParameter(Constants.InternalProperties.SiteUrl, true);
            }
            return siteURL;
        }

        private ContentType GetDocumentSetContentType(ContentTypeCollection contentColl)
        {
            foreach (ContentType ct in contentColl)
            {
                if (ct.Id.StringValue.StartsWith(Constants.SharePointProperties.DocSetContentType))
                {
                    return ct;
                }
            }
            return null;
        }
    }
}
