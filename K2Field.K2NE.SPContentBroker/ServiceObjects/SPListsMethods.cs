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

namespace K2Field.K2NE.SPContentBroker
{
    public class SPLists : ServiceObjectBase
    {
        public SPLists(SPContentBroker broker)
            : base(broker)
        {

        }

        //SPList is configured just to create service object of SP Lists
        public override List<ServiceObject> DescribeServiceObjects()
        {
            List<ServiceObject> SOs = new List<ServiceObject>();
            using (ClientContext context = InitializeContext(base.SiteURL))
            {
                Web spWeb = context.Web;
                ListCollection lists = spWeb.Lists;
                context.Load(lists);

                context.ExecuteQuery();
                foreach (List list in lists)
                {

                    if (list.Hidden == false || (list.Hidden && base.IncludeHiddenLists))
                    {

                        ServiceObject so = Helper.CreateServiceObject(list.Title, list.Title, list.Description);

                        so.MetaData.DisplayName = list.Title;
                        so.MetaData.Description = list.Description;
                        if (list.BaseType == BaseType.DocumentLibrary)
                        {
                            so.MetaData.ServiceProperties.Add(Constants.InternalProperties.ServiceFolder, "Document Libraries");
                        }
                        else
                        {
                            so.MetaData.ServiceProperties.Add(Constants.InternalProperties.ServiceFolder, "List Items");
                        }

                        so.MetaData.ServiceProperties.Add(Constants.InternalProperties.ListId, list.Id);
                        so.MetaData.ServiceProperties.Add(Constants.InternalProperties.ListTitle, list.Title);
                        so.MetaData.ServiceProperties.Add(Constants.InternalProperties.IsFolderEnabled, list.EnableFolderCreation);
                        so.MetaData.ServiceProperties.Add(Constants.InternalProperties.ListBaseType, list.BaseType);


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

                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Hidden, f.Hidden);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Title, fieldTitle);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.InternalName, f.InternalName);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Id, f.Id);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.ReadOnly, AssignReadonly(f));
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Required, f.Required);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.FieldTypeKind, _fieldtype);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, f.TypeAsString);
                                prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
                                so.Properties.Add(prop);
                            }
                        }

                        so.Properties.Add(NewProperty(Constants.SOProperties.DestinationURL, Constants.SOProperties.DestinationURL_DisplayName, true, SoType.Text, "The destination URL"));
                        so.Properties.Add(NewProperty(Constants.SOProperties.DestinationListTitle, Constants.SOProperties.DestinationListTitle_DisplayName, true, SoType.Text, "The system name of the list"));
                        so.Properties.Add(NewProperty(Constants.SOProperties.DestinationFolder, Constants.SOProperties.DestinationFolder_DisplayName, true, SoType.Text, "The foldername to apply this operation to"));
                        so.Properties.Add(NewProperty(Constants.SOProperties.FolderName, Constants.SOProperties.FolderName, true, SoType.Text, "The foldername to apply this operation to"));
                        so.Properties.Add(NewProperty(Constants.SOProperties.Recursively, Constants.SOProperties.Recursively, true, SoType.YesNo, "Recursively do this operation"));


                        AddGetByIdMethod(so);
                        AddCreateMethod(so);
                        AddDeleteMethod(so);
                        AddGetListMethod(so);
                        AddUpdateMethod(so);
                        if (so.IsSPFolderEnabled())
                        {
                            AddCreateFolderMethod(so);
                            AddDeleteFolderMethod(so);
                        }

                        //Non library methods
                        if (! so.IsDocumentLibrary())
                        {
                            AddMoveListItemMethod(so);
                            AddCopyListItemMethod(so);
                        }

                        SOs.Add(so);

                    }
                }

                return SOs;
            }
        }


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
                    string fileNameString = string.Format("{0}{1}", fd.InternalName, Constants.SOProperties.FileNameSuffix);
                    string fileURLString = string.Format("{0}{1}", fd.InternalName, Constants.SOProperties.FileURLSuffix);

                    Property fileNameProp = NewProperty(fileNameString, fileNameString, true, SoType.Text,
                        string.Format("{0} {1}", fd.InternalName, Constants.SOProperties.FileNameSuffix));
                    fileNameProp.MetaData.ServiceProperties.Add(Constants.InternalProperties.IsFileChild, true);

                    Property fileURLProp = NewProperty(fileURLString, fileURLString, true, SoType.Text,
                        string.Format("{0} {1}", fd.InternalName, Constants.SOProperties.FileURLSuffix));
                    fileURLProp.MetaData.ServiceProperties.Add(Constants.InternalProperties.IsFileChild, true);

                    so.Properties.Add(fileNameProp);
                    so.Properties.Add(fileURLProp);
                    break;
            }
        }

        private Property CreateFieldProperty(string nameSuffix, SoType soType, Field fd, bool readOnly)
        {
            string displayName = string.Format("{0} ({1})", fd.Title, nameSuffix);
            string internalName = string.Concat(fd.InternalName, nameSuffix);
            Property prop = Helper.CreateSpecificProperty(internalName, displayName, fd.Description, soType);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Hidden, fd.Hidden);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Title, fd.Title);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.InternalName, fd.InternalName);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Id, fd.Id);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.ReadOnly, readOnly);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Required, fd.Required);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.FieldTypeKind, fd.FieldTypeKind);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, fd.TypeAsString);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
            return prop;
        }

        private Property NewProperty(string internalName, string DisplayName, bool _internal, SoType soType, string description)
        {
            Property prop = Helper.CreateSpecificProperty(internalName, DisplayName, description, soType);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.InternalName, internalName);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Title, DisplayName);
            prop.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, _internal);
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
        private void AddGetByIdMethod(ServiceObject so)
        {
            Method mRead = Helper.CreateMethod(Constants.Methods.GetItembyId, "Retrieve one item by it's ID", MethodType.Read);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mRead);
            }
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mRead.InputProperties.Add(prop);
                    mRead.Validation.RequiredProperties.Add(prop);
                }

                if (!prop.IsInternal())
                {
                    mRead.ReturnProperties.Add(prop);

                }
                if (so.IsDocumentLibrary() && prop.IsFileName())
                {
                    mRead.ReturnProperties.Add(prop);
                }

            }


            so.Methods.Add(mRead);
        }

        private void AddCreateMethod(ServiceObject so)
        {
            Method mCreate = Helper.CreateMethod(Constants.Methods.CreatelistItem, "Create a list item", MethodType.Create);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCreate);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0 && !prop.IsInternal())
                {
                    mCreate.ReturnProperties.Add(prop);
                }

                if (!prop.IsReadOnly() && !prop.IsInternal())
                {
                    mCreate.InputProperties.Add(prop);
                    if (prop.IsRequired())
                    {
                        mCreate.Validation.RequiredProperties.Add(prop);
                    }
                }
                if (so.IsSPFolderEnabled() == true &&
                    string.Compare(prop.Name, Constants.SOProperties.FolderName, true) == 0)
                {
                    mCreate.InputProperties.Add(prop);
                }
            }

            so.Methods.Add(mCreate);
        }


        private void AddUpdateMethod(ServiceObject so)
        {
            Method mUpdate = Helper.CreateMethod(Constants.Methods.UpdatelistItem, "Update list item by it's ID", MethodType.Update);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mUpdate);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mUpdate.InputProperties.Add(prop);
                    mUpdate.Validation.RequiredProperties.Add(prop);
                    mUpdate.ReturnProperties.Add(prop);
                }
                if (!prop.IsReadOnly() && !prop.IsInternal() && !prop.IsFile())
                {
                    mUpdate.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mUpdate);
        }

        private void AddDeleteMethod(ServiceObject so)
        {
            Method mDelete = Helper.CreateMethod(Constants.Methods.DeleteItem, "Delete one item by it's ID", MethodType.Delete);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mDelete);
            }


            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mDelete.InputProperties.Add(prop);
                    mDelete.Validation.RequiredProperties.Add(prop);
                }
            }

            so.Methods.Add(mDelete);
        }

        private void AddGetListMethod(ServiceObject so)
        {
            Method mLists = Helper.CreateMethod(Constants.Methods.ListAllItems, "Retrieve item lists", MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mLists);
            }

            foreach (Property prop in so.Properties)
            {
                if (!prop.IsInternal() &&
                    !prop.IsFile())
                {
                    mLists.InputProperties.Add(prop);
                    mLists.ReturnProperties.Add(prop);
                }
                if (prop.IsFile())
                {
                    mLists.ReturnProperties.Add(prop);
                }
                if (string.Compare(prop.Name, Constants.SOProperties.FolderName, true) == 0)
                {
                    mLists.InputProperties.Add(prop);
                }
                if (string.Compare(prop.Name, Constants.SOProperties.Recursively, true) == 0)
                {
                    mLists.InputProperties.Add(prop);
                }
                if (so.IsDocumentLibrary() && prop.IsFileName())
                {
                    mLists.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mLists);
        }

        private void AddMoveListItemMethod(ServiceObject so)
        {
            Method mMove = Helper.CreateMethod(Constants.Methods.MovelistItem, "Move list item by it's ID", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mMove);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mMove.InputProperties.Add(prop);
                    mMove.Validation.RequiredProperties.Add(prop);
                    mMove.ReturnProperties.Add(prop);
                }

                if (string.Compare(prop.Name, Constants.SOProperties.DestinationURL, true) == 0)
                {
                    mMove.InputProperties.Add(prop);
                    mMove.Validation.RequiredProperties.Add(prop);

                }
                if (string.Compare(prop.Name, Constants.SOProperties.DestinationListTitle, true) == 0)
                {
                    mMove.InputProperties.Add(prop);
                    mMove.Validation.RequiredProperties.Add(prop);

                }
                if (string.Compare(prop.Name, Constants.SOProperties.DestinationFolder, true) == 0)
                {
                    mMove.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mMove);
        }
        private void AddCopyListItemMethod(ServiceObject so)
        {
            Method mCopy = Helper.CreateMethod(Constants.Methods.CopylistItem, "Copy list item by it's ID", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCopy);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mCopy.InputProperties.Add(prop);
                    mCopy.Validation.RequiredProperties.Add(prop);
                    mCopy.ReturnProperties.Add(prop);
                }

                if (string.Compare(prop.Name, Constants.SOProperties.DestinationURL, true) == 0)
                {
                    mCopy.InputProperties.Add(prop);
                    mCopy.Validation.RequiredProperties.Add(prop);

                }
                if (string.Compare(prop.Name, Constants.SOProperties.DestinationListTitle, true) == 0)
                {
                    mCopy.InputProperties.Add(prop);
                    mCopy.Validation.RequiredProperties.Add(prop);

                }
                if (string.Compare(prop.Name, Constants.SOProperties.DestinationFolder, true) == 0)
                {
                    mCopy.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mCopy);
        }

        private void AddCreateFolderMethod(ServiceObject so)
        {
            Method mCreatFolder = Helper.CreateMethod(Constants.Methods.CreateFolder, "create folder in list", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCreatFolder);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.FolderName, true) == 0)
                {
                    mCreatFolder.InputProperties.Add(prop);
                    mCreatFolder.Validation.RequiredProperties.Add(prop);
                }
            }
            so.Methods.Add(mCreatFolder);
        }
        private void AddDeleteFolderMethod(ServiceObject so)
        {
            Method mDeleteFolder = Helper.CreateMethod(Constants.Methods.DeleteFolder, "Delete folder in list", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mDeleteFolder);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.FolderName, true) == 0)
                {
                    mDeleteFolder.InputProperties.Add(prop);
                    mDeleteFolder.Validation.RequiredProperties.Add(prop);
                }
                if (string.Compare(prop.Name, Constants.SOProperties.Recursively, true) == 0)
                {
                    mDeleteFolder.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mDeleteFolder);
        }


        private void AddGetLisItemsByFoldertMethod(ServiceObject so)
        {
            Method mLists = Helper.CreateMethod(Constants.Methods.GetItemsByFolder, "Retrieve list items in a folder", MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mLists);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.FolderName, true) == 0)
                {
                    mLists.InputProperties.Add(prop);
                    mLists.Validation.RequiredProperties.Add(prop);
                }
                if (!prop.IsInternal())
                {
                    mLists.InputProperties.Add(prop);
                    mLists.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mLists);
        }


        private void GetListItemById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(list);
                context.Load(listItem);
                context.Load(listItem, i => i.File);
                context.ExecuteQuery();

                DataRow dataRow = results.NewRow();

                foreach (Property prop in serviceObject.Properties)
                {
                    if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                    {
                        Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                    }
                    if (prop.IsFile())
                    {
                        var fileRef = listItem.File.ServerRelativeUrl;
                        context.ExecuteQuery();
                        var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, fileRef);
                        var fileName = (string)listItem.File.Name;
                        using (var fileStream = new System.IO.MemoryStream())
                        {
                            fileInfo.Stream.CopyTo(fileStream);

                            byte[] fileByte = fileStream.ToArray();

                            FileProperty propFile = new FileProperty(prop.Name, new MetaData(), fileName, System.Convert.ToBase64String(fileByte));

                            dataRow[prop.Name] = propFile.Value;//fileString.ToString();
                            dataRow[string.Format("{0}{1}", prop.Name, Constants.SOProperties.FileNameSuffix)] = fileName;
                            dataRow[string.Format("{0}{1}", prop.Name, Constants.SOProperties.FileURLSuffix)] = new StringBuilder(siteURL).Append(fileRef).ToString();

                        }
                    }
                }
                results.Rows.Add(dataRow);
            }
        }


        private string GetSiteURL()
        {
            string siteURL = base.SiteURL;
            if (base.IsDynamicSiteURL)
            {
                siteURL = base.GetStringParameter(Constants.InternalProperties.SiteUrl, true);
            }
            return siteURL;
        }

        private void CreateListItem()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            DataRow dataRow = results.NewRow();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                FieldCollection fieldColl = list.Fields;

                context.Load(list, d => d.RootFolder.Name);
                context.Load(fieldColl);
                context.ExecuteQuery();

                ListItem newItem = null;

                //get folder path
                string folderPath = string.Empty;
                if (serviceObject.Properties[Constants.SOProperties.FolderName].Value != null)
                {
                    folderPath = serviceObject.Properties[Constants.SOProperties.FolderName].Value.ToString();
                }
                if (!serviceObject.IsDocumentLibrary())
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    itemCreateInfo.FolderUrl = string.Empty;
                    if (!string.IsNullOrEmpty(folderPath))
                    {
                        itemCreateInfo.FolderUrl = string.Format("{0}/lists/{1}/{2}", siteURL, list.Title, folderPath.Trim('/'));
                    }

                    newItem = list.AddItem(itemCreateInfo);
                }
                else
                {
                    FileCreationInformation fileCreateInfo = new FileCreationInformation();
                    //TODO: XmlDocument would load everything into memory.
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.LoadXml(serviceObject.Properties[Constants.InternalProperties.FileLeafRef].Value.ToString());
                    foreach (XmlNode xn in xmlDocument)
                    {
                        XmlNode name = xn.SelectSingleNode("name");
                        XmlNode content = xn.SelectSingleNode("content");
                        if (name != null)
                        {
                            if (string.IsNullOrEmpty(folderPath))
                            {
                                fileCreateInfo.Url = string.Concat(GetSiteURL(), "/", list.RootFolder.Name, "/", name.InnerText);
                            }
                            else
                            {
                                fileCreateInfo.Url = string.Concat(GetSiteURL(), "/", list.RootFolder.Name, "/", folderPath, "/", name.InnerText);
                            }
                        }
                        if (content != null)
                        {
                           // byte[] fileContent = Convert.FromBase64String(content.InnerText);
                            //fileCreateInfo.Content = fileContent;
                            fileCreateInfo.ContentStream = new System.IO.MemoryStream(Convert.FromBase64String(content.InnerText));
                        }
                    }
                    fileCreateInfo.Overwrite = true;

                    File newFile = null;
                    if (string.IsNullOrEmpty(folderPath))
                    {
                        newFile = list.RootFolder.Files.Add(fileCreateInfo);
                    }
                    else
                    {
                        Folder currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.Name, folderPath.Trim('/')));
                        context.Load(currentFolder);
                        context.ExecuteQuery();

                        newFile = currentFolder.Files.Add(fileCreateInfo);
                    }

                    context.Load(newFile);
                    context.ExecuteQuery();

                    newItem = newFile.ListItemAllFields;

                    context.Load(newItem);
                    context.ExecuteQuery();
                }

                foreach (Property prop in serviceObject.Properties)
                {
                    if (prop.Value != null && !prop.IsFile() && string.Compare(prop.Name, Constants.SOProperties.FolderName, true) != 0)
                    {
                        Helpers.SPHelper.AssignFieldValue(newItem, prop);
                    }
                }

                newItem.Update();
                context.ExecuteQuery();

                dataRow[Constants.SOProperties.ID] = newItem.Id;

                results.Rows.Add(dataRow);

            }
        }

        private void UpdateListItem()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(listItem);

                foreach (Property prop in serviceObject.Properties)
                {

                    if (!string.IsNullOrEmpty(Convert.ToString(prop.Value)) && string.Compare(prop.Name, Constants.SOProperties.ID, true) != 0 && !prop.IsFile())
                    {
                        Helpers.SPHelper.AssignFieldValue(listItem, prop);
                    }
                }

                listItem.Update();
                context.ExecuteQuery();

                DataRow dataRow = results.NewRow();
                dataRow[Constants.SOProperties.ID] = listItem.Id;
                results.Rows.Add(dataRow);


            }
        }

        private void DeleteListItem()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                listItem.DeleteObject();
                context.ExecuteQuery();

            }
        }

        private void GetListItems()
        {

            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName);
            bool recursive = base.GetBoolProperty(Constants.SOProperties.Recursively);
            string folderPath = string.Empty;


            Properties searchProperties = Helpers.SPHelper.GetSearchFields(serviceObject.Properties);
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields);
                context.Load(list.RootFolder);
                context.ExecuteQuery();


                if (!string.IsNullOrEmpty(folderName))
                {
                    folderPath = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName);
                    Folder folder;
                    try
                    {
                        folder = spWeb.GetFolderByServerRelativeUrl(folderPath);
                        context.Load(folder);
                        context.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                    }
                }
                else
                {
                    folderPath = list.RootFolder.ServerRelativeUrl;
                }

                CamlQuery camlQuery = Helpers.SPHelper.CreateCamlQuery(searchProperties, "BeginsWith", list, recursive);
                camlQuery.FolderServerRelativeUrl = folderPath;

                IQueryable<ListItem> queryable = list.GetItems(camlQuery).IncludeWithDefaultProperties<ListItem>(item => item.ContentType);

                foreach (Field f in list.Fields)
                {
                    queryable.Concat<ListItem>((IEnumerable<ListItem>)queryable.Include<ListItem>(item => item[f.Title]));
                }

                IEnumerable<ListItem> source = context.LoadQuery(queryable);
                context.ExecuteQuery();
                foreach (ListItem listItem in source)
                {
                    DataRow dataRow = results.NewRow();

                    context.Load(listItem, i => i.File);
                    context.ExecuteQuery();

                    //DataRow dataRow = results.NewRow();
                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                        {
                            Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                        }
                        if (prop.IsFile())
                        {
                            var fileRef = listItem.File.ServerRelativeUrl;
                            var fileName = (string)listItem.File.Name;
                            dataRow[string.Format("{0}{1}", prop.Name, Constants.SOProperties.FileNameSuffix)] = fileName;
                            dataRow[string.Format("{0}{1}", prop.Name, Constants.SOProperties.FileURLSuffix)] = new StringBuilder(siteURL).Append(fileRef).ToString();
                        }
                    }
                    results.Rows.Add(dataRow);
                }
            }
        }

        private void ExecuteMoveListItem()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string destinationURL = base.GetStringProperty(Constants.SOProperties.DestinationURL, true);
            string destinationListTitle = base.GetStringProperty(Constants.SOProperties.DestinationListTitle, true);
            string destinationFolder = base.GetStringProperty(Constants.SOProperties.DestinationFolder, false);
            string parentListTitle = serviceObject.GetListTitle();

            ListItem parentListItem = GetSPListItem(siteURL, id, parentListTitle);

            ListItem destinationListItem = CopyListItem(destinationURL, destinationListTitle, destinationFolder, parentListItem);

            //Delete List Item from the Parent List
            DeleteParentListItem(siteURL, id, parentListTitle);

            DataRow dataRow = results.NewRow();
            dataRow[Constants.SOProperties.ID] = destinationListItem.Id;
            results.Rows.Add(dataRow);
        }

        private ListItem GetSPListItem(string siteURL, int itemId, string listTitle)
        {
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(itemId);
                context.Load(listItem);
                context.ExecuteQuery();
                return listItem;
            }
        }

        private ListItem CopyListItem(string siteURL, string listTitle, string folder, ListItem parentListItem)
        {
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list.Fields);
                context.Load(list.RootFolder);
                context.ExecuteQuery();
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                if (!string.IsNullOrEmpty(folder))
                {
                    itemCreateInfo.FolderUrl = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folder);
                }

                ListItem newItem = list.AddItem(itemCreateInfo);

                foreach (Field field in list.Fields)
                {
                    if (!field.Hidden && !field.ReadOnlyField && !string.IsNullOrEmpty(field.InternalName) && string.Compare(field.InternalName, Constants.SOProperties.ID, true) != 0 && string.Compare(field.InternalName, "Attachments", true) != 0 && parentListItem.FieldValues.ContainsKey(field.InternalName) && parentListItem[field.InternalName] != null)
                    {

                        newItem[field.InternalName] = parentListItem[field.InternalName];

                    }
                }

                newItem.Update();
                context.ExecuteQuery();
                return newItem;
            }
        }

        private void DeleteParentListItem(string siteURL, int itemId, string listTitle)
        {
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(itemId);
                listItem.DeleteObject();
                context.ExecuteQuery();

            }
        }

        private void ExecuteCopyListItem()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);

            string destinationURL = base.GetStringProperty(Constants.SOProperties.DestinationURL, true);
            string destinationListTitle = base.GetStringProperty(Constants.SOProperties.DestinationListTitle, true);
            string destinationFolder = base.GetStringProperty(Constants.SOProperties.DestinationFolder, false);
            string parentListTitle = serviceObject.GetListTitle();

            //GetParentListItem
            ListItem parentListItem = GetSPListItem(siteURL, id, parentListTitle);
            //CopyItem to the destination List
            ListItem destinationListItem = CopyListItem(destinationURL, destinationListTitle, destinationFolder, parentListItem);


            DataRow dataRow = results.NewRow();
            dataRow[Constants.SOProperties.ID] = destinationListItem.Id;
            results.Rows.Add(dataRow);
        }

        private void CreateFolder()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string listName = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listName);
                context.Load(list);
                context.ExecuteQuery();

                //TODO: get this information from the ServiceObject
                if (!list.EnableFolderCreation)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderCreationIsNotSupported, listName));
                }

                CreateFolderInternal(spWeb, list, list.RootFolder, string.Empty, folderName);
            }
        }

        private static Folder CreateFolderInternal(Web web, List list, Folder parentFolder, string parentFolderPath, string fullFolderUrl)
        {
            string[] folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            parentFolderPath += string.Concat(folderUrl, "/");

            Folder curFolder;
            try
            {
                //check is folder exists 
                curFolder = parentFolder.Folders.GetByUrl(folderUrl);
                web.Context.Load(curFolder);
                web.Context.ExecuteQuery();
            }
            catch
            {
                //if not, create a new folder
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.LeafName = parentFolderPath.Trim('/');
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                ListItem newItem = list.AddItem(itemCreateInfo);
                newItem[Constants.InternalProperties.Title] = folderUrl;
                newItem.Update();
                web.Context.ExecuteQuery();

                //get current folder
                curFolder = parentFolder.Folders.GetByUrl(folderUrl);
                web.Context.Load(curFolder);
                web.Context.ExecuteQuery();
            }
            if (folderUrls.Length > 1)
            {
                string subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return CreateFolderInternal(web, list, curFolder, parentFolderPath, subFolderUrl);
            }
            return curFolder;
        }

        private void DeleteFolder()
        {
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
            bool recursive = base.GetBoolProperty(Constants.SOProperties.Recursively);
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(serviceObject.MetaData.GetServiceElement<string>(Constants.InternalProperties.ListTitle));
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                Folder folder;
                try
                {
                    string folderRelativeUrl = string.Concat(list.RootFolder.ServerRelativeUrl, '/', folderName);
                    folder = spWeb.GetFolderByServerRelativeUrl(folderRelativeUrl);
                    context.Load(folder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                }

                if (recursive && folder.ItemCount > 0)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderIsNotEmpty, folderName));
                }

                folder.DeleteObject();

                context.ExecuteQuery();
            }
        }

        private void GetItemsbyFolder()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            string siteURL = GetSiteURL();
            string listTitle = serviceObject.GetListTitle();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string folderPath = string.Empty;

            Properties searchProperties = Helpers.SPHelper.GetSearchFields(serviceObject.Properties);
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list.Fields);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                if (!string.IsNullOrEmpty(folderName))
                {
                    folderPath = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName);
                }
                else
                {
                    folderPath = list.RootFolder.ServerRelativeUrl;
                }

                CamlQuery calmQuery = Helpers.SPHelper.CreateCamlQuery(searchProperties, "BeginsWith", list);
                //Adding Folder Path
                calmQuery.FolderServerRelativeUrl = folderPath;

                IQueryable<ListItem> queryable = list.GetItems(calmQuery).IncludeWithDefaultProperties(item => item.ContentType);

                foreach (Field f in list.Fields)
                {
                    queryable.Concat((IEnumerable<ListItem>)queryable.Include(item => item[f.Title]));
                }

                IEnumerable<ListItem> source = context.LoadQuery(queryable);
                context.ExecuteQuery();
                foreach (ListItem listItem in source)
                {
                    DataRow dataRow = results.NewRow();
                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (listItem.FieldValues.ContainsKey(prop.Name))
                        {
                            Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                        }

                    }
                    results.Rows.Add(dataRow);
                }
            }
        }




        public override void Execute()
        {
            switch (ServiceBroker.Service.ServiceObjects[0].Methods[0].Name)
            {
                case Constants.Methods.GetItembyId:
                    GetListItemById();
                    break;
                case Constants.Methods.CreatelistItem:
                    CreateListItem();
                    break;
                case Constants.Methods.UpdatelistItem:
                    UpdateListItem();
                    break;
                case Constants.Methods.DeleteItem:
                    DeleteListItem();
                    break;
                case Constants.Methods.ListAllItems:
                    GetListItems();
                    break;
                case Constants.Methods.MovelistItem:
                    ExecuteMoveListItem();
                    break;
                case Constants.Methods.CopylistItem:
                    ExecuteCopyListItem();
                    break;
                case Constants.Methods.CreateFolder:
                    CreateFolder();
                    break;
                case Constants.Methods.DeleteFolder:
                    DeleteFolder();
                    break;
                case Constants.Methods.GetItemsByFolder:
                    GetItemsbyFolder();
                    break;
            }
        }
    }
}
