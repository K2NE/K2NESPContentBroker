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

namespace K2Field.K2NE.SPContentBroker.ServiceObjects
{
    public partial class SPServiceObject : ServiceObjectBase
    {
        #region CreateDocument
        private void AddCreateDocumentMethod(ServiceObject so)
        {
            Method mCreateDocument = Helper.CreateMethod(Constants.Methods.CreateDocument, "Create document with metadata", MethodType.Create);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCreateDocument);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0 && !prop.IsInternal())
                {
                    mCreateDocument.ReturnProperties.Add(prop);
                }

                if (!prop.IsReadOnly() && !prop.IsInternal())
                {
                    mCreateDocument.InputProperties.Add(prop);
                    if (prop.IsRequired())
                    {
                        mCreateDocument.Validation.RequiredProperties.Add(prop);
                    }
                }
                if ((so.IsSPFolderEnabled() == true &&
                    prop.IsFolderName()) || prop.IsOverwriteExistingDocument())
                {
                    mCreateDocument.InputProperties.Add(prop);
                }

                if (prop.IsLinkToItem() || prop.IsFileName())
                {
                    mCreateDocument.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCreateDocument);
        }

        private void ExecuteCreateDocument()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderPath = base.GetStringProperty(Constants.SOProperties.FolderName);
            bool overwriteExistingDocument = base.GetBoolProperty(Constants.SOProperties.OverwriteExistingDocument);

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            DataRow dataRow = results.NewRow();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                FieldCollection fieldColl = list.Fields;
                
                context.Load(list, d => d.RootFolder.Name, d => d.DefaultDisplayFormUrl);
                context.Load(fieldColl);
                context.ExecuteQuery();

                ListItem newItem = null;
                
                File newFile = null;

                FileProperty fileProperty = (FileProperty)serviceObject.Properties.Where(p => p.IsFile()).FirstOrDefault();
                
                newFile = AddFile(context, spWeb, list, fileProperty.FileName, Convert.FromBase64String(fileProperty.Content), folderPath, overwriteExistingDocument);
                
                newItem = newFile.ListItemAllFields;

                context.Load(newItem);
                context.ExecuteQuery();

                foreach (Property prop in serviceObject.Properties)
                {
                    if (prop.Value != null && !prop.IsFile() && !prop.IsFolderName() && !prop.IsOverwriteExistingDocument())
                    {
                        Helpers.SPHelper.AssignFieldValue(newItem, prop);
                    }
                }

                newItem.Update();
                context.ExecuteQuery();

                dataRow[Constants.SOProperties.ID] = newItem.Id;
                dataRow[Constants.SOProperties.FileName] = newFile.Name;
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, newItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;
                results.Rows.Add(dataRow);
                
            }
        }
        #endregion

        #region DeleteDocumentById
        private void AddDeleteDocumentByIdMethod(ServiceObject so)
        {
            Method mDeleteDocument = Helper.CreateMethod(Constants.Methods.DeleteDocumentById, "Delete one document by it's ID", MethodType.Delete);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mDeleteDocument);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mDeleteDocument.InputProperties.Add(prop);
                    mDeleteDocument.Validation.RequiredProperties.Add(prop);
                }
            }

            so.Methods.Add(mDeleteDocument);
        }

        private void ExecuteDeleteDocumentById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

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
        #endregion

        #region GetDocumentById
        private void AddGetDocumentByIdMethod(ServiceObject so)
        {
            Method mGetDocumentById = Helper.CreateMethod(Constants.Methods.GetDocumentById, "Retrieve one document with metadata by it's ID", MethodType.Read);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetDocumentById);
            }
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mGetDocumentById.InputProperties.Add(prop);
                    mGetDocumentById.Validation.RequiredProperties.Add(prop);
                }

                if (!prop.IsInternal())
                {
                    mGetDocumentById.ReturnProperties.Add(prop);

                }
                if (prop.IsFileName() || prop.IsLinkToItem())
                {
                    mGetDocumentById.ReturnProperties.Add(prop);
                }

            }
            so.Methods.Add(mGetDocumentById);
        }

        private void ExecuteGetDocumentById()
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
                context.Load(list, l => l.DefaultDisplayFormUrl);
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
                        FileProperty propFile = new FileProperty(prop.Name, new MetaData(), listItem.File.Name, 
                            System.Convert.ToBase64String(GetFile(context, listItem)));

                        dataRow[prop.Name] = propFile.Value;
                    }
                    if(prop.IsFileName())
                    {
                        dataRow[prop.Name] = listItem.File.Name;
                    }
                    if(prop.IsLinkToItem())
                    {
                        string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                        dataRow[prop.Name] = strurl;
                    }
                }
                results.Rows.Add(dataRow);
            }
        }
        #endregion

        #region GetDocuments
        private void AddGetDocumentsMethod(ServiceObject so)
        {
            Method mGetDocuments = Helper.CreateMethod(Constants.Methods.GetDocuments, "Retrieve documents and related items", MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetDocuments);
            }

            foreach (Property prop in so.Properties)
            {
                if (!prop.IsInternal() &&
                    !prop.IsFile())
                {
                    mGetDocuments.InputProperties.Add(prop);
                    mGetDocuments.ReturnProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName() || prop.IsRecursivelyName()))
                {
                    mGetDocuments.InputProperties.Add(prop);
                }
                if (prop.IsFileName() || prop.IsLinkToItem() || prop.IsFile())
                {
                    mGetDocuments.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mGetDocuments);
        }

        private void ExecuteGetDocuments()
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
                context.Load(list, l => l.Fields , l => l.DefaultDisplayFormUrl);
                context.Load(list.RootFolder);
                context.ExecuteQuery();
                
                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, recursive, "BeginsWith",true);
                
                foreach (ListItem listItem in source)
                {
                    if(listItem.FileSystemObjectType == FileSystemObjectType.File)
                    {
                        DataRow dataRow = results.NewRow();                 
                        foreach (Property prop in serviceObject.Properties)
                        {
                            if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                            {
                                Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                            }
                            if (prop.IsFile())
                            {
                                var fileRef = listItem.FieldValues[Constants.SharePointProperties.FileRef].ToString();
                                // OpenBinaryStream method works with Oauth authentication
                                var fileInfo = listItem.File.OpenBinaryStream();
                                context.ExecuteQuery();
                                //var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, fileRef);
                                var fileName = (string)listItem.FieldValues[Constants.SharePointProperties.FileLeafRef].ToString();
                                using (var fileStream = new System.IO.MemoryStream())
                                {
                                    fileInfo.Value.CopyTo(fileStream);

                                    byte[] fileByte = fileStream.ToArray();

                                    FileProperty propFile = new FileProperty(prop.Name, new MetaData(), fileName, System.Convert.ToBase64String(fileByte));

                                    dataRow[prop.Name] = propFile.Value;
                                }
                            }
                            if (prop.IsFileName())
                            {
                                var fileName = (string)listItem.FieldValues[Constants.SharePointProperties.FileLeafRef].ToString();
                                dataRow[prop.Name] = fileName;
                            }

                            if (prop.IsLinkToItem())
                            {
                                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                                dataRow[prop.Name] = strurl;
                            }
                        }
                        results.Rows.Add(dataRow);
                    }
                  
                   
                }
            }
        }
        #endregion

        #region RenameDocumentById
        private void AddRenameDocumentByIdMethod(ServiceObject so)
        {
            Method mRenameDocument = Helper.CreateMethod(Constants.Methods.RenameDocumentById, "Rename document list item's ID", MethodType.Read);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mRenameDocument);
            }
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mRenameDocument.InputProperties.Add(prop);
                    mRenameDocument.Validation.RequiredProperties.Add(prop);
                    mRenameDocument.ReturnProperties.Add(prop);
                }
                if(prop.IsNewFileName())
                {
                    mRenameDocument.InputProperties.Add(prop);
                    mRenameDocument.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mRenameDocument.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mRenameDocument);
        }

        private void ExecuteRenameDocumentById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string newFileName = base.GetStringProperty(Constants.SOProperties.NewFileName);

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(list);
                context.Load(list, l => l.DefaultDisplayFormUrl);
                context.Load(listItem);
                context.Load(listItem, i => i.File);
                context.ExecuteQuery();

                DataRow dataRow = results.NewRow();
                
                foreach (Property prop in serviceObject.Properties)
                {
                    if (prop.IsFile())
                    {
                        listItem[prop.Name] = newFileName;
                        listItem.Update();
                        context.ExecuteQuery();
                    }
                    if (prop.IsLinkToItem())
                    {
                        string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                        dataRow[prop.Name] = strurl;
                    }
                }
                dataRow[Constants.SOProperties.ID] = id;
                results.Rows.Add(dataRow);
            }
        }
        #endregion

        #region CopyDocumentByName
        private void AddCopyDocumentByNameMethod(ServiceObject so)
        {
            Method mCopyDocumentByName = Helper.CreateMethod(Constants.Methods.CopyDocumentByName, "Copy document by it's name", MethodType.Read);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCopyDocumentByName);
            }

            foreach (Property prop in so.Properties)
            {
                if(prop.IsDestinationURL() || prop.IsDestinationLibrary() || prop.IsDestinationFolder())
                {
                    mCopyDocumentByName.InputProperties.Add(prop);
                    if(!prop.IsDestinationFolder())
                    {
                        mCopyDocumentByName.Validation.RequiredProperties.Add(prop);
                    }
                }
                if (prop.IsFileName())
                {
                    mCopyDocumentByName.InputProperties.Add(prop);
                    mCopyDocumentByName.Validation.RequiredProperties.Add(prop);
                    mCopyDocumentByName.ReturnProperties.Add(prop);
                }
                if ((so.IsSPFolderEnabled() && prop.IsFolderName()) || prop.IsOverwriteExistingDocument())
                {
                    mCopyDocumentByName.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mCopyDocumentByName.ReturnProperties.Add(prop);
                }
            }
            
            so.Methods.Add(mCopyDocumentByName);
        }

        private void ExecuteCopyDocumentByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            
            string destinationSiteURL = base.GetStringProperty(Constants.SOProperties.DestinationURL, true);
            string destinationLibrary = base.GetStringProperty(Constants.SOProperties.DestinationLibrary, true);
            string destinationFolder = base.GetStringProperty(Constants.SOProperties.DestinationFolder, false);
            bool overwriteExistingDocument = base.GetBoolProperty(Constants.SOProperties.OverwriteExistingDocument);
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            string fileName = base.GetStringProperty(Constants.SOProperties.FileName, true);
            string folderPath = string.Empty;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            //add FileLeafRef property for filter by File Name
            Property fileLeafRefProperty = NewProperty(Constants.InternalProperties.FileLeafRef, Constants.InternalProperties.FileLeafRef, true, SoType.Text, "");
            fileLeafRefProperty.Value = serviceObject.Properties.Where(p => p.IsFileName()).FirstOrDefault().Value;
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, "Text");
            Properties searchProperties = new Properties { fileLeafRefProperty };
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, false, "Eq");

                foreach (ListItem listItem in source)
                {
                    context.Load(listItem);
                    context.Load(listItem, l => l.File);
                    context.ExecuteQuery();
                    DataRow dataRow = results.NewRow();

                    //create destination context
                    using (ClientContext destContext = InitializeContext(destinationSiteURL))
                    {
                        Web destSpWeb = destContext.Web;
                        List destList = destSpWeb.Lists.GetByTitle(destinationLibrary);
                        destContext.Load(destList);
                        destContext.Load(destList, l => l.RootFolder, l => l.DefaultDisplayFormUrl);
                        destContext.ExecuteQuery();

                        //get current file content
                        byte[] currentFileContent = GetFile(context, listItem);

                        //add file
                        File newFile = AddFile(destContext, destSpWeb, destList, listItem.File.Name,
                            currentFileContent, destinationFolder, overwriteExistingDocument);
                        ListItem newItem = newFile.ListItemAllFields;
                        destContext.Load(newItem);
                        destContext.ExecuteQuery();

                        //copy metaData
                        SPHelper.CopyItem(context, destContext, listItem, newItem);

                        dataRow[Constants.SOProperties.FileName] = newFile.Name;
                        string strurl = BuildListItemLink(destContext.Url, destList.DefaultDisplayFormUrl, newItem.Id);
                        dataRow[Constants.SOProperties.LinkToItem] = strurl;
                    }

                    results.Rows.Add(dataRow);
                }
            }
        }
        #endregion

        #region MoveDocumentByName
        private void AddMoveDocumentByNameMethod(ServiceObject so)
        {
            Method mMoveDocumentByName = Helper.CreateMethod(Constants.Methods.MoveDocumentByName, "Move document by it's name", MethodType.Read);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mMoveDocumentByName);
            }

            foreach (Property prop in so.Properties)
            {
                if (prop.IsDestinationURL() || prop.IsDestinationLibrary() || prop.IsDestinationFolder())
                {
                    mMoveDocumentByName.InputProperties.Add(prop);
                    if (!prop.IsDestinationFolder())
                    {
                        mMoveDocumentByName.Validation.RequiredProperties.Add(prop);
                    }
                }
                if (prop.IsFileName())
                {
                    mMoveDocumentByName.InputProperties.Add(prop);
                    mMoveDocumentByName.Validation.RequiredProperties.Add(prop);
                    mMoveDocumentByName.ReturnProperties.Add(prop);
                }
                if ((so.IsSPFolderEnabled() && prop.IsFolderName()) || prop.IsOverwriteExistingDocument())
                {
                    mMoveDocumentByName.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mMoveDocumentByName.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mMoveDocumentByName);
        }

        private void ExecuteMoveDocumentByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string destinationSiteURL = base.GetStringProperty(Constants.SOProperties.DestinationURL, true);
            string destinationLibrary = base.GetStringProperty(Constants.SOProperties.DestinationLibrary, true);
            string destinationFolder = base.GetStringProperty(Constants.SOProperties.DestinationFolder, false);
            bool overwriteExistingDocument = base.GetBoolProperty(Constants.SOProperties.OverwriteExistingDocument);
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            string fileName = base.GetStringProperty(Constants.SOProperties.FileName, true);
            string folderPath = string.Empty;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            //add FileLeafRef property for filter by File Name
            Property fileLeafRefProperty = NewProperty(Constants.InternalProperties.FileLeafRef, Constants.InternalProperties.FileLeafRef, true, SoType.Text, "");
            fileLeafRefProperty.Value = serviceObject.Properties.Where(p => p.IsFileName()).FirstOrDefault().Value;
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, "Text");
            Properties searchProperties = new Properties { fileLeafRefProperty };
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, false, "Eq");

                foreach (ListItem listItem in source)
                {
                    context.Load(listItem);
                    context.Load(listItem, l => l.File);
                    context.ExecuteQuery();
                    DataRow dataRow = results.NewRow();

                    //create destination context
                    using (ClientContext destContext = InitializeContext(destinationSiteURL))
                    {
                        Web destSpWeb = destContext.Web;
                        List destList = destSpWeb.Lists.GetByTitle(destinationLibrary);
                        destContext.Load(destList);
                        destContext.Load(destList, l => l.RootFolder, l => l.DefaultDisplayFormUrl);
                        destContext.ExecuteQuery();

                        //get current file content
                        byte[] currentFileContent = GetFile(context, listItem);

                        //add file
                        File newFile = AddFile(destContext, destSpWeb, destList, listItem.File.Name,
                            currentFileContent, destinationFolder, overwriteExistingDocument);
                        ListItem newItem = newFile.ListItemAllFields;
                        destContext.Load(newItem);
                        destContext.ExecuteQuery();

                        //copy metaData
                        SPHelper.CopyItem(context, destContext, listItem, newItem);

                        if (!IsFileTheSame(siteURL, destinationSiteURL, listTitle, destinationLibrary, listItem.Id, newItem.Id))
                        {
                            listItem.DeleteObject();
                            context.ExecuteQuery();
                        }

                        dataRow[Constants.SOProperties.FileName] = newFile.Name;
                        string strurl = BuildListItemLink(destContext.Url, destList.DefaultDisplayFormUrl, newItem.Id);
                        dataRow[Constants.SOProperties.LinkToItem] = strurl;
                    }

                    results.Rows.Add(dataRow);
                }
            }
        }
        #endregion

        private byte[] GetFile(ClientContext context, ListItem item)
        {
            context.Load(item.File, f => f.ServerRelativeUrl);
            context.ExecuteQuery();

            var fileRef = item.File.ServerRelativeUrl;
            
            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, fileRef);
            using (var fileStream = new System.IO.MemoryStream())
            {
                fileInfo.Stream.CopyTo(fileStream);

                byte[] fileByte = fileStream.ToArray();

                return fileByte;
            }
        }

        private File AddFile(ClientContext context, Web spWeb, List list, string fileName, byte[] fileContent, string folderPath, bool overwriteExistingDocument)
        {
            context.Load(list.RootFolder, rf => rf.ServerRelativeUrl);
            context.ExecuteQuery();

            using (System.IO.MemoryStream streamContent = new System.IO.MemoryStream(fileContent))
            {
                //create file base on input property
                FileCreationInformation fileCreateInfo = new FileCreationInformation();

                if (string.IsNullOrEmpty(folderPath)) 
                {
                    fileCreateInfo.Url = string.Concat(list.RootFolder.ServerRelativeUrl, "/", fileName);
                }
                else
                {
                    fileCreateInfo.Url = string.Concat(list.RootFolder.ServerRelativeUrl, "/", folderPath, '/', fileName);
                }
                fileCreateInfo.ContentStream = streamContent;

                fileCreateInfo.Overwrite = overwriteExistingDocument;

                //add folder into root folder or in a specific one
                File newFile = null;
                if (string.IsNullOrEmpty(folderPath))
                {
                    newFile = list.RootFolder.Files.Add(fileCreateInfo);
                }
                else
                {
                    Folder currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', folderPath));
                    context.Load(currentFolder);
                    context.ExecuteQuery();

                    newFile = currentFolder.Files.Add(fileCreateInfo);
                }

                context.Load(newFile);
                context.ExecuteQuery();
                return newFile;
            }            
        }

        private File AddFileByFolderPathWithRoot(ClientContext context, Web spWeb, List list, string fileName, byte[] fileContent, string folderPathWithRoot, bool overwriteExistingDocument)
        {
            using (System.IO.MemoryStream streamContent = new System.IO.MemoryStream(fileContent))
            {
                //create file base on input property
                FileCreationInformation fileCreateInfo = new FileCreationInformation();

                if (string.IsNullOrEmpty(folderPathWithRoot))
                {
                    fileCreateInfo.Url = string.Concat(list.RootFolder.ServerRelativeUrl, "/", fileName);
                }
                else
                {
                    fileCreateInfo.Url = string.Concat(folderPathWithRoot, '/', fileName);
                }
                fileCreateInfo.ContentStream = streamContent;

                fileCreateInfo.Overwrite = overwriteExistingDocument;

                //add folder into root folder or in a specific one
                File newFile = null;
                if (string.IsNullOrEmpty(folderPathWithRoot))
                {
                    newFile = list.RootFolder.Files.Add(fileCreateInfo);
                }
                else
                {
                    Folder currentFolder = spWeb.GetFolderByServerRelativeUrl(folderPathWithRoot);
                    context.Load(currentFolder);
                    context.ExecuteQuery();

                    newFile = currentFolder.Files.Add(fileCreateInfo);
                }

                context.Load(newFile);
                context.ExecuteQuery();
                return newFile;
            }
        }

        private bool IsFileTheSame(string siteURL, string destSiteURL, string libraryName, string destLibraryName, int itemID, int destItemID)
        {
            if (string.Compare(siteURL.Trim('/'), destSiteURL.Trim('/')) == 0 &&
                string.Compare(libraryName.Trim('/'), destLibraryName.Trim('/')) == 0 &&
                itemID == destItemID)
                return true;
            else
                return false;
        }

        private string BuildListItemLink(string siteUrl, string relativeDefaultDisplayFormUrl, int itemId)
        {
            return string.Concat(new Uri(new Uri(siteUrl), relativeDefaultDisplayFormUrl).AbsoluteUri, "?ID=", itemId);
        }

        #region CheckInDocument
        private File CheckInDocument(ClientContext context, File targetFile, string checkInComment, bool retainCheckout, bool useCheckedInVersion)
        {
            if (targetFile == null)
                return (File)null;
            if (targetFile.CheckOutType == CheckOutType.None)
            {
                if (useCheckedInVersion)
                    return targetFile;
                throw new Exception(Constants.ErrorMessages.DocumentNotCheckedOut);
            }
            targetFile.CheckIn(checkInComment, CheckinType.MajorCheckIn);
            context.ExecuteQuery();
            if (retainCheckout)
            {
                targetFile.CheckOut();
                context.ExecuteQuery();
            }
            return targetFile;
        }

        private void AddCheckInDocumentByNameMethod(ServiceObject so)
        {
            Method mCheckInDocumentByName = Helper.CreateMethod(Constants.Methods.CheckInDocumentByName, "Check In document by it's name", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCheckInDocumentByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFileName())
                {
                    mCheckInDocumentByName.InputProperties.Add(prop);
                    mCheckInDocumentByName.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsCheckInComments())
                {
                    mCheckInDocumentByName.InputProperties.Add(prop);
                }
                if (prop.IsRetainCheckout())
                {
                    mCheckInDocumentByName.InputProperties.Add(prop);
                }
                if (prop.IsUseCheckedInVersion())
                {
                    mCheckInDocumentByName.InputProperties.Add(prop);
                }
                if ((so.IsSPFolderEnabled() && prop.IsFolderName()))
                {
                    mCheckInDocumentByName.InputProperties.Add(prop);
                }

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mCheckInDocumentByName.ReturnProperties.Add(prop);
                }

                if (prop.IsLinkToItem())
                {
                    mCheckInDocumentByName.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCheckInDocumentByName);
        }

        private void ExecuteCheckInDocumentByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            string fileName = base.GetStringProperty(Constants.SOProperties.FileName, true);
            string checkInComments = base.GetStringProperty(Constants.SOProperties.CheckInComment, false);
            bool retainCheckout = base.GetBoolProperty(Constants.SOProperties.RetainCheckOut);
            bool useCheckedInVersion = base.GetBoolProperty(Constants.SOProperties.UseCheckedInVersion);

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            //add FileLeafRef property for filter by File Name
            Property fileLeafRefProperty = NewProperty(Constants.InternalProperties.FileLeafRef, Constants.InternalProperties.FileLeafRef, true, SoType.Text, "");
            fileLeafRefProperty.Value = serviceObject.Properties.Where(p => p.IsFileName()).FirstOrDefault().Value;
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, "Text");
            Properties searchProperties = new Properties { fileLeafRefProperty };
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.RootFolder );
                context.Load(list, l => l.DefaultDisplayFormUrl);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, false, "Eq");

                ListItem listItem = source.FirstOrDefault();

                if (listItem == null)
                {
                    throw new Exception(Constants.ErrorMessages.RequiredDocNotFound);
                }
                context.Load(listItem);
                context.Load(listItem, l => l.File);
                context.ExecuteQuery();
                DataRow dataRow = results.NewRow();

                CheckInDocument(context, listItem.File, checkInComments, retainCheckout, useCheckedInVersion);

                dataRow[Constants.SOProperties.ID] = listItem.Id;
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;



                results.Rows.Add(dataRow);
            }
        }

        private void AddCheckInDocumentByIdMethod(ServiceObject so)
        {
            Method mCheckInDocumentById = Helper.CreateMethod(Constants.Methods.CheckInDocumentById, "Check In document by Id", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCheckInDocumentById);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mCheckInDocumentById.InputProperties.Add(prop);
                    mCheckInDocumentById.Validation.RequiredProperties.Add(prop);
                    mCheckInDocumentById.ReturnProperties.Add(prop);
                }

                if (prop.IsCheckInComments())
                {
                    mCheckInDocumentById.InputProperties.Add(prop);
                }
                if (prop.IsRetainCheckout())
                {
                    mCheckInDocumentById.InputProperties.Add(prop);
                }
                if (prop.IsUseCheckedInVersion())
                {
                    mCheckInDocumentById.InputProperties.Add(prop);
                }
                if ((so.IsSPFolderEnabled() && prop.IsFolderName()))
                {
                    mCheckInDocumentById.InputProperties.Add(prop);
                }

                if (prop.IsLinkToItem())
                {
                    mCheckInDocumentById.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCheckInDocumentById);
        }

        private void ExecuteCheckInDocumentById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string checkInComments = base.GetStringProperty(Constants.SOProperties.CheckInComment, false);
            bool retainCheckout = base.GetBoolProperty(Constants.SOProperties.RetainCheckOut);
            bool useCheckedInVersion = base.GetBoolProperty(Constants.SOProperties.UseCheckedInVersion);

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(list);
                context.Load(list, l => l.DefaultDisplayFormUrl);
                context.Load(listItem);
                context.Load(listItem, i => i.File);
                context.ExecuteQuery();

                DataRow dataRow = results.NewRow();

                CheckInDocument(context, listItem.File, checkInComments, retainCheckout, useCheckedInVersion);

                dataRow[Constants.SOProperties.ID] = listItem.Id;
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;

                results.Rows.Add(dataRow);
            }
        }
        #endregion


        #region CheckOutDocument
        private File CheckOutDocument(ClientContext context, File file, bool useCheckedOutVersion)
        {
            if (file == null)
                return (Microsoft.SharePoint.Client.File)null;
            if (file.CheckOutType == CheckOutType.Online)
            {
                if (useCheckedOutVersion)
                {
                    return file;
                }
                throw new Exception(Constants.ErrorMessages.DocumentAlreadyCheckedOut);
            }
            file.CheckOut();
            context.ExecuteQuery();
            return file;
        }

        private void AddCheckOutDocumentByNameMethod(ServiceObject so)
        {
            Method mCheckOutDocumentByName = Helper.CreateMethod(Constants.Methods.CheckOutDocumentByName, "Check Out document by it's name", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCheckOutDocumentByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFileName())
                {
                    mCheckOutDocumentByName.InputProperties.Add(prop);
                    mCheckOutDocumentByName.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsUseCheckedOutVersion())
                {
                    mCheckOutDocumentByName.InputProperties.Add(prop);
                }

                if ((so.IsSPFolderEnabled() && prop.IsFolderName()))
                {
                    mCheckOutDocumentByName.InputProperties.Add(prop);
                }

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mCheckOutDocumentByName.ReturnProperties.Add(prop);
                }

                if (prop.IsLinkToItem())
                {
                    mCheckOutDocumentByName.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCheckOutDocumentByName);
        }

        private void ExecuteCheckOutDocumentByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            string fileName = base.GetStringProperty(Constants.SOProperties.FileName, true);
            bool useCheckedOutVersion = base.GetBoolProperty(Constants.SOProperties.UseCheckedOutVersion);

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            //add FileLeafRef property for filter by File Name
            Property fileLeafRefProperty = NewProperty(Constants.InternalProperties.FileLeafRef, Constants.InternalProperties.FileLeafRef, true, SoType.Text, "");
            fileLeafRefProperty.Value = serviceObject.Properties.Where(p => p.IsFileName()).FirstOrDefault().Value;
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
            fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, "Text");
            Properties searchProperties = new Properties { fileLeafRefProperty };
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.Load(list, l => l.DefaultDisplayFormUrl);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, false, "Eq");

                ListItem listItem = source.FirstOrDefault();

                if (listItem == null)
                {
                    throw new Exception(Constants.ErrorMessages.RequiredDocNotFound);
                }
                context.Load(listItem);
                context.Load(listItem, l => l.File);
                context.ExecuteQuery();
                DataRow dataRow = results.NewRow();

                CheckOutDocument(context, listItem.File, useCheckedOutVersion);

                dataRow[Constants.SOProperties.ID] = listItem.Id;
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;

                results.Rows.Add(dataRow);
            }
        }

        private void AddCheckOutDocumentByIdMethod(ServiceObject so)
        {
            Method mCheckOutDocumentById = Helper.CreateMethod(Constants.Methods.CheckOutDocumentById, "Check Out document by Id", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCheckOutDocumentById);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mCheckOutDocumentById.InputProperties.Add(prop);
                    mCheckOutDocumentById.Validation.RequiredProperties.Add(prop);
                    mCheckOutDocumentById.ReturnProperties.Add(prop);
                }

                if (prop.IsUseCheckedOutVersion())
                {
                    mCheckOutDocumentById.InputProperties.Add(prop);
                }
                if ((so.IsSPFolderEnabled() && prop.IsFolderName()))
                {
                    mCheckOutDocumentById.InputProperties.Add(prop);
                }

                if (prop.IsLinkToItem())
                {
                    mCheckOutDocumentById.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCheckOutDocumentById);
        }

        private void ExecuteCheckOutDocumentById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);

            bool useCheckedOutVersion = base.GetBoolProperty(Constants.SOProperties.UseCheckedInVersion);

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(list);
                context.Load(list, l => l.DefaultDisplayFormUrl);
                context.Load(listItem);
                context.Load(listItem, i => i.File);
                context.ExecuteQuery();

                DataRow dataRow = results.NewRow();

                CheckOutDocument(context, listItem.File, useCheckedOutVersion);

                dataRow[Constants.SOProperties.ID] = listItem.Id;
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;

                results.Rows.Add(dataRow);
            }
        }
        #endregion

    }

}
