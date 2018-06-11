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
    public partial class SPServiceObject: ServiceObjectBase
    {
        #region CreateFolder
        private void AddCreateFolderMethod(ServiceObject so)
        {
            Method mCreateFolder = Helper.CreateMethod(Constants.Methods.CreateFolder, "Create folder in list/library", MethodType.Create);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCreateFolder);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName())
                {
                    mCreateFolder.InputProperties.Add(prop);
                    mCreateFolder.Validation.RequiredProperties.Add(prop);
                }
                if(prop.IsLinkToItem())
                {
                    mCreateFolder.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mCreateFolder);
        }

        private void ExecuteCreateFolder()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string listName = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            DataRow dataRow = results.NewRow();
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listName);
                context.Load(list);
                context.Load(list, l => l.DefaultDisplayFormUrl);
                context.ExecuteQuery();
                
                if (!serviceObject.IsSPFolderEnabled())
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderCreationIsNotSupported, listName));
                }

                Folder folder = CreateFolderInternal(spWeb, list, list.RootFolder, string.Empty, folderName);

                context.Load(folder.ListItemAllFields);
                context.ExecuteQuery();

                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, folder.ListItemAllFields.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;

                results.Rows.Add(dataRow);
            }
        }
        #endregion

        #region DeleteFolder
        private void AddDeleteFolderMethod(ServiceObject so)
        {
            Method mDeleteFolder = Helper.CreateMethod(Constants.Methods.DeleteFolder, "Delete folder in a list/library", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mDeleteFolder);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName())
                {
                    mDeleteFolder.InputProperties.Add(prop);
                    mDeleteFolder.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsRecursivelyName())
                {
                    mDeleteFolder.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mDeleteFolder);
        }

        private void ExecuteDeleteFolder()
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

                Folder folder = GetFolderByName(context, list, folderName);

                if (!recursive && folder.ItemCount > 0)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderIsNotEmpty, folderName));
                }

                folder.DeleteObject();

                context.ExecuteQuery();
            }
        }
        #endregion

        #region RenameFolder
        private void AddRenameFolderMethod(ServiceObject so)
        {
            Method mRenameFolder = Helper.CreateMethod(Constants.Methods.RenameFolder, "Rename folder in list/library", MethodType.Read);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mRenameFolder);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName() || prop.IsDestinationFolder())
                {
                    mRenameFolder.InputProperties.Add(prop);
                    mRenameFolder.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mRenameFolder.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mRenameFolder);
        }

        private void ExecuteRenameFolder()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string destinationFolderName = base.GetStringProperty(Constants.SOProperties.DestinationFolder, true);

            folderName = folderName.Trim('/');
            
            DataRow dataRow = results.NewRow();
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.RootFolder, l => l.DefaultDisplayFormUrl);
                context.ExecuteQuery();

                Folder currentFolder = GetFolderByName(context, list, folderName);

                ListItem folderItem = currentFolder.ListItemAllFields;
                context.Load(folderItem);
                context.ExecuteQuery();

                folderItem[Constants.InternalProperties.FileLeafRef] = destinationFolderName.Trim('/').Split('/')[0];
                folderItem.Update();
                context.ExecuteQuery();

                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, folderItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;

                results.Rows.Add(dataRow);
            }
        }
        #endregion

        #region MoveFolder
        private void AddMoveFolderMethod(ServiceObject so)
        {
            Method mMoveFolder = Helper.CreateMethod(Constants.Methods.MoveFolder, "Move folder in list/library", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mMoveFolder);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName() || prop.IsDestinationURL() || prop.IsDestinationListLibrary())
                {
                    mMoveFolder.InputProperties.Add(prop);
                    mMoveFolder.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsDestinationFolder())
                {
                    mMoveFolder.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mMoveFolder.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mMoveFolder);
        }

        private void ExecuteMoveFolder()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string destinationSiteURL = base.GetStringProperty(Constants.SOProperties.DestinationURL, true);
            string destinationListLibrary = base.GetStringProperty(Constants.SOProperties.DestinationListLibrary, true);
            string destinationFolder = base.GetStringProperty(Constants.SOProperties.DestinationFolder, false);
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
                        
            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();

            using (ClientContext sourceContext = InitializeContext(siteURL))
            {
                List sourceList = sourceContext.Web.Lists.GetByTitle(listTitle);
                sourceContext.Load(sourceList);
                sourceContext.Load(sourceList, l => l.Fields, l => l.RootFolder);
                sourceContext.ExecuteQuery();

                //get currentFolder in source list
                Folder sourceFolder = GetFolderByName(sourceContext, sourceList, folderName);

                //create destination context
                using (ClientContext destContext = InitializeContext(destinationSiteURL))
                {
                    List destList = destContext.Web.Lists.GetByTitle(destinationListLibrary);
                    destContext.Load(destList);
                    destContext.Load(destList, l => l.RootFolder, l => l.DefaultDisplayFormUrl);
                    destContext.ExecuteQuery();

                    //get or create destination folder
                    Folder destFolder = CreateFolderInternal(destContext.Web, destList, destList.RootFolder, string.Empty, 
                        string.Concat(destinationFolder, '/', folderName.Split('/').Last()));

                    DataRow dataRow = results.NewRow();

                    CopyFolder(sourceContext, destContext, sourceList, destList, sourceFolder, destFolder);

                    destContext.Load(destFolder.ListItemAllFields);
                    destContext.ExecuteQuery();

                    string strurl = BuildListItemLink(destContext.Url, destList.DefaultDisplayFormUrl, destFolder.ListItemAllFields.Id);
                    dataRow[Constants.SOProperties.LinkToItem] = strurl;

                    results.Rows.Add(dataRow);
                }
            }
        }
        #endregion

        private bool CopyFolder(ClientContext sourceContext, ClientContext destContext, List sourceList, List destList, Folder sourceFolder, Folder destFolder)
        {
            IEnumerable<ListItem> sourceItems = FilterListItemsByServerRelativeUrl(sourceContext, sourceContext.Web, sourceList, new Properties(), 
                sourceFolder.ServerRelativeUrl, false, "Eq");

            foreach(ListItem sourceListItem in sourceItems)
            {
                ListItem destListItem = null;
                if(sourceList.BaseType != BaseType.DocumentLibrary)
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    itemCreateInfo.FolderUrl = string.Empty;
                    itemCreateInfo.FolderUrl = string.Format("{0}{1}", destContext.Url, destFolder.ServerRelativeUrl);

                    destListItem = destList.AddItem(itemCreateInfo);
                    destContext.ExecuteQuery();
                }
                else
                {
                    sourceContext.Load(sourceListItem.File, f => f.Name);
                    sourceContext.ExecuteQuery();
                    //get current file content
                    byte[] currentFileContent = GetFile(sourceContext, sourceListItem);

                    //add file
                    File newFile = AddFileByFolderPathWithRoot(destContext, destContext.Web, destList, sourceListItem.File.Name,
                        currentFileContent, destFolder.ServerRelativeUrl, true);

                    destListItem = newFile.ListItemAllFields;
                    destContext.Load(destListItem);
                    destContext.ExecuteQuery();
                }

                //copy metaData
                SPHelper.CopyItem(sourceContext, destContext, sourceListItem, destListItem);
            }
            FolderCollection childSourceFolders = sourceFolder.Folders;
            sourceContext.Load(childSourceFolders);
            sourceContext.ExecuteQuery();

            foreach(Folder childSourceFolder in childSourceFolders)
            {
                Folder childDestFolder = CreateSubFolderInternalWithoutRecursion(destContext, destList, destFolder, childSourceFolder.Name);
                CopyFolder(sourceContext, destContext, sourceList, destList, childSourceFolder, childDestFolder);
            }
            
            return true;
        }

        private Folder GetFolderByName(ClientContext context, List list, string folderName)
        {
            Folder currentFolder = null;

            try
            {
                currentFolder = context.Web.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName));
                context.Load(currentFolder);
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
            }

            return currentFolder;
        }

        private Folder CreateFolderInternal(Web web, List list, Folder parentFolder, string parentFolderPath, string fullFolderUrl)
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

        private Folder CreateSubFolderInternalWithoutRecursion(ClientContext context, List list, Folder parentFolder, string folderName)
        {
            string folderPath = string.Concat(parentFolder.ServerRelativeUrl, '/'+ folderName);

            Folder curFolder;
            try
            {
                //check is folder exists 
                curFolder = parentFolder.Folders.GetByUrl(parentFolder.ServerRelativeUrl);
                context.Load(curFolder);
                context.ExecuteQuery();
            }
            catch
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.LeafName = folderName;
                itemCreateInfo.FolderUrl = parentFolder.ServerRelativeUrl;
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                ListItem newItem = list.AddItem(itemCreateInfo);
                newItem[Constants.InternalProperties.Title] = folderName;
                newItem.Update();
                context.ExecuteQuery();

                //get current folder
                curFolder = parentFolder.Folders.GetByUrl(folderPath);
                context.Load(curFolder);
                context.ExecuteQuery();
            }
            return curFolder;
        }
    }
}
