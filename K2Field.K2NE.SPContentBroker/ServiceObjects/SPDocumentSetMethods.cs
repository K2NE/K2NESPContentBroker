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
using Microsoft.SharePoint.Client.DocumentSet;

namespace K2Field.K2NE.SPContentBroker.ServiceObjects
{
    public partial class SPServiceObject : ServiceObjectBase
    {
        #region CreateDocSet

        private void AddCreateDocumentSetByNameMethod(ServiceObject so)
        {
            Method mCreateDocSet = Helper.CreateMethod(Constants.Methods.CreateDocumentSet, "Create a Document set by name", MethodType.Create);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCreateDocSet);
            }

            foreach (Property prop in so.Properties)
            {
                if (!prop.IsInternal() && !prop.IsFile() && !prop.IsReadOnly())
                {
                    mCreateDocSet.InputProperties.Add(prop);
                }

                if (prop.IsDocSetName())
                {
                    mCreateDocSet.InputProperties.Add(prop);
                    mCreateDocSet.Validation.RequiredProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName()))
                {
                    mCreateDocSet.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mCreateDocSet.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCreateDocSet);
        }
        private void ExecuteCreateDocSetByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string DocSetName = base.GetStringProperty(Constants.SOProperties.DocSetName, true);
            string FolderName = string.Empty;
            if (serviceObject.IsSPFolderEnabled())
            {
                FolderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            }

            DataRow dataRow = results.NewRow();
            using (ClientContext context = InitializeContext(GetSiteURL()))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                Folder parentFolder;
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.Load(list, d => d.DefaultDisplayFormUrl, d => d.ItemCount);
                context.Load(list.ContentTypes);
                context.ExecuteQuery();

                if (string.IsNullOrEmpty(FolderName))
                {
                    parentFolder = list.RootFolder;
                }
                else
                {
                    parentFolder = spWeb.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', FolderName));
                    context.Load(parentFolder);
                    context.ExecuteQuery();
                }

                Microsoft.SharePoint.Client.ClientResult<string> FileLink = new ClientResult<string>();
                foreach (ContentType ct in list.ContentTypes)
                {
                    if (ct.Id.StringValue.StartsWith(Constants.SharePointProperties.DocSetContentType))
                    {
                        FileLink = DocumentSet.Create(context, parentFolder, DocSetName, ct.Id);
                        break;
                    }
                }
                context.ExecuteQuery();

                Folder documentSet = null;
                documentSet = context.Web.GetFolderByServerRelativeUrl(FileLink.Value);
                context.Load(documentSet,c=>c.ListItemAllFields);
                context.ExecuteQuery();

                ListItem listItem = documentSet.ListItemAllFields;

                if (listItem == null)
                {
                    throw new Exception(Constants.ErrorMessages.RequiredDocNotFound);
                }

                foreach (Property prop in serviceObject.Properties)
                {
                    if (prop.Value != null && !prop.IsDocSetName() && !prop.IsFolderName())
                    {
                        Helpers.SPHelper.AssignFieldValue(listItem, prop);
                    }
                }
                listItem.Update();
                context.ExecuteQuery();

                if (FileLink != null)
                {
                    dataRow[Constants.SOProperties.LinkToItem] = FileLink.Value;
                    results.Rows.Add(dataRow);
                }

            }
        }
        #endregion
        #region UpdateDocSet
        private void AddUpdateDocumentSetByNameMethod(ServiceObject so)
        {
            Method mUpdateDocSet = Helper.CreateMethod(Constants.Methods.UpdateDocumentSetByName, "Update a Document set by name", MethodType.Update);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mUpdateDocSet);
            }

            foreach (Property prop in so.Properties)
            {
                if (!prop.IsInternal() && !prop.IsFile() && !prop.IsReadOnly())
                {
                    mUpdateDocSet.InputProperties.Add(prop);
                }

                if (prop.IsDocSetName())
                {
                    mUpdateDocSet.InputProperties.Add(prop);
                    mUpdateDocSet.Validation.RequiredProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName()))
                {
                    mUpdateDocSet.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mUpdateDocSet.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mUpdateDocSet);
        }
        private void ExecuteUpdateDocSetByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            DataRow dataRow = results.NewRow();
            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string DocSetName = base.GetStringProperty(Constants.SOProperties.DocSetName, true);
            string FolderName = string.Empty;
            if (serviceObject.IsSPFolderEnabled())
            {
                FolderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            }
           
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.Load(list, d => d.DefaultDisplayFormUrl, d => d.ItemCount);
                context.ExecuteQuery();

                if (list != null && list.ItemCount > 0)
                {
                    Folder documentSet = null;
                    if(string.IsNullOrEmpty(FolderName))
                    {
                        documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name,'/', DocSetName));
                    }
                    else
                    {
                        documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', FolderName, '/', DocSetName));
                    }
                   
                    context.Load(documentSet, c => c.ListItemAllFields);
                    context.ExecuteQuery();

                    ListItem listItem = documentSet.ListItemAllFields;

                    if (listItem == null)
                    {
                        throw new Exception(Constants.ErrorMessages.RequiredDocNotFound);
                    }

                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (prop.Value != null && !prop.IsDocSetName() && !prop.IsFolderName())
                        {
                            Helpers.SPHelper.AssignFieldValue(listItem, prop);
                        }
                    }
                    listItem.Update();
                    context.ExecuteQuery();


                    string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                    dataRow[Constants.SOProperties.LinkToItem] = strurl;
                    results.Rows.Add(dataRow);

                }
            }
        }
        #endregion
        #region GetDocSet
        private void AddGetDocSetMethod(ServiceObject so)
        {
            Method mGetDocSet = Helper.CreateMethod(Constants.Methods.GetDocumentSetByName, "Get Document set", MethodType.Read);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetDocSet);
            }

            foreach (Property prop in so.Properties)
            {
                if (!prop.IsInternal() &&
                    !prop.IsFile())
                {

                    mGetDocSet.ReturnProperties.Add(prop);
                }

                if (prop.IsDocSetName())
                {
                    mGetDocSet.InputProperties.Add(prop);
                    mGetDocSet.Validation.RequiredProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName()))
                {
                    mGetDocSet.InputProperties.Add(prop);
                }

                if (prop.IsLinkToItem())
                {
                    mGetDocSet.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mGetDocSet);
        }

        private void ExecuteGetDocSetByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            DataRow dataRow = results.NewRow();
            string listTitle = serviceObject.GetListTitle();
            string DocSetName = base.GetStringProperty(Constants.SOProperties.DocSetName, true);
            string FolderName = string.Empty;
            if (serviceObject.IsSPFolderEnabled())
            {
                FolderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            }


            using (ClientContext context = InitializeContext(GetSiteURL()))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.DefaultDisplayFormUrl);
                context.Load(list.RootFolder);
                context.Load(list.RootFolder.Folders);
                context.ExecuteQuery();

                Folder documentSet = null;
                if (string.IsNullOrEmpty(FolderName))
                {
                    documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', DocSetName));
                }
                else
                {
                    documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', FolderName, '/', DocSetName));
                }

                context.Load(documentSet, c => c.ListItemAllFields);
                context.ExecuteQuery();

                ListItem listItem = documentSet.ListItemAllFields;

              
                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                        {
                            Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
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
        #endregion
        #region GetDocumentSets
        private void AddGetDocumentSetsMethod(ServiceObject so)
        {
            Method mGetDocumentSets = Helper.CreateMethod(Constants.Methods.GetDocumentSets, "Retrieve document Sets", MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetDocumentSets);
            }

            foreach (Property prop in so.Properties)
            {
                if (prop.IsDocSetName())
                {
                    mGetDocumentSets.InputProperties.Add(prop);
                    mGetDocumentSets.ReturnProperties.Add(prop);
                }
                if (!prop.IsInternal() &&
                    !prop.IsFile())
                {
                    mGetDocumentSets.InputProperties.Add(prop);
                    mGetDocumentSets.ReturnProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName() || prop.IsRecursivelyName()))
                {
                    mGetDocumentSets.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mGetDocumentSets.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mGetDocumentSets);
        }

        private void ExecuteGetDocumentSets()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string DocSetName = base.GetStringProperty(Constants.SOProperties.DocSetName, false);
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName);
            bool recursive = base.GetBoolProperty(Constants.SOProperties.Recursively);
            string folderPath = string.Empty;
            Properties searchProperties = Helpers.SPHelper.GetSearchFields(serviceObject.Properties);

            if (!string.IsNullOrEmpty(DocSetName))
            {
                Property fileLeafRefProperty = NewProperty(Constants.InternalProperties.FileLeafRef, Constants.InternalProperties.FileLeafRef, true, SoType.Text, "");
                fileLeafRefProperty.Value = DocSetName;
                fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.Internal, false);
                fileLeafRefProperty.MetaData.ServiceProperties.Add(Constants.InternalProperties.SPFieldType, "Text");
                searchProperties.Add(fileLeafRefProperty);
            }

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.DefaultDisplayFormUrl);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, recursive, "BeginsWith", true);

                foreach (ListItem listItem in source)
                {
                    if (listItem.FileSystemObjectType == FileSystemObjectType.Folder)
                    {
                        DataRow dataRow = results.NewRow();
                        foreach (Property prop in serviceObject.Properties)
                        {
                            if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                            {
                                Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                            }
                            if (prop.IsDocSetName())
                            {
                                dataRow[prop.Name] = (string)listItem.FieldValues[Constants.SharePointProperties.FileLeafRef].ToString();
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

        #region RenameDocumentSet
        private void AddRenameDocumentSetMethod(ServiceObject so)
        {
            Method mRenameDocSet = Helper.CreateMethod(Constants.Methods.RenameDocumentSetByName, "Rename a Document set", MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mRenameDocSet);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsDocSetName() || prop.IsNewDocumentSetName())
                {
                    mRenameDocSet.InputProperties.Add(prop);
                    mRenameDocSet.Validation.RequiredProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName()))
                {
                    mRenameDocSet.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mRenameDocSet.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mRenameDocSet);
        }
        private void ExecuteRenameDocSet()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            DataRow dataRow = results.NewRow();
            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string docSetName = base.GetStringProperty(Constants.SOProperties.DocSetName, true);
            string newDocSetName = base.GetStringProperty(Constants.SOProperties.DocSetNewName, true);
            string FolderName = string.Empty;
            if (serviceObject.IsSPFolderEnabled())
            {
                FolderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            }
            

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.Load(list, d => d.DefaultDisplayFormUrl, d => d.ItemCount);
                context.ExecuteQuery();

                if (list != null && list.ItemCount > 0)
                {
                    Folder documentSet = null;
                    if (string.IsNullOrEmpty(FolderName))
                    {
                        documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', docSetName));
                    }
                    else
                    {
                        documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', FolderName, '/', docSetName));
                    }

                    context.Load(documentSet, c => c.ListItemAllFields);
                    context.ExecuteQuery();

                    ListItem listItem = documentSet.ListItemAllFields;

                    if (listItem == null)
                    {
                        throw new Exception(Constants.ErrorMessages.RequiredDocNotFound);
                    }

                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (prop.Value != null && prop.IsNewDocumentSetName())
                        {
                            listItem[Constants.SharePointProperties.FileLeafRef] = prop.Value;
                        }
                    }
                    listItem.Update();
                    context.ExecuteQuery();

                    string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                    dataRow[Constants.SOProperties.LinkToItem] = strurl;
                    results.Rows.Add(dataRow);

                }
            }
        }
        #endregion

        #region DeleteDocumentSet
        private void AddDeleteDocumentSetMethod(ServiceObject so)
        {
            Method mDeleteDocSet = Helper.CreateMethod(Constants.Methods.DeleteDocumentSetByName, "Delete a Document set", MethodType.Delete);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mDeleteDocSet);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsDocSetName())
                {
                    mDeleteDocSet.InputProperties.Add(prop);
                    mDeleteDocSet.Validation.RequiredProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName()))
                {
                    mDeleteDocSet.InputProperties.Add(prop);
                }
            }

            so.Methods.Add(mDeleteDocSet);
        }
        private void ExecuteDeleteDocSet()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;
            DataRow dataRow = results.NewRow();
            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string docSetName = base.GetStringProperty(Constants.SOProperties.DocSetName, true);
            string FolderName = string.Empty;
            if (serviceObject.IsSPFolderEnabled())
            {
                FolderName = base.GetStringProperty(Constants.SOProperties.FolderName, false);
            }
           
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, l => l.Fields, l => l.RootFolder);
                context.Load(list, d => d.ItemCount);
                context.ExecuteQuery();

                if (list != null && list.ItemCount > 0)
                {
                    Folder documentSet = null;
                    if (string.IsNullOrEmpty(FolderName))
                    {
                        documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', docSetName));
                    }
                    else
                    {
                        documentSet = context.Web.GetFolderByServerRelativeUrl(string.Concat(list.RootFolder.Name, '/', FolderName, '/', docSetName));
                    }

                    context.Load(documentSet, c => c.ListItemAllFields);
                    context.ExecuteQuery();

                    ListItem listItem = documentSet.ListItemAllFields;

                    if (listItem == null)
                    {
                        throw new Exception(Constants.ErrorMessages.RequiredDocNotFound);
                    }

                    listItem.DeleteObject();
                    context.ExecuteQuery();

                }
            }
        }
        #endregion
    }
}
