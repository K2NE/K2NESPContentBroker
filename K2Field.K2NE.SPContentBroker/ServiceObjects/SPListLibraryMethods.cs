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
        #region GetItemById
        private void AddGetItemByIdMethod(ServiceObject so)
        {
            Method mGetItemById = Helper.CreateMethod(Constants.Methods.GetItemById, "Retrieve metadata for one item by it's ID", MethodType.Read);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetItemById);
            }
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mGetItemById.InputProperties.Add(prop);
                    mGetItemById.Validation.RequiredProperties.Add(prop);
                }

                if (!prop.IsInternal() && !prop.IsFile())
                {
                    mGetItemById.ReturnProperties.Add(prop);

                }
                if (so.IsDocumentLibrary() && prop.IsFileName())
                {
                    mGetItemById.ReturnProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mGetItemById.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mGetItemById);
        }

        private void ExecuteGetItemById()
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
                context.Load(list, i => i.DefaultDisplayFormUrl);
                context.Load(listItem);
                context.Load(listItem, i => i.File);
                context.ExecuteQuery();
                string test = list.DefaultDisplayFormUrl;
                
                DataRow dataRow = results.NewRow();

                foreach (Property prop in serviceObject.Properties)
                {
                    if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                    {
                        Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                    }
                    if (prop.IsFileName())
                    {
                        var fileName = (string)listItem.File.Name;
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
        #endregion

        #region CreateItem
        private void AddCreateItemMethod(ServiceObject so)
        {
            Method mCreateItem = Helper.CreateMethod(Constants.Methods.CreateItem, "Create a list item", MethodType.Create);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mCreateItem);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0 && !prop.IsInternal())
                {
                    mCreateItem.ReturnProperties.Add(prop);
                }

                if (!prop.IsReadOnly() && !prop.IsInternal())
                {
                    mCreateItem.InputProperties.Add(prop);
                    if (prop.IsRequired())
                    {
                        mCreateItem.Validation.RequiredProperties.Add(prop);
                    }
                }
                if (so.IsSPFolderEnabled() == true &&
                    prop.IsFolderName())
                {
                    mCreateItem.InputProperties.Add(prop);
                }

                if (prop.IsLinkToItem())
                {
                    mCreateItem.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mCreateItem);
        }

        private void ExecuteCreateItem()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string folderPath = base.GetStringProperty(Constants.SOProperties.FolderName);

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
                                
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.FolderUrl = string.Empty;
                if (!string.IsNullOrEmpty(folderPath))
                {
                    itemCreateInfo.FolderUrl = string.Format("{0}/lists/{1}/{2}", siteURL, list.Title, folderPath.Trim('/'));
                }

                newItem = list.AddItem(itemCreateInfo);

                foreach (Property prop in serviceObject.Properties)
                {
                    if (prop.Value != null && !prop.IsFile() && !prop.IsFolderName())
                    {
                        Helpers.SPHelper.AssignFieldValue(newItem, prop);
                    }
                }

                newItem.Update();
                context.ExecuteQuery();

                dataRow[Constants.SOProperties.ID] = newItem.Id;
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, newItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;

                results.Rows.Add(dataRow);
            }
        }
        #endregion

        #region UpdateItemById
        private void AddUpdateItemByIdMethod(ServiceObject so)
        {
            Method mUpdateItemById = Helper.CreateMethod(Constants.Methods.UpdateItemById, "Update list or library metadata by it's ID", MethodType.Update);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mUpdateItemById);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mUpdateItemById.InputProperties.Add(prop);
                    mUpdateItemById.Validation.RequiredProperties.Add(prop);
                    mUpdateItemById.ReturnProperties.Add(prop);
                }
                if (!prop.IsReadOnly() && !prop.IsInternal() && !prop.IsFile())
                {
                    mUpdateItemById.InputProperties.Add(prop);
                }
                if (prop.IsLinkToItem())
                {
                    mUpdateItemById.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mUpdateItemById);
        }

        private void ExecuteUpdateItemById()
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
                context.Load(list, i => i.DefaultDisplayFormUrl);
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
                string strurl = BuildListItemLink(context.Url, list.DefaultDisplayFormUrl, listItem.Id);
                dataRow[Constants.SOProperties.LinkToItem] = strurl;
                results.Rows.Add(dataRow);
            }
        }
        #endregion

        #region DeleteItemById
        private void AddDeleteItemByIdMethod(ServiceObject so)
        {
            Method mDeleteItemById = Helper.CreateMethod(Constants.Methods.DeleteItemById, "Delete one item by it's ID", MethodType.Delete);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mDeleteItemById);
            }

            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mDeleteItemById.InputProperties.Add(prop);
                    mDeleteItemById.Validation.RequiredProperties.Add(prop);
                }
            }

            so.Methods.Add(mDeleteItemById);
        }

        private void ExecuteDeleteItemById()
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

        #region GetItemByTitle
        private void AddGetItemByTitleMethod(ServiceObject so)
        {
            Method mGetItemByTitle = Helper.CreateMethod(Constants.Methods.GetItemByTitle, "Retrieve metadata for one item by it's Title", MethodType.List);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetItemByTitle);
            }
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.Title, true) == 0)
                {
                    mGetItemByTitle.InputProperties.Add(prop);
                    mGetItemByTitle.Validation.RequiredProperties.Add(prop);
                }

                if (so.IsSPFolderEnabled() == true &&(
                    prop.IsFolderName() || prop.IsRecursivelyName()))
                {
                    mGetItemByTitle.InputProperties.Add(prop);
                }

                if ((!prop.IsInternal() && !prop.IsFile()) || prop.IsLinkToItem())
                {
                    mGetItemByTitle.ReturnProperties.Add(prop);

                }
            }

            so.Methods.Add(mGetItemByTitle);
        }

        private void ExecuteGetItemByTitle()
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
                context.Load(list, l => l.Fields, l => l.DefaultDisplayFormUrl);
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

                CamlQuery camlQuery = Helpers.SPHelper.CreateCamlQuery(searchProperties, "Eq", list, recursive);
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
        }
        #endregion

        #region GetItemByName
        private void AddGetItemByNameMethod(ServiceObject so)
        {
            Method mGetItemByName = Helper.CreateMethod(Constants.Methods.GetItemByName, "Retrieve metadata for one item by it's document name", MethodType.List);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetItemByName);
            }
            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.FileName, true) == 0)
                {
                    mGetItemByName.InputProperties.Add(prop);
                    mGetItemByName.Validation.RequiredProperties.Add(prop);
                }

                if (so.IsSPFolderEnabled() == true && (
                    prop.IsFolderName() || prop.IsRecursivelyName()))
                {
                    mGetItemByName.InputProperties.Add(prop);
                }

                if ((!prop.IsInternal() && !prop.IsFile()) || prop.IsLinkToItem() || (so.IsDocumentLibrary() && prop.IsFileName()))
                {
                    mGetItemByName.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mGetItemByName);
        }

        private void ExecuteGetItemByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName);
            bool recursive = base.GetBoolProperty(Constants.SOProperties.Recursively);
            string folderPath = string.Empty;

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
                context.Load(list, l => l.Fields, l => l.DefaultDisplayFormUrl);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, recursive, "Eq");
                
                foreach (ListItem listItem in source)
                {
                    DataRow dataRow = results.NewRow();

                    context.Load(listItem, i => i.File);
                    context.ExecuteQuery();

                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                        {
                            Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                        }
                        if (prop.IsFileName())
                        {
                            var fileName = (string)listItem.File.Name;
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
        #endregion

        #region GetItems
        private void AddGetItemsMethod(ServiceObject so)
        {
            Method mGetItems = Helper.CreateMethod(Constants.Methods.GetItems, "Retrieve items list", MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetItems);
            }

            foreach (Property prop in so.Properties)
            {
                if (!prop.IsInternal() &&
                    !prop.IsFile())
                {
                    mGetItems.InputProperties.Add(prop);
                    mGetItems.ReturnProperties.Add(prop);
                }
                if (so.IsSPFolderEnabled() && (prop.IsFolderName() || prop.IsRecursivelyName()))
                {
                    mGetItems.InputProperties.Add(prop);
                }
                if ((so.IsDocumentLibrary() && prop.IsFileName()) || prop.IsLinkToItem())
                {
                    mGetItems.ReturnProperties.Add(prop);
                }
            }

            so.Methods.Add(mGetItems);
        }

        private void ExecuteGetItems()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName);
            bool recursive = base.GetBoolProperty(Constants.SOProperties.Recursively);
            string listtype = serviceObject.MetaData.GetServiceElement<string>(Constants.InternalProperties.ListBaseType).ToString();
            Properties searchProperties = Helpers.SPHelper.GetSearchFields(serviceObject.Properties);
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.Fields, l => l.DefaultDisplayFormUrl);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                IEnumerable<ListItem> source = FilterListItems(context, spWeb, list, searchProperties, folderName, recursive, "BeginsWith");
                
                foreach (ListItem listItem in source)
                {
                    DataRow dataRow = results.NewRow();
                    
                    foreach (Property prop in serviceObject.Properties)
                    {
                        if (listItem.FieldValues.ContainsKey(prop.Name) && !prop.IsFile())
                        {
                            Helpers.SPHelper.AddFieldValue(dataRow, prop, listItem);
                        }
                        if (prop.IsFileName() && serviceObject.IsDocumentLibrary())
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
        #endregion

        private IEnumerable<ListItem> FilterListItems(ClientContext context, Web spWeb, List list, Properties searchProperties, string folderName, bool recursive, string compareType, bool IsDocumentLibrary = false)
        {
            string folderPath = String.Empty;
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

            CamlQuery camlQuery = Helpers.SPHelper.CreateCamlQuery(searchProperties, compareType, list, recursive);
            camlQuery.FolderServerRelativeUrl = folderPath;

            IQueryable<ListItem> queryable;
            if (IsDocumentLibrary)
            {
                //populating file for all the listitems.
                queryable = list.GetItems(camlQuery).IncludeWithDefaultProperties<ListItem>(item => item.ContentType, item => item.File);
            }
            else
            {
                queryable = list.GetItems(camlQuery).IncludeWithDefaultProperties<ListItem>(item => item.ContentType);
            }

            foreach (Field f in list.Fields)
            {
                queryable.Concat<ListItem>((IEnumerable<ListItem>)queryable.Include<ListItem>(item => item[f.Title]));
            }

            IEnumerable<ListItem> source = context.LoadQuery(queryable);
            context.ExecuteQuery();
            return source;
        }

        private IEnumerable<ListItem> FilterListItemsByServerRelativeUrl(ClientContext context, Web spWeb, List list, Properties searchProperties, string folderServerRelativeUrl, bool recursive, string compareType)
        {
            string folderPath = String.Empty;
            if (!string.IsNullOrEmpty(folderServerRelativeUrl))
            {
                folderPath = folderServerRelativeUrl;
                Folder folder;
                try
                {
                    folder = spWeb.GetFolderByServerRelativeUrl(folderPath);
                    context.Load(folder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderServerRelativeUrl), ex);
                }
            }
            else
            {
                folderPath = list.RootFolder.ServerRelativeUrl;
            }

            CamlQuery camlQuery = Helpers.SPHelper.CreateCamlQuery(searchProperties, compareType, list, recursive);
            camlQuery.FolderServerRelativeUrl = folderPath;

            IQueryable<ListItem> queryable = list.GetItems(camlQuery).IncludeWithDefaultProperties<ListItem>(item => item.ContentType);

            foreach (Field f in list.Fields)
            {
                queryable.Concat<ListItem>((IEnumerable<ListItem>)queryable.Include<ListItem>(item => item[f.Title]));
            }

            IEnumerable<ListItem> source = context.LoadQuery(queryable);
            context.ExecuteQuery();

            return source;
        }
    }
}
