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
        #region Get Content Type by Name
        private void AddGetContentTypeByNameMethod(ServiceObject so)
        {
            Method mGetContentTypeByName = Helper.CreateMethod(Constants.Methods.GetContentTypeByName, "Retrieve metadata for one item by it's Name", MethodType.Read);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetContentTypeByName);
            }


            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ContentTypeName, true) == 0)
                {
                    mGetContentTypeByName.InputProperties.Add(prop);
                    mGetContentTypeByName.Validation.RequiredProperties.Add(prop);
                }

                mGetContentTypeByName.ReturnProperties.Add(prop);

            }

            so.Methods.Add(mGetContentTypeByName);
        }

        private void ExecuteGetContentTypeByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string ctName = base.GetStringProperty(Constants.SOProperties.ContentTypeName, true);

            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.AvailableContentTypes;
                context.Load(ctColl);
                context.ExecuteQuery();

                var cType = ctColl.FirstOrDefault(c => c.Name == ctName);

                if (cType != null)
                {
                    DataRow dataRow = results.NewRow();

                    PopulateDataRow(cType, dataRow);

                    results.Rows.Add(dataRow);
                }

            }
        }

        private static void PopulateDataRow(ContentType cType, DataRow dataRow)
        {
            dataRow[Constants.SOProperties.ContentTypeName] = cType.Name;
            dataRow[Constants.SOProperties.ContentTypeGroup] = cType.Group;
            dataRow[Constants.SOProperties.ContentTypeReadOnly] = cType.ReadOnly;
            dataRow[Constants.SOProperties.ContentTypeHidden] = cType.Hidden;
            dataRow[Constants.SOProperties.ContentTypeCount] = 0;
            dataRow[Constants.SOProperties.ContentTypeID] = cType.StringId;
        }
        #endregion

        #region Get Content Type by ID
        private void AddGetContentTypeByIdMethod(ServiceObject so)
        {
            Method mGetContentTypeById = Helper.CreateMethod(Constants.Methods.GetContentTypeById, "Retrieve metadata for one item by it's ID", MethodType.Read);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetContentTypeById);
            }


            foreach (Property prop in so.Properties)
            {
                if (string.Compare(prop.Name, Constants.SOProperties.ContentTypeID, true) == 0)
                {
                    mGetContentTypeById.InputProperties.Add(prop);
                    mGetContentTypeById.Validation.RequiredProperties.Add(prop);
                }

                mGetContentTypeById.ReturnProperties.Add(prop);

            }

            so.Methods.Add(mGetContentTypeById);
        }

        private void ExecuteGetContentTypeById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string ctId = base.GetStringProperty(Constants.SOProperties.ContentTypeID, true);


            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.AvailableContentTypes;
                context.Load(ctColl);
              

                var cType = ctColl.GetById(ctId);
                context.Load(cType);
                context.ExecuteQuery();

                if (cType != null)
                {
                    DataRow dataRow = results.NewRow();

                    PopulateDataRow(cType, dataRow);

                    results.Rows.Add(dataRow);
                }

            }
        }
        #endregion

        #region Get Content Types
        private void AddGetContentTypesMethod(ServiceObject so)
        {
            Method mGetContentTypeById = Helper.CreateMethod(Constants.Methods.GetContentTypes, "Retrieve metadata for one item by it's ID", MethodType.List);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetContentTypeById);
            }


            foreach (Property prop in so.Properties)
            {

                mGetContentTypeById.InputProperties.Add(prop);

                mGetContentTypeById.ReturnProperties.Add(prop);

            }

            so.Methods.Add(mGetContentTypeById);
        }

        private void ExecuteGetContentTypes()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string ctHidden = base.GetStringProperty(Constants.SOProperties.ContentTypeHidden, false);
            string ctGroup = base.GetStringProperty(Constants.SOProperties.ContentTypeGroup, false);

            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.AvailableContentTypes;
                context.Load(ctColl);

                IQueryable<ContentType> filterQuery = ctColl;
                filterQuery = ComposeQuery(ctHidden, ctGroup, string.Empty, filterQuery);

                IEnumerable<ContentType> _results = context.LoadQuery<ContentType>(filterQuery);
                              
                context.ExecuteQuery();


                foreach (var item in _results)
                {
                    DataRow dataRow = results.NewRow();

                    PopulateDataRow(item, dataRow);

                    results.Rows.Add(dataRow);
                }

            }
        }

        #endregion
        
        #region Get Content Types By Parent
        private void AddGetContentTypesByParentMethod(ServiceObject so)
        {
            Method mGetContentTypeById = Helper.CreateMethod(Constants.Methods.GetContentTypesByParent, "Retrieve metadata for one item by it's ID", MethodType.List);
            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetContentTypeById);
            }

            Helper.AddStringParameter(mGetContentTypeById, Constants.SOProperties.ContentTypeParent);

            foreach (Property prop in so.Properties)
            {
               
                mGetContentTypeById.ReturnProperties.Add(prop);

            }

            so.Methods.Add(mGetContentTypeById);
        }

        private void ExecuteGetContentTypesByParent()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string parentCtName = base.GetStringParameter(Constants.SOProperties.ContentTypeParent, true);          
            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.AvailableContentTypes;
                context.Load(ctColl);

                IQueryable<ContentType> filterQuery = ctColl.Where(ct => parentCtName == ct.Parent.Name);               

                IEnumerable<ContentType> _results = context.LoadQuery<ContentType>(filterQuery);
            
                context.ExecuteQuery();


                foreach (var item in _results)
                {
                    DataRow dataRow = results.NewRow();

                    PopulateDataRow(item, dataRow);

                    results.Rows.Add(dataRow);
                }

            }
        }
        #endregion

        #region private methods

        private static IQueryable<ContentType> ComposeQuery(string ctHidden, string ctGroup, string ctId, IQueryable<ContentType> filterQuery)
        {

            if (!String.IsNullOrWhiteSpace(ctHidden) && !String.IsNullOrWhiteSpace(ctGroup))
            {
                bool isHidden = Convert.ToBoolean(ctHidden.Trim());
                filterQuery = filterQuery.Where(ct => ct.Hidden == isHidden && ct.Group == ctGroup);
                return filterQuery;
            }
           
            else if(!String.IsNullOrWhiteSpace(ctHidden))
            {
                bool isHidden = Convert.ToBoolean(ctHidden.Trim());
                filterQuery = filterQuery.Where(ct => ct.Hidden == isHidden);
                return filterQuery;
            }

            else if(!String.IsNullOrWhiteSpace(ctGroup))
            {

                filterQuery = filterQuery.Where(ct => ct.Group == ctGroup);
                return filterQuery;
            }

            else { 
            return filterQuery;
            }

        }
        #endregion
    }
}
