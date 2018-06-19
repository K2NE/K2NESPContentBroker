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

            string ctName = String.Empty;

            ctName = base.GetStringProperty(Constants.SOProperties.ContentTypeName, true).Trim();

            if (String.IsNullOrWhiteSpace(ctName))
            {
                return;
            }


            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.ContentTypes;
                context.Load(ctColl);
                context.ExecuteQuery();

                var cType = ctColl.FirstOrDefault(c => c.Name == ctName);

                if (cType != null)
                {
                    DataRow dataRow = results.NewRow();

                    dataRow[Constants.SOProperties.ContentTypeName] = cType.Name;
                    dataRow[Constants.SOProperties.ContentTypeGroup] = cType.Group;
                    dataRow[Constants.SOProperties.ContentTypeReadOnly] = cType.ReadOnly;
                    dataRow[Constants.SOProperties.ContentTypeHidden] = cType.Hidden;
                    dataRow[Constants.SOProperties.ContentTypeCount] = 0;
                    dataRow[Constants.SOProperties.ContentTypeID] = cType.StringId;

                    results.Rows.Add(dataRow);
                }               
                
            }
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

            string ctId = String.Empty;

            ctId = base.GetStringProperty(Constants.SOProperties.ContentTypeID, true).Trim();

            if (String.IsNullOrWhiteSpace(ctId))
            {
                return;
            }


            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.ContentTypes;
                context.Load(ctColl);
                //context.ExecuteQuery();

                var cType = ctColl.GetById(ctId);
                context.Load(cType);
                context.ExecuteQuery();

                if (cType != null)
                {
                    DataRow dataRow = results.NewRow();

                    dataRow[Constants.SOProperties.ContentTypeName] = cType.Name;
                    dataRow[Constants.SOProperties.ContentTypeGroup] = cType.Group;
                    dataRow[Constants.SOProperties.ContentTypeReadOnly] = cType.ReadOnly;
                    dataRow[Constants.SOProperties.ContentTypeHidden] = cType.Hidden;
                    dataRow[Constants.SOProperties.ContentTypeCount] = 0;
                    dataRow[Constants.SOProperties.ContentTypeID] = cType.StringId;

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

           

            string siteURL = GetSiteURL();

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                context.Load(spWeb);

                var ctColl = spWeb.ContentTypes;
                context.Load(ctColl);
                //context.ExecuteQuery();               
                context.ExecuteQuery();


                foreach (var item in ctColl)
                {
                    DataRow dataRow = results.NewRow();

                    dataRow[Constants.SOProperties.ContentTypeName] = item.Name;
                    dataRow[Constants.SOProperties.ContentTypeGroup] = item.Group;
                    dataRow[Constants.SOProperties.ContentTypeReadOnly] = item.ReadOnly;
                    dataRow[Constants.SOProperties.ContentTypeHidden] = item.Hidden;
                    dataRow[Constants.SOProperties.ContentTypeCount] = 0;
                    dataRow[Constants.SOProperties.ContentTypeID] = item.StringId;

                    results.Rows.Add(dataRow);
                }               

            }
        }
        #endregion
    }
}
