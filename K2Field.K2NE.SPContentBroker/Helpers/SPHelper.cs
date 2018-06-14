using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Linq;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using System.Globalization;
using Microsoft.SharePoint.Taxonomy;

namespace K2Field.K2NE.SPContentBroker.Helpers
{
    public class SPHelper
    {
        public static void CopyItem(ClientContext sourceContext, ClientContext destinationContext, ListItem sourceItem, ListItem destinationItem)
        {
            sourceContext.Load(sourceItem.ParentList, l => l.Fields);
            sourceContext.ExecuteQuery();
            destinationContext.Load(destinationItem.ParentList, l => l.Fields);
            destinationContext.ExecuteQuery();

            foreach (Field destinationField in destinationItem.ParentList.Fields)
            {
                if (string.Compare(destinationField.InternalName, Constants.SharePointProperties.Attachments) != 0 && string.Compare(destinationField.InternalName, Constants.SharePointProperties.ContentType) != 0)
                {
                        Field sourceField = sourceItem.ParentList.Fields.Where(f => string.Compare(f.InternalName, destinationField.InternalName) == 0
                        && f.FieldTypeKind == destinationField.FieldTypeKind && !f.ReadOnlyField && !f.Hidden).FirstOrDefault();

                        if (sourceField != null)
                        {
                                destinationItem[destinationField.InternalName] = sourceItem[destinationField.InternalName];
                        }
                }
            }
            destinationItem.Update();
            destinationContext.ExecuteQuery();
        }

        public static void AddFieldValue(DataRow newRow, Property prop, ListItem listItem)
        {


            object fieldValue = listItem.FieldValues[prop.Name];
            if (fieldValue == null)
            {
                return;
            }
            FieldType ft = prop.MetaData.GetServiceElement<FieldType>(Constants.InternalProperties.FieldTypeKind);

            switch (ft)
            {
                case FieldType.AllDayEvent:
                case FieldType.Attachments:
                case FieldType.CrossProjectLink:
                case FieldType.Recurrence:
                case FieldType.Boolean:
                    newRow[prop.Name] = Convert.ToBoolean(fieldValue);
                    break;
                case FieldType.DateTime:
                    newRow[prop.Name] = TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(fieldValue.ToString()), TimeZoneInfo.Local);
                    break;
                case FieldType.Counter:
                case FieldType.MaxItems:
                case FieldType.ModStat:
                case FieldType.WorkflowEventType:
                case FieldType.WorkflowStatus:
                case FieldType.Integer:
                    newRow[prop.Name] = Convert.ToInt32(fieldValue);
                    break;
                case FieldType.Number:
                case FieldType.Currency:
                    newRow[prop.Name] = Decimal.Parse(fieldValue.ToString(), NumberStyles.Any, NumberFormatInfo.CurrentInfo);
                    break;
                case FieldType.Note:
                case FieldType.Computed:
                    newRow[prop.Name] = fieldValue.ToString();
                    break;              
                case FieldType.ThreadIndex:
                case FieldType.Threading:
                case FieldType.File: //This coulmn contains only reference of the actual file TODO : Either change the SOType in Maphelper or we might have to change the code here.
                    //newRow[prop.Name] = fieldValue.ToString();
                    break;
                case FieldType.Guid:
                    newRow[prop.Name] = Guid.Parse(fieldValue.ToString());
                    break;
                case FieldType.ContentTypeId:
                    newRow[prop.Name] = Convert.ToString(fieldValue);
                    break;
                case FieldType.Lookup:
                    //TODO: Move LookupMulti to constant
                    if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType), "LookupMulti") == 0)
                    {
                        addMultiLookupFieldValue(newRow, prop, fieldValue);
                    }
                    else
                    {
                        FieldLookupValue fieldLookupValue = fieldValue as FieldLookupValue;
                        if (fieldLookupValue != null)
                        {
                            if (!string.IsNullOrEmpty(fieldLookupValue.LookupValue))
                            {
                                newRow[prop.Name] = string.Format("{0};#{1}", fieldLookupValue.LookupId, fieldLookupValue.LookupValue);
                                newRow[prop.Name + Constants.InternalProperties.Suffix_Value] = fieldLookupValue.LookupValue.ToString();
                            }
                            newRow[prop.Name + Constants.InternalProperties.Suffix_ID] = fieldLookupValue.LookupId.ToString();
                        }

                    }

                    break;
                case FieldType.User:
                    if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType), "UserMulti") == 0)
                    {
                        addMultiUserFieldValue(newRow, prop, fieldValue);
                    }
                    else
                    {
                        FieldUserValue fieldUserValue = fieldValue as FieldUserValue;
                        if (fieldUserValue != null)
                        {
                            if (!string.IsNullOrEmpty(fieldUserValue.LookupValue))
                            {
                                newRow[prop.Name] = string.Format("{0};#{1}", fieldUserValue.LookupId, fieldUserValue.LookupValue);
                                newRow[prop.Name + Constants.InternalProperties.Suffix_Value] = fieldUserValue.LookupValue.ToString();
                            }
                            newRow[prop.Name + Constants.InternalProperties.Suffix_ID] = fieldUserValue.LookupId.ToString();
                        }

                    }

                    break;
                case FieldType.URL:
                    HyperlinkProperty hyperlinkProperty = new HyperlinkProperty();
                    FieldUrlValue fieldUrlValue = fieldValue as FieldUrlValue;
                    if (fieldUrlValue.Url.Length > 0)
                    {
                        hyperlinkProperty.Display = fieldUrlValue.Description;
                        hyperlinkProperty.Link = fieldUrlValue.Url;
                        newRow[prop.Name] = hyperlinkProperty.Value.ToString();
                    }
                    break;
                //Not Required as Fieldtype is saved according to its outPutType
                //case FieldType.Calculated:
                //    FieldCalculated calcVar = fieldValue as FieldCalculated;
                //    AddCalculatedFieldValue(newRow,prop,calcVar.OutputType,fieldValue);
                //    break;
                case FieldType.MultiChoice:
                    string[] multiChoice = fieldValue as string[];
                    if (multiChoice.Length > 0)
                    {
                        string UserMultiValue = string.Empty;
                        StringBuilder strMulti = new StringBuilder();
                        foreach (string field in multiChoice)
                        {
                            strMulti = strMulti.AppendFormat("{0};", field.ToString());
                        }
                        if (strMulti.Length > 0)
                        {
                            UserMultiValue = strMulti.ToString().Substring(0, strMulti.Length - 1);
                        }
                        newRow[prop.Name] = UserMultiValue.ToString();
                    }
                    break;
                case FieldType.Invalid:
                     if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType), Constants.SharePointProperties.TaxanomyFieldType) == 0)
                    {                       
                        addTaxonomyFieldValue(newRow, prop, fieldValue);
                    }
                    else if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType), Constants.SharePointProperties.TaxanomyFieldTypeMulti) == 0)
                    {
                        addMultiTaxonomyFieldValue(newRow, prop, fieldValue);
                    }
                    else
                    {
                        newRow[prop.Name] = fieldValue.ToString();
                    }
                    break;
                //TODO case FieldType.Invalid: Depending upon TypeAsString value. But this needs to be implemented case by case basis.
                case FieldType.Text:
                case FieldType.Choice:
                default:
                    newRow[prop.Name] = fieldValue.ToString();
                    break;

            }
        }

        private static void addMultiLookupFieldValue(DataRow newRow, Property prop, object fieldValue)
        {
            string LookupMultiComplete = string.Empty;
            string LookupMultiValue = string.Empty;
            string LookupMultiID = string.Empty;
            StringBuilder strMultiComplete = new StringBuilder();
            StringBuilder strMultiID = new StringBuilder();
            StringBuilder strMultiValue = new StringBuilder();
            FieldLookupValue[] fieldLookupValueArray = fieldValue as FieldLookupValue[];
            foreach (FieldLookupValue field in fieldLookupValueArray)
            {
                if (field != null)
                {
                    if (!string.IsNullOrEmpty(field.LookupValue))
                    {
                        strMultiComplete = strMultiComplete.AppendFormat("{0};#{1};", field.LookupId, field.LookupValue);

                        strMultiValue = strMultiValue.AppendFormat("{0};", field.LookupValue);
                    }
                    strMultiID = strMultiID.AppendFormat("{0};", field.LookupId);
                }
            }
            if (strMultiComplete.Length > 0)
            {
                LookupMultiComplete = strMultiComplete.ToString().Substring(0, strMultiComplete.Length - 1);

            }
            if (strMultiID.Length > 0)
            {
                LookupMultiID = strMultiID.ToString().Substring(0, strMultiID.Length - 1);
            }
            if (strMultiValue.Length > 0)
            {
                LookupMultiValue = strMultiValue.ToString().Substring(0, strMultiValue.Length - 1);
            }
            newRow[prop.Name] = LookupMultiComplete;
            newRow[prop.Name + Constants.InternalProperties.Suffix_ID] = LookupMultiID;
            newRow[prop.Name + Constants.InternalProperties.Suffix_Value] = strMultiValue;
        }

        private static void addMultiUserFieldValue(DataRow newRow, Property prop, object fieldValue)
        {
            string LookupMultiComplete = string.Empty;
            string LookupMultiValue = string.Empty;
            string LookupMultiID = string.Empty;
            StringBuilder strMultiComplete = new StringBuilder();
            StringBuilder strMultiID = new StringBuilder();
            StringBuilder strMultiValue = new StringBuilder();

            FieldUserValue[] fieldUserValueArray = fieldValue as FieldUserValue[];
            foreach (FieldUserValue field in fieldUserValueArray)
            {
                if (field != null)
                {
                    if (!string.IsNullOrEmpty(field.LookupValue))
                    {
                        strMultiComplete = strMultiComplete.AppendFormat("{0};#{1};", field.LookupId, field.LookupValue);

                        strMultiValue = strMultiValue.AppendFormat("{0};", field.LookupValue);
                    }
                    strMultiID = strMultiID.AppendFormat("{0};", field.LookupId);
                }

            }

            if (strMultiComplete.Length > 0)
            {
                LookupMultiComplete = strMultiComplete.ToString().Substring(0, strMultiComplete.Length - 1);

            }
            if (strMultiID.Length > 0)
            {
                LookupMultiID = strMultiID.ToString().Substring(0, strMultiID.Length - 1);
            }
            if (strMultiValue.Length > 0)
            {
                LookupMultiValue = strMultiValue.ToString().Substring(0, strMultiValue.Length - 1);
            }
            newRow[prop.Name] = LookupMultiComplete;
            newRow[prop.Name + Constants.InternalProperties.Suffix_ID] = LookupMultiID;
            newRow[prop.Name + Constants.InternalProperties.Suffix_Value] = strMultiValue;
        }

        private static void addTaxonomyFieldValue(DataRow newRow, Property prop, object FieldValue)
        {
             TaxonomyFieldValue taxanomyValue = FieldValue as TaxonomyFieldValue;

             if (taxanomyValue != null)
            {
                newRow[prop.Name] = taxanomyValue.TermGuid;
                newRow[prop.Name + Constants.InternalProperties.Suffix_Value] = taxanomyValue.Label;
            }
        }

        private static void addMultiTaxonomyFieldValue(DataRow newRow, Property prop, object FieldValue)
        {
            string LookupMultiComplete = string.Empty;
            string LookupMultiValue = string.Empty;
            StringBuilder strMultiComplete = new StringBuilder();
            StringBuilder strMultiValue = new StringBuilder();
           
            TaxonomyFieldValueCollection taxanomyValues = FieldValue as TaxonomyFieldValueCollection;

            foreach (TaxonomyFieldValue taxanomyValue in taxanomyValues)
            {
                if (taxanomyValue != null)
                {
                    strMultiComplete = strMultiComplete.AppendFormat("{0};", taxanomyValue.TermGuid);

                    strMultiValue = strMultiValue.AppendFormat("{0};", taxanomyValue.Label);
                }
            }

            if (strMultiComplete.Length > 0)
            {
                LookupMultiComplete = strMultiComplete.ToString().Substring(0, strMultiComplete.Length - 1);

            }

            if (strMultiValue.Length > 0)
            {
                LookupMultiValue = strMultiValue.ToString().Substring(0, strMultiValue.Length - 1);
            }

            newRow[prop.Name] = LookupMultiComplete;
            newRow[prop.Name + Constants.InternalProperties.Suffix_Value] = strMultiValue;

        }

        //private void addTaxonomyFieldValue(DataRow newRow, Property prop, object fieldValue)
        //{
        //    string[] strArray = fieldValue.ToString().Split('|');
        //    newRow[prop.Name] = string.Concat(strArray.GetValue(1).ToString(), ";#", strArray.GetValue(0).ToString());
        //}

        public static void AssignFieldValue(ListItem listItem, Property prop)
        {
            FieldType ft = (FieldType)prop.MetaData.GetServiceElement<FieldType>(Constants.InternalProperties.FieldTypeKind);
            // https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.fieldtype.aspx
            switch (ft)
            {
                case FieldType.AllDayEvent:
                case FieldType.Attachments:
                case FieldType.CrossProjectLink:
                case FieldType.Recurrence:
                case FieldType.Boolean:
                    listItem[prop.Name] = Convert.ToBoolean(prop.Value);
                    break;
                case FieldType.DateTime:

                    listItem[prop.Name] = TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(prop.Value.ToString()), TimeZoneInfo.Local);
                    break;
                case FieldType.Counter:
                case FieldType.MaxItems:
                case FieldType.ModStat:
                case FieldType.WorkflowEventType:
                case FieldType.WorkflowStatus:
                case FieldType.Integer:
                    listItem[prop.Name] = Convert.ToInt32(prop.Value);
                    break;
                case FieldType.Number:
                case FieldType.Currency:
                    listItem[prop.Name] = Decimal.Parse(prop.Value.ToString(), NumberStyles.Any, NumberFormatInfo.CurrentInfo);
                    break;
                case FieldType.Note:
                case FieldType.Computed:
                case FieldType.ThreadIndex:
                case FieldType.Threading:
                case FieldType.File:
                    listItem[prop.Name] = prop.Value.ToString();
                    break;
                case FieldType.Guid:
                    listItem[prop.Name] = Guid.Parse(prop.Value.ToString());
                    break;
                case FieldType.Lookup:
                    //TODO: Move LookupMulti to constant
                    if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType), "LookupMulti") == 0)
                    {
                        assignLookupFieldTypeMulti(listItem, prop);
                    }
                    else
                    {
                        int LookupId = -1;
                        if (int.TryParse(prop.Value.ToString(), out LookupId))
                        {
                            listItem[(prop.Name).Substring(0, (prop.Name).Length - 2)] = LookupId;//This is done as Property name has a ID suffix to its name
                        }
                    }
                    break;
                case FieldType.User:
                    //TODO: Move UserMulti to constant
                    if (string.Compare(prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType), "UserMulti") == 0)
                    {
                        assignUserFieldTypeMulti(listItem, prop);
                    }
                    else
                    {
                        listItem[(prop.Name).Substring(0, (prop.Name).Length - 2)] = getUserId(prop.Value.ToString());//This is done as Property name has a ID suffix to its name
                    }
                    break;
                case FieldType.URL:
                    if (!string.IsNullOrEmpty(prop.Value as string))
                    {
                        FieldUrlValue _url = new FieldUrlValue();
                        HyperlinkProperty hProp = prop as HyperlinkProperty;
                        if (prop != null)
                        {
                            _url.Url = hProp.Link;
                            _url.Description = hProp.Display;
                            listItem[prop.Name] = _url;
                        }
                    }
                    break;
                case FieldType.MultiChoice:
                    string[] multiChoice = prop.Value.ToString().Split(new char[] { ';' });
                    listItem[prop.Name] = multiChoice;
                    break;
                //TODO case FieldType.Invalid: Depending upon TypeAsString value. But this needs to be implemented case by case basis.
                case FieldType.Text:
                case FieldType.Choice:
                default:
                    listItem[prop.Name] = prop.Value.ToString();
                    break;
            }
        }

        public static Properties GetSearchFields(Properties inputs)
        {
            Properties properties = new Properties();
            foreach (Property input in inputs)
            {
                if (input.Value != null)
                {
                    properties.Add(input);
                }
            }
            return properties;
        }


        public static CamlQuery CreateCamlQuery(Properties props, string comparison, List list)
        {

            StringBuilder queryBuilder = new StringBuilder();
            queryBuilder.Append("<View>");
            if (list.Fields.Count > 0)
            {
                queryBuilder.Append("<ViewFields>");
                foreach (Field field in (IEnumerable<Field>)list.Fields)
                {
                    queryBuilder.AppendFormat("<FieldRef Name=\"{0}\" />", field.InternalName);
                }
                queryBuilder.Append("</ViewFields>");
            }
            if (props != null && props.Count != 0)
            {
                buildQuery(props, comparison, queryBuilder);
            }
            queryBuilder.Append("</View>");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = queryBuilder.ToString();
            return camlQuery;
        }

        public static CamlQuery CreateCamlQuery(Properties props, string comparison, List list, bool recursive)
        {

            StringBuilder queryBuilder = new StringBuilder();

            if (recursive)
            {
                queryBuilder.Append("<View Scope='Recursive'>");
            }
            else
            {
                queryBuilder.Append("<View Scope='FilesOnly'>");
            }

            if (list.Fields.Count > 0)
            {
                queryBuilder.Append("<ViewFields>");
                foreach (Field field in (IEnumerable<Field>)list.Fields)
                {
                    queryBuilder.AppendFormat("<FieldRef Name=\"{0}\" />", field.InternalName);
                }
                queryBuilder.Append("</ViewFields>");
            }
            if (props != null && props.Count != 0)
            {
                buildQuery(props, comparison, queryBuilder);
            }
            queryBuilder.Append("</View>");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = queryBuilder.ToString();
            return camlQuery;
        }


        private static void buildQuery(Properties props, string comparison, StringBuilder sb)
        {

            string typeAsString = string.Empty;

            sb.Append("<Query>");
            sb.Append("<Where>");


            int searchPropertiesCount = props.Where(item => item.Value != null && item.MetaData.GetServiceElement<Boolean>(Constants.InternalProperties.Internal) == false).Count();
            
            if (searchPropertiesCount > 1)
            {
                for (int i = 1; i < searchPropertiesCount; i++)
                {
                    sb.Append("<And>");
                }
            }

            int propCount = 1;
            foreach (Property prop in props)
            {
                if (prop.Value != null && prop.MetaData.GetServiceElement<Boolean>(Constants.InternalProperties.Internal) == false)
                {
                    string tmpValue = prop.Value.ToString();

                    //TODO: move to fieldType enum
                    typeAsString = prop.MetaData.GetServiceElement<string>(Constants.InternalProperties.SPFieldType);

                    string str = GetComparison(comparison, typeAsString, tmpValue, out tmpValue);
                    sb.AppendFormat("<{0}>", str);
                    sb.AppendFormat("<FieldRef Name=\"{0}\" />", prop.Name);
                    sb.AppendFormat("<Value Type=\"{0}\">{1}</Value>", typeAsString, tmpValue);
                    sb.AppendFormat("</{0}>", str);
                    
                    if (propCount > 1 && propCount != searchPropertiesCount)
                    {
                        sb.Append("</And>");
                    }
                    propCount++;
                }
            }
            if (searchPropertiesCount > 1)
            {
                sb.Append("</And>");
            }
            sb.Append("</Where>");
            sb.Append("</Query>");
        }


        private static string GetComparison(string comparison, string typeAsString, string fieldValue, out string returnValue)
        {
            if (fieldValue.StartsWith("*", StringComparison.OrdinalIgnoreCase) && fieldValue.EndsWith("*", StringComparison.OrdinalIgnoreCase))
            {
                fieldValue = fieldValue.Replace("*", string.Empty);
                comparison = "Contains";
            }
            if (fieldValue.EndsWith("*", StringComparison.OrdinalIgnoreCase))
            {
                fieldValue = fieldValue.Replace("*", string.Empty);
                comparison = "BeginsWith";
            }
            if (typeAsString.Equals("Boolean", StringComparison.OrdinalIgnoreCase) || typeAsString.Equals("Attachments", StringComparison.OrdinalIgnoreCase))
            {
                if (!fieldValue.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    comparison = "Neq";
                }
                else
                {
                    comparison = "Eq";
                }
                fieldValue = "0";
            }
            if (typeAsString.Equals("DateTime", StringComparison.OrdinalIgnoreCase) || typeAsString.Equals("Computed", StringComparison.OrdinalIgnoreCase))
            {
                comparison = "Eq";
            }
            returnValue = fieldValue;
            return comparison;
        }


        private static void assignLookupFieldTypeMulti(ListItem listItem, Property prop)
        {
            string[] strLookUpValues = prop.Value.ToString().Split(new char[] { ';' });
            List<FieldLookupValue> fieldLookupValues = new List<FieldLookupValue>(strLookUpValues.Length);

            for (int i = 0; i < strLookUpValues.Length; i++)
            {
                if (!string.IsNullOrEmpty(strLookUpValues[i]))
                {
                    FieldLookupValue fieldLookupValue = new FieldLookupValue()
                    {
                        LookupId = int.Parse(strLookUpValues[i])
                    };
                    fieldLookupValues.Add(fieldLookupValue);
                }
            }
            listItem[(prop.Name).Substring(0, (prop.Name).Length - 2)] = fieldLookupValues.ToArray(); //This is done as Property name has a ID suffix to its name
        }

        private static void assignUserFieldTypeMulti(ListItem listItem, Property prop)
        {
            string[] strUserValues = prop.Value.ToString().Split(new char[] { ';' });
            List<FieldUserValue> fieldUserValues = new List<FieldUserValue>(strUserValues.Length);

            for (int i = 0; i < strUserValues.Length; i++)
            {
                if (!string.IsNullOrEmpty(strUserValues[i]))
                {
                    FieldUserValue fieldUserValue = new FieldUserValue()
                    {
                        LookupId = getUserId(strUserValues[i])
                    };

                    fieldUserValues.Add(fieldUserValue);
                }
            }
            listItem[(prop.Name).Substring(0, (prop.Name).Length - 2)] = fieldUserValues.ToArray(); //This is done as Property name has a ID suffix to its name
        }

        private static int getUserId(string userString)
        {

            if (userString.Contains(";#"))
            {
                return int.Parse(userString.Split(";#".ToCharArray()).GetValue(0).ToString());
            }
            else
            {
                return int.Parse(userString);
            }
        }

        private static HyperlinkProperty getHyperLinkValue(Property prop)
        {
            return prop as HyperlinkProperty;
        }

    }
}
