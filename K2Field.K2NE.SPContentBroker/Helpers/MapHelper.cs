using System;
using System.Collections.Generic;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

namespace K2Field.K2NE.SPContentBroker.Helpers
{
    /// <summary>
    /// MapHelper class is a simple static helper class that's used to handle supportive
    /// methods on the TypeMappings class. The TypeMappings are use dto see what simple types we support.
    /// </summary>
    public static class MapHelper
    {
        #region Private Field And Filling method
        private static TypeMappings _map;

        private static Dictionary<string, SoType> spType2SoType = new Dictionary<string, SoType>() {
            {"Attachments", SoType.File},
            {"BusinessData", SoType.Text},
            {"Computed", SoType.Text},
            {"Currency", SoType.Decimal},
            {"File", SoType.File},
            {"Lookup", SoType.Text},
            {"LookupMulti", SoType.Text},
            {"Number", SoType.Number},
            //TaxonomyFieldType
            //TaxonomyFieldTypeMulti
            //User
            //UserMulti

        };


        private static Dictionary<SoType, string> soType2SystemTypeMapping = new Dictionary<SoType, string>() {
            {SoType.AutoGuid, typeof(Guid).ToString() },
            {SoType.Autonumber, typeof(int).ToString() },
            {SoType.DateTime, typeof(DateTime).ToString() },
            {SoType.Decimal, typeof(Decimal).ToString() },
            {SoType.Default, typeof(string).ToString() },
            //{SoType.File, typeof(byte[]).ToString() },
            {SoType.File, typeof(string).ToString() },
            {SoType.Guid, typeof(Guid).ToString() },
            {SoType.HyperLink, typeof(string).ToString() },
            {SoType.Image, typeof(byte[]).ToString() },
            {SoType.Memo, typeof(string).ToString() },
            {SoType.MultiValue, typeof(string).ToString() },
            {SoType.Number, typeof(int).ToString() },
            {SoType.Text, typeof(string).ToString() },
            {SoType.Xml, typeof(string).ToString() },
            {SoType.YesNo, typeof(bool).ToString() }
        };

        private static TypeMappings CreateTypeMappings()
        {
            TypeMappings map = new TypeMappings();
            map.Add(typeof(System.Int16), SoType.Number);
            map.Add(typeof(System.Int32), SoType.Number);
            map.Add(typeof(System.Int64), SoType.Number);
            map.Add(typeof(System.UInt16), SoType.Number);
            map.Add(typeof(System.UInt32), SoType.Number);
            map.Add(typeof(System.UInt64), SoType.Number);
            map.Add(typeof(System.Boolean), SoType.YesNo);
            map.Add(typeof(System.Char), SoType.Text);
            map.Add(typeof(System.DateTime), SoType.DateTime);
            map.Add(typeof(System.Decimal), SoType.Decimal);
            map.Add(typeof(System.Single), SoType.Decimal);
            map.Add(typeof(System.Double), SoType.Decimal);
            map.Add(typeof(System.Guid), SoType.Guid);
            //map.Add(typeof(System.Byte), SoType.File);
            //map.Add(typeof(System.SByte), SoType.File);
            map.Add(typeof(System.String), SoType.File);
            map.Add(typeof(System.String), SoType.Text);

            map.Add(typeof(Nullable<System.Int16>), SoType.Number);
            map.Add(typeof(Nullable<System.Int32>), SoType.Number);
            map.Add(typeof(Nullable<System.Int64>), SoType.Number);
            map.Add(typeof(Nullable<System.UInt16>), SoType.Number);
            map.Add(typeof(Nullable<System.UInt32>), SoType.Number);
            map.Add(typeof(Nullable<System.UInt64>), SoType.Number);
            map.Add(typeof(Nullable<System.Boolean>), SoType.YesNo);
            map.Add(typeof(Nullable<System.Char>), SoType.Text);
            map.Add(typeof(Nullable<System.DateTime>), SoType.DateTime);
            map.Add(typeof(Nullable<System.Decimal>), SoType.Decimal);
            map.Add(typeof(Nullable<System.Single>), SoType.Decimal);
            map.Add(typeof(Nullable<System.Double>), SoType.Decimal);
            map.Add(typeof(Nullable<System.Guid>), SoType.Guid);
            //map.Add(typeof(Nullable<System.Byte>), SoType.File);
            //map.Add(typeof(Nullable<System.SByte>), SoType.File);

            return map;
        }
        #endregion Private Field And Filling method

        #region Public properties
        /// <summary>
        /// Returns the TypeMappings that this object contains.
        /// </summary>
        public static TypeMappings Map
        {
            get
            {
                if (_map == null)
                {
                    _map = CreateTypeMappings();
                }
                return _map;
            }
        }
        #endregion Public properties

        #region Public methods

        /// <summary>
        /// Retrieves the SOType for the given .NET Type.
        /// </summary>
        public static SoType GetSoTypeByType(Type type)
        {
            return Map[type.FullName.ToLower()];
        }


        /// <summary>
        /// Retrieve the .NET Type (typeof(Type).toString()) for a given SOType.
        /// </summary>
        /// <param name="soType">the SOType</param>
        /// <returns>A typeof(T).toString() for the given SOType.</returns>
        public static string GetTypeBySoType(SoType soType)
        {
            return soType2SystemTypeMapping[soType];
        }

        /// <summary>
        /// Checks to see if the type is in the map. If it is, than this must be a simple type.
        /// </summary>
        public static bool IsSimpleMapableType(Type type)
        {
            if (Map.Contains(type.FullName.ToLower()))
            {
                return true;
            }
            return false;

        }
        #endregion Public methods


        internal static SoType SPTypeField(Microsoft.SharePoint.Client.Field fd, out Microsoft.SharePoint.Client.FieldType fieldtype)
        {
            if (fd.FieldTypeKind == Microsoft.SharePoint.Client.FieldType.Calculated)
            {
                Microsoft.SharePoint.Client.FieldCalculated calcVar = (Microsoft.SharePoint.Client.FieldCalculated)fd;
                return SPTypeToSoType(calcVar.OutputType, out  fieldtype);
            }            
            else
                return SPTypeToSoType(fd.FieldTypeKind, out fieldtype);

        }

        //Reference - https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.spfieldtype.aspx
       // https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listitem.aspx       
        // Other useful link https://msdn.microsoft.com/en-us/library/office/jj245356.aspx
        private static SoType SPTypeToSoType(Microsoft.SharePoint.Client.FieldType ft, out Microsoft.SharePoint.Client.FieldType _fieldtype)
        {
            _fieldtype = ft; //This is done so that Fieldtype of the field calculated is saved as its output type. This will help us at the time of data conversion
            
            switch (ft)
            {
                case Microsoft.SharePoint.Client.FieldType.Calculated: //Calculated field type is can be further broken to its output type
                    throw new Exception("Can't call SPTypeToSOType on Calculated field.");
                case Microsoft.SharePoint.Client.FieldType.ContentTypeId:
                    return SoType.Text;
                case Microsoft.SharePoint.Client.FieldType.Counter:
                    return SoType.Autonumber;
                case Microsoft.SharePoint.Client.FieldType.Boolean:
                    return SoType.YesNo;
                case Microsoft.SharePoint.Client.FieldType.Currency:
                    return SoType.Decimal;                            
                case Microsoft.SharePoint.Client.FieldType.DateTime:
                    return SoType.DateTime;
                case Microsoft.SharePoint.Client.FieldType.File: //Modifiying As the field will return only the reference value ref: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listitem.aspx
                    return SoType.File;
                case Microsoft.SharePoint.Client.FieldType.Guid:
                    return SoType.Guid;
                case Microsoft.SharePoint.Client.FieldType.Integer:
                    return SoType.Number;
                case Microsoft.SharePoint.Client.FieldType.Lookup:
                    return SoType.Text;                           
                case Microsoft.SharePoint.Client.FieldType.Note:     
                    return SoType.Memo;
                case Microsoft.SharePoint.Client.FieldType.Number:
                    return SoType.Decimal;
                case Microsoft.SharePoint.Client.FieldType.Text:
                    return SoType.Text;
                case Microsoft.SharePoint.Client.FieldType.URL:
                    return SoType.HyperLink;                   
                case Microsoft.SharePoint.Client.FieldType.User:
                    return SoType.Text;                          
                case Microsoft.SharePoint.Client.FieldType.Attachments:   //The field is just to indicate if there is any attachments  
                case Microsoft.SharePoint.Client.FieldType.AllDayEvent: //The field is just to indicate if calendar even is for all day
                    return SoType.YesNo;
                case Microsoft.SharePoint.Client.FieldType.MaxItems:
                    return SoType.Number;

                //TODO: take care of the following types
                // these all need to be tested and/or assinged a correct type.
                // It might be that some fields should never map - for example 'error' seems something that we would never expect.
                // It might also be that some items can return different types, for example Calculated or Computed might be string and int? Would that depend on the computation? If that's the case, then we simply can't put this mapping here in place.

                case Microsoft.SharePoint.Client.FieldType.Computed: //All custom columns are fieldType as Computed,but no option to find output type.
                    return SoType.Text;

                case Microsoft.SharePoint.Client.FieldType.Choice:      //Text
                case Microsoft.SharePoint.Client.FieldType.MultiChoice: //System.string[]
                case Microsoft.SharePoint.Client.FieldType.ThreadIndex:
                case Microsoft.SharePoint.Client.FieldType.Threading:   
                    return SoType.Text;              
                case Microsoft.SharePoint.Client.FieldType.ModStat:
                case Microsoft.SharePoint.Client.FieldType.WorkflowEventType:
                case Microsoft.SharePoint.Client.FieldType.WorkflowStatus:
                    return SoType.Number;
                case Microsoft.SharePoint.Client.FieldType.CrossProjectLink:
                case Microsoft.SharePoint.Client.FieldType.Recurrence:
                    return SoType.YesNo;
                case Microsoft.SharePoint.Client.FieldType.Invalid:    // Columns with typeAsString value as "Taxanomy , HashTags,Likes" have field type  Invalid.Thus datatype will remain as string but additional step during conversion.
                    return SoType.Text;
                //This has been defaulted to type string,But couldnot confirm thus if required this could be change.
                case Microsoft.SharePoint.Client.FieldType.Geolocation:
                case Microsoft.SharePoint.Client.FieldType.GridChoice:    //Output type is array but whether string or Int?          
                case Microsoft.SharePoint.Client.FieldType.OutcomeChoice:
                case Microsoft.SharePoint.Client.FieldType.PageSeparator: //Its a placeholder.
                case Microsoft.SharePoint.Client.FieldType.Error: //This might not be used in actual scenario
                    return SoType.Text;       
                default:
                    throw new NotSupportedException(string.Format("Unknown field type: {0}", ft));
                    //return SoType.Text;
                               
            }

        }
    }
}
