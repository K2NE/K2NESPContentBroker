using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using Microsoft.SharePoint.Client;

namespace K2Field.K2NE.SPContentBroker.Helpers
{
    /// <summary>
    /// The broker generally adds a ServiceObject per list/library. This happens in the DiscoverServiceObject method.
    /// The discoverServiceObject method also adds additional MetaData properties to the service object, so that we can 
    /// later use these properties to quickly understand what serviceObject it is. A good example is the list title as we use that often but
    /// do not want to constantly query sharepoint for this.
    /// </summary>
    internal static class ServiceObjectExtensions
    {

        public static string GetListTitle(this ServiceObject serviceObject)
        {
            return serviceObject.MetaData.GetServiceElement<string>(Constants.InternalProperties.ListTitle);
        }

        public static bool IsDocumentLibrary(this ServiceObject serviceObject)
        {
            return (serviceObject.MetaData.GetServiceElement<BaseType>(Constants.InternalProperties.ListBaseType) == BaseType.DocumentLibrary);
        }

        public static bool IsDocumentSetLibrary(this ServiceObject serviceObject)
        {
            return (serviceObject.MetaData.GetServiceElement<bool>(Constants.InternalProperties.IsListDocumentSet) == true);
        }

        public static bool IsSPFolderEnabled(this ServiceObject serviceObject)
        {
            return serviceObject.MetaData.GetServiceElement<bool>(Constants.InternalProperties.IsFolderEnabled);
        }
    }
}
