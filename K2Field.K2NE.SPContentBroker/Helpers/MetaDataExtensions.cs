using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

namespace K2Field.K2NE.SPContentBroker.Helpers
{
    /// <summary>
    /// Helper class to allow additional values to be stored on a ServiceObject's metadata.
    /// </summary>
    internal static class MetaDataExtensions
    {
        public static void AddServiceElement(this MetaData metaData, string elementName, object elementValue)
        {
            if (metaData == null)
            {
                throw new ArgumentNullException("metaData");
            }

            metaData.ServiceProperties.Add(elementName, elementValue);
        }


        /// <summary>
        /// Retrieves a specific service element's metadata valu.
        /// </summary>
        /// <typeparam name="T">Type of the variable</typeparam>
        /// <param name="metaData"></param>
        /// <param name="elementName">Name to be retrieved</param>
        /// <returns>The value of the metadata element</returns>
        public static T GetServiceElement<T>(this MetaData metaData, string elementName)
        {
            if (metaData == null)
            {
                throw new ArgumentNullException("metaData");
            }

            if (string.IsNullOrEmpty(elementName))
            {
                throw new ArgumentNullException("elementName");
            }

            object value = metaData.ServiceProperties[elementName];
            if (string.IsNullOrEmpty(value.ToString()))
            {
                throw new Exception(string.Format("ServiceElement {0} is a required property, but it's string.empty.", elementName));
            }

            return (T)Convert.ChangeType(value, typeof(T));
        }

    }
}
