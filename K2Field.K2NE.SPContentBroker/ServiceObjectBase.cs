using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.Hosting.Client.BaseAPI;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using SourceCode.EnvironmentSettings.Client;
using System.Threading;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Data.SqlClient;
using K2Field.K2NE.SPContentBroker.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.Security;

namespace K2Field.K2NE.SPContentBroker
{
    public abstract class ServiceObjectBase
    {

        #region Protected Methods and properties that are useful for the child class
        /// <summary>
        /// The serviceBroker object that is currently being used/executed. Property values (etc) will be taken from this object.
        /// </summary>
        protected SPContentBroker ServiceBroker
        {
            get;
            private set;
        }





        public string SiteURL
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.SiteUrl))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.SiteUrl));
                }

                return this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.SiteUrl].ToString();
            }
        }

        public int RequestTimeout
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.RequestTimeout))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.RequestTimeout));
                }

                return int.Parse(this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.RequestTimeout].ToString());
            }
        }

        public bool IncludeHiddenLists
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.IncludeHiddenLists))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.IncludeHiddenLists));
                }

                return bool.Parse(this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.IncludeHiddenLists].ToString());
            }
        }


        public bool IncludeHiddenFields
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.IncludeHiddenFields))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.IncludeHiddenFields));
                }

                return bool.Parse(this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.IncludeHiddenFields].ToString());
            }
        }

        public bool ShowInternalNames
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.ShowInternalNames))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.ShowInternalNames));
                }

                return bool.Parse(this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.ShowInternalNames].ToString());
            }
        }


        public bool IsDynamicSiteURL
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.IsDynamic))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.IsDynamic));
                }

                return bool.Parse(this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.IsDynamic].ToString());
            }
        }

        public bool Office365
        {
            get
            {
                if (!this.ServiceBroker.Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.Office365))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.Office365));
                }

                return bool.Parse(this.ServiceBroker.Service.ServiceConfiguration[Constants.ConfigurationProperties.Office365].ToString());
            }
        }

        protected ClientContext InitializeContext(string siteUrl)
        {
            ClientContext context = new ClientContext(siteUrl);
            context.RequestTimeout = this.RequestTimeout;
            try
            {

                switch (this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode)
                {
                    case AuthenticationMode.OAuth:
                        context.AuthenticationMode = ClientAuthenticationMode.Anonymous;
                        context.FormDigestHandlingEnabled = false;
                        context.ExecutingWebRequest +=
                            delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                            {
                                webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] = string.Concat("Bearer ", this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.OAuthToken);
                            };
                        break;

                    case AuthenticationMode.Static:
                        if (!Office365)
                        {
                            if (this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName.Contains("\\"))
                            {
                                string[] splits = this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName.Split('\\');
                                string domain = splits[0];
                                string user = splits[1];
                                context.Credentials = new System.Net.NetworkCredential(user, this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password, domain);
                            }
                            else
                            {
                                context.Credentials = new System.Net.NetworkCredential(this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName, this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password);
                            }
                            context.ExecutingWebRequest += context_ExecutingWebRequest;
                        }
                        else
                        {
                            var login = this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName;
                            var password = this.ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password;
                            var securePassword = new SecureString();
                            foreach (char c in password)
                            {
                                securePassword.AppendChar(c);
                            }

                            context.Credentials = new SharePointOnlineCredentials(login, securePassword);
                        }
                        

                        break;
                    case AuthenticationMode.Impersonate:
                    case AuthenticationMode.ServiceAccount:
                        if (!Office365)
                        {
                            context = new ClientContext(siteUrl);
                        }
                        else
                        {
                            throw new Exception(Constants.ErrorMessages.SAImpersonateWithOffice365NotSupported);
                        }
                        break;

                    default:
                        throw new ArgumentException("Authentication mode is not supported.");


                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to initialize SP Context", ex);
            }
            return context;
        }

        void context_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            try
            {
                e.WebRequestExecutor.RequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to add X-FORMS_BASED_AUTH_ACCEPT header to request headers", ex);
            }

        }




        /// <summary>
        /// Returns the FQN for the user calling the SMO.
        /// </summary>
        protected string CallingFQN
        {
            get
            {
                //TODO: FQN string might not read thread, or at least support the other authentication modes.
                switch (ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode)
                {
                    case AuthenticationMode.Impersonate:
                    case AuthenticationMode.ServiceAccount:
                        return ServiceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName;
                    default:
                        return "K2:" + System.Threading.Thread.CurrentPrincipal.Identity.Name;
                }

            }
        }


        #region Protected helper methods for property value retrieval
        /// <summary>
        /// Returns the value of the string property with the given value from the current SErviceObject.
        /// Method should always be called in the context of a 'Execute()'.
        /// 
        /// Setting isRequired might cause this method to throw an exception if the property was not found or the value is string empty.
        /// 
        /// If isRequired is not set, String.Empty is returned when the property does not exist.
        /// </summary>
        /// <param name="name">Name of the property to retrieve.</param>
        /// <param name="isRequired"></param>
        /// <returns></returns>

        protected string GetStringProperty(string name, bool isRequired = false)
        {
            Property p = ServiceBroker.Service.ServiceObjects[0].Properties[name];
            if (p == null)
            {
                if (isRequired)
                    throw new ArgumentException(string.Format(Constants.ErrorMessages.RequiredPropertyNotFound, name));
                return string.Empty;
            }
            string val = p.Value as string;
            if (isRequired && string.IsNullOrEmpty(val))
                throw new ArgumentException(string.Format("{0} is required but is empty.", name));

            return val;
        }



        protected string GetStringParameter(string name, bool isRequired = false)
        {
            MethodParameter p = ServiceBroker.Service.ServiceObjects[0].Methods[0].MethodParameters[name];
            if (p == null)
            {
                if (isRequired)
                    throw new ArgumentException(string.Format(Constants.ErrorMessages.RequiredParameterNotFound, name));
                return string.Empty;
            }
            string val = p.Value as string;
            if (isRequired && string.IsNullOrEmpty(val))
                throw new ArgumentException(string.Format("{0} is required but is empty.", name));

            return val;
        }


        /// <summary>
        /// Retrieve an integer property with the given name from the current ServiceObject.
        /// Method should always be called in the context of a 'Execute()'.
        /// 
        /// The isRequired value is optional. Setting it to true will cause an exception if the property is not found or
        /// if the value could not be parsed to an integer.
        /// If false or nothing is supplied, the method returns 0 (zero) when it cannot determin the value.
        /// </summary>
        /// <param name="name">Name of the property to retrieve.</param>
        /// <param name="isRequred">Specify if the field must exist and needs to be parsable to int.</param>
        /// <returns>0 or the value of the property.</returns>
        protected int GetIntProperty(string name, bool isRequred = false)
        {
            Property p = ServiceBroker.Service.ServiceObjects[0].Properties[name];
            if (p == null)
            {
                if (isRequred)
                    throw new ArgumentException(string.Format(Constants.ErrorMessages.RequiredPropertyNotFound, name));
                return 0;
            }
            string val = p.Value as string;
            int ret;
            if (int.TryParse(val, out ret))
                return ret;
            if (isRequred)
                throw new ArgumentException(string.Format("{0} could not be parsed to a Integer", name));

            return 0;
        }

        protected short GetShortProperty(string name, bool isRequired = false)
        {
            Property p = ServiceBroker.Service.ServiceObjects[0].Properties[name];
            if (p == null)
            {
                if (isRequired)
                    throw new ArgumentException(string.Format(Constants.ErrorMessages.RequiredPropertyNotFound, name));
                return 0;
            }
            string val = p.Value as string;
            short ret;
            if (short.TryParse(val, out ret))
                return ret;
            if (isRequired)
                throw new ArgumentException(string.Format("{0} could not be parsed to a Integer", name));

            return 0;
        }

        protected bool GetBoolProperty(string name)
        {
            Property p = ServiceBroker.Service.ServiceObjects[0].Properties[name];
            if (p == null)
            {
                return false;
            }
            string val = p.Value as string;
            bool ret;

            if (string.IsNullOrEmpty(val))
            {
                return false;
            }

            //bool.TryParse() always returns false for these values.
            if (string.Compare(val.Trim(), "1") == 0 || string.Compare(val.Trim(), "yes") == 0)
            {
                return true;
            }

            if (bool.TryParse(val, out ret))
            {
                return ret;
            }
            return false;
        }

        protected byte GetByteProperty(string name, bool isRequired = false)
        {
            Property p = ServiceBroker.Service.ServiceObjects[0].Properties[name];
            if (p == null)
            {
                if (isRequired)
                    throw new ArgumentException(string.Format(Constants.ErrorMessages.RequiredPropertyNotFound, name));
                return 0;
            }
            string val = p.Value as string;
            byte ret;
            if (byte.TryParse(val, out ret))
                return ret;
            if (isRequired)
                throw new ArgumentException(string.Format("{0} could not be parsed to a Byte.", name));
            return 0;
        }

        protected Guid GetGuidProperty(string name, bool isRequired = false)
        {
            Property p = ServiceBroker.Service.ServiceObjects[0].Properties[name];
            if (p == null)
            {
                if (isRequired)
                    throw new ArgumentException(string.Format(Constants.ErrorMessages.RequiredPropertyNotFound, name));
                return Guid.Empty;
            }
            string val = p.Value as string;
            Guid ret;
            if (Guid.TryParse(val, out ret))
                return ret;
            if (isRequired)
                throw new ArgumentException(string.Format("{0} could not be parsed to a Guid.", name));
            return Guid.Empty;
        }

        #endregion Protected helper methods for property value retrieval


        #endregion Protected Methods and properties that are useful for the child class


        #region Abstract methods
        /// <summary>
        /// Public method, required because we want to set _serviceBroker.
        /// </summary>
        /// <param name="broker"></param>
        public ServiceObjectBase(SPContentBroker broker)
        {
            ServiceBroker = broker;
        }

        /// <summary>
        /// Method to return ServiceObjects. This is then used to describe service objects in the main class.
        /// A list is returned, so you can return multiple service objects.
        /// </summary>
        /// <returns></returns>
        public abstract List<ServiceObject> DescribeServiceObjects();



        /// <summary>
        /// Gets called when executing the service object.
        /// </summary>
        public abstract void Execute();

        #endregion Abstract methods




    }
}
