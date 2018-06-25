using System;
using System.Collections.Generic;
using System.Text;
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using SourceCode.Hosting.Server.Interfaces;
using Microsoft.SharePoint.Client;
using K2Field.K2NE.SPContentBroker.Helpers;
using K2Field.K2NE.SPContentBroker.ServiceObjects;

namespace K2Field.K2NE.SPContentBroker
{
    public class SPContentBroker : ServiceAssemblyBase, IHostableType
    {
        #region Private Properties
        private static readonly object serviceObjectToTypeLock = new object();
        private static readonly object serviceObjectLock = new object();
        private static Dictionary<string, Type> _serviceObjectToType = new Dictionary<string, Type>();
        private List<ServiceObjectBase> _serviceObjects;
        private object syncobject = new object();
        #endregion Private Properties


        #region Public properties for ServiceObjectBase's child classes.
        public Logger HostServiceLogger { get; private set; }
        public IIdentityService IdentityService { get; private set; }
        public ISecurityManager SecurityManager { get; private set; }
        #endregion Public properties for ServiceObjectBase's child classes.



        #region Private Methods




        /// <summary>
        /// Cache of all service object that we have. We load this into a static object in the hope to re-use it as often as possible.
        /// </summary>
        private IEnumerable<ServiceObjectBase> ServiceObjectClasses
        {
            get
            {
                if (_serviceObjects == null)
                {
                    lock (serviceObjectLock)
                    {
                        if (_serviceObjects == null)
                        {
                            _serviceObjects = new List<ServiceObjectBase>
                            {
                                new SPServiceObject(this)
                            };

                        }
                    }
                }
                return _serviceObjects;
            }
        }

        /// <summary>
        /// helper property to get the type of the service object, to be able to initialize a specific instance of it.
        /// </summary>
        private Dictionary<string, Type> ServiceObjectToType
        {
            get
            {
                lock (serviceObjectToTypeLock)
                {
                    _serviceObjectToType = new Dictionary<string, Type>();
                    foreach (ServiceObjectBase soBase in ServiceObjectClasses)
                    {

                        List<ServiceObject> serviceObjs = soBase.DescribeServiceObjects();
                        foreach (ServiceObject so in serviceObjs)
                        {
                            _serviceObjectToType.Add(so.Name, soBase.GetType());
                        }
                    }
                }
                return _serviceObjectToType;
            }
        }


        ///<summary>
        ///helper method to get the type of the service object, to be able to initialize a specific instance of it.
        ///</summary>
        private Type GetServiceObjectByType(string serviceObjectName)
        {
            string searchKey = string.Concat(this.Service.Guid.ToString(), "_", serviceObjectName);
            if (!_serviceObjectToType.ContainsKey(searchKey))
            {
                lock (serviceObjectToTypeLock)
                {

                    foreach (ServiceObjectBase soBase in ServiceObjectClasses)
                    {
                        List<ServiceObject> serviceObjs = soBase.DescribeServiceObjects();
                        foreach (ServiceObject so in serviceObjs)
                        {
                            _serviceObjectToType.Add(string.Concat(this.Service.Guid.ToString(), "_", so.Name), soBase.GetType());
                        }
                    }
                }
                return _serviceObjectToType[searchKey];
            }
            else
            {
                return _serviceObjectToType[searchKey];
            }
        }


        private ServiceFolder InitializeServiceFolder(string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException("folderName");
            }

            List<string> folderList = new List<string>(folderName.Split('\\'));

            return EnsureServiceFolder(Service.ServiceFolders, folderList);

        }

        private ServiceFolder EnsureServiceFolder(ServiceFolders serviceFolders, List<string> folderList)
        {
            string folderToEnsure = folderList[0];
            folderList.RemoveAt(0);

            foreach (ServiceFolder folder in serviceFolders)
            {
                if (string.Compare(folder.Name, folderToEnsure) == 0) // we found the folder
                {
                    if (folderList.Count == 0) // we're at the end of our path
                    {
                        return folder;
                    }

                    // we're not at the end of our path, so go deeper.
                    return EnsureServiceFolder(folder.ServiceFolders, folderList);
                }
            }

            // We didn't find the folder, so create it and go deeper.
            ServiceFolder newSf = new ServiceFolder(folderToEnsure, new MetaData(folderToEnsure, folderToEnsure));
            serviceFolders.Create(newSf);
            if (folderList.Count != 0)
            {
                return EnsureServiceFolder(newSf.ServiceFolders, folderList);
            }
            return newSf;

        }



        #endregion

        #region Constructor
        /// <summary>
        /// A new instance is called for every new connection that is created to the K2 server. A new instance of this class is not created
        /// when the connection remains open. One connection can be used for multiple things.
        /// </summary>
        public SPContentBroker()
        {
        }
        #endregion


        public string SiteURL
        {
            get
            {
                if (!Service.ServiceConfiguration.Contains(Constants.ConfigurationProperties.SiteUrl))
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.ConfigOptionNotFound, Constants.ConfigurationProperties.SiteUrl));
                }

                return Service.ServiceConfiguration[Constants.ConfigurationProperties.SiteUrl].ToString();
            }
        }


        #region Public overrides for ServiceAssemblyBase
        public override string DescribeSchema()
        {
            try
            {
                Service.Name = "K2NESPContentBroker";
                Service.MetaData.DisplayName = string.Format("K2NE's SPContentBroker - {0}", SiteURL);
                Service.MetaData.Description = "A Service Broker that integration with SharePoint CSOM.";
                ServicePackage.IsSuccessful = true;


                foreach (ServiceObjectBase entry in ServiceObjectClasses)
                {
                    List<ServiceObject> serviceObjects = entry.DescribeServiceObjects();
                    foreach (ServiceObject so in serviceObjects)
                    {
                        string serviceFolder = so.MetaData.GetServiceElement<string>(Constants.InternalProperties.ServiceFolder);
                        Service.ServiceObjects.Create(so);
                        string dictKey = string.Concat(this.Service.Guid.ToString(), "_", so.Name);
                        if (!_serviceObjectToType.ContainsKey(dictKey))
                        {
                            _serviceObjectToType.Add(dictKey, entry.GetType());
                        }

                        if (!string.IsNullOrEmpty(serviceFolder))
                        {
                            ServiceFolder sf = InitializeServiceFolder(serviceFolder);
                            sf.Add(so);
                        }
                    }
                }
                return base.DescribeSchema();
            }
            catch (Exception ex)
            {
                HandleToplevelException(ex);
            }
            return base.DescribeSchema();
        }
        public override void Execute()
        {
            ServiceObject so = Service.ServiceObjects[0];
            try
            {
                if (base.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.ServiceAccount)
                {
                    System.Security.Principal.WindowsIdentity.Impersonate(IntPtr.Zero);
                }

                //TODO: improve performance? http://bloggingabout.net/blogs/vagif/archive/2010/04/02/don-t-use-activator-createinstance-or-constructorinfo-invoke-use-compiled-lambda-expressions.aspx

                // This creates an instance of the object responsible to handle the execution.
                // We can't cache the instance itself, as that gives threading issue because the 
                // object can be re-used by the k2 host server for multiple different SMO calls
                // so we always need to know which ServiceObject we actually want to execute and 
                // create an instance first. This is  "late" initalization. We can also not keep a list of 
                // service objects that have been instanciated around in memory as this would be to resource 
                // intensive and slow (as we would constantly initialize all).
                if (so == null || string.IsNullOrEmpty(so.Name))
                {
                    throw new ApplicationException("ServiceObject is not set.");
                }

                Type soType = GetServiceObjectByType(so.Name);
                object[] constParams = new object[] { this };
                ServiceObjectBase soInstance = Activator.CreateInstance(soType, constParams) as ServiceObjectBase;

                soInstance.Execute();
                ServicePackage.IsSuccessful = true;
            }
            catch (Exception ex)
            {
                HandleToplevelException(ex);
            }
        }

        private void HandleToplevelException(Exception ex)
        {
            StringBuilder error = new StringBuilder();
            error.AppendFormat("Exception.Message: {0}\n", ex.Message);
            error.AppendFormat("Exception.StackTrace: {0}\n", ex.StackTrace);

            Exception innerEx = ex;
            int i = 0;
            while (innerEx.InnerException != null)
            {
                error.AppendFormat("{0} InnerException.Message: {1}\n", i, innerEx.InnerException.Message);
                error.AppendFormat("{0} InnerException.StackTrace: {1}\n", i, innerEx.InnerException.StackTrace);
                innerEx = innerEx.InnerException;
                i++;
            }
            ServicePackage.ServiceMessages.Add(error.ToString(), MessageSeverity.Error);
            ServicePackage.IsSuccessful = false;
        }
        public override string GetConfigSection()
        {
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.SiteUrl, true, "https://sharepointsite/");
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.IsDynamic, true, false);
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.RequestTimeout, true, 180000);
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.IncludeHiddenLists, true, false);
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.IncludeHiddenFields, true, false);
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.ShowInternalNames, true, true);
            Service.ServiceConfiguration.Add(Constants.ConfigurationProperties.Office365, true, false);
            return base.GetConfigSection();
        }
        public void Init(IServiceMarshalling serviceMarshalling, IServerMarshaling serverMarshaling)
        {
            lock (syncobject)
            {
                if (HostServiceLogger == null)
                {
                    HostServiceLogger = new Logger(serviceMarshalling.GetHostedService(typeof(SourceCode.Logging.ILogger)) as SourceCode.Logging.ILogger);
                    HostServiceLogger.LogDebug("Logger loaded from ServiceMarshalling");
                }

                if (IdentityService == null)
                {
                    IdentityService = serviceMarshalling.GetHostedService(typeof(IIdentityService)) as IIdentityService;
                }
                if (SecurityManager == null)
                {
                    SecurityManager = serverMarshaling.GetSecurityManagerContext();
                }

            }


        }
        public override void Extend() { }
        public void Unload()
        {
            HostServiceLogger.Dispose();
        }
        #endregion Public overrides for ServiceAssemblyBase


    }
}
