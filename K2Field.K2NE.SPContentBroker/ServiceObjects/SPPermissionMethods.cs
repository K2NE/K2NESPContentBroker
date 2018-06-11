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
        #region ItemPermissions

        #region BreakItemInheritanceById
        private void AddBreakItemInheritanceByIdMethod(ServiceObject so)
        {
            Method mBreakItemInheritanceById = Helper.CreateMethod(Constants.Methods.BreakItemInheritanceById, "Break item inheritance in list/library by id", 
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mBreakItemInheritanceById);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mBreakItemInheritanceById.InputProperties.Add(prop);
                    mBreakItemInheritanceById.Validation.RequiredProperties.Add(prop);
                }
            }
            so.Methods.Add(mBreakItemInheritanceById);
        }

        private void ExecuteBreakItemInheritanceById()
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
                context.Load(listItem);
                context.ExecuteQuery();

                listItem.BreakRoleInheritance(false, true);
                context.ExecuteQuery();
            }
        }
        #endregion

        #region ResetItemInheritanceById
        private void AddResetItemInheritanceByIdMethod(ServiceObject so)
        {
            Method mResetItemInheritanceById = Helper.CreateMethod(Constants.Methods.ResetItemInheritanceById, "Reset item inheritance in list/library by id",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mResetItemInheritanceById);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mResetItemInheritanceById.InputProperties.Add(prop);
                    mResetItemInheritanceById.Validation.RequiredProperties.Add(prop);
                }
            }
            so.Methods.Add(mResetItemInheritanceById);
        }

        private void ExecuteResetItemInheritanceById()
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
                context.Load(listItem);
                context.ExecuteQuery();

                listItem.ResetRoleInheritance();
                context.ExecuteQuery();
            }
        }
        #endregion

        #region AddItemPermissionById
        private void AddAddItemPermissionByIdMethod(ServiceObject so)
        {
            Method mAddItemPermissionById = Helper.CreateMethod(Constants.Methods.AddItemPermissionById, "Add item permissions in list/library by id",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mAddItemPermissionById);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0 || prop.IsPermissionColumn())
                {
                    mAddItemPermissionById.InputProperties.Add(prop);
                    mAddItemPermissionById.Validation.RequiredProperties.Add(prop);
                }
                if(prop.IsUserLogins() || prop.IsGroupLogins())
                {
                    mAddItemPermissionById.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mAddItemPermissionById);
        }

        private void ExecuteAddItemPermissionById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string userLogins = base.GetStringProperty(Constants.SOProperties.UserLogins);
            string groupLogins = base.GetStringProperty(Constants.SOProperties.GroupLogins);
            string permission = base.GetStringProperty(Constants.SOProperties.Permission, true);

            if(string.IsNullOrEmpty(userLogins) && string.IsNullOrEmpty(groupLogins))
            {
                throw new ApplicationException(Constants.ErrorMessages.UserGroupLoginsAreEmptyException);
            }

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(listItem);
                context.ExecuteQuery();
                
                var role = new RoleDefinitionBindingCollection(context);
                role.Add(context.Web.RoleDefinitions.GetByName(permission));

                if (!string.IsNullOrEmpty(userLogins))
                {
                    foreach (User user in GetUsersByLoginString(context, userLogins))
                    {
                        listItem.RoleAssignments.Add(user, role);
                        listItem.Update();
                        context.ExecuteQuery();
                    }
                }

                if (!string.IsNullOrEmpty(groupLogins))
                {
                    foreach (Group group in GetGroupsByLoginString(context, groupLogins))
                    {
                        listItem.RoleAssignments.Add(group, role);
                        listItem.Update();
                        context.ExecuteQuery();
                    }
                }
            }
        }
        #endregion

        #region RemoveItemPermissionById
        private void AddRemoveItemPermissionByIdMethod(ServiceObject so)
        {
            Method mRemoveItemPermissionById = Helper.CreateMethod(Constants.Methods.RemoveItemPermissionById, "Remove item permissions in list/library by id",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mRemoveItemPermissionById);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0)
                {
                    mRemoveItemPermissionById.InputProperties.Add(prop);
                    mRemoveItemPermissionById.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsUserLogins() || prop.IsGroupLogins())
                {
                    mRemoveItemPermissionById.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mRemoveItemPermissionById);
        }

        private void ExecuteRemoveItemPermissionById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string userLogins = base.GetStringProperty(Constants.SOProperties.UserLogins);
            string groupLogins = base.GetStringProperty(Constants.SOProperties.GroupLogins);

            if (string.IsNullOrEmpty(userLogins) && string.IsNullOrEmpty(groupLogins))
            {
                throw new ApplicationException(Constants.ErrorMessages.UserGroupLoginsAreEmptyException);
            }

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(listItem);
                context.ExecuteQuery();

                RoleAssignmentCollection roles = listItem.RoleAssignments;
                context.Load(roles, role => role.Include(roleassigned => roleassigned.Member.LoginName, roleassigned => roleassigned.Member));
                context.ExecuteQuery();

                if (!string.IsNullOrEmpty(userLogins))
                {
                    foreach (User user in GetUsersByLoginString(context, userLogins))
                    {
                        context.Load(user, u => u.LoginName);
                        context.ExecuteQuery();
                        
                        foreach (RoleAssignment role in roles)
                        {
                            if (role.Member.LoginName == user.LoginName)
                            {
                                role.DeleteObject();
                                break;
                            }
                        }
                    }
                }

                if (!string.IsNullOrEmpty(groupLogins))
                {
                    foreach(Group group in GetGroupsByLoginString(context, groupLogins))
                    {
                        context.Load(group, g => g.LoginName);
                        context.ExecuteQuery();

                        foreach (RoleAssignment role in roles)
                        {
                            if (role.Member.LoginName == group.LoginName)
                            {
                                role.DeleteObject();
                                break;
                            }
                        }
                    }
                }

                context.ExecuteQuery();
            }
        }
        #endregion

        #region GetItemPermissionById
        private void AddGetItemPermissionByIdMethod(ServiceObject so)
        {
            Method mGetItemPermissionById = Helper.CreateMethod(Constants.Methods.GetItemPermissionById, "Get item permissions in list/library by id",
                MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetItemPermissionById);
            }

            foreach (Property prop in so.Properties)
            {

                if (string.Compare(prop.Name, Constants.SOProperties.ID, true) == 0 || prop.IsUserOrGroup())
                {
                    mGetItemPermissionById.InputProperties.Add(prop);
                    mGetItemPermissionById.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsPermissionColumn())
                {
                    mGetItemPermissionById.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mGetItemPermissionById);
        }

        private void ExecuteGetItemPermissionById()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            int id = base.GetIntProperty(Constants.SOProperties.ID, true);
            string userOrGroup = base.GetStringProperty(Constants.SOProperties.UserOrGroup, true);
            string permissions = string.Empty;
            
            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;
                
                List list = spWeb.Lists.GetByTitle(listTitle);
                ListItem listItem = list.GetItemById(id);
                context.Load(listItem);
                context.ExecuteQuery();

                RoleAssignmentCollection roles = listItem.RoleAssignments;
                context.Load(roles, role => role.Include(roleassigned => roleassigned.Member.LoginName, roleassigned => roleassigned.Member));
                context.ExecuteQuery();

                foreach (RoleAssignment role in roles)
                {
                    if (role.Member.LoginName == GetUserOrGroupLogin(context, userOrGroup))
                    {
                        RoleDefinitionBindingCollection roleDefinitions = role.RoleDefinitionBindings;
                        context.Load(roleDefinitions);
                        context.ExecuteQuery();

                        foreach (var roleDefinition in roleDefinitions)
                        {
                            DataRow dataRow = results.NewRow();
                            dataRow[Constants.SOProperties.Permission] = roleDefinition.Name;
                            results.Rows.Add(dataRow);
                        }
                    }
                }
            }
        }
        #endregion

        #endregion

        #region FolderPermissions

        #region BreakFolderInheritanceByName
        private void AddBreakFolderInheritanceByNameMethod(ServiceObject so)
        {
            Method mBreakFolderInheritanceByName = Helper.CreateMethod(Constants.Methods.BreakFolderInheritanceByName, "Break folder inheritance in list/library by name",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mBreakFolderInheritanceByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName())
                {
                    mBreakFolderInheritanceByName.InputProperties.Add(prop);
                    mBreakFolderInheritanceByName.Validation.RequiredProperties.Add(prop);
                }
            }
            so.Methods.Add(mBreakFolderInheritanceByName);
        }

        private void ExecuteBreakFolderInheritanceByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, d => d.RootFolder);
                context.ExecuteQuery();

                Folder currentFolder = null;
                try
                {
                    currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName.Trim('/')));
                    context.Load(currentFolder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                }

                currentFolder.ListItemAllFields.BreakRoleInheritance(false, true);
                context.ExecuteQuery();
            }
        }
        #endregion

        #region ResetFolderInheritanceByName
        private void AddResetFolderInheritanceByNameMethod(ServiceObject so)
        {
            Method mResetFolderInheritanceByName = Helper.CreateMethod(Constants.Methods.ResetFolderInheritanceByName, "Reset folder inheritance in list/library by name",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mResetFolderInheritanceByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName())
                {
                    mResetFolderInheritanceByName.InputProperties.Add(prop);
                    mResetFolderInheritanceByName.Validation.RequiredProperties.Add(prop);
                }
            }
            so.Methods.Add(mResetFolderInheritanceByName);
        }

        private void ExecuteResetFolderInheritanceByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, d => d.RootFolder);
                context.ExecuteQuery();

                Folder currentFolder = null;
                try
                {
                    currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName.Trim('/')));
                    context.Load(currentFolder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                }

                currentFolder.ListItemAllFields.ResetRoleInheritance();
                context.ExecuteQuery();
            }
        }
        #endregion

        #region AddFolderPermissionByName
        private void AddAddFolderPermissionByNameMethod(ServiceObject so)
        {
            Method mAddFolderPermissionByName = Helper.CreateMethod(Constants.Methods.AddFolderPermissionByName, "Add folder permissions in list/library by name",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mAddFolderPermissionByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName() || prop.IsPermissionColumn())
                {
                    mAddFolderPermissionByName.InputProperties.Add(prop);
                    mAddFolderPermissionByName.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsUserLogins() || prop.IsGroupLogins())
                {
                    mAddFolderPermissionByName.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mAddFolderPermissionByName);
        }

        private void ExecuteAddFolderPermissionByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName= base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string userLogins = base.GetStringProperty(Constants.SOProperties.UserLogins);
            string groupLogins = base.GetStringProperty(Constants.SOProperties.GroupLogins);
            string permission = base.GetStringProperty(Constants.SOProperties.Permission, true);

            if (string.IsNullOrEmpty(userLogins) && string.IsNullOrEmpty(groupLogins))
            {
                throw new ApplicationException(Constants.ErrorMessages.UserGroupLoginsAreEmptyException);
            }

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, d => d.RootFolder);
                context.ExecuteQuery();

                Folder currentFolder = null;
                try
                {
                    currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName));
                    context.Load(currentFolder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                }

                var role = new RoleDefinitionBindingCollection(context);
                role.Add(context.Web.RoleDefinitions.GetByName(permission));

                if (!string.IsNullOrEmpty(userLogins))
                {
                    foreach (User user in GetUsersByLoginString(context, userLogins))
                    {
                        currentFolder.ListItemAllFields.RoleAssignments.Add(user, role);
                    }
                }

                if (!string.IsNullOrEmpty(groupLogins))
                {
                    foreach (Group group in GetGroupsByLoginString(context, groupLogins))
                    {
                        currentFolder.ListItemAllFields.RoleAssignments.Add(group, role);
                    }
                }
                currentFolder.Update();
                context.ExecuteQuery();
            }
        }
        #endregion

        #region RemoveFolderPermissionByName
        private void AddRemoveFolderPermissionByNameMethod(ServiceObject so)
        {
            Method mRemoveFolderPermissionByName = Helper.CreateMethod(Constants.Methods.RemoveFolderPermissionByName, "Remove folder permissions in list/library by name",
                MethodType.Execute);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mRemoveFolderPermissionByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName())
                {
                    mRemoveFolderPermissionByName.InputProperties.Add(prop);
                    mRemoveFolderPermissionByName.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsUserLogins() || prop.IsGroupLogins())
                {
                    mRemoveFolderPermissionByName.InputProperties.Add(prop);
                }
            }
            so.Methods.Add(mRemoveFolderPermissionByName);
        }

        private void ExecuteRemoveFolderPermissionByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName= base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string userLogins = base.GetStringProperty(Constants.SOProperties.UserLogins);
            string groupLogins = base.GetStringProperty(Constants.SOProperties.GroupLogins);

            if (string.IsNullOrEmpty(userLogins) && string.IsNullOrEmpty(groupLogins))
            {
                throw new ApplicationException(Constants.ErrorMessages.UserGroupLoginsAreEmptyException);
            }

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list, d => d.RootFolder);
                context.ExecuteQuery();

                Folder currentFolder = null;
                try
                {
                    currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName));
                    context.Load(currentFolder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                }

                RoleAssignmentCollection roles = currentFolder.ListItemAllFields.RoleAssignments;
                context.Load(roles, role => role.Include(roleassigned => roleassigned.Member.LoginName, roleassigned => roleassigned.Member));
                context.ExecuteQuery();

                if (!string.IsNullOrEmpty(userLogins))
                {
                    foreach (User user in GetUsersByLoginString(context, userLogins))
                    {
                        context.Load(user, u => u.LoginName);
                        context.ExecuteQuery();

                        foreach (RoleAssignment role in roles)
                        {
                            if (role.Member.LoginName == user.LoginName)
                            {
                                role.DeleteObject();
                                break;
                            }
                        }
                    }
                }

                if (!string.IsNullOrEmpty(groupLogins))
                {
                    foreach (Group group in GetGroupsByLoginString(context, groupLogins))
                    {
                        context.Load(group, g => g.LoginName);
                        context.ExecuteQuery();

                        foreach (RoleAssignment role in roles)
                        {
                            if (role.Member.LoginName == group.LoginName)
                            {
                                role.DeleteObject();
                                break;
                            }
                        }
                    }
                }

                context.ExecuteQuery();
            }
        }
        #endregion

        #region GetFolderPermissionByName
        private void AddGetFolderPermissionByIdMethod(ServiceObject so)
        {
            Method mGetFolderPermissionByName = Helper.CreateMethod(Constants.Methods.GetFolderPermissionByName, "Get folder permissions in list/library by name",
                MethodType.List);

            if (base.IsDynamicSiteURL)
            {
                Helper.AddSiteURLParameter(mGetFolderPermissionByName);
            }

            foreach (Property prop in so.Properties)
            {

                if (prop.IsFolderName() || prop.IsUserOrGroup())
                {
                    mGetFolderPermissionByName.InputProperties.Add(prop);
                    mGetFolderPermissionByName.Validation.RequiredProperties.Add(prop);
                }
                if (prop.IsPermissionColumn())
                {
                    mGetFolderPermissionByName.ReturnProperties.Add(prop);
                }
            }
            so.Methods.Add(mGetFolderPermissionByName);
        }

        private void ExecuteGetFolderPermissionByName()
        {
            ServiceObject serviceObject = ServiceBroker.Service.ServiceObjects[0];
            serviceObject.Properties.InitResultTable();
            DataTable results = base.ServiceBroker.ServicePackage.ResultTable;

            string listTitle = serviceObject.GetListTitle();
            string siteURL = GetSiteURL();
            string folderName = base.GetStringProperty(Constants.SOProperties.FolderName, true);
            string userOrGroup = base.GetStringProperty(Constants.SOProperties.UserOrGroup, true);
            string permissions = string.Empty;

            using (ClientContext context = InitializeContext(siteURL))
            {
                Web spWeb = context.Web;

                List list = spWeb.Lists.GetByTitle(listTitle);
                context.Load(list);
                context.Load(list, l => l.RootFolder, l => l.DefaultDisplayFormUrl);
                context.ExecuteQuery();

                Folder currentFolder = null;

                try
                {
                    currentFolder = spWeb.GetFolderByServerRelativeUrl(string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderName));
                    context.Load(currentFolder);
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format(Constants.ErrorMessages.FolderWasNotFound, folderName), ex);
                }

                RoleAssignmentCollection roles = currentFolder.ListItemAllFields.RoleAssignments;
                context.Load(roles, role => role.Include(roleassigned => roleassigned.Member.LoginName, roleassigned => roleassigned.Member));
                context.ExecuteQuery();

                foreach (RoleAssignment role in roles)
                {
                    
                    if (role.Member.LoginName == GetUserOrGroupLogin(context, userOrGroup))
                    {
                        RoleDefinitionBindingCollection roleDefinitions = role.RoleDefinitionBindings;
                        context.Load(roleDefinitions);
                        context.ExecuteQuery();

                        foreach (var roleDefinition in roleDefinitions)
                        {
                            DataRow dataRow = results.NewRow();
                            dataRow[Constants.SOProperties.Permission] = roleDefinition.Name;
                            results.Rows.Add(dataRow);
                        }
                    }
                }

                
            }
        }
        #endregion

        #endregion

        private List<User> GetUsersByLoginString(ClientContext context, string userLogins)
        {
            List<User> users = new List<User>();

            foreach (string userLogin in userLogins.Split(';'))
            {
                users.Add(context.Web.EnsureUser(userLogin.ToLower()));
            }

            return users;
        }

        private List<Group> GetGroupsByLoginString(ClientContext context, string groupLogins)
        {
            List<Group> groups = new List<Group>();

            foreach (string userLogin in groupLogins.Split(';'))
            {
                groups.Add(context.Web.SiteGroups.GetByName(userLogin));
            }

            return groups;
        }

        private string GetUserOrGroupLogin(ClientContext context, string userOrGroup)
        {
            Group group = null;
            User user = null;
            try
            {
                group = context.Web.SiteGroups.GetByName(userOrGroup);
                context.Load(group);
                context.Load(group, g => g.LoginName);
                context.ExecuteQuery();
                return group.LoginName;
            }
            catch
            {
                try
                {
                    user = context.Web.SiteUsers.GetByLoginName("i:0#.w|" + userOrGroup);
                    context.Load(user);
                    context.Load(user, u => u.LoginName);
                    context.ExecuteQuery();
                    return user.LoginName;
                }
                catch(Exception iex)
                {
                    throw new ApplicationException(Constants.ErrorMessages.UserGroupLoginAreIncorrectException, iex);
                }
            }
        }
    }
}
