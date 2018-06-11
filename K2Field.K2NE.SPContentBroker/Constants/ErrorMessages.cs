using System;

namespace K2Field.K2NE.SPContentBroker.Constants
{
    public static class ErrorMessages
    {
        public const string RequiredPropertyNotFound = "{0} is a required property, but does not exist.";
        public const string RequiredParameterNotFound = "{0} is a required parameter, but does not exist.";
        public const string PropertyNotFound = "The property with name '{0}', could not be found.";
        public const string ConfigOptionNotFound = "The Service Instance Configuration option '{0}' could not be found. Please specify it.";
        public const string FolderWasNotFound = "The folder \"{0}\" was not found in current list.";
        public const string FolderIsNotEmpty = "The folder \"{0}\" is not empty.";
        public const string FolderCreationIsNotSupported = "Folders are not supported for \"{0}\".";
        public const string UserGroupLoginsAreEmptyException = "You should set at least one user or group.";
        public const string UserGroupLoginAreIncorrectException = "Provided user's or group's login is incorrect.";
        public const string SAImpersonateWithOffice365NotSupported = "Impossible to use Impersonate/Service Account authentication mode for Office 365";
        public const string DocumentNotCheckedOut = "The document is not checked-out";
        public const string DocumentAlreadyCheckedOut = "Document is already checked out";
        public const string RequiredDocNotFound = "Document not found";
    }
}
