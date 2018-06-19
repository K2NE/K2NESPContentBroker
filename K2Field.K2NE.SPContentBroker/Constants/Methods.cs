using System;

namespace K2Field.K2NE.SPContentBroker.Constants
{

    public static class Methods
    {
        public const string GetItembyId = "Get Item by Id";
        public const string CreatelistItem = "Create list Item";
        public const string UpdatelistItem = "Update list Item";
        public const string DeleteItem = "Delete item";
        public const string ListAllItems = "List all items";
        public const string MovelistItem = "Move list Item";
        public const string CopylistItem = "Copy list Item";
        //public const string CreateFolder = "Create Folder";
        //public const string DeleteFolder = "Delete Folder";
        public const string GetItemsByFolder = "Get Items by folder";

        #region List/Library methods
        public const string GetItemById = "Get Item By Id";
        public const string CreateItem = "Create Item";
        public const string UpdateItemById = "Update Item By Id";
        public const string DeleteItemById = "Delete Item By Id";
        public const string GetItemByTitle = "Get Item By Title";
        public const string GetItemByName = "Get Item By Name";
        public const string GetItems = "Get Items";
        #endregion

        #region Document library methods
        public const string CreateDocument = "Create Document";
        public const string DeleteDocumentById = "Delete Document By Id";
        public const string GetDocumentById = "Get Document By Id";
        public const string GetDocuments = "Get Documents";
        public const string CopyDocumentByName = "Copy Document By Name";
        public const string MoveDocumentByName = "Move Document By Name";
        public const string RenameDocumentById = "Rename Document By Id";
        public const string CheckInDocumentByName = "Check In Document By Name";
        public const string CheckInDocumentById = "Check In Document By Id";
        public const string CheckOutDocumentByName = "Check Out Document By Name";
        public const string CheckOutDocumentById = "Check Out Document By Id";
        public const string GetDocumentItems = "Get Document Items";
        #endregion

       
        #region Folder methods
        public const string CreateFolder = "Create Folder";
        public const string DeleteFolder = "Delete Folder";
        public const string RenameFolder = "Rename Folder";
        public const string MoveFolder = "Move Folder";
        #endregion

        #region Document set methods
        public const string CreateDocumentSet = "Create Document Set";
        public const string DeleteDocumentSetByName = "Delete Document Set By Name";
        public const string UpdateDocumentSetByName = "Update Document Set By Name";
        public const string GetDocumentSetByName = "Get Document Set By Name";
        public const string GetDocumentSets = "Get Document Sets";
        public const string RenameDocumentSetByName = "Rename Document Set By Name";
        #endregion

        #region Permission methods
        public const string BreakItemInheritanceById = "Break Item Inheritance By Id";
        public const string BreakFolderInheritanceByName = "Break Folder Inheritance By Name";
        public const string ResetFolderInheritanceByName = "Reset Folder Inheritance By Name";
        public const string ResetItemInheritanceById = "Reset Item Inheritance By Id";
        public const string AddItemPermissionById = "Add Item Permission By Id";
        public const string AddFolderPermissionByName = "Add Folder Permission By Name";
        public const string RemoveFolderPermissionByName = "Remove Folder Permission By Name";
        public const string RemoveItemPermissionById = "Remove Item Permission By Id";
        public const string GetFolderPermissionByName = "Get Folder Permission Name";
        public const string GetItemPermissionById = "Get Item Permission By Id";
        #endregion

        #region Content Type Gallery methods
        public const string GetContentTypeByName = "Get Content Type By Name";
        public const string GetContentTypeById = "Get Content Type By Id";
        public const string GetContentTypes = "Get Content Types";
        #endregion
    }
}
