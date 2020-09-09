using Domino;
using G1ANT.Addon.LotusNotes.Services;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class ViewModel : IEnumerable<DocumentModel>
    {
        private readonly LotusNotesWrapper wrapper;
        private readonly NotesView view;

        //dynamic Aliases { get; set; }
        //NotesViewEntryCollection AllEntries { get; }
        //bool AutoUpdate { get; set; }
        //COLORS BackgroundColor { get; set; }
        //int ColumnCount { get; }
        //dynamic ColumnNames { get; }
        //dynamic Columns { get; }
        public DateTime Created => view.Created;
        //int HeaderLines { get; }
        //public string HttpURL { get; }
        public bool IsCalendar => view.IsCalendar;
        public bool IsCategorized => view.IsCategorized;
        public bool IsConflict => view.IsConflict;
        public bool IsDefaultView => view.IsDefaultView;
        public bool IsFolder => view.IsFolder;
        public bool IsHierarchical => view.IsHierarchical;
        public bool IsModified => view.IsModified;
        public bool IsPrivate => view.IsPrivate;
        public DateTime LastModified => view.LastModified;
        public string Name { get => view.Name; set => view.Name = value; }
        public string NotesURL => view.NotesURL;
        //NotesDatabase Parent { get; }
        //bool ProtectReaders { get; set; }
        //dynamic Readers { get; set; }
        //int RowLines { get; }
        //SPACING SPACING { get; set; }
        //int TopLevelEntryCount { get; }
        public string Unid => view.UniversalID;
        //bool IsProhibitDesignRefresh { get; set; }
        //string SelectionFormula { get; set; }
        public int EntryCount => view.EntryCount;
        //dynamic LockHolders { get; }
        //string ViewInheritedName { get; }



        public ViewModel(LotusNotesWrapper wrapper, NotesView view)
        {
            this.wrapper = wrapper;
            this.view = view;
        }


        public IEnumerator<DocumentModel> GetEnumerator() => GetDocuments().GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();


        public void Clear() => view.Clear();
        public void Refresh() => view.Refresh();
        public void Remove() => view.Remove();

        //NotesViewNavigator CreateViewNav(int lCacheSize = 0);
        //NotesViewNavigator CreateViewNavMaxLevel(int lLevel, int lCacheSize = 0);
        //NotesViewNavigator CreateViewNavFrom(object pIUnk, int lCacheSize = 0);
        //NotesViewNavigator CreateViewNavFromCategory(string pName, int lCacheSize = 0);
        //NotesViewNavigator CreateViewNavFromChildren(object pIUnk, int lCacheSize = 0);
        //NotesViewNavigator CreateViewNavFromDescendants(object pIUnk, int lCacheSize = 0);

        //int FTSearch(string pQuery, int lMaxDocs = 0);

        //NotesDocumentCollection GetAllDocumentsByKey(object Keys, bool bExact = false);
        //NotesViewEntryCollection GetAllEntriesByKey(object Keys, bool bExact = false);

        //NotesDocument GetChild(NotesDocument pICurrent);

        //NotesViewColumn GetColumn(int lColumnNumber);
        //NotesViewColumn CreateColumn(int pos = -1, string Name = "", string Formula = "");
        //NotesViewColumn CopyColumn(object nameIndexObj, int dst = -1);
        //void RemoveColumn(object nameIndex);

        //NotesDocument GetDocumentByKey(object Keys, bool bExact = false);
        //NotesViewEntry GetEntryByKey(object Keys, bool bExact = false);

        private DocumentModel CreateModelOrNull(NotesDocument notesDocument) => notesDocument != null ? new DocumentModel(wrapper, notesDocument) : null;

        public DocumentModel GetFirstDocument() => CreateModelOrNull(view.GetFirstDocument());
        public DocumentModel GetLastDocument() => CreateModelOrNull(view.GetLastDocument());
        public DocumentModel GetNextDocument(DocumentModel currentDocument) => CreateModelOrNull(view.GetNextDocument(currentDocument.GetNotesDocument()));
        public DocumentModel GetDocumentByIndex(int index) => CreateModelOrNull(view.GetNthDocument(index));


        public IEnumerable<DocumentModel> GetDocuments()
        {
            var document = GetFirstDocument();
            while (document != null)
            {
                yield return document;
                document = GetNextDocument(document);
            }
        }


        //NotesDocument GetNextSibling(NotesDocument pICurrent);
        //NotesDocument GetParentDocument(NotesDocument pICurrent);
        //NotesDocument GetPrevDocument(NotesDocument pICurrent);
        //NotesDocument GetPrevSibling(NotesDocument pICurrent);

        //void SetAliases(string Aliases = "");
        //bool Lock(ref object pName, bool bProvisionalOK = false);
        //bool LockProvisional(ref object pName);
        //void Unlock();

        public override string ToString() => JObject.FromObject(this).ToString();
    }
}
