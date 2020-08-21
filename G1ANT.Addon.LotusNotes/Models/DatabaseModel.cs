using Domino;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class DatabaseModel
    {
        private readonly NotesDatabase database;

        public bool IsOpen => database.IsOpen;


        public DatabaseModel(NotesDatabase database)
        {
            this.database = database;
        }

        private ViewModel CreateModelOrNull(NotesView view) => view != null ? new ViewModel(view) : null;
        private DocumentModel CreateModelOrNull(NotesDocument document) => document != null ? new DocumentModel(document) : null;


        public DocumentModel GetDocumentByUrl(string url) => new DocumentModel(database.GetDocumentByURL(url));
        public DocumentModel GetDocumentByUnid(string unid) => new DocumentModel(database.GetDocumentByUNID(unid));
        public DocumentModel GetDocumentById(string noteId) => new DocumentModel(database.GetDocumentByID(noteId));

        public NotesDocument CreateDocument() => database.CreateDocument();

        public IEnumerable<DocumentModel> GetDocumentsFromFolder(string folderName)
        {
            var view = database.GetView(folderName) ?? throw new ArgumentException($"Folder {folderName} not found", nameof(folderName));
            return new ViewModel(view);
        }


        public void Open() => database.Open();



        public ViewModel GetView(string viewName) => CreateModelOrNull(database.GetView(viewName));

        /// <summary>
        /// Returns names of views marked as folders (IsFolder is set)
        /// </summary>
        /// <returns></returns>
        public IReadOnlyCollection<string> GetFolderNames()
        {
            var folders = ((IEnumerable)database.Views)
                .OfType<NotesView>()
                .Where(v => v.IsFolder)
                .Select(v => v.Name)
                .ToList();
            return folders;
        }
    }
}
