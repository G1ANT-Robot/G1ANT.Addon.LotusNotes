using Domino;
using G1ANT.Addon.LotusNotes.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class DatabaseModel
    {
        private readonly LotusNotesWrapper wrapper;
        private readonly NotesDatabase database;

        public bool IsOpen => database.IsOpen;


        public DatabaseModel(LotusNotesWrapper wrapper, NotesDatabase database)
        {
            this.wrapper = wrapper ?? throw new ArgumentNullException(nameof(wrapper));
            this.database = database ?? throw new ArgumentNullException(nameof(database));
        }


        private ViewModel CreateModelOrNull(NotesView view) => view != null ? new ViewModel(wrapper, view) : null;


        public DocumentModel GetDocumentByUrl(string url) => new DocumentModel(wrapper, database.GetDocumentByURL(url));
        public DocumentModel GetDocumentByUnid(string unid) => new DocumentModel(wrapper, database.GetDocumentByUNID(unid));
        public DocumentModel GetDocumentById(string noteId) => new DocumentModel(wrapper, database.GetDocumentByID(noteId));

        public NotesDocument CreateDocument() => database.CreateDocument();

        public IEnumerable<DocumentModel> GetDocumentsFromFolder(string folderName)
        {
            var view = database.GetView(folderName) ?? throw new ArgumentException($"Folder {folderName} not found", nameof(folderName));
            return new ViewModel(wrapper, view);
        }


        public void Open() => database.Open();


        public ViewModel GetView(string viewName)
        {
            return CreateModelOrNull(database.GetView(viewName))
                ?? throw new ApplicationException($"View (folder) {viewName} not found");
        }

        /// <summary>
        /// Returns names of views marked as folders (IsFolder is set)
        /// </summary>
        /// <returns></returns>
        public IReadOnlyCollection<string> GetFolderNames()
        {
            var cats = database.Categories;

            var folders = ((IEnumerable)database.Views)
                .OfType<NotesView>()
                .Where(v => v.IsFolder)
                .Select(v => v.Name)
                .ToList();
            return folders;
        }
    }
}
