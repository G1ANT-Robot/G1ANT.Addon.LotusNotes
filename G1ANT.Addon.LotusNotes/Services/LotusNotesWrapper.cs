/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using Domino;
using G1ANT.Addon.LotusNotes.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Services
{
    public class LotusNotesWrapper : IDisposable
    {
        private NotesSession session;
        private NotesDatabase serverDatabase;

        public int Id { get; set; }

        /// <summary>
        /// Set to false to avoid background conversion of email body to rich text and enable Mime Entity operations.
        /// Remember to set to true after finishing processing of the email.
        /// </summary>
        public bool ConvertEmailToRichText { get => session.ConvertMime; set => session.ConvertMime = value; }

        public LotusNotesWrapper(int id)
        {
            Id = id;
        }

        public void Close()
        {
            session = null;
            serverDatabase = null;
        }


        public void Dispose() => Close();

        /// <summary>
        /// Connects to server using the last login used by the IBM Notes app. If you want to change
        /// the login you must manually sign-in using different login with IBM Notes app.
        /// </summary>
        /// <param name="password"></param>
        /// <param name="server"></param>
        /// <param name="databaseFile"></param>
        public void Connect(string password, string server, string databaseFile)
        {
            session = new NotesSession();
            session.Initialize(password);
            serverDatabase = session.GetDatabase(server, databaseFile, false);
            // note to myself: that won't work: serverDatabase = session.CurrentDatabase; - unimplemented exception


            if (!serverDatabase.IsOpen)
            {
                serverDatabase.Open();
            }
        }


        public void SendEmail(string to, string subject, string richTextMessage, string cc = "")
        {
            SendEmail(new[] { to }, subject, richTextMessage, string.IsNullOrEmpty(cc) ? null : new[] { cc });
        }

        public void SendEmail(string[] to, string subject, string richTextMessage, string[] cc = null, bool saveMessageOnSend = true)
        {
            var notesDocument = serverDatabase.CreateDocument();
            notesDocument.SaveMessageOnSend = saveMessageOnSend;

            //notesDocument.ReplaceItemValue(ItemFieldNames.Form, "Memo"); // ???

            notesDocument.ReplaceItemValue(ItemFieldNames.SendTo, to);
            if (cc != null && cc.Any())
                notesDocument.ReplaceItemValue(ItemFieldNames.CopyTo, cc);
            notesDocument.ReplaceItemValue(ItemFieldNames.Subject, subject);

            var richTextItem = notesDocument.CreateRichTextItem(ItemFieldNames.Body);
            richTextItem.AppendText(richTextMessage);

            var oItemValue = notesDocument.GetItemValue(ItemFieldNames.SendTo);
            notesDocument.Send(false, ref oItemValue);
        }


        /// <summary>
        /// Returns names of views marked as folders (IsFolder is set)
        /// </summary>
        /// <returns></returns>
        public IReadOnlyCollection<string> GetFolderNames()
        {
            var folders = ((IEnumerable)serverDatabase.Views)
                .OfType<NotesView>()
                .Where(v => v.IsFolder)
                .Select(v => v.Name)
                .ToList();
            return folders;
        }

        public IEnumerable<DocumentModel> GetDocumentsFromFolder(string folderName)
        {
            var view = serverDatabase.GetView(folderName) ?? throw new ArgumentException($"Folder {folderName} not found", nameof(folderName));
            
            var document = view.GetFirstDocument();
            while (document != null)
            {
                var model = new DocumentModel(document);

                yield return model;
                document = view.GetNextDocument(document);
            }

        }


        public DocumentModel GetDocumentByURL(string url) => new DocumentModel(serverDatabase.GetDocumentByURL(url));
        public DocumentModel GetDocumentByUNID(string unid) => new DocumentModel(serverDatabase.GetDocumentByUNID(unid));
        public DocumentModel GetDocumentByID(string noteId) => new DocumentModel(serverDatabase.GetDocumentByID(noteId));



        public IReadOnlyCollection<AttachmentModel> GetAttachments(DocumentModel document) => document.GetAttachments();

        public void Save(DocumentModel document, bool force = true, bool makeResponse = false, bool markRead = false) => document.Save(force, makeResponse, markRead);

        public void Remove(DocumentModel document, bool force = true) => document.Remove(force);
        public void RemovePermanently(DocumentModel document, bool force = true) => document.RemovePermanently(force);

        public void RemoveFromFolder(DocumentModel document, string folder) => document.RemoveFromFolder(folder);
        public void PutInFolder(DocumentModel document, string folder) => document.PutInFolder(folder);

        public void MakeResponse(DocumentModel response, DocumentModel respondTo) => response.MakeResponse(respondTo);
    }
}

