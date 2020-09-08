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
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Services
{
    public class LotusNotesWrapper : IDisposable
    {
        private NotesSession session;
        private DatabaseModel database;

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
            database = null;
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

            database = new DatabaseModel(this, session.GetDatabase(server, databaseFile, false));
            //// note to myself: that won't work: serverDatabase = session.CurrentDatabase; - unimplemented exception

            if (!database.IsOpen)
                database.Open();
        }



        public void SendEmail(string[] to, string subject, string message, string[] cc = null, string[] bcc = null, string[] attachmentPaths = null, bool saveMessageOnSend = true)
        {
            ValidateRecipients(to);

            session.ConvertMime = false;

            var notesDocument = database.CreateDocument();
            notesDocument.SaveMessageOnSend = saveMessageOnSend;

            //notesDocument.ReplaceItemValue(ItemFieldNames.Form, "Memo"); // "Form" seems to be a name of sender, but is not visible anywhere

            notesDocument.ReplaceItemValue(ItemFieldNames.SendTo, to);
            if (cc?.Any() == true)
                notesDocument.ReplaceItemValue(ItemFieldNames.CopyTo, cc);
            if (bcc?.Any() == true)
                notesDocument.ReplaceItemValue(ItemFieldNames.BlindCopyTo, bcc);

            notesDocument.ReplaceItemValue(ItemFieldNames.Subject, subject);

            SetHtmlContent(notesDocument, message);
            //SetTextContent(notesDocument);

            if (attachmentPaths?.Any() == true)
            {
                int count = 0;
                foreach (var attachmentPath in attachmentPaths)
                {
                    var attachmentName = $"Attachment{count}";
                    var notesRtf = notesDocument.CreateRichTextItem(attachmentName);
                    var embedObj = notesRtf.EmbedObject(EMBED_TYPE.EMBED_ATTACHMENT, "", attachmentPath, attachmentName);
                    count++;
                }
            }

            var oItemValue = notesDocument.GetItemValue(ItemFieldNames.SendTo);

            notesDocument.Send(false, oItemValue);
            session.ConvertMime = true;
        }

        private void SetTextContent(NotesDocument notesDocument, string richTextMessage)
        {
            var richTextItem = notesDocument.CreateRichTextItem(ItemFieldNames.Body);
            richTextItem.AppendText(richTextMessage);
        }

        private void SetHtmlContent(NotesDocument notesDocument, string message)
        {
            var body = notesDocument.CreateMIMEEntity();
            var stream = session.CreateStream();
            stream.WriteText(message);
            body.SetContentFromText(stream, "text/html;charset=utf-8", MIME_ENCODING.ENC_IDENTITY_8BIT);
            stream.Close(); // ?
        }

        private static void ValidateRecipients(string[] to)
        {
            if (!to.Any())
                throw new ArgumentException("`To` argument is empty");
            if (to.Any(v => string.IsNullOrEmpty(v)))
                throw new ArgumentException("Some elements of argument `to` are empty");
        }


        /// <summary>
        /// Returns names of views marked as folders (IsFolder is set)
        /// </summary>
        /// <returns></returns>
        public IReadOnlyCollection<string> GetFolderNames() => database.GetFolderNames();


        public DocumentModel GetDocumentByIndex(ViewModel view, int index) => view.GetDocumentByIndex(index);
        public DocumentModel GetDocumentByUrl(string url) => database.GetDocumentByUrl(url);
        public DocumentModel GetDocumentByUnid(string unid) => database.GetDocumentByUnid(unid);
        public DocumentModel GetDocumentById(string noteId) => database.GetDocumentById(noteId);


        public IReadOnlyCollection<AttachmentModel> GetAttachments(DocumentModel document) => document.GetAttachments();

        public void Save(DocumentModel document, bool force = true, bool makeResponse = false, bool markRead = false) => document.Save(force, makeResponse, markRead);

        public void Remove(DocumentModel document, bool force = true) => document.Remove(force);
        public void RemovePermanently(DocumentModel document, bool force = true) => document.RemovePermanently(force);

        public void RemoveFromFolder(DocumentModel document, string folder) => document.RemoveFromFolder(folder);
        public void PutInFolder(DocumentModel document, string folder) => document.PutInFolder(folder);

        public void MakeResponse(DocumentModel response, DocumentModel respondTo) => response.MakeResponse(respondTo);

        public ViewModel GetView(string viewName) => database.GetView(viewName);
    }
}
