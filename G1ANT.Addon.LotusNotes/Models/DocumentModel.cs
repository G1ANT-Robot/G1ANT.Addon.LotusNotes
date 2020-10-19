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
using G1ANT.Addon.LotusNotes.Services;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class DocumentModel : PropertiesContainer
    {
        private readonly NotesDocument document;

        public bool HasMime { get; private set; }
        public string[] Authors { get; }
        //public dynamic ColumnValues { get; }
        public DateTime Created { get; }
        public NotesEmbeddedObject[] EmbeddedObjects { get; }
        //public dynamic EncryptionKeys { get; set; }
        //public bool EncryptOnSend { get; set; }
        //public dynamic FolderReferences { get; }
        //public int FTSearchScore { get; }
        public bool HasEmbedded { get; }
        public string HttpURL { get; }
        public bool IsDeleted { get; }
        public bool IsNewNote { get; }
        public bool IsProfile { get; }
        public bool IsResponse { get; }
        public bool IsSigned { get; }
        public bool IsUIDocOpen { get; }
        public bool IsValid { get; }
        public Lazy<DocumentItemModel[]> Items { get; }
        public string Key { get; }
        public dynamic LastAccessed { get; }
        public dynamic LastModified { get; }
        public string NameOfProfile { get; }
        public string NoteID { get; }
        public string NotesURL { get; }
        //public NotesDatabase ParentDatabase { get; }
        //public string ParentDocumentUNID { get; }
        public NotesView ParentView { get; }
        public NotesDocumentCollection Responses { get; }
        public bool SaveMessageOnSend { get; set; }
        public bool SentByAgent { get; }
        public string Signer { get; }
        public bool SignOnSend { get; set; }
        public int Size { get; }
        public string Unid { get; set; }
        //public string Verifier { get; }
        //public dynamic LockHolders { get; }
        public bool IsEncrypted { get; }
        public string Subject { get; }
        public string From { get; }
        public string[] To { get; }
        public string[] Cc { get; }
        public string ContentText { get; private set; }
        public string ContentHtml { get; private set; }


        public DocumentModel(LotusNotesWrapper wrapper, NotesDocument document)
        {
            this.document = document;

            LoadEmailContent(wrapper);

            Authors = ((IEnumerable)document.Authors).Cast<string>().ToArray();
            Created = document.Created;
            EmbeddedObjects = document.EmbeddedObjects;
            HasEmbedded = document.HasEmbedded;
            HttpURL = document.HttpURL;
            IsDeleted = document.IsDeleted;
            IsNewNote = document.IsNewNote;
            IsProfile = document.IsProfile;
            IsResponse = document.IsResponse;
            IsSigned = document.IsSigned;
            IsUIDocOpen = document.IsUIDocOpen;
            IsValid = document.IsValid;
            Key = document.Key;
            LastAccessed = document.LastAccessed;
            LastModified = document.LastModified;
            NameOfProfile = document.NameOfProfile;
            NoteID = document.NoteID;
            NotesURL = document.NotesURL;
            ParentView = document.ParentView;
            Responses = document.Responses;
            SentByAgent = document.SentByAgent;
            Signer = document.Signer;
            SignOnSend = document.SignOnSend;
            Size = document.Size;
            Unid = document.UniversalID;
            IsEncrypted = document.IsEncrypted;

            Items = new Lazy<DocumentItemModel[]>(() => GetItems(document));

            Subject = document.GetItemValue(ItemFieldNames.Subject)[0];
            From = document.GetItemValue(ItemFieldNames.From).ToString();
            To = ((IEnumerable)document.GetItemValue(ItemFieldNames.SendTo)).Cast<string>().ToArray();
            Cc = ((IEnumerable)document.GetItemValue(ItemFieldNames.CopyTo)).Cast<string>().ToArray();
        }


        private DocumentItemModel[] GetItems(NotesDocument document)
        {
            return GetNoteItems().Select(ni => new DocumentItemModel(ni)).ToArray();
        }


        private void LoadEmailContent(LotusNotesWrapper wrapper)
        {
            var convertEmailToRichText = wrapper.ConvertEmailToRichText;
            try
            {
                wrapper.ConvertEmailToRichText = false;
                HasMime = document.HasItem("$NoteHasNativeMIME");
                if (HasMime)
                {
                    ContentHtml = GetMimeContent("text/html");
                    ContentText = GetMimeContent("text/plain");
                }
            }
            finally
            {
                wrapper.ConvertEmailToRichText = convertEmailToRichText;
            }

            if (string.IsNullOrEmpty(ContentText))
            {
                var noteItem = GetNoteItems().FirstOrDefault(i => i.Name == ItemFieldNames.Body);
                ContentText = noteItem == null ? null : new DocumentItemModel(noteItem).Value;
            }
        }

        private string GetMimeContent(string contentType)
        {
            var contentTypeParts = contentType.Split('/');
            var mime = document.GetMIMEEntity("Body");
            if (mime != null)
            {
                if (mime.ContentType == "multipart")
                {
                    var child = mime.GetFirstChildEntity();
                    while (child != null)
                    {
                        if (child.ContentType == contentTypeParts[0] && child.ContentSubType == contentTypeParts[1])
                        {
                            return child.ContentAsText;
                        }
                        var innerChildOrSibling = child.GetFirstChildEntity();
                        if (innerChildOrSibling == null)
                        {
                            innerChildOrSibling = child.GetNextSibling();
                            if (innerChildOrSibling == null)
                            {
                                innerChildOrSibling = child.GetParentEntity();
                                if (innerChildOrSibling != null)
                                {
                                    innerChildOrSibling = innerChildOrSibling.GetNextSibling();
                                }
                            }
                        }
                        child = innerChildOrSibling;
                    }
                }
                else
                {
                    return mime.ContentAsText;
                }
            }

            return "";
        }

        internal NotesDocument GetNotesDocument() => document;

        private IEnumerable<NotesItem> GetNoteItems() => ((IEnumerable)document.Items).OfType<NotesItem>();

        public IReadOnlyCollection<AttachmentModel> GetAttachments()
        {
            if (document.HasEmbedded && document.HasItem("$File"))
            {
                return GetNoteItems()
                    .Where(AttachmentModel.IsAttachmentItem)
                    .Select(item => new AttachmentModel(document, item))
                    .ToList();
            }

            return new Collection<AttachmentModel>();
        }



        public void Save(bool force = true, bool makeResponse = false, bool markRead = false) => document.Save(force, makeResponse, markRead);

        public void Remove(bool force = true) => document.Remove(force);
        public void RemovePermanently(bool force = true) => document.RemovePermanently(force);

        public void RemoveFromFolder(string folder) => document.RemoveFromFolder(folder);
        public void PutInFolder(string folder) => document.PutInFolder(folder, false);

        public void MakeResponse(DocumentModel model) => document.MakeResponse(model.document);

        public override string ToString() => JObject.FromObject(this).ToString();
    }
}

