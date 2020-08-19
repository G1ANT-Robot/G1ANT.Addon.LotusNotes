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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Services
{
    public class DocumentModel
    {
        private readonly NotesDocument document;

        public dynamic Authors { get; }
        //public dynamic ColumnValues { get; }
        public dynamic Created { get; }
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
        public IReadOnlyCollection<DocumentItemModel> Items { get; }
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
        public string UniversalID { get; set; }
        //public string Verifier { get; }
        //public dynamic LockHolders { get; }
        public bool IsEncrypted { get; }

        public DocumentModel(NotesDocument document)
        {
            this.document = document;

            Authors = document.Authors;
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
            Items = ((object[])document.Items).Cast<NotesItem>().Select(ni => new DocumentItemModel(ni)).ToList();
            //Items = document.Items;
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
            UniversalID = document.UniversalID;
            IsEncrypted = document.IsEncrypted;
        }


        public IReadOnlyCollection<AttachmentModel> GetAttachments()
        {
            if (document.HasEmbedded && document.HasItem("$File"))
            {
                return ((IEnumerable)document.Items)
                    .OfType<NotesItem>()
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

    }




    public class AttachmentModel
    {
        private readonly NotesEmbeddedObject embeddedObject;

        public string FileName { get; }
        public string Class { get; }
        public int FileSize { get; }
        //bool FitBelowFields { get; set; }
        //bool FitToWindow { get; set; }
        public string Name { get; }
        public dynamic Object { get; }
        //NotesRichTextItem Parent { get; }
        //bool RunReadOnly { get; set; }
        public string Source { get; }
        public EMBED_TYPE EmbedType { get; }
        public object[] Verbs { get; }


        public static bool IsAttachmentItem(NotesItem item) => item.type == IT_TYPE.ATTACHMENT;

        public AttachmentModel(NotesDocument document, NotesItem item)
        {
            if (!IsAttachmentItem(item))
                throw new ArgumentException("Item is not an attachment", nameof(item));
            
            FileName = ((IEnumerable)item.Values).Cast<string>().First();
            embeddedObject = document.GetAttachment(FileName);

            Class = embeddedObject.Class;
            FileSize = embeddedObject.FileSize;
            Name = embeddedObject.Name;
            Object = embeddedObject.Object;
            Source = embeddedObject.Source;
            EmbedType = embeddedObject.type;
            Verbs = embeddedObject.Verbs;
        }

        public object Activate(bool show) => embeddedObject.Activate(show);
        public void DoVerb(string verb) => embeddedObject.DoVerb(verb);
        public void ExtractFile(string path) => embeddedObject.ExtractFile(path);
        public void Remove() => embeddedObject.Remove();
    }
}

