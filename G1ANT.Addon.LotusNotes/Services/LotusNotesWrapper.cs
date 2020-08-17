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
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Services
{
    internal static class ItemFieldNames
    {
        public const string Form = "Form";
        public const string SendTo = "SendTo";
        public const string CopyTo = "CopyTo";
        public const string Subject = "Subject";
        public const string Body = "Body";
    }

    public class LotusNotesWrapper
    {
        private NotesSession lotusNotesServerSession;
        private NotesDatabase serverDatabase;

        public int Id { get; set; }

        public LotusNotesWrapper(int id)
        {
            Id = id;
        }

        /// <summary>
        /// Connects to server using the last login used by the IBM Notes app. If you want to change
        /// the login you must manually sign-in using different login with IBM Notes app.
        /// </summary>
        /// <param name="password"></param>
        /// <param name="server"></param>
        /// <param name="databaseFile"></param>
        public void Connect(string password = "Rhenus.123", string server = "DE9899SL8/RETHMANN", string databaseFile = @"mail_de\zzm-de0849-robot-test-40453.nsf")
        {
            lotusNotesServerSession = new NotesSession();
            lotusNotesServerSession.Initialize(password);
            serverDatabase = lotusNotesServerSession.GetDatabase(server, databaseFile, false);
            if (!serverDatabase.IsOpen)
            {
                serverDatabase.Open();
            }
        }


        public void SendEmail(string to, string subject, string richTextMessage, string cc = "")
        {
            SendEmail(new[] { to }, subject, richTextMessage, string.IsNullOrEmpty(cc) ? null : new[] { cc });
        }

        public void SendEmail(string[] to, string subject, string richTextMessage, string[] cc = null)
        {
            var notesDocument = serverDatabase.CreateDocument();

            notesDocument.ReplaceItemValue(ItemFieldNames.Form, "Memo");

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


        public class DocumentItemModel
        {
            public NotesDateTime DateTimeValue { get; set; }
            public bool IsAuthors { get; set; }
            public bool IsEncrypted { get; set; }
            public bool IsNames { get; set; }
            public bool IsProtected { get; set; }
            public bool IsReaders { get; set; }
            public bool IsSigned { get; set; }
            public bool IsSummary { get; set; }
            public dynamic LastModified { get; }
            public string Name { get; }
            public DocumentModel Parent { get; }
            public bool SaveToDisk { get; set; }
            public string Text { get; }
            public IT_TYPE Type { get; }
            public int ValueLength { get; }
            public object[] Values { get; }
            public string Value { get; }

            public DocumentItemModel(NotesItem item)
            {
                DateTimeValue = item.DateTimeValue;
                IsAuthors = item.IsAuthors;
                IsEncrypted = item.IsEncrypted;
                IsNames = item.IsNames;
                IsProtected = item.IsProtected;
                IsReaders = item.IsReaders;
                IsSigned = item.IsSigned;
                IsSummary = item.IsSummary;
                LastModified = item.LastModified;
                Name = item.Name;
                //Parent = new DocumentModel(item.Parent);
                SaveToDisk = item.SaveToDisk;
                Text = item.Text;
                Type = item.type;
                ValueLength = item.ValueLength;
                if (Values is IEnumerable)
                    Values = item.Values;
                else// if (Values != null)
                    Value = Values?.ToString();

                //            Values = ((IEnumerable)(item.Values)).Cast<string>().ToList();

                //var items = ((object[])document.Items).Cast<NotesItem>();

                //if (document.HasEmbedded && document.HasItem("$File"))
                //{
                //    foreach (var item in items)
                //    {
                //        if (item.type == IT_TYPE.ATTACHMENT)
                //        {
                //            var v = ((IEnumerable)(item.Values)).Cast<string>().ToList();
                //            //item.Values
                //            //document.GetAttachment()
                //        }
                //    }
                //}            
            }
        }

        public class DocumentModel
        {
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
        }

        public IEnumerable<DocumentModel> GetDocumentsFromFolder(string folderName)
        {
            var view = serverDatabase.GetView(folderName);
            var document = view.GetFirstDocument();
            while (document != null)
            {
                var model = new DocumentModel(document);
                //Console.WriteLine($"\t\tDocument {document.Key}");
                //var items = ((object[])document.Items).Cast<NotesItem>();

                //if (document.HasEmbedded && document.HasItem("$File"))
                //{
                //    foreach (var item in items)
                //    {
                //        if (item.type == IT_TYPE.ATTACHMENT)
                //        {
                //            var v = ((IEnumerable)(item.Values)).Cast<string>().ToList();
                //            //item.Values
                //            //document.GetAttachment()
                //        }
                //    }
                //}


                //foreach (var item in items)
                //{
                //    if (item is NotesRichTextItem nrti)
                //    {
                //        var mimeEntity = nrti.GetMIMEEntity();
                //        Console.WriteLine($"###\r\n\t\t\t\t{nrti.Name}: {nrti.GetUnformattedText()}");
                //    }
                //    else
                //        Console.WriteLine($"***\r\n\t\t\t\t{item.Name}: {item.Text}");
                //}

                yield return model;
                document = view.GetNextDocument(document);
            }

        }

    }
}

