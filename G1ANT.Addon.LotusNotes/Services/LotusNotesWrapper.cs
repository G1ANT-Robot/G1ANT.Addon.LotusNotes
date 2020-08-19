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

        public void SendEmail(string[] to, string subject, string richTextMessage, string[] cc = null, bool saveMessageOnSend = true)
        {
            var notesDocument = serverDatabase.CreateDocument();
            notesDocument.SaveMessageOnSend = saveMessageOnSend;

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

