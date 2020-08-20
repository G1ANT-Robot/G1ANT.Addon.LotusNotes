using G1ANT.Addon.LotusNotes.Services;
using System.Linq;

namespace G1ANT.Addon.LotusNotes
{
    class test
    {
        public static void Main()
        {
            var wrapper = LotusNotesManager.CreateWrapper();


            wrapper.Connect("Rhenai-Robot", "DE9899SL8/RETHMANN", @"mail_de\zzm-de0849-robot-test-40453.nsf");
            //wrapper.ConvertEmailToRichText = false;

            var folders = wrapper.GetFolderNames();

            var inboxDocuments = wrapper.GetDocumentsFromFolder("($Inbox)");

            var email = inboxDocuments.First();
            
            var subject = email.Subject;
            var from = email.From;
            var to = email.To;

            var items = email.Items.ToList();
            var attachments = email.GetAttachments();

            email = wrapper.GetDocumentByUNID(email.UniversalID);


            //var doc = email.document;
            //var entity = doc.GetMIMEEntity();

            //var ce = entity.GetFirstChildEntity();
            //while (ce != null)
            //{
            //    var s = doc.GetItemValue("Subject");
            //    var t = ce.ContentAsText;
            //    ce = entity.GetNextSibling();
            //}




            //doc.CloseMIMEEntities();
        }
    }
}
