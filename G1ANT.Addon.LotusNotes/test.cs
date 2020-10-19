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

            var view = wrapper.GetView("($Inbox)");
            var inboxDocuments = view.GetDocuments();

            //var inboxDocuments = wrapper.GetDocumentsFromFolder("($Inbox)");

            //wrapper.ConvertEmailToRichText = false;
            var email = inboxDocuments.First();

            var subject = email.Subject;
            var from = email.From;
            var to = email.To;

            var items = email.Items.Value.ToList();
            var attachments = email.GetAttachments();

            email = wrapper.GetDocumentByUnid(email.Unid);

            wrapper.SendEmail(new string[] { "lukasz.fronczyk@g1ant.com" }, "zażółć gęślą jaźń", "<body>body<p>zażółć gęślą jaźń</p><a href='http://g1ant.com'>link test</a></body>");


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
