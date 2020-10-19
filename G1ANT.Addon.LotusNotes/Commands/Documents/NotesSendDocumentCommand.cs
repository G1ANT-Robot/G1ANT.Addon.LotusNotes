/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.LotusNotes.Services;
using G1ANT.Language;
using System;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.sendemail", Tooltip = "Send document (email)")]
    public class NotesSendDocumentCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "")]
            public Structure To { get; set; }

            [Argument(Tooltip = "")]
            public Structure Cc { get; set; }

            [Argument(Tooltip = "")]
            public Structure Bcc { get; set; }

            [Argument(Tooltip = "")]
            public Structure PathsToAttachments { get; set; }

            [Argument(Required = true, Tooltip = "")]
            public TextStructure Subject { get; set; }

            [Argument(Required = true, Tooltip = "")]
            public TextStructure Message { get; set; }

            [Argument(Tooltip = "")]
            public BooleanStructure SaveMessageOnSend { get; set; } = new BooleanStructure(true);
        }

        public NotesSendDocumentCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var to = GetStructureValueAsArray(arguments.To);
            var cc = GetStructureValueAsArray(arguments.Cc);
            var bcc = GetStructureValueAsArray(arguments.Bcc);
            var attachments = GetStructureValueAsArray(arguments.PathsToAttachments);

            LotusNotesManager.CurrentWrapper.SendEmail(
                to,
                arguments.Subject.Value,
                arguments.Message.Value,
                cc,
                bcc,
                attachments,
                arguments.SaveMessageOnSend.Value
            );
        }

        private static string[] GetStructureValueAsArray(Structure structure)
        {
            if (structure?.Object == null)
                return new string[0];

            if (structure is ListStructure toList)
                return toList.Value.Cast<string>().ToArray();
            else if (structure is TextStructure text)
                return new string[] { text.Value };
            else
                throw new ArgumentException("Argument is not a text nor list");
        }
    }
}
