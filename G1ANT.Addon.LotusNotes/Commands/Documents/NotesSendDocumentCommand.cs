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
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.sendemail", Tooltip = "Send document (email)")]
    public class NotesSendDocumentCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "")]
            public TextStructure To { get; set; }

            [Argument(Tooltip = "")]
            public ListStructure AdditionalTo { get; set; }

            [Argument(Tooltip = "")]
            public TextStructure Cc { get; set; }

            [Argument(Tooltip = "")]
            public ListStructure AdditionalCc { get; set; }

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
            var to = OptionalJoinElementAndList(arguments.To, arguments.AdditionalTo);
            var cc = OptionalJoinElementAndList(arguments.Cc, arguments.AdditionalCc);

            LotusNotesManager.CurrentWrapper.SendEmail(
                to,
                arguments.Subject.Value,
                arguments.Message.Value,
                cc,
                arguments.SaveMessageOnSend.Value
            );
        }

        private static List<string> OptionalJoinElementAndList(TextStructure element, ListStructure list)
        {
            var to = new List<string>();

            if (!string.IsNullOrWhiteSpace(element?.Value))
                to.Add(element.Value);

            if (list?.Value?.Any() == true)
                to.AddRange(list.Value.Select(t => t.ToString()));

            return to;
        }
    }
}
