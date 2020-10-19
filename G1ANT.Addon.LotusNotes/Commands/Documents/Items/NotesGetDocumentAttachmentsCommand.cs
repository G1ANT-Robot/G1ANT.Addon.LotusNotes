/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.SAP.Structures;
using G1ANT.Language;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Commands.Documents.Items
{
    [Command(Name = "notes.getemailattachments", Tooltip = "Get document (email) attachments")]
    public class NotesGetDocumentAttachmentsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Document (email) with attachments")]
            public NotesDocumentStructure Document { get; set; }

            [Argument(Tooltip = "Name of a variable where result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesGetDocumentAttachmentsCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var attachments = arguments.Document.Value.GetAttachments();

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new ListStructure(attachments.Select(a => new NotesAttachmentStructure(a))));
        }
    }
}
