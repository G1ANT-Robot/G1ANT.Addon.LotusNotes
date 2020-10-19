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

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.removeemail", Tooltip = "Remove document (email)")]
    public class NotesRemoveDocumentCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Document (email) to be removed")]
            public NotesDocumentStructure Document { get; set; }

            [Argument(Tooltip = "True to force removal. False by default.")]
            public BooleanStructure Force { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "True to remove document permanently. False by default.")]
            public BooleanStructure Permanent { get; set; } = new BooleanStructure(false);
        }

        public NotesRemoveDocumentCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var document = arguments.Document.Value;
            
            if (arguments.Permanent.Value)
                document.RemovePermanently(arguments.Force.Value);
            else
                document.Remove(arguments.Force.Value);
        }
    }
}
