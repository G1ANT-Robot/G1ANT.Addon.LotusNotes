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
using System;

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.removefromfolder", Tooltip = "Remove document (email) from view (folder)")]
    public class NotesRemoveDocumentFromViewCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Document (email) to be removed")]
            public NotesDocumentStructure Document { get; set; }

            [Argument(Required = true, Tooltip = "Name of view (folder) where from document should be removed")]
            public TextStructure Folder { get; set; } = new TextStructure("($Inbox)");
        }

        public NotesRemoveDocumentFromViewCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments.Folder.Value))
                throw new ArgumentNullException(nameof(arguments.Folder));

            var document = arguments.Document.Value;
            document.RemoveFromFolder(arguments.Folder.Value);
        }
    }
}
