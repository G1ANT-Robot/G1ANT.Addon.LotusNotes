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
using G1ANT.Addon.SAP.Structures;
using G1ANT.Language;
using System;

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.getemailbyindex", Tooltip = "Get document (emails) by its index in parent view (folder).")]
    public class NotesGetDocumentByIndexCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Index")]
            public IntegerStructure Index { get; set; }

            [Argument(Required = true, Tooltip = "Name of view (folder) that is parent of desired document (email). Default is inbox (`($Inbox)`), list of all folders can be obtained using `notes.getfoldernames` command")]
            public TextStructure Folder { get; set; } = new TextStructure("($Inbox)");

            [Argument(Tooltip = "Name of a variable where document (email) will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesGetDocumentByIndexCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments.Folder.Value))
                throw new ArgumentNullException(nameof(arguments.Folder));

            var view = LotusNotesManager.CurrentWrapper.GetView(arguments.Folder.Value);
            if (view == null)
                throw new ApplicationException($"$View (folder) {arguments.Folder} not found");

            var document = LotusNotesManager.CurrentWrapper.GetDocumentByIndex(view, arguments.Index.Value);
            if (document == null)
                throw new ApplicationException("Document (email) not found");

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new NotesDocumentStructure(document));
        }
    }
}
