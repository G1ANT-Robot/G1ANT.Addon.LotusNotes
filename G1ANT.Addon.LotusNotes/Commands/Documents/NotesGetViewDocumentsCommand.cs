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
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.getemails", Tooltip = "Get documents (emails) from specified view (folder)")]
    public class NotesGetViewDocumentsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of view (folder) which contents will be returned. Default is inbox (`($Inbox)`), list of all folders can be obtained using `notes.getfoldernames` command")]
            public TextStructure Folder { get; set; } = new TextStructure("($Inbox)");

            [Argument(Tooltip = "Name of a variable where documents (`liststructure` with `notesemailstructure`) will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesGetViewDocumentsCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments.Folder.Value))
                throw new ArgumentNullException(nameof(arguments.Folder));

            var view = LotusNotesManager.CurrentWrapper.GetView(arguments.Folder.Value);
            var documents = view.GetDocuments().ToList();

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new ListStructure(documents.Select(d => new NotesDocumentStructure(d))));
        }
    }
}
