/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.LotusNotes.Models;
using G1ANT.Addon.LotusNotes.Services;
using G1ANT.Addon.SAP.Structures;
using G1ANT.Language;
using System;

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.getemail", Tooltip = "Get document (emails) by its index in parent view (folder) or by Universal Notes Identifier (UNID - this is the value from email's property `UniversalID`)")]
    public class NotesGetDocumentCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Index (offset) of email in its folder (view)")]
            public IntegerStructure Index { get; set; }

            [Argument(Tooltip = "Name of view (folder) that is parent of desired document (email). Default is inbox (`($Inbox)`), list of all folders can be obtained using `notes.getfoldernames` command. Required to get email by index")]
            public TextStructure Folder { get; set; } = new TextStructure("($Inbox)");

            [Argument(Tooltip = "Value of UNID")]
            public TextStructure Unid { get; set; }

            [Argument(Tooltip = "Name of a variable where document (email) will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesGetDocumentCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            DocumentModel document = arguments.Unid != null
                ? GetDocumentByUnid(arguments.Unid?.Value)
                : GetDocumentByIndex(arguments.Index?.Value, arguments.Folder?.Value);

            if (document == null)
                throw new ApplicationException("Document (email) not found");

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new NotesDocumentStructure(document));
        }

        private static DocumentModel GetDocumentByIndex(int? index, string folder)
        {
            ValidateArguments(index, folder);

            return LotusNotesManager.CurrentWrapper.GetDocumentByIndex(
                GetViewByName(folder),
                index.Value
            );
        }

        private static ViewModel GetViewByName(string folder)
        {
            var view = LotusNotesManager.CurrentWrapper.GetView(folder);
            if (view == null)
                throw new ApplicationException($"$View (folder) {folder} not found");
            return view;
        }

        private static void ValidateArguments(int? index, string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
                throw new ArgumentNullException(nameof(folder));
            if (index == null)
                throw new ArgumentNullException(nameof(index));
        }

        private static DocumentModel GetDocumentByUnid(string unid)
        {
            if (string.IsNullOrEmpty(unid))
                throw new ArgumentNullException(nameof(unid));

            return LotusNotesManager.CurrentWrapper.GetDocumentByUnid(unid);
        }
    }
}
