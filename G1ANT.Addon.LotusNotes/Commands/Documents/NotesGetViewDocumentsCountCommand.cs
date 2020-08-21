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

namespace G1ANT.Addon.LotusNotes.Commands.Documents
{
    [Command(Name = "notes.getemailcount", Tooltip = "Get number of documents (emails) in specified view (folder)")]
    public class NotesGetViewDocumentsCountCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of view (folder) which contents will be counted. Default is inbox (`($Inbox)`), list of all folders can be obtained using `notes.getfoldernames` command")]
            public TextStructure Folder { get; set; } = new TextStructure("($Inbox)");

            [Argument(Tooltip = "Name of a variable where result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesGetViewDocumentsCountCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments.Folder.Value))
                throw new ArgumentNullException(nameof(arguments.Folder));

            var view = LotusNotesManager.CurrentWrapper.GetView(arguments.Folder.Value);
            if (view == null)
                throw new ApplicationException($"$View (folder) {arguments.Folder} not found");

            var count = view.EntryCount;

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new IntegerStructure(count));
        }
    }
}
