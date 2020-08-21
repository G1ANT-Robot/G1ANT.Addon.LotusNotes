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

namespace G1ANT.Addon.LotusNotes.Commands
{
    [Command(Name = "notes.getfoldernames", Tooltip = "Get names of Notes' folders (folder is a view with IsFolder flag set)")]
    public class LotusGetFolderNamesCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a variable where folder names (`liststructure` with `text`) will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public LotusGetFolderNamesCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var folderNames = LotusNotesManager.CurrentWrapper.GetFolderNames();
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new ListStructure(folderNames));
        }
    }
}
