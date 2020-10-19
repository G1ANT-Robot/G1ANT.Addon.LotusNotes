
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
    [Command(Name = "notes.close", Tooltip = "Connects to Lotus Notes/IBM Notes server and remote database. Before you can use this command you must have Notes client installed and configured, because Notes interop uses user login stored on your host by Notes client. That's also explaining why there's no argument `login`.")]
    public class NotesCloseCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Id of Lotus Notes Manager returned")]
            public IntegerStructure Id { get; set; }
        }

        public NotesCloseCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            LotusNotesManager.Remove(arguments.Id.Value);
        }
    }
}
