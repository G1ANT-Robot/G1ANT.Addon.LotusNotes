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
    [Command(Name = "notes.connect", Tooltip = "Connects to Lotus Notes/IBM Notes server and remote database. Before you can use this command you must have Notes client installed and configured, because Notes interop uses user login stored on your host by Notes client. That's also explaining why there's no argument `login`.")]
    public class NotesConnectCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Lotus Notes/IBM Notes server name")]
            public TextStructure Server { get; set; }

            [Argument(Required = true, Tooltip = "Password")]
            public TextStructure Password { get; set; }

            [Argument(Required = true, Tooltip = "Path to database file including file name and extension (.nsf)")]
            public TextStructure DatabaseFile { get; set; }

            [Argument(Tooltip = "Name of a variable where newly created instance of Lotus Notes Manager will be stored. You can use `notes.switch` command to switch between managers and work with multiple Notes instances")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesConnectCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var wrapper = LotusNotesManager.CreateWrapper();
            wrapper.Connect(arguments.Password.Value, arguments.Server.Value, arguments.DatabaseFile.Value);
            LotusNotesManager.RegisterAndSetAsCurrentWrapper(wrapper);

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new IntegerStructure(wrapper.Id));
        }
    }
}
