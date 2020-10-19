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
using System.IO;

namespace G1ANT.Addon.LotusNotes.Commands.Documents.Items
{
    [Command(Name = "notes.saveattachment", Tooltip = "Save attachment to a file")]
    public class NotesSaveAttachmentCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Attachment")]
            public NotesAttachmentStructure Attachment { get; set; }

            [Argument(Tooltip = "True to overwrite if file already exists (default). False to raise error in such case.")]
            public BooleanStructure Overwrite { get; set; } = new BooleanStructure(true);

            [Argument(Required = true, Tooltip = "Path where the attachment will be saved. If it does not contain file name then attachment's file name will be used")]
            public TextStructure Path { get; set; }
        }

        public NotesSaveAttachmentCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var attachment = arguments.Attachment.Value;

            var path = arguments.Path.Value;
            if (Directory.Exists(path))
                path = Path.Combine(path, attachment.Name);

            if (arguments.Overwrite.Value && File.Exists(path))
                File.Delete(path);

            attachment.ExtractFile(path);
        }
    }
}
