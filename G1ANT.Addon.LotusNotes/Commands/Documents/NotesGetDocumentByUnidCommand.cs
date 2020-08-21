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
    [Command(Name = "notes.getemailbyunid", Tooltip = "Get document (emails) by its Unique Notes Identifier (unid). This value is stored in UniversalID property.")]
    public class NotesGetDocumentByUnidCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Value of UNID")]
            public TextStructure Unid { get; set; }

            [Argument(Tooltip = "Name of a variable where document (email) will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public NotesGetDocumentByUnidCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments.Unid.Value))
                throw new ArgumentNullException(nameof(arguments.Unid));

            var document = LotusNotesManager.CurrentWrapper.GetDocumentByUnid(arguments.Unid.Value);
            if (document == null)
                throw new ApplicationException("Document (email) not found");

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new NotesDocumentStructure(document));
        }
    }
}
