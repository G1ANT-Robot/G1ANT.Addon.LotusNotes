using G1ANT.Addon.LotusNotes.Models;
using G1ANT.Language;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.SAP.Structures
{
    [Structure(Name = "notesemail", Default = "", AutoCreate = false, Tooltip = "This structure stores Notes document (email) information")]
    public class NotesDocumentStructure : StructureTyped<DocumentModel>
    {
        //private const string ChildrenIndex = "children";

        public NotesDocumentStructure(string value, string format = "", AbstractScripter scripter = null) :
            base(value, format, scripter)
        {
        }

        public NotesDocumentStructure(object value, string format = null, AbstractScripter scripter = null)
            : base(value, format, scripter)
        {
        }

        public override IList<string> Indexes => Value.GetPropertyNames().ToList();


        public override Structure Get(string index = "")
        {
            index = index.ToLower();
            if (string.IsNullOrWhiteSpace(index))
                return this;

            var propertyValue = Value.GetPropertyValue(index);
            if (propertyValue is IEnumerable enumerable && !(propertyValue is string))
                return new ListStructure(enumerable.Cast<object>().Select(e => Scripter.Structures.CreateStructure(e)).ToList());
            return Scripter.Structures.CreateStructure(propertyValue);

            throw new ArgumentException($"Unknown index '{index}'");
        }

        public override void Set(Structure structure, string index = null)
        {
            if (structure == null || structure.Object == null)
                throw new ArgumentNullException(nameof(structure));
            else
            {
                Value.SetPropertyValue(index, structure.Object);
            }
        }

        private string serializedObject = null;
        public override string ToString(string format)
        {
            return serializedObject ?? (serializedObject = JObject.FromObject(Value).ToString());
        }


        protected override DocumentModel Parse(string value, string format = null)
        {
            throw new NotSupportedException();
        }
    }
}
