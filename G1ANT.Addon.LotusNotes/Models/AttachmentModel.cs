/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using Domino;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class AttachmentModel : PropertiesContainer
    {
        private readonly NotesEmbeddedObject embeddedObject;

        //public string FileName { get; }
        public int FileSize { get; }
        public string Name { get; }
        public string Source { get; }
        public EMBED_TYPE EmbedType { get; }


        public static bool IsAttachmentItem(NotesItem item) => item.type == IT_TYPE.ATTACHMENT;

        public AttachmentModel(NotesDocument document, NotesItem item)
        {
            if (!IsAttachmentItem(item))
                throw new ArgumentException("Item is not an attachment", nameof(item));

            var name = ((IEnumerable)item.Values).Cast<string>().First();
            embeddedObject = document.GetAttachment(name);

            FileSize = embeddedObject.FileSize;
            Name = embeddedObject.Name;
            Source = embeddedObject.Source;
            EmbedType = embeddedObject.type;
        }

        public object Activate(bool show) => embeddedObject.Activate(show);
        //public void DoVerb(string verb) => embeddedObject.DoVerb(verb);
        public void ExtractFile(string path) => embeddedObject.ExtractFile(path);
        public void Remove() => embeddedObject.Remove();

        public override string ToString() => JObject.FromObject(this).ToString();
    }
}

