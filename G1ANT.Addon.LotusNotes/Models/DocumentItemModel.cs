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
using System.Collections;
using System.Linq;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class DocumentItemModel
    {
        public NotesDateTime DateTimeValue { get; set; }
        public bool IsAuthors { get; set; }
        public bool IsEncrypted { get; set; }
        public bool IsNames { get; set; }
        public bool IsProtected { get; set; }
        public bool IsReaders { get; set; }
        public bool IsSigned { get; set; }
        public bool IsSummary { get; set; }
        public dynamic LastModified { get; }
        public string Name { get; }
        //public DocumentModel Parent { get; }
        public bool SaveToDisk { get; set; }
        public string Text { get; }
        public IT_TYPE Type { get; }
        public int ValueLength { get; }
        public object[] Values { get; }
        public string Value { get; }

        public DocumentItemModel(NotesItem item)
        {
            DateTimeValue = item.DateTimeValue;
            IsAuthors = item.IsAuthors;
            IsEncrypted = item.IsEncrypted;
            IsNames = item.IsNames;
            IsProtected = item.IsProtected;
            IsReaders = item.IsReaders;
            IsSigned = item.IsSigned;
            IsSummary = item.IsSummary;
            LastModified = item.LastModified;
            Name = item.Name;
            //Parent = new DocumentModel(item.Parent);
            SaveToDisk = item.SaveToDisk;
            try { Text = item.Text; } catch { }
            Type = item.type;
            ValueLength = item.ValueLength;

            if (item.Values is object[])
                Values = (object[])item.Values;
            else if (item.Values != null)
                Values = new object[] { (object)item.Values };

            if (Values.Count() == 1)
                Value = Values.FirstOrDefault()?.ToString();
        }


        public override string ToString() => $"{Type}: {Name}";
    }
}

