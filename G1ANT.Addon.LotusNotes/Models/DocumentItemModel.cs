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

namespace G1ANT.Addon.LotusNotes.Services
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
        public DocumentModel Parent { get; }
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
            Text = item.Text;
            Type = item.type;
            ValueLength = item.ValueLength;
            if (item.Values is IEnumerable && !(item.Values is string))
                Values = item.Values;
            else// if (Values != null)
                Value = Values?.ToString();

            //            Values = ((IEnumerable)(item.Values)).Cast<string>().ToList();

            //var items = ((object[])document.Items).Cast<NotesItem>();

            //if (document.HasEmbedded && document.HasItem("$File"))
            //{
            //    foreach (var item in items)
            //    {
            //        if (item.type == IT_TYPE.ATTACHMENT)
            //        {
            //            var v = ((IEnumerable)(item.Values)).Cast<string>().ToList();
            //            //item.Values
            //            //document.GetAttachment()
            //        }
            //    }
            //}            
        }
    }
}

