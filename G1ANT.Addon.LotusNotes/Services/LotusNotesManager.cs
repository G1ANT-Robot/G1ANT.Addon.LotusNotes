/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System;
using System.Collections.Generic;
using System.Linq;


namespace G1ANT.Addon.LotusNotes.Services
{
    public static class LotusNotesManager
    {
        private static List<LotusNotesWrapper> wrappers = new List<LotusNotesWrapper>();

        public static LotusNotesWrapper CurrentWrapper { get; private set; }

        public static int GetNextId() => wrappers.Select(w => w.Id).DefaultIfEmpty(0).FirstOrDefault();


        public static LotusNotesWrapper AddWrapper()
        {
            int assignedId = GetNextId();
            CurrentWrapper = new LotusNotesWrapper(assignedId);
            wrappers.Add(CurrentWrapper);
            return CurrentWrapper;
        }

        public static void SwitchWrapper(int id)
        {
            var instanceToSwitchTo = wrappers.FirstOrDefault(x => x.Id == id);
            CurrentWrapper = instanceToSwitchTo ?? throw new ArgumentException($"No wrapper found with id: {id}");
        }

        public static void Remove(int id)
        {
            var wrapperToRemove = wrappers.FirstOrDefault(x => x.Id == id);

            if (wrapperToRemove != null)
            {
                //wrapperToRemove.Close();
                wrappers.Remove(wrapperToRemove);
                CurrentWrapper = wrappers.FirstOrDefault();
            }
        }
    }
}
