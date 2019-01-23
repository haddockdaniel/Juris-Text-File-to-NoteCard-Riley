using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class NoteCard
    {
        public NoteCard()
        {
            partyName = "";
            partyType = "";
            synopsis = "";
        }

        public string Client { get; set; }
        public string matter { get; set; }
        public string clientID { get; set; }
        public string matterID { get; set; }
        public string clientName { get; set; }
        public string partyType { get; set; }
        public string partyName { get; set; }
        public string synopsis { get; set; }

    }
}
