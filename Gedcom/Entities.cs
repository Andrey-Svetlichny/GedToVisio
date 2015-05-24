using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace GedToVisio.Gedcom
{
    public class Header
    {
        public string Source;
        public string SourceVersion;
        public string SourceName;
        public string SourceCorporation;
        public string Destination;
        public string Date;
        public string File;
        public string CharacterEncoding;
        public string GedcomVersion;
        public string GedcomForm;
    }

    public class GedcomObject
    {
        public string Uid;
        public string Id;        
    }

    [DebuggerDisplay("{GivenName} {Surname}")]
    public class Individual : GedcomObject
    {
        public string GivenName;
        public string Surname;
        public string Suffix;
        public string Sex;
        public string BirthDate;
        public string BirthPlace;
        public string Occupation;
        public string Description;
        public string Nationality;
        public string DiedDate;
        public string DiedPlace;
        public string DiedCause;
        public string ParentFamilyId;
        public string SpouseFamilyId;
        public List<string> Notes = new List<string>();
    }

    [DebuggerDisplay("Family")]
    public class Family : GedcomObject
    {
        public string HusbandId;
        public string WifeId;
        public string MarriageDate;
        public string MarriagePlace;
        public List<string> Children = new List<string>();
        public List<string> Notes = new List<string>();
    }

    public class Note
    {
        public string Id;
        public string Text;
    }

    public enum GedcomRecordEnum
    {
        None,
        Header,
        Individual,
        Family,
        Note
    }

    public enum GedcomSubRecordEnum
    {
        None,
        HeaderSource,
        HeaderGedcom,
        IndividualName,
        IndividualBirth,
        IndividualDeath,
        FamilyChildren,
        FamilyMarriage
    }
}
