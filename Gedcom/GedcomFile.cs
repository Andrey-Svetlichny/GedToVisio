using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GedToVisio.Gedcom
{
    public class GedcomFile
    {
        private Header _header;
        public Header Header
        {
            get { return _header; }
            private set { _header = value; }
        }

        public List<Individual> Individuals { get; private set; }
        public List<Family> Families { get; private set; }
        public List<Note> Notes { get; private set; }

        private Individual _currentIndividual;
        private Family _currentFamily;
        private Note _currentNote;
        private GedcomRecordEnum _currentRecord = GedcomRecordEnum.None;
        private GedcomSubRecordEnum _currentSubRecord = GedcomSubRecordEnum.None;

        public GedcomFile()
        {
            Header = new Header();
            Individuals = new List<Individual>();
            Families = new List<Family>();
            Notes = new List<Note>();
        }

        public void Load(string filename, Encoding encoding)
        {
            var reader = new StreamReader(filename, encoding);
            string line = "";
            char[] separators = new char[1] {' '};
            while (!reader.EndOfStream)
            {
                line = reader.ReadLine().Replace("'", "''");
                while (line.IndexOf("  ") > 0)
                {
                    line = line.Replace("  ", " ");
                }
                string[] lineArray = line.Split(separators, 3);
                switch (lineArray[0])
                {
                    case "0":
                        ProcessRootLevel(lineArray);
                        break;
                    case "1":
                        ProcessLevel1(lineArray);
                        break;
                    case "2":
                        ProcessLevel2(lineArray);
                        break;
                }

            }
        }

        private void ProcessRootLevel(string[] lineArray)
        {
            switch (_currentRecord)
            {
                case GedcomRecordEnum.Individual:
                    Individuals.Add(_currentIndividual);
                    break;
                case GedcomRecordEnum.Family:
                    Families.Add(_currentFamily);
                    break;
                case GedcomRecordEnum.Note:
                    Notes.Add(_currentNote);
                    break;
            }

            if (lineArray[1] == "HEAD")
            {
                    _currentRecord = GedcomRecordEnum.Header;
                    _currentSubRecord = GedcomSubRecordEnum.None;
            } else if (lineArray[1].IndexOf("@") >= 0) {
                switch (lineArray[2])
                {
                    case "INDI":
                        _currentRecord = GedcomRecordEnum.Individual;
                        _currentIndividual = new Individual { Id = lineArray[1] };
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "FAM":
                        _currentRecord = GedcomRecordEnum.Family;
                        _currentFamily = new Family { Id = lineArray[1] };
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "NOTE":
                        _currentRecord = GedcomRecordEnum.Note;
                        _currentNote = new Note { Id = lineArray[1] };
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                }
            }
        }

        private void ProcessLevel1(string[] lineArray)
        {
            var s1 = lineArray[1];
            var s2 = lineArray.Length > 2 ? lineArray[2] : "";
            if (_currentRecord == GedcomRecordEnum.Header)
            {
                switch (s1)
                {
                    case "SOUR":
                        _header.Source = s2;
                        _currentSubRecord = GedcomSubRecordEnum.HeaderSource;
                        break;
                    case "DEST":
                        _header.Destination = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "DATE":
                        _header.Date = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "FILE":
                        _header.File = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "CHAR":
                        _header.CharacterEncoding = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "GEDC":
                        _currentSubRecord = GedcomSubRecordEnum.HeaderGedcom;
                        break;
                }
            }
            else if (_currentRecord == GedcomRecordEnum.Individual)
            {
                switch (s1)
                {
                    case "_UID":
                        _currentIndividual.Uid = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "NAME":
                        _currentIndividual.GivenName = s2;
                        _currentSubRecord = GedcomSubRecordEnum.IndividualName;
                        break;
                    case "SEX":
                        _currentIndividual.Sex = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "BIRT":
                        _currentSubRecord = GedcomSubRecordEnum.IndividualBirth;
                        break;
                    case "DEAT":
                        _currentSubRecord = GedcomSubRecordEnum.IndividualDeath;
                        break;
                    case "OCCU":
                        _currentIndividual.Occupation = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "DSCR":
                        _currentIndividual.Description = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "NATI":
                        _currentIndividual.Nationality = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "NOTE":
                        _currentIndividual.Notes.Add(s2);
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                }
            }
            else if (_currentRecord == GedcomRecordEnum.Family)
            {
                switch (s1)
                {
                    case "_UID":
                        _currentFamily.Uid = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "HUSB":
                        _currentFamily.HusbandId = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "WIFE":
                        _currentFamily.WifeId = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "CHIL":
                        _currentFamily.Children.Add(s2);
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "MARR":
                        _currentSubRecord = GedcomSubRecordEnum.FamilyMarriage;
                        break;
                    case "NOTE":
                        _currentFamily.Notes.Add(s2);
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                }
            }
            else if (_currentRecord == GedcomRecordEnum.Note)
            {
                switch (s1)
                {
                    case "CONC":
                        _currentNote.Text = s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                    case "CONT":
                        _currentNote.Text += s2;
                        _currentSubRecord = GedcomSubRecordEnum.None;
                        break;
                }
            }
        }

        private void ProcessLevel2(string[] lineArray)
        {
            if (_currentSubRecord == GedcomSubRecordEnum.HeaderSource)
            {
                switch (lineArray[1])
                {
                    case "VERS":
                        _header.SourceVersion = lineArray[2];
                        break;
                    case "NAME":
                        _header.SourceName = lineArray[2];
                        break;
                    case "CORP":
                        _header.SourceCorporation = lineArray[2];
                        break;
                }
            }
            else if (_currentSubRecord == GedcomSubRecordEnum.HeaderGedcom)
            {
                switch (lineArray[1])
                {
                    case "VERS":
                        _header.GedcomVersion = lineArray[2];
                        break;
                    case "FORM":
                        _header.GedcomForm = lineArray[2];
                        break;
                }
            }
            else if (_currentSubRecord == GedcomSubRecordEnum.IndividualName)
            {
                switch (lineArray[1])
                {
                    case "GIVN":
                        _currentIndividual.GivenName = lineArray[2];
                        break;
                    case "SURN":
                        _currentIndividual.Surname = lineArray[2];
                        break;
                    case "NSFX":
                        _currentIndividual.Suffix = lineArray[2];
                        break;
                }
            }
            else if (_currentSubRecord == GedcomSubRecordEnum.IndividualBirth)
            {
                switch (lineArray[1])
                {
                    case "DATE":
                        _currentIndividual.BirthDate = lineArray[2];
                        break;
                    case "PLAC":
                        _currentIndividual.BirthPlace = lineArray[2];
                        break;
                }
            }
            else if (_currentSubRecord == GedcomSubRecordEnum.IndividualDeath)
            {
                switch (lineArray[1])
                {
                    case "DATE":
                        _currentIndividual.DiedDate = lineArray[2];
                        break;
                    case "PLAC":
                        _currentIndividual.DiedPlace = lineArray[2];
                        break;
                    case "CAUS":
                        _currentIndividual.DiedCause = lineArray[2];
                        break;
                }
            }
            else if (_currentSubRecord == GedcomSubRecordEnum.FamilyMarriage)
            {
                switch (lineArray[1])
                {
                    case "DATE":
                        _currentFamily.MarriageDate = lineArray[2];
                        break;
                    case "PLAC":
                        _currentFamily.MarriagePlace = lineArray[2];
                        break;
                }
            }
        }

        public void Save(string filename, Encoding encoding)
        {
            var writer = new StreamWriter(filename, false, encoding);
            Individual currentGedcomIndividual;
            Family currentGedcomFamily;

            writer.WriteLine("0 HEAD");
            if ((_header.Source.Length > 0) || (_header.SourceCorporation.Length) > 0 || (_header.SourceName.Length > 0))
            {
                writer.WriteLine("1 SOUR " + _header.Source);
                if (_header.SourceVersion.Length > 0)
                {
                    writer.WriteLine("2 VERS " + _header.SourceVersion);
                }
                if (_header.SourceName.Length > 0)
                {
                    writer.WriteLine("2 NAME " + _header.SourceName);
                }
                if (_header.SourceCorporation.Length > 0)
                {
                    writer.WriteLine("2 CORP " + _header.SourceCorporation);
                }
            }
            if (_header.Destination.Length > 0)
            {
                writer.WriteLine("1 DEST " + _header.Destination);
            }
            if (_header.Date.Length > 0)
            {
                writer.WriteLine("1 DATE " + _header.Date);
            }
            if (!string.IsNullOrEmpty(_header.File))
            {
                writer.WriteLine("1 FILE " + _header.File);
            }
            if (_header.CharacterEncoding.Length > 0)
            {
                writer.WriteLine("1 CHAR " + _header.CharacterEncoding);
            }
            if ((_header.GedcomVersion.Length > 0) || (_header.GedcomForm.Length > 0))
            {
                writer.WriteLine("1 GEDC");
                if (_header.GedcomVersion.Length > 0)
                {
                    writer.WriteLine("1 VERS " + _header.GedcomVersion);
                }
                if (_header.GedcomForm.Length > 0)
                {
                    writer.WriteLine("1 FORM " + _header.GedcomForm);
                }
            }

            for (int i = 0; i < Individuals.Count; i++)
            {
                currentGedcomIndividual = (Individual)(Individuals[i]);
                writer.WriteLine("0 " + currentGedcomIndividual.Id + " INDI");
                writer.WriteLine("1 " + currentGedcomIndividual.GivenName + "/" + currentGedcomIndividual.Surname + "/");
                if (currentGedcomIndividual.GivenName.Length > 0)
                {
                    writer.WriteLine("2 GIVN " + currentGedcomIndividual.GivenName);
                }
                if (currentGedcomIndividual.Surname.Length > 0)
                {
                    writer.WriteLine("2 SURN " + currentGedcomIndividual.Surname);
                }
                if (currentGedcomIndividual.Sex.Length > 0)
                {
                    writer.WriteLine("1 SEX " + currentGedcomIndividual.Sex);
                }
                if (currentGedcomIndividual.Occupation.Length > 0)
                {
                    writer.WriteLine("1 OCCU " + currentGedcomIndividual.Occupation);
                }
                if ((currentGedcomIndividual.BirthDate.Length > 0) || (currentGedcomIndividual.BirthPlace.Length > 0))
                {
                    writer.WriteLine("1 BIRT");
                    if (currentGedcomIndividual.BirthDate.Length > 0)
                    {
                        writer.WriteLine("2 DATE " + currentGedcomIndividual.BirthDate);
                    }
                    if (currentGedcomIndividual.BirthDate.Length > 0)
                    {
                        writer.WriteLine("2 PLAC " + currentGedcomIndividual.BirthPlace);
                    }
                }
                if ((currentGedcomIndividual.DiedDate.Length > 0) || (currentGedcomIndividual.DiedPlace.Length > 0))
                {
                    writer.WriteLine("1 DEAT");
                    if (currentGedcomIndividual.DiedDate.Length > 0)
                    {
                        writer.WriteLine("2 DATE " + currentGedcomIndividual.DiedDate);
                    }
                    if (currentGedcomIndividual.DiedDate.Length > 0)
                    {
                        writer.WriteLine("2 PLAC " + currentGedcomIndividual.DiedPlace);
                    }
                }
                if (currentGedcomIndividual.Nationality.Length > 0)
                {
                    writer.WriteLine("1 NATI " + currentGedcomIndividual.Nationality);
                }
                if (currentGedcomIndividual.ParentFamilyId.Length > 0)
                {
                    writer.WriteLine("1 FAMC " + currentGedcomIndividual.ParentFamilyId);
                }
                if (currentGedcomIndividual.SpouseFamilyId.Length > 0)
                {
                    writer.WriteLine("1 FAMS " + currentGedcomIndividual.SpouseFamilyId);
                }

            }
            for (int i = 0; i < Families.Count; i++)
            {
                currentGedcomFamily = (Family)(Families[i]);
                writer.WriteLine("0 " + currentGedcomFamily.Id + " FAM");
                if (currentGedcomFamily.HusbandId.Length > 0)
                {
                    writer.WriteLine("1 HUSB " + currentGedcomFamily.HusbandId);
                }
                if (currentGedcomFamily.WifeId.Length > 0)
                {
                    writer.WriteLine("1 WIFE " + currentGedcomFamily.WifeId);
                }
                if ((currentGedcomFamily.MarriageDate.Length > 0) || (currentGedcomFamily.MarriagePlace.Length > 0))
                {
                    writer.WriteLine("1 MARR");
                    if (currentGedcomFamily.MarriageDate.Length > 0)
                    {
                        writer.WriteLine("2 DATE " + currentGedcomFamily.MarriageDate);
                    }
                    if (currentGedcomFamily.MarriagePlace.Length > 0)
                    {
                        writer.WriteLine("2 PLAC " + currentGedcomFamily.MarriagePlace);
                    }
                }
                for (int j = 0; j < currentGedcomFamily.Children.Count; j++)
                {
                    writer.WriteLine("1 CHIL " + currentGedcomFamily.Children[j].ToString());
                }
            }
            writer.WriteLine("0 TRLR");
            writer.Close();
        }

    }
}
