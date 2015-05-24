using System;
using GedToVisio.Gedcom;
using GedToVisio.Properties;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{
    /// <summary>
    /// Отображает генеалогические объекты в документе MS Visio.
    /// </summary>
    public class VisioRenderer
    {
        public enum ConnectionPoint
        {
            Left = 0,
            Rigth = 1,
            Center = -1
        }

        readonly Page _page;

        private static double ScaleX(int x)
        {
            return Settings.Default.OffsetX + x * Settings.Default.ScaleX;
        }

        private static double ScaleY(int y)
        {
            return Settings.Default.OffsetY + y * Settings.Default.ScaleY;
        }


        public VisioRenderer()
        {
            Application application;
            try
            {
                application = new Microsoft.Office.Interop.Visio.Application();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {                
                throw new Exception("MS Visio not found", ex);
            }
            application.Documents.Add("");
            _page = application.Documents[1].Pages[1];            
        }


        public void Connect(Shape shapeFrom, ConnectionPoint shapeFromConnectionPoint, Shape shapeTo, ConnectionPoint shapeToConnectionPoint)
        {
            VisioHelper.Connect(shapeFrom, shapeFromConnectionPoint == ConnectionPoint.Center ? null : (short?)shapeFromConnectionPoint, 
                shapeTo, shapeToConnectionPoint == ConnectionPoint.Center ? null : (short?)shapeToConnectionPoint);
        }


        public void Move(Shape shape, int x, int y)
        {
            shape.SetCenter(ScaleX(x), ScaleY(y));           
        }


        public Shape Render(Individual indi, int x, int y, bool hasParents)
        {
            var sx = ScaleX(x);
            var sy = ScaleY(y);
            double height = Settings.Default.IndiHeight;
            double width = Settings.Default.IndiWidth;
            string femaleRounding = Settings.Default.IndiFemaleRounding;
            short indiCharacterSize = Settings.Default.IndiCharacterSize;

            var shape = VisioHelper.DrawRectangle(_page, sx - width / 2, sy - height / 2, sx + width / 2, sy + height / 2);
            shape.Characters.CharProps[(short)VisCellIndices.visCharacterSize] = indiCharacterSize;
            if (indi.Sex == "F")
            {
                shape.CellsU["Rounding"].FormulaU = femaleRounding;
            }

            shape.Text = ShapeText(indi, hasParents);

            VisioHelper.SetCustomProperty(shape, "_UID", "Unique Identification Number", indi.Uid);
            VisioHelper.SetCustomProperty(shape, "ID", "gedcom ID", indi.Id);
            VisioHelper.SetCustomProperty(shape, "GivenName", indi.GivenName);
            VisioHelper.SetCustomProperty(shape, "Surname", indi.Surname);
            VisioHelper.SetCustomProperty(shape, "Sex", indi.Sex);
            VisioHelper.SetCustomProperty(shape, "BirthDate", indi.BirthDate);
            VisioHelper.SetCustomProperty(shape, "DiedDate", indi.DiedDate);
            VisioHelper.SetCustomProperty(shape, "Notes", string.Join("\n", indi.Notes));

            return shape;
        }

        private string ShapeText(Individual indi, bool hasParents)
        {
            var notes = string.Join("\n", indi.Notes);
            var text = string.Format("{0}{1}\n{2}{3}", 
                // имя
                indi.GivenName,
                // фамилия, если нет предков
                hasParents || string.IsNullOrEmpty(indi.Surname) ? "" : string.Format(" {0}", indi.Surname),
                // даты жизни
                string.IsNullOrEmpty(indi.BirthDate) && string.IsNullOrEmpty(indi.DiedDate)
                    ? ""
                    : string.Format("{0} - {1}", FormatDate(indi.BirthDate), FormatDate(indi.DiedDate)),
                // примечания
                string.IsNullOrEmpty(notes) ? "" : "\n" + notes
                );
            return text;
        }

        public Shape Render(Family fam, int x, int y)
        {
            var sx = ScaleX(x);
            var sy = ScaleY(y);
            double height = Settings.Default.FamilyHeight;
            double width = Settings.Default.FamilyWidth;
            short familyCharacterSize = Settings.Default.FamilyCharacterSize;

            var shape = _page.DrawOval(sx - width / 2, sy - height / 2, sx + width / 2, sy + height / 2);
            shape.Text = FormatDate(fam.MarriageDate);
            shape.Characters.CharProps[(short)VisCellIndices.visCharacterSize] = familyCharacterSize;

            VisioHelper.SetCustomProperty(shape, "_UID", "Unique Identification Number", fam.Uid);
            VisioHelper.SetCustomProperty(shape, "ID", "gedcom ID", fam.Id);
            VisioHelper.SetCustomProperty(shape, "MarriageDate", fam.MarriageDate);

            return shape;
        }

        private string FormatDate(string date)
        {
            return date == null ? null : date.Replace("ABT ", "~");
        }

    }
}
