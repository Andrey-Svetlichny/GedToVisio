using System.Collections.Generic;
using GedToVisio.Gedcom;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{

    /// <summary>
    /// Определяет порядок семей по оси Y (внутри поколения).
    /// Не визуальный, используется только для упорядочивания.
    /// </summary>
    public class OrderLink
    {
        public VisualObjectFamily Upper { get; set; } 
        public VisualObjectFamily Lower { get; set; }

        public OrderLink(VisualObjectFamily upper, VisualObjectFamily lower)
        {
            Upper = upper;
            Lower = lower;
        }

        public override string ToString()
        {
            return string.Format("{0} => {1}", Upper, Lower);
        }
    }

    /// <summary>
    /// Визуальный объект. Обертка вокруг Individual.
    /// </summary>
    public class VisualObjectIndividual : VisualObject
    {
        /// <summary>
        /// Семья, в которой он родился.
        /// </summary>
        public VisualObjectFamily Family { get; set; }

        /// <summary>
        /// Семьи, которые он создал.
        /// </summary>
        public List<VisualObjectFamily> ChildFamilies { get; set; }

        public string FullName()
        {
            var indi = (Individual)GedcomObject;
            return string.Format("{0} {1}", indi.GivenName, indi.Surname);
        }

        public override string ToString()
        {
            return LevelCoordinatesInfo() + " " + FullName();
        }
    }

    /// <summary>
    /// Визуальный объект. Обертка вокруг Family.
    /// </summary>
    public class VisualObjectFamily : VisualObject
    {
        /// <summary>
        /// Муж.
        /// </summary>
        public VisualObjectIndividual Husband { get; set; }

        /// <summary>
        /// Жена.
        /// </summary>
        public VisualObjectIndividual Wife { get; set; }

        /// <summary>
        /// Дети.
        /// </summary>
        public List<VisualObjectIndividual> Children { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1} - {2}", LevelCoordinatesInfo(), Husband == null ? "" : Husband.FullName(), Wife == null ? "" : Wife.FullName());
        } 
    }

    /// <summary>
    /// Базовый визуальный объект. Обертка вокруг GedcomObject.
    /// Имеет координаты X,Y, ссылку на Shape и ссылки на дочерние и родительские объекты в дереве.
    /// </summary>
    public class VisualObject
    {
        public string Id { get; set; }

        /// <summary>
        /// Уровень в дереве, начиная от корня (поколение).
        /// </summary>
        public int Level { get; set; }

        public int X { get; set; }
        public int Y { get; set; }
        public Shape Shape { get; set; }
        public Gedcom.GedcomObject GedcomObject { get; set; }

        public List<VisualObject> ParentObjects { get; set; }

        public List<VisualObject> ChildObjects { get; set; }


        public string LevelCoordinatesInfo()
        {
            return string.Format("[{0}] {1} {2}", Level, X, Y);
        }
    }
}
