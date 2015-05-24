using System;
using System.Collections.Generic;
using System.Linq;
using GedToVisio.Gedcom;

namespace GedToVisio.Visio
{
    /// <summary>
    /// Экспортирует в Visio, располагая поколения по оси X.
    /// </summary>
    public class VisioExport
    {
        private readonly List<VisualObjectIndividual> _visualObjectsIndividual = new List<VisualObjectIndividual>();
        private readonly List<VisualObjectFamily> _visualObjectsFamily = new List<VisualObjectFamily>();
        /// <summary>
        /// Все визуальные объекты (люди и семьи).
        /// </summary>
        public List<VisualObject> VisualObjects {
            get
            {
                return _visualObjectsIndividual.Cast<VisualObject>().Union(_visualObjectsFamily).ToList();
            }
        }

        private readonly List<OrderLink> _orderLinks = new List<OrderLink>();
        readonly VisioRenderer _renderer;


        public VisioExport()
        {
            _renderer = new VisioRenderer();
        }


        public void Add(Individual indi)
        {
            _visualObjectsIndividual.Add(new VisualObjectIndividual
            {
                Id = indi.Id, 
                GedcomObject = indi,
                ChildFamilies = new List<VisualObjectFamily>(),
                ChildObjects = new List<VisualObject>(), 
                ParentObjects = new List<VisualObject>()
            });
        }

        public void Add(Family family)
        {
            _visualObjectsFamily.Add(new VisualObjectFamily
            {
                Id = family.Id, 
                GedcomObject = family, 
                Children = new List<VisualObjectIndividual>(),
                ChildObjects = new List<VisualObject>(), 
                ParentObjects = new List<VisualObject>()
            });
        }


        /// <summary>
        /// Строит дерево загруженых объектов, заполняя Parents, Children, Husband, Wife.
        /// </summary>
        public void BuildTree()
        {
            var visualObjectIndividualsDict = _visualObjectsIndividual.ToDictionary(o => o.GedcomObject.Id);

            // обходим все семьи
            foreach (var visualObjectFamily in _visualObjectsFamily)
            {
                var fam = (Family)visualObjectFamily.GedcomObject;

                if (fam.HusbandId != null && visualObjectIndividualsDict.ContainsKey(fam.HusbandId))
                {
                    var husband = visualObjectIndividualsDict[fam.HusbandId];
                    visualObjectFamily.Husband = husband;
                    visualObjectFamily.ParentObjects.Add(husband);
                    husband.ChildFamilies.Add(visualObjectFamily);
                    husband.ChildObjects.Add(visualObjectFamily);
                }

                if (fam.WifeId != null && visualObjectIndividualsDict.ContainsKey(fam.WifeId))
                {
                    var wife = visualObjectIndividualsDict[fam.WifeId];
                    visualObjectFamily.Wife = wife;
                    visualObjectFamily.ParentObjects.Add(wife);
                    wife.ChildFamilies.Add(visualObjectFamily);
                    wife.ChildObjects.Add(visualObjectFamily);
                }

                foreach (var childId in fam.Children)
                {
                    if (visualObjectIndividualsDict.ContainsKey(childId))
                    {
                        var child = visualObjectIndividualsDict[childId];
                        child.Family = visualObjectFamily;
                        visualObjectFamily.Children.Add(child);
                        visualObjectFamily.ChildObjects.Add(child);
                        child.ParentObjects.Add(visualObjectFamily);
                    }
                }
            }
        }


        /// <summary>
        /// Расположить объекты на листе.
        /// </summary>
        public void Arrange()
        {
            CalcLevel();
            CreateOrderLinks();

            // arrange X on 1x1 grid
            VisualObjects.ForEach(o => o.X = o.Level);

            // arrange Y on 1x1 grid

            // упорядочиваем семьи на основании _orderLinks
            _visualObjectsFamily.ForEach(o => o.Y = 0);
            while (true)
            {
                var links = _orderLinks.Where(orderLink => orderLink.Upper.Y <= orderLink.Lower.Y).ToList();
                if (links.Count == 0)
                    break;
                links.ForEach(l => l.Upper.Y++);
            }

            // разделяем семьи, которые внутри одного поколения имеют одинаковый Y
            var famGroups = _visualObjectsFamily.GroupBy(o => o.Level).OrderBy(g => g.Key).ToList();
            foreach (var famGroup in famGroups)
            {
                // семьи внутри поколения, упорядоченные по Y
                var families = famGroup.OrderBy(o => o.Y).ToList();
                for (int i = 0; i < families.Count; i++)
                {
                    families[i].Y = i;
                }
            }

            foreach (var famGroup in famGroups)
            {
                int shift = 0;
                // семьи внутри поколения, упорядоченные по Y
                var families = famGroup.OrderBy(o => o.Y).ToList();
                foreach (var family in families)
                {
                    // количество детей в семье
                    var childCount = family.Children.Count;
                    family.Y += shift + childCount / 2;
                    if (childCount > 1)
                    {
                        shift += childCount - 1;
                    }

                    // дети
                    for (int j = 0; j < family.Children.Count; j++)
                    {
                        family.Children[j].Y = family.Y + j - childCount / 2;
                    }                   
                }
            }


            // размещаем родителей, у которых нет семьи
            // семьи по поколениям от младшего поколения к старшему
            foreach (var famGroup in famGroups.OrderByDescending(g => g.Key))
            {
                // семьи внутри поколения, упорядоченные по Y
                var families = famGroup.OrderBy(o => o.Y).ToList();
                foreach (var family in families)
                {
                    var husband = family.Husband;
                    var wife = family.Wife;

                    if (wife != null && wife.Family == null)
                    {
                        wife.Y = family.Y - 1;                        
                        var indiToShift = _visualObjectsIndividual.Where(o => o.Level == wife.Level && o.Y >= wife.Y && o != wife).ToList();
                        var famToShift = _visualObjectsFamily.Where(o => o.Level == wife.Level - 1 && o.Y >= wife.Y).ToList();
                        famToShift.ForEach(o => o.Y++);
                        indiToShift.ForEach(o => o.Y++);
                    }

                    if (husband != null && husband.Family == null)
                    {
                        husband.Y = family.Y;
                        var indiToShift = _visualObjectsIndividual.Where(o => o.Level == husband.Level && o.Y >= husband.Y && o != husband).ToList();
                        var famToShift = _visualObjectsFamily.Where(o => o.Level == husband.Level - 1 && o.Y >= husband.Y).ToList();
                        famToShift.ForEach(o => o.Y++);
                        indiToShift.ForEach(o => o.Y++);
                    }
                }
            }

            Render();

            var graphOptimizer = new GraphOptimizer(VisualObjects, _renderer);
            graphOptimizer.OptimizeSimple();
        }



        private void CreateOrderLinks()
        {
            // определяем OrderLinks - семья отца выше, чем семья матери
            foreach (var family in _visualObjectsFamily)
            {
                if (family.Husband == null || family.Wife == null) continue;
                if (family.Husband.Family == null || family.Wife.Family == null) continue;

                AddOrderLink(new OrderLink(family.Husband.Family, family.Wife.Family));
            }

            // определяем OrderLinks - все семьи одного человека в порядке возникновения
            foreach (var individual in _visualObjectsIndividual)
            {
                var families = individual.ChildFamilies
                    .OrderBy(o => ParseGedcomDate(((Family)o.GedcomObject).MarriageDate))
                    .ToList();
                for (int i = 0; i < families.Count - 1; i++)
                {
                    var upper = individual.ChildFamilies[i];
                    var lower = individual.ChildFamilies[i + 1];
                    AddOrderLink(new OrderLink(upper, lower));
                }
            }

            // определяем OrderLinks - дети в семье по старшинству, 
            // затем дети, даты рождения которых неизвестны, по алфавиту
            foreach (var family in _visualObjectsFamily)
            {
                // дети у которых есть свои семьи
                var childsWithChildFam = family.Children.Where(c => c.ChildFamilies.Count > 0).ToList();
                var orderedChild = childsWithChildFam
                    .OrderBy(o => ParseGedcomDate(((Individual) o.GedcomObject).BirthDate))
                    .ToList();

                for (int i = 0; i < orderedChild.Count - 1; i++)
                {
                    // берем все пары детей по старшинству
                    // все семьи старшего (в паре) ребенка
                    var upper = orderedChild[i].ChildFamilies;
                    // все семьи младшего (в паре) ребенка
                    var lower = orderedChild[i + 1].ChildFamilies;
                    foreach (var up in upper)
                    {
                        foreach (var lo in lower)
                        {
                            AddOrderLink(new OrderLink(up, lo));
                        }
                    }
                }
            }

            // если одна семья располагается выше другой,
            // то семьи, созданные детьми из первой, располагаются выше,
            // чем семьи, созданные детьми из второй, если это не общая семья
            var orderLinksCopy = _orderLinks.ToList();
            foreach (var link in orderLinksCopy)
            {
                var upper = link.Upper.Children.SelectMany(o => o.ChildFamilies).ToList();
                var lower = link.Lower.Children.SelectMany(o => o.ChildFamilies).ToList();
                foreach (var up in upper)
                {
                    foreach (var lo in lower)
                    {
                        var orderLink = new OrderLink(up, lo);
                        AddOrderLink(orderLink);
                    }
                }
            }
        }

        /// <summary>
        /// Добавляет link в _orderLinks.
        /// Если это дубликат или добавление приводит к появлению циклических графов - ничего не делает.
        /// </summary>
        void AddOrderLink(OrderLink link)
        {
            if (link.Upper == link.Lower)
            {
                return;
            }
            if (_orderLinks.Any(o => o.Upper == link.Upper && o.Lower == link.Lower))
            {
                // такой линк уже есть
                return;
            }
            if (_orderLinks.Any(o => o.Upper == link.Lower && o.Lower == link.Upper))
            {
                // обратный линк уже есть
                return;
            }

            if (FindPath(link.Lower, link.Upper))
            {
                // обратный путь уже есть
                return;
            }
            _orderLinks.Add(link);
        }

        /// <summary>
        /// Ищет путь в OrderLinks от Upper к Lower
        /// </summary>
        bool FindPath(VisualObjectFamily source, VisualObjectFamily target)
        {
            var orderLinksCopy = _orderLinks.ToList();
            var bound = new List<VisualObjectFamily> { source };
            while (orderLinksCopy.Count > 0 && bound.Count > 0)
            {
                var boundLinks = orderLinksCopy.Where(o => bound.Contains(o.Upper)).ToList();
                boundLinks.ForEach(o => orderLinksCopy.Remove(o));
                bound = boundLinks.Select(o => o.Lower).ToList();
                if (bound.Contains(target))
                {
                    return true;
                }
            }
            return false;
        }


        /// <summary>
        /// Определяет уровни дерева, соответствующие поколениям (и семьям).
        /// Для циклических графов (например, если жениться на дочери своей жены), возможно зацикливание.
        /// Но у меня таких данных нет :)
        /// </summary>
        public void CalcLevel()
        {
            bool hasChanges;
            do
            {
                hasChanges = false;
                // "раздвигаем" перекрывающиеся уровни
                foreach (var o in VisualObjects)
                {
                    var clildLevel = o.Level + 1;
                    foreach (var child in o.ChildObjects.Where(child => child.Level < clildLevel))
                    {
                        child.Level = clildLevel;
                        hasChanges = true;
                    }
                }
                // "придвигаем" объекты, которые отстоят слишком далеко от детей
                foreach (var o in VisualObjects)
                {
                    if (o.ChildObjects.Any())
                    {
                        var level = o.ChildObjects.Max(c => c.Level) - 1;
                        if (level > o.Level)
                        {
                            o.Level = level;
                            hasChanges = true;
                        }
                    }
                }                
            } while (hasChanges);
        }


        public void Render()
        {
            foreach (var indi in _visualObjectsIndividual)
            {
                var hasParents = indi.ParentObjects.Count > 0;
                var shape = _renderer.Render((Individual)indi.GedcomObject, indi.X, indi.Y, hasParents);
                indi.Shape = shape;               
            }

            foreach (var family in _visualObjectsFamily)
            {
                var shape = _renderer.Render((Family)family.GedcomObject, family.X, family.Y);
                family.Shape = shape;
                foreach (var parent in family.ParentObjects)
                {
                    _renderer.Connect(shape, VisioRenderer.ConnectionPoint.Center,  parent.Shape, VisioRenderer.ConnectionPoint.Left);
                }

                foreach (var children in family.ChildObjects)
                {
                    _renderer.Connect(shape, VisioRenderer.ConnectionPoint.Center,  children.Shape, VisioRenderer.ConnectionPoint.Rigth);
                }
            }
        }


        private DateTime ParseGedcomDate(string date)
        {
            if (date == null)
            {
                return new DateTime();
            }
            date = date.Replace("ABT ", "");
            date = date.Replace("AFT ", "");
            date = date.Replace("BEF ", "");
            date = date.Replace("Приблизительно в ", "");
            date = date.Replace("№", "");
            if (date.Length <= 4)
            {
                // только год
                return new DateTime(int.Parse(date), 1, 1);
            }
            return DateTime.Parse(date);
        }

    }
}
