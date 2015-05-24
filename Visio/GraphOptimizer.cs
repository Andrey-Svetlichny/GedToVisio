using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace GedToVisio.Visio
{
    /// <summary>
    /// Оптимизирует генеалогический граф (уменьшает длинну связей), перемещая объекты по оси Y.
    /// </summary>
    public class GraphOptimizer
    {
        private readonly VisioRenderer _renderer;
        private readonly List<VisualObject> _visualObjects;


        /// <summary>
        /// Перемещение единичного VisualObject.
        /// </summary>
        class Move
        {
            public VisualObject VisualObject { get; set; }
            public int OrigY { get; set; }
            public int NewY { get; set; }
            public Move(VisualObject visualObject, int shiftY)
            {
                VisualObject = visualObject;
                OrigY = visualObject.Y;
                NewY = visualObject.Y + shiftY;
            }

            public void Do()
            {
                VisualObject.Y = NewY;
            }
            public void Undo()
            {
                VisualObject.Y = OrigY;
            }
        }


        /// <summary>
        /// Вариант перемещения одного или нескольких объектов.
        /// </summary>
        class MoveVariaint
        {
            public List<Move> Moves { get; set; }

            public double Cost { get; set; }

            public MoveVariaint(List<Move> moves)
            {
                Moves = moves;
            }

            public void Do()
            {
                Moves.ForEach(o => o.Do());
            }
            public void Undo()
            {
                Moves.ForEach(o => o.Undo());
            }
        }

        public GraphOptimizer(List<VisualObject> visualObjects, VisioRenderer renderer)
        {
            _visualObjects = visualObjects;
            _renderer = renderer;
        }

        public void OptimizeSimple()
        {
            while (true)
            {
                // all available variants to move

                var groupsToMove = new List<List<VisualObject>>();
                // single objects
                groupsToMove.AddRange(_visualObjects.Select(o => new List<VisualObject> { o }));
                // objects with parents
                foreach (var visualObject in _visualObjects)
                {
                    var objWirhParents = visualObject.ParentObjects.ToList();
                    objWirhParents.Add(visualObject);

                    groupsToMove.Add(objWirhParents);
                }

                // objects with childs
                foreach (var visualObject in _visualObjects)
                {
                    var objWirhChilds = visualObject.ChildObjects.ToList();
                    objWirhChilds.Add(visualObject);

                    groupsToMove.Add(objWirhChilds);
                }

                var moveVariaints = new List<MoveVariaint>();
                // move up
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, +1)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, +2)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, +3)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, +4)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, +5)));
                // move down
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, -1)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, -2)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, -3)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, -4)));
                moveVariaints.AddRange(groupsToMove.Select(o => CalcMoveVariaint(o, -5)));



                var bestVariant = FindBestVariant(moveVariaints);
                if (bestVariant.Cost >= 0)
                    break;

                // apply
                bestVariant.Do();

                // visualize
                foreach (var move in bestVariant.Moves)
                {
                    _renderer.Move(move.VisualObject.Shape, move.VisualObject.X, move.VisualObject.Y);
                }
                Thread.Sleep(300);
            }
        }


        private MoveVariaint CalcMoveVariaint(List<VisualObject> visualObjectGroup, int shift)
        {
            var shiftReverse = shift > 0 ? -1 : 1;
            var moves = visualObjectGroup.Select(o => new Move(o, shift)).ToList();
            var visualObjectNotInGroup = _visualObjects.Except(visualObjectGroup).ToList();
            var overlappedObjects = moves
                .SelectMany(m => visualObjectNotInGroup.Where(o => o.X == m.VisualObject.X && o.Y == m.NewY))
                .ToList();
            overlappedObjects = shift > 0 ? overlappedObjects.OrderByDescending(o => o.Y).ToList() : overlappedObjects.OrderBy(o => o.Y).ToList();
            var overlappedMoves = new List<Move>();
            foreach (var oo in overlappedObjects)
            {
                var shiftY = shiftReverse;
                while (
                    // перекрывается с подвинутым объектом в группе
                    moves.Any(m => oo.X == m.VisualObject.X && oo.Y + shiftY == m.NewY)
                        // перекрывается с любым объектом не из группы
                    || visualObjectNotInGroup.Any(o => oo.X == o.X && oo.Y + shiftY == o.Y)
                    )
                {
                    shiftY += shiftReverse;
                }

                var move = new Move(oo, shiftY);
                overlappedMoves.Add(move);
                move.Do(); // сдвигаем на время, чтобы проверить перекрытие
            }
            overlappedMoves.ForEach(o => o.Undo()); // возвращаем
            moves.AddRange(overlappedMoves);
            var moveVariaint = new MoveVariaint(moves);
            return moveVariaint;
        }

        private MoveVariaint FindBestVariant(List<MoveVariaint> moveVariaints)
        {
            var origCost = Cost();
            foreach (var moveVariaint in moveVariaints)
            {
                moveVariaint.Do();
                moveVariaint.Cost = Cost() - origCost;
                moveVariaint.Undo();
            }

            var bestVariant = moveVariaints.OrderBy(o => o.Cost).First();
            return bestVariant;
        }

        double Cost()
        {
            // штраф за совпадающие позиции
            int overlappingCount = _visualObjects.Sum(visualObject => _visualObjects.Count(o => o.X == visualObject.X && o.Y == visualObject.Y && o != visualObject)) / 2;

            // длинна связей
            double len = _visualObjects.Sum(o => o.ChildObjects.Sum(c => Math.Sqrt((c.X - o.X) * (c.X - o.X) + (c.Y - o.Y) * (c.Y - o.Y))));
            return overlappingCount * 10 + len;
        }
    }
}
