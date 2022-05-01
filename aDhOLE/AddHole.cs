using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace aDhOLE
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().
                Where(x => x.Title.Contains("OB")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Succeeded;
            }

            var familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Succeeded;
            }


            var ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            var pipes = new FilteredElementCollector(ovDoc)
             .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            var view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Succeeded;
            }

            
            GreatHolesDuct(arDoc, familySymbol, ducts, view3D, out ReferenceIntersector refIntersector, out Transaction tr);
            

            CreatHolesPipe(arDoc, familySymbol, pipes, view3D, out ReferenceIntersector refIntersector1, out Transaction tr1);

            return Result.Succeeded;




        }

        private static void CreatHolesPipe(Document arDoc, FamilySymbol familySymbol, List<Pipe> pipes, View3D view3D, out ReferenceIntersector refIntersector1, out Transaction tr1)
        {
             refIntersector1 = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)),
               FindReferenceTarget.Element, view3D);

            Transaction tr0 = new Transaction(arDoc);
            tr0.Start("Расстановка отверстий");
            if (!familySymbol.IsActive)
            { familySymbol.Activate(); }

            tr0.Commit();

            tr1 = new Transaction(arDoc);
            tr1.Start("Расстановка отверстий");

            foreach (Pipe p in pipes)
            {
                
                var curve = (p.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersection = refIntersector1.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersection)
                {
                    double proximity = refer.Proximity;
                    double d1 = 0;
                    var reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);
                    var hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level,
                        StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    Parameter outD=p.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    width.Set(outD.AsDouble());
                    height.Set(outD.AsDouble());
                }
            }

            tr1.Commit();
        }

        private static void GreatHolesDuct(Document arDoc, FamilySymbol familySymbol, List<Duct> ducts, View3D view3D, out ReferenceIntersector refIntersector, out Transaction tr)
        {
            refIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)),
                            FindReferenceTarget.Element, view3D);

            Transaction tr0 = new Transaction(arDoc);
            tr0.Start("Расстановка отверстий");
            if (!familySymbol.IsActive)
            { familySymbol.Activate(); }

            tr0.Commit();

            tr = new Transaction(arDoc);
            tr.Start("Расстановка отверстий");

            foreach (Duct d in ducts)
            {
                var curve = (d.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersection = refIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersection)
                {
                    double proximity = refer.Proximity;
                    var reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);
                    var hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level,
                        StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(d.Diameter);
                    height.Set(d.Diameter);
                }
            }

            tr.Commit();
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }

    }
}
