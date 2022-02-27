using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationAddHole
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document; 
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault(); 
            if (ovDoc == null)			
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;		
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc) 
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();

            if (familySymbol == null)			
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстие\"");
                return Result.Cancelled;
            }

            List<Element> ducts = new FilteredElementCollector(ovDoc)  
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .Select(x => x as Element)
                .ToList();

            List<Element> pipes = new FilteredElementCollector(ovDoc) 
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .Select(x => x as Element)
                .ToList();

            List<Element> elementsOV = new List<Element>();
            elementsOV.AddRange(ducts);
            elementsOV.AddRange(pipes);

            View3D view3D = new FilteredElementCollector(arDoc) 	
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate) 			
                .FirstOrDefault();

            if (view3D == null)				
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction tr = new Transaction(arDoc);
            tr.Start("Расстановка отверстий");

            if (!familySymbol.IsActive)
            { familySymbol.Activate(); }       
            tr.Commit();


            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");
            foreach (Element el in elementsOV)
            {
                Pipe pipe = el as Pipe;
                Duct duct = el as Duct;

                Line curve = pipe == null ? (duct.Location as LocationCurve).Curve as Line : (pipe.Location as LocationCurve).Curve as Line;	
                XYZ point = curve.GetEndPoint(0);	
                XYZ direction = curve.Direction;	


                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) 
                    .Where(x => x.Proximity <= curve.Length)  
                    .Distinct(new ReferenceWithContextElementEqualityComparer())  
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;  
                    Reference reference = refer.GetReference();   

                    Element host = arDoc.GetElement(reference.ElementId);  
                    Level level = arDoc.GetElement(host.LevelId) as Level;	

                    XYZ pointHole = point + (direction * proximity);  

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, host, level, StructuralType.NonStructural);	
                    Parameter width = hole.LookupParameter("Ширина_1");
                    Parameter height = hole.LookupParameter("Высота_1");
                    if (duct != null)
                    {
                        width.Set(duct.Diameter);   
                        height.Set(duct.Diameter);  

                    }
                    else if (pipe != null)
                    {
                        width.Set(pipe.Diameter);       
                        height.Set(pipe.Diameter);      
                    }
                }
            }
            transaction.Commit();

            return Result.Succeeded;
        }
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
