using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace AutoDimPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AxesColumnDimensions : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uIApplication = commandData.Application;
            UIDocument uIDocument = uIApplication.ActiveUIDocument;
            Document document = uIDocument.Document;

            //Geometry Options
            Options options = document.Application.Create.NewGeometryOptions();
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            options.View = document.ActiveView;

            //Retrieve COlumns from Revit model
            IList<Element> columns = new FilteredElementCollector(document, document.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType()
                .ToElements().ToList();

            //Retrieve grids
            IList<Grid> grids = new FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType()
                .Cast<Grid>().ToList();

            IList<Grid> horizontalGrids = grids.Where(x => Math.Abs(Math.Abs(GetGridLine(x, options).Direction.X)-1)<0.000001).ToList();
            IList<Grid> verticalGrids = grids.Where(x => Math.Abs(Math.Abs(GetGridLine(x, options).Direction.Y)-1)<0.00001).ToList();

            DimensionType dimensionType = new FilteredElementCollector(document).OfClass(typeof(DimensionType))
                .Cast<DimensionType>().FirstOrDefault();

            using (Transaction t = new Transaction(document, "Create Dimensions"))
            {
                t.Start();
                dimensionType.LookupParameter("Text Size").Set(0.00656168);

                CreateGridDimensions(horizontalGrids, document, options , dimensionType);
                CreateGridDimensions(verticalGrids, document, options , dimensionType);               

                foreach (Element column in columns)
                {
                    DimensionWithHorizontalGrid(column, horizontalGrids, document, options , dimensionType);
                    DimensionWithVerticalGrid(column, verticalGrids, document, options , dimensionType);
                }
                t.Commit();
            }
            return Result.Succeeded;
        }

        private Line GetGridLine(Grid grid , Options options)
        {
            Line line = null;
            foreach (GeometryObject obj in grid.get_Geometry(options))
            {
                if (obj is Line)
                {
                    line = obj as Line;
                }
            }
            return line;
        }

        private FaceArray GetColumnFcaes(Element element, Options options)
        {            
            GeometryElement elementGeo = element.get_Geometry(options);
            FaceArray faceArray = new FaceArray();
            foreach (GeometryObject item in elementGeo)
            {
                GeometryInstance elementGeoInstance = item as GeometryInstance;
                if (elementGeoInstance != null)
                {
                    GeometryElement elementGeoIn = elementGeoInstance.GetSymbolGeometry();
                    foreach (GeometryObject elementObj in elementGeoIn)
                    {
                        Solid elementSolid = elementObj as Solid;
                        if (elementSolid != null)
                        {
                            foreach (Face face in elementSolid.Faces)
                            {
                                faceArray.Append(face);
                            }
                        }
                    }
                }
            }
            return faceArray;
        }

        private void CreateGridDimensions(IList<Grid> grids, Document document, Options options , DimensionType dimensionType)
        {
            ReferenceArray referenceArray = new ReferenceArray();
            foreach (Grid grid in grids)
            {
                referenceArray.Append(GetGridLine(grid, options).Reference);
            }
            XYZ[] points = new XYZ[2];
            points[0] = (GetGridLine(grids[0], options)).Origin;
            points[1] = (GetGridLine(grids[1], options)).Origin;
            Dimension dim = document.Create.NewDimension(
                        options.View, Line.CreateBound(points[0], points[1]), referenceArray,dimensionType);
        }

        private void DimensionWithHorizontalGrid(Element column, IList<Grid> grids, Document document, Options option ,
            DimensionType dimensionType)
        {
            double columnY = (column.Location as LocationPoint).Point.Y;
            double minDistance = 100.0;
            Grid nearestGrid = null;
            foreach (Grid grid in grids)
            {
                double gridY = GetGridLine(grid, option).Origin.Y;
                double verticalDistance = Math.Abs(gridY - columnY);
                if (verticalDistance < minDistance)
                {
                    minDistance = verticalDistance;
                    nearestGrid = grid;
                }
            }
            DrawDimension(column, nearestGrid, document, option , dimensionType);
        }

        private void DimensionWithVerticalGrid(Element column, IList<Grid> grids, Document document, Options option
            , DimensionType dimensionType)
        {
            double columnY = (column.Location as LocationPoint).Point.X;
            double minDistance = 100.0;
            Grid nearestGrid = null;
            foreach (Grid grid in grids)
            {
                double gridY = GetGridLine(grid, option).Origin.X;
                double verticalDistance = Math.Abs(gridY - columnY);
                if (verticalDistance < minDistance)
                {
                    minDistance = verticalDistance;
                    nearestGrid = grid;
                }
            }
            DrawDimension(column, nearestGrid, document, option , dimensionType);
        }

        private void DrawDimension(Element column , Grid nearestGrid , Document document , Options option , DimensionType dimensionType)
        {
            Line line = GetGridLine(nearestGrid, option);
            //FaceArray faceArray = GetColumnFcaes(column, option);
            GeometryElement element = column.get_Geometry(option);
            var last = element.FirstOrDefault(e => e as Solid != null);
            Solid solid = last as Solid;
            if (solid != null)
            {
                FaceArray faceArray = solid.Faces;
                XYZ[] points = new XYZ[2];
                ReferenceArray referenceArray = new ReferenceArray();
                foreach (Face face in faceArray)
                {
                    XYZ faceNormal = face.ComputeNormal(face.GetBoundingBox().Max);
                    LocationPoint locationPoint = column.Location as LocationPoint;
                    faceNormal = new XYZ(faceNormal.X * Math.Cos(locationPoint.Rotation) - faceNormal.Y * Math.Sin(locationPoint.Rotation)
                                        , faceNormal.X * Math.Sin(locationPoint.Rotation) + faceNormal.Y * Math.Cos(locationPoint.Rotation), faceNormal.Z);
                    if (Math.Abs(faceNormal.DotProduct(line.Direction)) < 0.000001
                        && Math.Abs(faceNormal.DotProduct(XYZ.BasisZ)) < 0.000001)
                    {
                        referenceArray.Append(face.Reference);
                        referenceArray.Append(line.Reference);
                        double localX = (face as PlanarFace).Origin.X * Math.Cos(locationPoint.Rotation) - (face as PlanarFace).Origin.Y * Math.Sin(locationPoint.Rotation);
                        double localY = (face as PlanarFace).Origin.X * Math.Sin(locationPoint.Rotation) + (face as PlanarFace).Origin.Y * Math.Cos(locationPoint.Rotation);

                        points[0] = new XYZ((column.Location as LocationPoint).Point.X + localX + (1.64 * Math.Abs(line.Direction.X) * faceNormal.Y),
                            (column.Location as LocationPoint).Point.Y + localY - (1.64 * Math.Abs(line.Direction.Y) * faceNormal.X), 0);
                        points[1] = new XYZ(line.Project(points[0]).XYZPoint.X,
                            line.Project(points[0]).XYZPoint.Y, 0);
                        Dimension dim = document.Create.NewDimension(
                        option.View, Line.CreateBound(points[0], points[1]), referenceArray, dimensionType);
                        referenceArray.Clear();
                        //Only one dimension is needed to determine the position of column relative to axis 
                        //End Loop when the required face is reached and dimension is drawn
                        break;
                    }
                }
            }
        }
    }
}