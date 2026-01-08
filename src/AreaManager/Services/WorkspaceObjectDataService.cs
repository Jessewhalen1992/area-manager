using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;

// Explicit aliases help avoid confusion with AutoCAD DatabaseServices.Table
using OdTable = Autodesk.Gis.Map.ObjectData.Table;
using OdRecord = Autodesk.Gis.Map.ObjectData.Record;
using OdRecords = Autodesk.Gis.Map.ObjectData.Records;

namespace AreaManager.Services
{
    public static class WorkspaceObjectDataService
    {
        private const string WorkspaceTableName = "WORKSPACENUM";
        private const string WorkspaceFieldName = "WORKSPACENUM";

        private static readonly HashSet<string> ValidLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "P-TEMP_ACCESS ROAD",
            "P-TEMP_BANK STABILIZATION",
            "P-TEMP_BORROW PIT",
            "P-TEMP_CAMP SITE",
            "P-TEMP_LOG DECK",
            "P-TEMP_PUSHOUT",
            "P-TEMP_REMOTE SUMP",
            "P-TEMP_WORKSPACE",
            "P-TEMP-BLUE",
            "P-WORKAREA"
        };

        public static void AddObjectDataToShapes()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            var activityResult = editor.GetString(new PromptStringOptions("\nWhat type of Activity? (W) (LD) etc: ")
            {
                AllowSpaces = false
            });

            if (activityResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nOperation cancelled.");
                return;
            }

            var activityPrefix = activityResult.StringResult?.Trim();
            if (string.IsNullOrWhiteSpace(activityPrefix))
            {
                editor.WriteMessage("\nActivity type is required. Operation cancelled.");
                return;
            }

            var polylineOptions = new PromptEntityOptions("\nSelect a Polyline: ");
            polylineOptions.SetRejectMessage("\nOnly polylines are supported.");
            polylineOptions.AddAllowedClass(typeof(Polyline), true);
            polylineOptions.AddAllowedClass(typeof(Polyline2d), true);
            polylineOptions.AddAllowedClass(typeof(Polyline3d), true);

            var polylineResult = editor.GetEntity(polylineOptions);
            if (polylineResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nOperation cancelled.");
                return;
            }

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var polylineEntity = transaction.GetObject(polylineResult.ObjectId, OpenMode.ForRead) as Entity;
                var vertices = ExtractPolylineVertices(transaction, polylineEntity);
                if (vertices.Count == 0)
                {
                    editor.WriteMessage("\nNo vertices found on the selected polyline.");
                    return;
                }

                var mapApp = HostMapApplicationServices.Application;
                var odTables = mapApp.ActiveProject.ODTables;

                if (!odTables.IsTableDefined(WorkspaceTableName))
                {
                    editor.WriteMessage($"\nObject data table '{WorkspaceTableName}' was not found.");
                    return;
                }

                // ObjectData.Table is a DisposableWrapper; dispose it when you're done.
                using (OdTable odTable = odTables[WorkspaceTableName])
                {
                    var fieldIndex = FindFieldIndex(odTable.FieldDefinitions, WorkspaceFieldName);
                    if (fieldIndex < 0)
                    {
                        editor.WriteMessage($"\nField '{WorkspaceFieldName}' was not found in '{WorkspaceTableName}'.");
                        return;
                    }

                    var candidates = CollectCandidateShapes(transaction, database);
                    if (candidates.Count == 0)
                    {
                        editor.WriteMessage("\nNo closed shapes found on the expected layers.");
                        return;
                    }

                    // assignments: entity -> WORKSPACENUM value
                    var assignments = new Dictionary<ObjectId, string>();
                    var matchedIds = new HashSet<ObjectId>();

                    for (var index = 0; index < vertices.Count; index++)
                    {
                        var vertexNumber = index + 1;
                        var targetValue = $"{activityPrefix}{vertexNumber}";
                        var matches = new List<ObjectId>();

                        var testPoint = vertices[index];

                        foreach (var candidate in candidates)
                        {
                            if (!IsPointInsideClosedCurve(candidate.Curve, testPoint))
                            {
                                continue;
                            }

                            matches.Add(candidate.ObjectId);
                        }

                        if (matches.Count == 0)
                        {
                            editor.WriteMessage($"\nNo closed shape found for vertex {vertexNumber}. Operation cancelled.");
                            return;
                        }

                        foreach (var objectId in matches)
                        {
                            assignments[objectId] = targetValue;
                            matchedIds.Add(objectId);
                        }
                    }

                    // Apply OD values to matched shapes
                    foreach (var assignment in assignments)
                    {
                        var entity = transaction.GetObject(assignment.Key, OpenMode.ForRead) as Entity;
                        if (entity == null)
                        {
                            continue;
                        }

                        using (OdRecords records = odTable.GetObjectTableRecords(
                                   0, entity.ObjectId, MapOpenMode.OpenForWrite, false))
                        {
                            if (records == null || records.Count == 0)
                            {
                                // Cross-version pattern:
                                // Record.Create() + Table.InitRecord() + AddRecord()
                                // (works for Map 3D 2015/2025 per Autodesk Map .NET dev guide) :contentReference[oaicite:1]{index=1}
                                using (OdRecord record = OdRecord.Create())
                                {
                                    odTable.InitRecord(record);

                                    using (var mapValue = record[fieldIndex])
                                    {
                                        mapValue.Assign(assignment.Value);
                                    }

                                    // Use ObjectId overload for best compatibility
                                    odTable.AddRecord(record, entity.ObjectId);
                                }
                            }
                            else
                            {
                                var recordToUpdate = records[0];
                                using (var mapValue = recordToUpdate[fieldIndex])
                                {
                                    mapValue.Assign(assignment.Value);
                                }

                                // Persist the update (recommended by Map OD docs) :contentReference[oaicite:2]{index=2}
                                records.UpdateRecord(recordToUpdate);
                            }
                        }
                    }

                    // Clear OD values for shapes on valid layers that were NOT matched (only if they start with activityPrefix)
                    var prefixComparison = StringComparison.OrdinalIgnoreCase;
                    foreach (var candidate in candidates)
                    {
                        if (matchedIds.Contains(candidate.ObjectId))
                        {
                            continue;
                        }

                        var entity = transaction.GetObject(candidate.ObjectId, OpenMode.ForRead) as Entity;
                        if (entity == null)
                        {
                            continue;
                        }

                        using (OdRecords records = odTable.GetObjectTableRecords(
                                   0, entity.ObjectId, MapOpenMode.OpenForWrite, false))
                        {
                            if (records == null || records.Count == 0)
                            {
                                continue;
                            }

                            var recordToUpdate = records[0];
                            using (var mapValue = recordToUpdate[fieldIndex])
                            {
                                var currentValue = mapValue.StrValue ?? string.Empty;
                                if (currentValue.StartsWith(activityPrefix, prefixComparison))
                                {
                                    mapValue.Assign(string.Empty);
                                    records.UpdateRecord(recordToUpdate);
                                }
                            }
                        }
                    }

                    editor.WriteMessage($"\nUpdated {assignments.Count} shape(s) with {activityPrefix} values.");
                }

                transaction.Commit();
            }
        }

        private static List<Point3d> ExtractPolylineVertices(Transaction transaction, Entity polylineEntity)
        {
            var vertices = new List<Point3d>();

            if (polylineEntity is Polyline polyline)
            {
                for (var i = 0; i < polyline.NumberOfVertices; i++)
                {
                    vertices.Add(polyline.GetPoint3dAt(i));
                }

                return vertices;
            }

            if (polylineEntity is Polyline2d polyline2d)
            {
                foreach (ObjectId vertexId in polyline2d)
                {
                    var vertex = transaction.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                    if (vertex != null)
                    {
                        vertices.Add(vertex.Position);
                    }
                }

                return vertices;
            }

            if (polylineEntity is Polyline3d polyline3d)
            {
                foreach (ObjectId vertexId in polyline3d)
                {
                    var vertex = transaction.GetObject(vertexId, OpenMode.ForRead) as PolylineVertex3d;
                    if (vertex != null)
                    {
                        vertices.Add(vertex.Position);
                    }
                }
            }

            return vertices;
        }

        private static List<ShapeCandidate> CollectCandidateShapes(Transaction transaction, Database database)
        {
            var candidates = new List<ShapeCandidate>();

            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId objectId in modelSpace)
            {
                var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                if (entity == null)
                {
                    continue;
                }

                if (!ValidLayers.Contains(entity.Layer ?? string.Empty))
                {
                    continue;
                }

                if (entity is Curve curve && curve.Closed)
                {
                    candidates.Add(new ShapeCandidate(objectId, curve));
                }
            }

            return candidates;
        }

        /// <summary>
        /// Cross-version point-in-closed-curve test (AutoCAD 2015+).
        /// AutoCAD DatabaseServices.Curve does NOT have GetPointContainment().
        /// We implement a robust even/odd ray test using IntersectWith.
        /// </summary>
        private static bool IsPointInsideClosedCurve(Curve curve, Point3d point)
        {
            if (curve == null || !curve.Closed)
            {
                return false;
            }

            // Quick extents check (cheap reject)
            Extents3d extents;
            bool hasExtents = TryGetExtents(curve, out extents);
            if (hasExtents)
            {
                if (point.X < extents.MinPoint.X || point.X > extents.MaxPoint.X ||
                    point.Y < extents.MinPoint.Y || point.Y > extents.MaxPoint.Y)
                {
                    return false;
                }
            }

            // Treat "on boundary" as inside (closest point distance <= tolerance)
            try
            {
                var closest = curve.GetClosestPointTo(point, false);
                if (closest.DistanceTo(point) <= Tolerance.Global.EqualPoint)
                {
                    return true;
                }
            }
            catch
            {
                // ignore and continue to ray test
            }

            // Ray cast: count intersections with a long line starting at the point.
            // If odd => inside, if even => outside.
            try
            {
                double farX;
                double farY;

                if (hasExtents)
                {
                    var width = Math.Abs(extents.MaxPoint.X - extents.MinPoint.X);
                    if (width < 1.0) width = 1.0;

                    var height = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);
                    if (height < 1.0) height = 1.0;

                    farX = extents.MaxPoint.X + width + 10.0;

                    // Offset Y a bit to reduce the chance of the ray going exactly through a vertex
                    farY = point.Y + (height * 0.1234567);
                }
                else
                {
                    farX = point.X + 1000000.0;
                    farY = point.Y + 123.4567;
                }

                if (farX <= point.X)
                {
                    farX = point.X + 1000000.0;
                }

                using (var ray = new Line(point, new Point3d(farX, farY, point.Z)))
                {
                    var hits = new Point3dCollection();
                    curve.IntersectWith(ray, Intersect.OnBothOperands, hits, IntPtr.Zero, IntPtr.Zero);

                    if (hits == null || hits.Count == 0)
                    {
                        return false;
                    }

                    // Remove duplicates / near-start hits
                    double tol = Math.Max(Tolerance.Global.EqualPoint, 1e-7);

                    var unique = new List<Point3d>();
                    foreach (Point3d hit in hits)
                    {
                        // Ignore hits extremely close to the start point
                        if (hit.DistanceTo(point) <= tol)
                        {
                            continue;
                        }

                        // Only count hits "in front" of the start (primarily along +X direction)
                        if (hit.X < point.X - tol)
                        {
                            continue;
                        }

                        bool already = false;
                        foreach (var u in unique)
                        {
                            if (u.DistanceTo(hit) <= tol)
                            {
                                already = true;
                                break;
                            }
                        }

                        if (!already)
                        {
                            unique.Add(hit);
                        }
                    }

                    return (unique.Count % 2) == 1;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetExtents(Entity entity, out Extents3d extents)
        {
            extents = new Extents3d();
            if (entity == null)
            {
                return false;
            }

            try
            {
                extents = entity.GeometricExtents;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int FindFieldIndex(FieldDefinitions fieldDefinitions, string fieldName)
        {
            for (var index = 0; index < fieldDefinitions.Count; index++)
            {
                var field = fieldDefinitions[index];
                if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    return index;
                }
            }

            return -1;
        }

        private class ShapeCandidate
        {
            public ShapeCandidate(ObjectId objectId, Curve curve)
            {
                ObjectId = objectId;
                Curve = curve;
            }

            public ObjectId ObjectId { get; }
            public Curve Curve { get; }
        }
    }
}
