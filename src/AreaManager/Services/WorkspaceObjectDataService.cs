using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;

// Explicit aliases to avoid confusion with AutoCAD Table
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
            if (document == null)
            {
                return;
            }

            var editor = document.Editor;
            var database = document.Database;

            // IMPORTANT: called from WinForms button (modeless) -> lock doc
            using (document.LockDocument())
            {
                // 1) Ask for prefix
                var activityResult = editor.GetString(new PromptStringOptions("\nWhat type of Activity? (W) (LD) etc: ")
                {
                    AllowSpaces = false
                });

                if (activityResult.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("\nOperation cancelled.");
                    return;
                }

                var activityPrefix = (activityResult.StringResult ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(activityPrefix))
                {
                    editor.WriteMessage("\nActivity type is required. Operation cancelled.");
                    return;
                }

                // 2) Select the polyline providing the vertex order
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

                // ------------------------------------------------------------
                // READ PHASE: compute vertices, candidates, and assignments
                // ------------------------------------------------------------
                List<Point3d> vertices;
                List<ObjectId> candidateIds;
                Dictionary<ObjectId, string> assignments;

                using (var readTr = database.TransactionManager.StartTransaction())
                {
                    var polylineEntity = readTr.GetObject(polylineResult.ObjectId, OpenMode.ForRead) as Entity;
                    vertices = ExtractPolylineVertices(readTr, polylineEntity);

                    if (vertices.Count == 0)
                    {
                        editor.WriteMessage("\nNo vertices found on the selected polyline.");
                        return;
                    }

                    candidateIds = CollectCandidateShapeIds(readTr, database);
                    if (candidateIds.Count == 0)
                    {
                        editor.WriteMessage("\nNo closed shapes found on the expected layers.");
                        return;
                    }

                    // Build assignments based on vertex order
                    assignments = BuildAssignmentsFromVertices(readTr, editor, candidateIds, vertices, activityPrefix);
                    if (assignments == null)
                    {
                        // already wrote message
                        return;
                    }

                    readTr.Commit();
                }

                // ------------------------------------------------------------
                // WRITE PHASE: clear then rewrite OD values
                // ------------------------------------------------------------
                var mapApp = HostMapApplicationServices.Application;
                var odTables = mapApp.ActiveProject.ODTables;

                if (!odTables.IsTableDefined(WorkspaceTableName))
                {
                    editor.WriteMessage($"\nObject data table '{WorkspaceTableName}' was not found.");
                    return;
                }

                int clearedCount = 0;
                int updatedCount = 0;
                int createdCount = 0;
                int writeFailCount = 0;
                int clearFailCount = 0;

                using (var writeTr = database.TransactionManager.StartTransaction())
                using (OdTable odTable = odTables[WorkspaceTableName])
                {
                    var fieldIndex = FindFieldIndex(odTable.FieldDefinitions, WorkspaceFieldName);
                    if (fieldIndex < 0)
                    {
                        editor.WriteMessage($"\nField '{WorkspaceFieldName}' was not found in '{WorkspaceTableName}'.");
                        return;
                    }

                    // Pass 1: clear all OD values that start with the prefix on candidate shapes
                    foreach (var id in candidateIds)
                    {
                        var clearResult = TryClearPrefixValue(writeTr, odTable, id, fieldIndex, activityPrefix);
                        if (clearResult == ClearResult.Cleared)
                        {
                            clearedCount++;
                        }
                        else if (clearResult == ClearResult.Failed)
                        {
                            clearFailCount++;
                        }
                    }

                    // Pass 2: write the new assignments in vertex order (creates OD for new shapes)
                    foreach (var kvp in assignments)
                    {
                        var writeResult = TryWriteValue(writeTr, odTable, kvp.Key, fieldIndex, kvp.Value);
                        if (writeResult == WriteResult.Updated)
                        {
                            updatedCount++;
                        }
                        else if (writeResult == WriteResult.Created)
                        {
                            createdCount++;
                        }
                        else if (writeResult == WriteResult.Failed)
                        {
                            writeFailCount++;
                        }
                    }

                    writeTr.Commit();
                }

                editor.WriteMessage(
                    $"\nCleared {clearedCount} shape(s) with values starting '{activityPrefix}'. " +
                    $"Updated {updatedCount}, Created {createdCount} new OD record(s).");

                if (writeFailCount > 0 || clearFailCount > 0)
                {
                    editor.WriteMessage(
                        $"\nWarning: {clearFailCount} clear failures, {writeFailCount} write failures. " +
                        "Common causes: layer locked, entity on an xref, or object not writeable.");
                }
            }
        }

        private static Dictionary<ObjectId, string> BuildAssignmentsFromVertices(
            Transaction tr,
            Editor editor,
            List<ObjectId> candidateIds,
            List<Point3d> vertices,
            string activityPrefix)
        {
            var assignments = new Dictionary<ObjectId, string>();

            for (int i = 0; i < vertices.Count; i++)
            {
                var vertexNumber = i + 1;
                var targetValue = activityPrefix + vertexNumber;
                var point = vertices[i];

                var matches = new List<ObjectId>();

                for (int c = 0; c < candidateIds.Count; c++)
                {
                    var candidateId = candidateIds[c];
                    if (!IsPointInsideClosedCurve(tr, candidateId, point))
                    {
                        continue;
                    }

                    matches.Add(candidateId);
                }

                if (matches.Count == 0)
                {
                    editor.WriteMessage($"\nNo closed shape found for vertex {vertexNumber}. Operation cancelled.");
                    return null;
                }

                // If multiple shapes contain the point, all get the same value.
                // (If you want EXACTLY one shape per vertex, say so and I’ll switch
                // to “smallest-area containing shape”.)
                foreach (var id in matches)
                {
                    assignments[id] = targetValue;
                }
            }

            return assignments;
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

        private static List<ObjectId> CollectCandidateShapeIds(Transaction transaction, Database database)
        {
            var candidates = new List<ObjectId>();

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

                var curve = entity as Curve;
                if (curve != null && curve.Closed)
                {
                    candidates.Add(objectId);
                }
            }

            return candidates;
        }

        private static bool IsPointInsideClosedCurve(Transaction tr, ObjectId curveId, Point3d point)
        {
            try
            {
                var curve = tr.GetObject(curveId, OpenMode.ForRead) as Curve;
                if (curve == null || !curve.Closed)
                {
                    return false;
                }

                // Quick extents reject
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

                // On-boundary -> inside
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
                    // continue
                }

                // Ray cast
                double farX;
                double farY;

                if (hasExtents)
                {
                    var width = Math.Abs(extents.MaxPoint.X - extents.MinPoint.X);
                    if (width < 1.0) width = 1.0;

                    var height = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);
                    if (height < 1.0) height = 1.0;

                    farX = extents.MaxPoint.X + width + 10.0;
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

                    double tol = Math.Max(Tolerance.Global.EqualPoint, 1e-7);
                    var unique = new List<Point3d>();

                    foreach (Point3d hit in hits)
                    {
                        if (hit.DistanceTo(point) <= tol)
                        {
                            continue;
                        }

                        if (hit.X < point.X - tol)
                        {
                            continue;
                        }

                        bool already = false;
                        for (int i = 0; i < unique.Count; i++)
                        {
                            if (unique[i].DistanceTo(hit) <= tol)
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

        private enum ClearResult
        {
            NotCleared,
            Cleared,
            Failed
        }

        /// <summary>
        /// Clears WORKSPACENUM value if it starts with the given prefix.
        /// IMPORTANT: opens entity ForWrite and uses entity overload of GetObjectTableRecords.
        /// </summary>
        private static ClearResult TryClearPrefixValue(Transaction tr, OdTable odTable, ObjectId objectId, int fieldIndex, string prefix)
        {
            try
            {
                var entity = tr.GetObject(objectId, OpenMode.ForWrite) as Entity;
                if (entity == null)
                {
                    return ClearResult.Failed;
                }

                using (OdRecords records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForWrite, false))
                {
                    if (records == null || records.Count == 0)
                    {
                        return ClearResult.NotCleared;
                    }

                    bool clearedAny = false;

                    foreach (OdRecord rec in records)
                    {
                        using (var mapValue = rec[fieldIndex])
                        {
                            var current = mapValue.StrValue ?? string.Empty;
                            if (!string.IsNullOrEmpty(current) &&
                                current.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                            {
                                mapValue.Assign(string.Empty);
                                records.UpdateRecord(rec);
                                clearedAny = true;
                            }
                        }
                    }

                    return clearedAny ? ClearResult.Cleared : ClearResult.NotCleared;
                }
            }
            catch
            {
                return ClearResult.Failed;
            }
        }

        private enum WriteResult
        {
            None,
            Updated,
            Created,
            Failed
        }

        /// <summary>
        /// Writes WORKSPACENUM value to the object (creates a record if needed).
        /// IMPORTANT: opens entity ForWrite and uses entity overload of GetObjectTableRecords / AddRecord.
        /// </summary>
        private static WriteResult TryWriteValue(Transaction tr, OdTable odTable, ObjectId objectId, int fieldIndex, string value)
        {
            try
            {
                var entity = tr.GetObject(objectId, OpenMode.ForWrite) as Entity;
                if (entity == null)
                {
                    return WriteResult.Failed;
                }

                using (OdRecords records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForWrite, false))
                {
                    if (records == null || records.Count == 0)
                    {
                        using (OdRecord record = OdRecord.Create())
                        {
                            odTable.InitRecord(record);

                            using (var mapValue = record[fieldIndex])
                            {
                                mapValue.Assign(value);
                            }

                            // This is the key change: attach OD via the ENTITY (opened ForWrite)
                            odTable.AddRecord(record, entity);

                            return WriteResult.Created;
                        }
                    }

                    // Update first record (typical case: exactly one record)
                    var recToUpdate = records[0];
                    using (var mapValue = recToUpdate[fieldIndex])
                    {
                        mapValue.Assign(value);
                    }

                    records.UpdateRecord(recToUpdate);
                    return WriteResult.Updated;
                }
            }
            catch
            {
                return WriteResult.Failed;
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
    }
}
