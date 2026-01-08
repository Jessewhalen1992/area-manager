using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;

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
            "P-TEMP-BLUE"
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

                var odTable = odTables[WorkspaceTableName];
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

                var assignments = new Dictionary<ObjectId, string>();
                var matchedIds = new HashSet<ObjectId>();

                for (var index = 0; index < vertices.Count; index++)
                {
                    var vertexNumber = index + 1;
                    var targetValue = $"{activityPrefix}{vertexNumber}";
                    var matches = new List<ObjectId>();

                    foreach (var candidate in candidates)
                    {
                        if (!IsPointInsideCurve(candidate.Curve, vertices[index]))
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

                foreach (var assignment in assignments)
                {
                    var entity = transaction.GetObject(assignment.Key, OpenMode.ForRead) as Entity;
                    if (entity == null)
                    {
                        continue;
                    }

                    using (var records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForWrite, false))
                    {
                        if (records == null || records.Count == 0)
                        {
                            using (var record = odTable.CreateRecord())
                            {
                                record.Init(odTable.FieldDefinitions);
                                using (var mapValue = record[fieldIndex])
                                {
                                    mapValue.Assign(assignment.Value);
                                }

                                odTable.AddRecord(record, entity);
                            }

                            continue;
                        }

                        var recordToUpdate = records[0];
                        using (var mapValue = recordToUpdate[fieldIndex])
                        {
                            mapValue.Assign(assignment.Value);
                        }
                    }
                }

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

                    using (var records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForWrite, false))
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
                            }
                        }
                    }
                }

                editor.WriteMessage($"\nUpdated {assignments.Count} shape(s) with {activityPrefix} values.");
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

        private static bool IsPointInsideCurve(Curve curve, Point3d point)
        {
            if (curve == null)
            {
                return false;
            }

            try
            {
                var containment = curve.GetPointContainment(point, out _);
                return containment == PointContainment.Inside || containment == PointContainment.OnBoundary;
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
