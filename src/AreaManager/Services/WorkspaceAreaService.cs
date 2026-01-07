using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
using AreaManager.Models;
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;

namespace AreaManager.Services
{
    public static class WorkspaceAreaService
    {
        private const string WorkspaceTableName = "WORKSPACENUM";
        private const string WorkspaceFieldName = "WORKSPACENUM";

        public static List<WorkspaceAreaRow> CalculateWorkspaceAreas(Editor editor, Database database)
        {
            var results = new List<WorkspaceAreaRow>();

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var workspacePolylines = GetWorkspacePolylines(transaction, database);

                foreach (var workspace in workspacePolylines)
                {
                    var existingCut = CalculateLayerAreaWithinBoundary(transaction, database, workspace.Polyline, "P-EXISTINGCUT");
                    var existingDispo = CalculateLayerAreaWithinBoundary(transaction, database, workspace.Polyline, "P-EXISTINGDISPO");
                    var total = existingCut + existingDispo;

                    results.Add(new WorkspaceAreaRow
                    {
                        WorkspaceId = workspace.WorkspaceId,
                        ExistingCutHa = existingCut,
                        ExistingDispositionHa = existingDispo,
                        TotalHa = total,
                        ExistingCutDisturbanceHa = existingCut,
                        NewCutDisturbanceHa = 0.0
                    });
                }

                transaction.Commit();
            }

            return results;
        }

        private static List<(string WorkspaceId, Polyline Polyline)> GetWorkspacePolylines(Transaction transaction, Database database)
        {
            var results = new List<(string, Polyline)>();
            var mapApp = HostMapApplicationServices.Application;
            var tables = mapApp.ActiveProject.ODTables;

            if (!tables.IsTableDefined(WorkspaceTableName))
            {
                return results;
            }

            var table = tables[WorkspaceTableName];

            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId objectId in modelSpace)
            {
                var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                if (entity == null)
                {
                    continue;
                }

                var records = table.GetObjectTableRecords(0, entity, MapOpenMode.OpenForRead, false);
                if (records.Count == 0)
                {
                    continue;
                }

                var record = records[0];
                var fieldIndex = FindFieldIndex(table.FieldDefinitions, WorkspaceFieldName);
                if (fieldIndex < 0)
                {
                    continue;
                }

                var value = record[fieldIndex].StrValue;
                var polyline = entity as Polyline;
                if (polyline != null && polyline.Closed && !string.IsNullOrWhiteSpace(value))
                {
                    results.Add((value, polyline));
                }
            }

            return results;
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

        private static double CalculateLayerAreaWithinBoundary(Transaction transaction, Database database, Polyline boundary, string layerName)
        {
            var area = 0.0;
            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId objectId in modelSpace)
            {
                var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                if (entity == null || !entity.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var curve = entity as Curve;
                if (curve == null || !curve.Closed)
                {
                    continue;
                }

                area += CalculateIntersectionArea(boundary, curve);
            }

            return area / 10000.0;
        }

        private static double CalculateIntersectionArea(Polyline boundary, Curve target)
        {
            var regionBoundary = CreateRegion(boundary);
            var regionTarget = CreateRegion(target);

            if (regionBoundary == null || regionTarget == null)
            {
                return 0.0;
            }

            using (regionBoundary)
            using (regionTarget)
            {
                regionBoundary.BooleanOperation(BooleanOperationType.BoolIntersect, regionTarget);
                return regionBoundary.Area;
            }
        }

        private static Region CreateRegion(Curve curve)
        {
            if (curve == null || !curve.Closed)
            {
                return null;
            }

            var curves = new DBObjectCollection { curve };
            var regions = Region.CreateFromCurves(curves);

            return regions.Count > 0 ? regions[0] as Region : null;
        }
    }
}
