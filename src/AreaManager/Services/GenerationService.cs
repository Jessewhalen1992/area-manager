using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AreaManager.Models;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;

// Alias the AutoCAD DatabaseServices.Table class to avoid ambiguity with
// Autodesk.Gis.Map.ObjectData.Table.
using AcadTable = Autodesk.AutoCAD.DatabaseServices.Table;

// Alias the Map constants OpenMode enumeration.
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;

namespace AreaManager.Services
{
    /// <summary>
    /// Provides entry points for generating both the temporary areas table and the crown area
    /// usage summary.
    /// </summary>
    public static class GenerationService
    {
        private const string WorkspaceTableName = "WORKSPACENUM";
        private const string WorkspaceFieldName = "WORKSPACENUM";

        public static void GenerateTemporaryAreasTable()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            var keywordOptions = new PromptKeywordOptions("\nARE YOU IN GROUND? ");
            keywordOptions.Keywords.Add("Yes");
            keywordOptions.Keywords.Add("No");
            keywordOptions.Keywords.Default = "Yes";
            var keywordResult = editor.GetKeywords(keywordOptions);

            if (keywordResult.Status != PromptStatus.OK ||
                keywordResult.StringResult.Equals("No", StringComparison.OrdinalIgnoreCase))
            {
                editor.WriteMessage("\nOperation cancelled.");
                return;
            }

            var pairs = BlockAttributeService.GetUniqueAttributePairs(editor, database);
            if (pairs.Count == 0)
            {
                editor.WriteMessage("\nNo matching blocks found.");
                return;
            }

            var tempRows = BuildTempAreaRows(pairs).ToList();

            if (!ComputeAreasForRows(editor, database, tempRows))
            {
                return;
            }

            InsertTable(database, editor, tempRows);
        }

        public static void GenerateWorkspaceAreasTable()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            var tableSelection = PromptForWorkspaceTable(editor);
            if (tableSelection == ObjectId.Null)
            {
                return;
            }

            var workspaceRows = WorkspaceAreaService.CalculateWorkspaceAreasFromTable(editor, database, tableSelection);
            if (workspaceRows.Count == 0)
            {
                editor.WriteMessage("\nNo workspace rows found in the selected table.");
                return;
            }

            InsertWorkspaceTable(database, editor, workspaceRows);
        }

        private static List<TempAreaRow> BuildTempAreaRows(IEnumerable<Tuple<string, string>> pairs)
        {
            var rows = new List<TempAreaRow>();

            foreach (var pair in pairs)
            {
                var identifierField = pair.Item1?.Trim() ?? string.Empty;
                var enterText = pair.Item2?.Trim() ?? string.Empty;

                var identifiers = identifierField.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (identifiers.Length == 0)
                {
                    var description = MapIdentifierDescription(identifierField);
                    var row = BuildTempAreaRow(description, identifierField, enterText);
                    rows.Add(row);
                    continue;
                }

                foreach (var id in identifiers)
                {
                    var description = MapIdentifierDescription(id);
                    var row = BuildTempAreaRow(description, id, enterText);
                    rows.Add(row);
                }
            }

            return rows
                .OrderBy(row => ParseIdentifierPrefix(row.Identifier))
                .ThenBy(row => ParseIdentifierNumber(row.Identifier))
                .ThenBy(row => row.Identifier, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static TempAreaRow BuildTempAreaRow(string description, string identifier, string enterText)
        {
            var row = new TempAreaRow
            {
                Description = description,
                Identifier = identifier,
                WithinExistingDisposition = "No"
            };

            if (enterText.IndexOf("/P=", StringComparison.OrdinalIgnoreCase) >= 0 ||
                enterText.IndexOf("\\P=", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                row.Width = "IRREGULAR";
                row.Length = "IRREGULAR";
                row.AreaHa = ExtractAreaFromIrregular(enterText);
                row.NewCutDisturbance = "0.000";
                row.ExistingCutDisturbance = row.AreaHa;
                return row;
            }

            if (enterText.Contains("x"))
            {
                var parts = enterText.Split('x');
                if (parts.Length >= 2)
                {
                    var widthStr = TextParsingService.ExtractLastNumber(parts[0]);
                    var lengthNumbers = TextParsingService.ExtractAllNumbers(parts[1])?
                        .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    double width = 0.0;
                    double length = 0.0;
                    double? explicitArea = null;

                    if (!string.IsNullOrWhiteSpace(widthStr))
                    {
                        width = TextParsingService.ParseDoubleOrDefault(widthStr);
                    }

                    if (lengthNumbers != null && lengthNumbers.Length > 0)
                    {
                        length = TextParsingService.ParseDoubleOrDefault(lengthNumbers[0]);
                        if (lengthNumbers.Length > 1)
                        {
                            explicitArea = TextParsingService.ParseDoubleOrDefault(lengthNumbers[lengthNumbers.Length - 1]);
                        }
                    }

                    if (width > 0.0 && length > 0.0)
                    {
                        var computedArea = (width * length) / 10000.0;

                        if (explicitArea.HasValue && Math.Abs(computedArea - explicitArea.Value) > 1e-6)
                        {
                            row.Width = "IRREGULAR";
                            row.Length = "IRREGULAR";
                            row.AreaHa = explicitArea.Value.ToString("0.000");
                            row.ExistingCutDisturbance = row.AreaHa;
                            row.NewCutDisturbance = "0.000";
                            return row;
                        }

                        row.Width = width.ToString("0.0");
                        row.Length = length.ToString("0.0");
                        row.AreaHa = computedArea.ToString("0.000");
                        row.ExistingCutDisturbance = row.AreaHa;
                        row.NewCutDisturbance = "0.000";
                        return row;
                    }
                }
            }

            row.Width = "IRREGULAR";
            row.Length = "IRREGULAR";
            row.AreaHa = TextParsingService.ExtractAllNumbers(enterText);
            row.ExistingCutDisturbance = row.AreaHa;
            row.NewCutDisturbance = "0.000";
            return row;
        }

        private static string ExtractAreaFromIrregular(string enterText)
        {
            var index = enterText.IndexOf("/P=", StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                index = enterText.IndexOf("\\P=", StringComparison.OrdinalIgnoreCase);
            }

            var areaText = index >= 0 ? enterText.Substring(index + 3) : enterText;
            var number = TextParsingService.ExtractLastNumber(areaText);
            if (IsNumeric(number))
            {
                var area = TextParsingService.ParseDoubleOrDefault(number);
                return area.ToString("0.000");
            }

            return "N/A";
        }

        private static bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }

        private static string MapIdentifierDescription(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier))
            {
                return string.Empty;
            }

            var upper = identifier.ToUpperInvariant();
            if (upper.StartsWith("W") && upper.Length > 1 && char.IsDigit(upper[1]))
            {
                return "WORKSPACE";
            }

            if (upper.StartsWith("LD"))
            {
                return "LOG DECK";
            }

            if (upper.StartsWith("AR"))
            {
                return "ACCESS ROAD";
            }

            if (upper.StartsWith("BP"))
            {
                return "BORROW PIT";
            }

            if (upper.StartsWith("BS"))
            {
                return "BANK STABILIZATION";
            }

            return string.Empty;
        }

        private static void InsertTable(Database database, Editor editor, IEnumerable<TempAreaRow> rows)
        {
            var promptPoint = editor.GetPoint("\nPick insertion point for Temporary Areas table: ");
            if (promptPoint.Status != PromptStatus.OK)
            {
                return;
            }

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(database);
            using (var docLock = doc?.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = TableService.CreateTempAreaTable(promptPoint.Value, rows);

                // Ensure defaults, apply final style, suppress title/header, THEN generate layout.
                table.SetDatabaseDefaults(database);
                ApplyTableStyle(transaction, database, editor, table, "Induction Bend");
                table.IsTitleSuppressed = true;
                table.IsHeaderSuppressed = true;

                // Generate the layout under the final style.
                table.GenerateLayout();

                // If the style still merged the first row (title-like), undo it and regenerate.
                var firstRow = table.Rows.Count > 0 ? table.Rows[0] : null;
                if (firstRow?.IsMerged.HasValue == true && firstRow.IsMerged.Value)
                {
                    table.UnmergeCells(firstRow);
                    table.GenerateLayout();
                }

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                modelSpace.AppendEntity(table);
                transaction.AddNewlyCreatedDBObject(table, true);
                transaction.Commit();
            }
        }

        private static void InsertWorkspaceTable(Database database, Editor editor, IEnumerable<WorkspaceAreaRow> rows)
        {
            var promptPoint = editor.GetPoint("\nPick insertion point for Crown Area Usage table: ");
            if (promptPoint.Status != PromptStatus.OK)
            {
                return;
            }

            var aggregatedRows = SummarizeWorkspaceRows(rows);

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(database);
            using (var docLock = doc?.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = TableService.CreateWorkspaceTotalsTable(promptPoint.Value, aggregatedRows);

                table.SetDatabaseDefaults(database);
                ApplyTableStyle(transaction, database, editor, table, "Induction Bend");
                table.IsTitleSuppressed = true;
                table.IsHeaderSuppressed = true;

                table.GenerateLayout();

                var first = table.Rows.Count > 0 ? table.Rows[0] : null;
                if (first?.IsMerged.HasValue == true && first.IsMerged.Value)
                {
                    table.UnmergeCells(first);
                    table.GenerateLayout();
                }

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                modelSpace.AppendEntity(table);
                transaction.AddNewlyCreatedDBObject(table, true);
                transaction.Commit();
            }
        }

        private static List<WorkspaceAreaRow> SummarizeWorkspaceRows(IEnumerable<WorkspaceAreaRow> rows)
        {
            var groups = new Dictionary<string, AggregatedValues>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                var description = MapIdentifierDescription(row.WorkspaceId);
                if (string.IsNullOrWhiteSpace(description))
                {
                    description = row.WorkspaceId?.Trim() ?? string.Empty;
                }

                if (!groups.TryGetValue(description, out var agg))
                {
                    agg = new AggregatedValues();
                    groups[description] = agg;
                }

                agg.ExistingCut += row.ExistingCutHa;
                agg.ExistingDisposition += row.ExistingDispositionHa;
                agg.Total += row.TotalHa;
                agg.ExistingCutDisturbance += row.ExistingCutDisturbanceHa;
                agg.NewCutDisturbance += row.NewCutDisturbanceHa;
            }

            var result = new List<WorkspaceAreaRow>();
            foreach (var kvp in groups)
            {
                var description = kvp.Key;
                var agg = kvp.Value;

                result.Add(new WorkspaceAreaRow
                {
                    WorkspaceId = description,
                    ExistingCutHa = agg.ExistingCut,
                    ExistingDispositionHa = agg.ExistingDisposition,
                    TotalHa = agg.Total,
                    ExistingCutDisturbanceHa = agg.ExistingCutDisturbance,
                    NewCutDisturbanceHa = agg.NewCutDisturbance
                });
            }

            if (result.Count > 1)
            {
                var grandTotalHa = result.Sum(r => r.TotalHa);
                var grandExistingDisposition = result.Sum(r => r.ExistingDispositionHa);
                var grandExistingCut = result.Sum(r => r.ExistingCutHa);
                var grandExistingCutDist = result.Sum(r => r.ExistingCutDisturbanceHa);
                var grandNewCutDist = result.Sum(r => r.NewCutDisturbanceHa);

                result.Add(new WorkspaceAreaRow
                {
                    WorkspaceId = string.Empty,
                    ExistingCutHa = grandExistingCut,
                    ExistingDispositionHa = grandExistingDisposition,
                    TotalHa = grandTotalHa,
                    ExistingCutDisturbanceHa = grandExistingCutDist,
                    NewCutDisturbanceHa = grandNewCutDist
                });
            }

            return result;
        }

        private class AggregatedValues
        {
            public double ExistingCut;
            public double ExistingDisposition;
            public double Total;
            public double ExistingCutDisturbance;
            public double NewCutDisturbance;
        }

        private static ObjectId PromptForWorkspaceTable(Editor editor)
        {
            var options = new PromptEntityOptions("\nSelect workspace source table: ")
            {
                AllowNone = false
            };
            options.SetRejectMessage("\nOnly tables are allowed.");
            options.AddAllowedClass(typeof(AcadTable), true);

            var result = editor.GetEntity(options);
            return result.Status == PromptStatus.OK ? result.ObjectId : ObjectId.Null;
        }

        private static void ApplyTableStyle(Transaction transaction, Database database, Editor editor, AcadTable table, string styleName)
        {
            // Always end up with a valid style. If the named style isn't present,
            // fall back to the drawing's current table style.
            ObjectId styleId = database.Tablestyle;

            var dictionary = (DBDictionary)transaction.GetObject(database.TableStyleDictionaryId, OpenMode.ForRead);
            if (dictionary.Contains(styleName))
            {
                styleId = dictionary.GetAt(styleName);
            }
            else
            {
                editor.WriteMessage($"\nWarning: Table style '{styleName}' was not found. Using current drawing table style.");
            }

            table.TableStyle = styleId;
        }

        private static string ParseIdentifierPrefix(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier))
            {
                return string.Empty;
            }

            var trimmed = identifier.Trim();
            var index = 0;
            while (index < trimmed.Length && char.IsLetter(trimmed[index]))
            {
                index++;
            }

            return trimmed.Substring(0, index).ToUpperInvariant();
        }

        private static int ParseIdentifierNumber(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier))
            {
                return int.MaxValue;
            }

            var trimmed = identifier.Trim();
            var index = 0;
            while (index < trimmed.Length && char.IsLetter(trimmed[index]))
            {
                index++;
            }

            var numberStart = index;
            while (index < trimmed.Length && char.IsDigit(trimmed[index]))
            {
                index++;
            }

            if (numberStart == index)
            {
                return int.MaxValue;
            }

            if (int.TryParse(trimmed.Substring(numberStart, index - numberStart), out var number))
            {
                return number;
            }

            return int.MaxValue;
        }

        private static bool ComputeAreasForRows(Editor editor, Database database, List<TempAreaRow> rows)
        {
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var mapApp = HostMapApplicationServices.Application;
                var odTables = mapApp.ActiveProject.ODTables;
                if (!odTables.IsTableDefined(WorkspaceTableName))
                {
                    return true;
                }

                var odTable = odTables[WorkspaceTableName];

                var boundaryMap = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
                var duplicateIds = new List<ObjectId>();

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                var fieldIndex = FindFieldIndex(odTable.FieldDefinitions, WorkspaceFieldName);
                if (fieldIndex < 0)
                {
                    return true;
                }

                // Build boundary map and detect duplicates (DISPOSE OD WRAPPERS!)
                foreach (ObjectId objectId in modelSpace)
                {
                    var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                    if (entity == null)
                    {
                        continue;
                    }

                    var layerName = entity.Layer ?? string.Empty;
                    if (layerName.Equals("P-EXISTINGCUT", StringComparison.OrdinalIgnoreCase) ||
                        layerName.Equals("P-EXISTINGDISPO", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    try
                    {
                        using (var records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForRead, false))
                        {
                            if (records == null || records.Count == 0)
                            {
                                continue;
                            }

                            var record = records[0];
                            var mapValue = record[fieldIndex];
                            if (mapValue == null)
                            {
                                continue;
                            }

                            using (mapValue) // MapValue is a DisposableWrapper
                            {
                                var value = mapValue.StrValue;
                                if (string.IsNullOrWhiteSpace(value))
                                {
                                    continue;
                                }

                                var key = value.Trim();
                                if (boundaryMap.ContainsKey(key))
                                {
                                    duplicateIds.Add(objectId);
                                    duplicateIds.Add(boundaryMap[key]);
                                }
                                else
                                {
                                    boundaryMap[key] = objectId;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Some entities/OD configurations can throw; ignore and continue.
                        continue;
                    }
                }

                if (duplicateIds.Count > 0)
                {
                    var distinct = duplicateIds.Distinct().ToArray();
                    editor.SetImpliedSelection(distinct);
                    editor.WriteMessage("\nDuplicate workspace boundaries detected. Please resolve duplicates and retry.");
                    return false;
                }

                // Precompute cut/disposition entities
                var cutEntities = new List<Tuple<ObjectId, string>>();
                foreach (ObjectId objectId in modelSpace)
                {
                    var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    var ln = entity.Layer ?? string.Empty;
                    if (ln.Equals("P-EXISTINGCUT", StringComparison.OrdinalIgnoreCase) ||
                        ln.Equals("P-EXISTINGDISPO", StringComparison.OrdinalIgnoreCase))
                    {
                        cutEntities.Add(Tuple.Create(objectId, ln));
                    }
                }

                foreach (var row in rows)
                {
                    double withinDispositionHa = 0.0;
                    double existingCutHa = 0.0;

                    if (boundaryMap.TryGetValue(row.Identifier?.Trim() ?? string.Empty, out var boundaryId))
                    {
                        var boundaryEntity = transaction.GetObject(boundaryId, OpenMode.ForRead) as Entity;
                        if (boundaryEntity != null)
                        {
                            Extents3d boundaryExtents;
                            try
                            {
                                boundaryExtents = boundaryEntity.GeometricExtents;
                            }
                            catch
                            {
                                boundaryExtents = new Extents3d();
                            }

                            foreach (var tuple in cutEntities)
                            {
                                var entId = tuple.Item1;
                                var ln = tuple.Item2;

                                var ent = transaction.GetObject(entId, OpenMode.ForRead) as Entity;
                                if (ent == null) continue;

                                Extents3d shapeExtents;
                                try
                                {
                                    shapeExtents = ent.GeometricExtents;
                                }
                                catch
                                {
                                    continue;
                                }

                                bool inside =
                                    shapeExtents.MinPoint.X >= boundaryExtents.MinPoint.X &&
                                    shapeExtents.MinPoint.Y >= boundaryExtents.MinPoint.Y &&
                                    shapeExtents.MaxPoint.X <= boundaryExtents.MaxPoint.X &&
                                    shapeExtents.MaxPoint.Y <= boundaryExtents.MaxPoint.Y;

                                if (!inside)
                                {
                                    continue;
                                }

                                double areaSq = 0.0;
                                using (var shapeRegion = CreateRegionFromEntity(ent))
                                {
                                    if (shapeRegion != null)
                                    {
                                        areaSq = Math.Abs(shapeRegion.Area);
                                    }
                                }

                                if (areaSq <= 0.0) continue;

                                double areaHa = areaSq / 10000.0;
                                existingCutHa += areaHa;

                                if (ln.Equals("P-EXISTINGDISPO", StringComparison.OrdinalIgnoreCase))
                                {
                                    withinDispositionHa += areaHa;
                                }
                            }
                        }
                    }

                    var totalAreaHa = TextParsingService.ParseDoubleOrDefault(row.AreaHa);
                    var newCutHa = totalAreaHa - existingCutHa;
                    if (newCutHa < 0) newCutHa = 0.0;

                    row.WithinExistingDisposition = withinDispositionHa.ToString("0.000");
                    row.ExistingCutDisturbance = existingCutHa.ToString("0.000");
                    row.NewCutDisturbance = newCutHa.ToString("0.000");
                }
            }

            return true;
        }

        private static Region CreateRegionFromEntity(Entity entity)
        {
            if (entity == null)
            {
                return null;
            }

            var curves = new DBObjectCollection();
            try
            {
                entity.Explode(curves);
                if (curves.Count == 0)
                {
                    return null;
                }

                var regions = Region.CreateFromCurves(curves);
                if (regions == null || regions.Count == 0)
                {
                    return null;
                }

                Region firstRegion = regions[0] as Region;
                for (int i = 1; i < regions.Count; i++)
                {
                    var reg = regions[i] as Region;
                    reg?.Dispose();
                }

                return firstRegion;
            }
            catch
            {
                return null;
            }
            finally
            {
                foreach (DBObject obj in curves)
                {
                    obj.Dispose();
                }
            }
        }

        private static double ComputeIntersectionArea(Region boundary, Region shape)
        {
            if (boundary == null || shape == null)
            {
                return 0.0;
            }

            using (var boundaryClone = boundary.Clone() as Region)
            {
                if (boundaryClone == null)
                {
                    return 0.0;
                }

                try
                {
                    boundaryClone.BooleanOperation(BooleanOperationType.BoolIntersect, shape);
                    return Math.Abs(boundaryClone.Area);
                }
                catch
                {
                    return 0.0;
                }
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
