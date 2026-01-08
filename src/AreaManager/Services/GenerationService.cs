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
// Autodesk.Gis.Map.ObjectData.Table.  Throughout this file we refer to
// database tables using the AcadTable alias.  Without this alias the
// compiler cannot distinguish between the two different Table types when
// resolving overloaded methods and type checks.
using AcadTable = Autodesk.AutoCAD.DatabaseServices.Table;

// Alias the Map constants OpenMode enumeration to avoid referencing the
// undefined Constants class.  This is used when reading object data
// records.
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;

namespace AreaManager.Services
{
    /// <summary>
    /// Provides entry points for generating both the temporary areas table and the crown area
    /// usage summary.  This updated version applies the "Induction Bend" table style to
    /// all tables and delegates area calculations to WorkspaceAreaService.
    /// </summary>
    public static class GenerationService
    {
        // Object data table and field names used to locate workspace boundary shapes. These
        // constants mirror those defined in WorkspaceAreaService but are duplicated here to
        // avoid making that service's internals public.  They identify the name of the
        // object data table and field that stores workspace identifiers on boundary
        // shapes.  When computing existing cut and disposition areas, entities with this
        // object data are treated as workspace boundaries.
        private const string WorkspaceTableName = "WORKSPACENUM";
        private const string WorkspaceFieldName = "WORKSPACENUM";
        public static void GenerateTemporaryAreasTable()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            // Prompt the user to confirm whether the calculation should proceed.  Certain
            // portions of the workflow require the design to be in the "ground" state.  If
            // the user answers No or cancels the prompt, exit without doing anything.
            var keywordOptions = new PromptKeywordOptions("\nARE YOU IN GROUND? ");
            keywordOptions.Keywords.Add("Yes");
            keywordOptions.Keywords.Add("No");
            keywordOptions.Keywords.Default = "Yes";
            var keywordResult = editor.GetKeywords(keywordOptions);
            if (keywordResult.Status != PromptStatus.OK || keywordResult.StringResult.Equals("No", StringComparison.OrdinalIgnoreCase))
            {
                editor.WriteMessage("\nOperation cancelled.");
                return;
            }

            // Gather all unique block attribute pairs from the selected blocks.  These pairs
            // contain the workspace identifier and dimensional information.  If none are
            // found, notify the user and abort.
            var pairs = BlockAttributeService.GetUniqueAttributePairs(editor, database);
            if (pairs.Count == 0)
            {
                editor.WriteMessage("\nNo matching blocks found.");
                return;
            }

            // Build the initial list of temporary area rows from the attribute pairs.  This
            // computes width, length and area based on the ENTER_TEXT value but does not yet
            // calculate existing cut, new cut, or within disposition areas from object data.
            var tempRows = BuildTempAreaRows(pairs).ToList();

            // Use object data and underlying drawing geometry to compute existing cut
            // disturbance, new cut disturbance and within disposition values for each row.
            // If duplicates are detected or an error occurs, the helper returns false and
            // we exit without creating the table.  Otherwise, the rows are updated in
            // place with the computed values.
            if (!ComputeAreasForRows(editor, database, tempRows))
            {
                return;
            }

            // Finally insert the completed table into model space.  Passing the updated
            // rows ensures the computed areas are displayed.  The table style will be
            // applied and header rows suppressed in InsertTable.
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

                // Split the identifier field on whitespace to handle cases where multiple
                // workspace IDs are specified in a single block (e.g. "W69 W70").  The
                // original Excel macro duplicates the row for each ID and preserves the
                // ENTER_TEXT value.  We replicate that behaviour here by generating a
                // TempAreaRow for each individual identifier.
                var identifiers = identifierField.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (identifiers.Length == 0)
                {
                    // No valid identifier; still produce a row with an empty ID
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
                // When the text contains an 'x', attempt to parse width, length and possibly an
                // explicit area from the text.  Users often include both dimensions and a
                // computed area separated by a newline (e.g. "10.0x50.0\P=0.040").  In such
                // cases the last numeric value represents the area and may not match the
                // product of width and length.  If the computed area differs from the
                // explicit area beyond a small tolerance, treat the entry as irregular.
                var parts = enterText.Split('x');
                if (parts.Length >= 2)
                {
                    // Determine the width by taking the last numeric value from the portion
                    // before 'x'.  The original Excel macro used ExtractLastNumber to pick
                    // the final number from strings like "2-20.0" which should yield 20.0.
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
                        // The first number after 'x' is the length.  Additional numbers (if
                        // present) may include an explicit area; use the last number for the
                        // area comparison.
                        length = TextParsingService.ParseDoubleOrDefault(lengthNumbers[0]);
                        if (lengthNumbers.Length > 1)
                        {
                            explicitArea = TextParsingService.ParseDoubleOrDefault(lengthNumbers[lengthNumbers.Length - 1]);
                        }
                    }

                    // If both width and length are non‑zero, compute the area.  Otherwise, fall
                    // through to the irregular branch below.
                    if (width > 0.0 && length > 0.0)
                    {
                        var computedArea = (width * length) / 10000.0;
                        // If an explicit area exists and it does not match the computed area
                        // within a tolerance, treat the entry as irregular and use the
                        // explicit area.  Otherwise, treat it as a regular rectangular
                        // workspace.
                        if (explicitArea.HasValue && Math.Abs(computedArea - explicitArea.Value) > 1e-6)
                        {
                            row.Width = "IRREGULAR";
                            row.Length = "IRREGULAR";
                            row.AreaHa = explicitArea.Value.ToString("0.000");
                            row.ExistingCutDisturbance = row.AreaHa;
                            row.NewCutDisturbance = "0.000";
                            return row;
                        }
                        else
                        {
                            row.Width = width.ToString("0.0");
                            row.Length = length.ToString("0.0");
                            row.AreaHa = computedArea.ToString("0.000");
                            row.ExistingCutDisturbance = row.AreaHa;
                            row.NewCutDisturbance = "0.000";
                            return row;
                        }
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

            // When modifying the drawing database outside of a command context (e.g. from
            // a Windows Forms event), the document must be locked to prevent other threads
            // from obtaining conflicting locks.  Without this, attempting to open the
            // ModelSpace record for write may throw an eLockViolation exception.  Acquire
            // a document lock associated with the provided database before starting the
            // transaction.  The lock will be released automatically when disposed.
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(database);
            using (var docLock = doc?.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = TableService.CreateTempAreaTable(promptPoint.Value, rows);
                // Apply the desired table style.  Because the row-level RowType property does
                // not exist on some AutoCAD versions (e.g. 2015 and 2025), avoid trying to
                // assign it.  Instead, suppress the automatic title and header rows that
                // certain table styles insert by default.  Using the obsolete
                // IsTitleSuppressed/IsHeaderSuppressed properties will work across a range of
                // AutoCAD versions.  After applying the style, set both properties to true.
                ApplyTableStyle(transaction, database, editor, table, "Induction Bend");
                table.IsTitleSuppressed = true;
                table.IsHeaderSuppressed = true;
                // The "Induction Bend" style can still merge the first row (treating it as
                // a title) even after suppressing the title and header flags.  To avoid
                // breaking custom merges such as those used for "IRREGULAR" width/length
                // values, only unmerge the first row if it has been merged by the style.
                var firstRow = table.Rows.Count > 0 ? table.Rows[0] : null;
                if (firstRow?.IsMerged.HasValue == true && firstRow.IsMerged.Value)
                {
                    table.UnmergeCells(firstRow);
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
            // Prompt the user for an insertion point.  Bail out if they cancel.
            var promptPoint = editor.GetPoint("\nPick insertion point for Crown Area Usage table: ");
            if (promptPoint.Status != PromptStatus.OK)
            {
                return;
            }

            // Summarise the incoming workspace rows by their description (e.g. WORKSPACE, LOG DECK).
            // This collapses the list of individual workspaces into a single row per activity type
            // and appends a final summary row that aggregates the totals across all activities.
            var aggregatedRows = SummarizeWorkspaceRows(rows);

            // Acquire a document lock before modifying the drawing database.  This prevents
            // other threads or commands from modifying the document concurrently and
            // avoids eLockViolation errors when opening ModelSpace for write.
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(database);
            using (var docLock = doc?.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                // Build the table using the aggregated rows.  The table itself has no headers and
                // sizes are assigned in the TableService.  After creation, apply the requested
                // table style and force every row to be treated as a data row to avoid automatic
                // merging of the first row by the style (which designates the first row as a title).
                var table = TableService.CreateWorkspaceTotalsTable(promptPoint.Value, aggregatedRows);
                ApplyTableStyle(transaction, database, editor, table, "Induction Bend");

                // Explicitly suppress any title or header rows that the table style might
                // automatically generate.  The row-level RowType property is not available
                // in earlier AutoCAD versions (e.g. 2015) so we use the table-level
                // suppression flags instead.  Marking both as true removes the need to
                // assign a DataRow type to every row and prevents the first row from being
                // merged across columns.
                table.IsTitleSuppressed = true;
                table.IsHeaderSuppressed = true;
                // After suppressing title and header rows, unmerge the first row if the
                // style has merged it into a single cell.  Do not unmerge subsequent
                // rows, as some may legitimately contain merged cells (e.g. irregular
                // width/length in the temporary areas table).  This prevents the first
                // row from spanning all columns while preserving intentional merges.
                var first = table.Rows.Count > 0 ? table.Rows[0] : null;
                if (first?.IsMerged.HasValue == true && first.IsMerged.Value)
                {
                    table.UnmergeCells(first);
                }

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                modelSpace.AppendEntity(table);
                transaction.AddNewlyCreatedDBObject(table, true);
                transaction.Commit();
            }
        }

        /// <summary>
        /// Collapses a collection of individual workspace rows into a list of summary rows keyed
        /// by their activity description (e.g. WORKSPACE, LOG DECK, ACCESS ROAD).  A final row
        /// with an empty identifier is appended to represent the grand totals across all
        /// activities.  The mapping logic for determining a description from a workspace ID
        /// replicates that used when generating the temporary areas table.
        /// </summary>
        private static List<WorkspaceAreaRow> SummarizeWorkspaceRows(IEnumerable<WorkspaceAreaRow> rows)
        {
            var groups = new Dictionary<string, AggregatedValues>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                var description = MapIdentifierDescription(row.WorkspaceId);
                // If the description could not be determined, use the workspace ID itself as the
                // grouping key.  This prevents unknown identifiers from being collapsed into a
                // single empty description bucket which would produce a blank row in the final
                // table.  Using the workspace ID maintains a unique group per unknown entry.
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
                // Build a summary row for this activity
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

            // Append a grand total row only when there are multiple activity groups.
            // When there is only a single activity type, the overall summary is omitted.
            if (result.Count > 1)
            {
                // Compute grand totals for the final row.  Only the TotalHa (and implicitly TotalAc)
                // columns are populated; other fields remain zero so the summary displays blanks for
                // within, outside, and disturbance values.  An empty WorkspaceId is used as a flag
                // for TableService to treat this as the final totals row.
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

        /// <summary>
        /// Helper structure for accumulating numeric values while summarising workspace rows.
        /// </summary>
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
            var dictionary = (DBDictionary)transaction.GetObject(database.TableStyleDictionaryId, OpenMode.ForRead);
            if (!dictionary.Contains(styleName))
            {
                editor.WriteMessage($"\nWarning: Table style '{styleName}' was not found.");
                return;
            }

            table.TableStyle = dictionary.GetAt(styleName);
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

        /// <summary>
        /// Computes existing cut disturbance, new cut disturbance and within disposition areas
        /// for each temporary row.  Instead of performing expensive boolean region intersections,
        /// this method simply sums the area of each cut/disposition entity whose bounding
        /// extents fall entirely within the workspace boundary.  The area of the shapes is
        /// computed using a region created from the entity itself.  This approach assumes
        /// that users draw cut/disposition shapes wholly inside their corresponding
        /// workspace boundaries.  If no boundary is found for a workspace, all computed
        /// areas are set to zero and the new cut disturbance equals the full area of the
        /// workspace.  Returns false on duplicate detection or if an unexpected error
        /// occurs.
        /// </summary>
        private static bool ComputeAreasForRows(Editor editor, Database database, List<TempAreaRow> rows)
        {
            // Start a transaction to read entities and object data.  We do not make
            // modifications, so the transaction is read‑only and will be aborted at the
            // end.
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

                // Build the boundary map and detect duplicates
                foreach (ObjectId objectId in modelSpace)
                {
                    var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                    if (entity == null)
                    {
                        continue;
                    }
                    // Skip cut/disposition layers; we only care about boundary shapes here
                    var layerName = entity.Layer ?? string.Empty;
                    if (layerName.Equals("P-EXISTINGCUT", StringComparison.OrdinalIgnoreCase) ||
                        layerName.Equals("P-EXISTINGDISPO", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    var records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForRead, false);
                    if (records.Count == 0 || fieldIndex < 0)
                    {
                        continue;
                    }
                    var record = records[0];
                    var value = record[fieldIndex].StrValue;
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
                // Process each temporary row
                foreach (var row in rows)
                {
                    double withinDispositionHa = 0.0;
                    double existingCutHa = 0.0;
                    if (boundaryMap.TryGetValue(row.Identifier?.Trim() ?? string.Empty, out var boundaryId))
                    {
                        var boundaryEntity = transaction.GetObject(boundaryId, OpenMode.ForRead) as Entity;
                        // Compute the bounding extents of the boundary.  If the call fails, skip computations.
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
                            // Check if this entity's extents are completely inside the boundary's extents
                            Extents3d shapeExtents;
                            try
                            {
                                shapeExtents = ent.GeometricExtents;
                            }
                            catch
                            {
                                continue;
                            }
                            bool inside = shapeExtents.MinPoint.X >= boundaryExtents.MinPoint.X &&
                                         shapeExtents.MinPoint.Y >= boundaryExtents.MinPoint.Y &&
                                         shapeExtents.MaxPoint.X <= boundaryExtents.MaxPoint.X &&
                                         shapeExtents.MaxPoint.Y <= boundaryExtents.MaxPoint.Y;
                            if (!inside)
                            {
                                continue;
                            }
                            // Compute the area of the shape (in drawing units squared) using a region
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

        /// <summary>
        /// Attempts to create a Region object from a given entity.  Only planar and
        /// closed entities such as circles, ellipses and closed polylines can be
        /// converted into regions.  Returns null if the conversion fails.
        /// </summary>
        private static Region CreateRegionFromEntity(Entity entity)
        {
            if (entity == null)
            {
                return null;
            }
            // Collect curve geometries from the entity.  Some entities like polylines
            // can be exploded into a set of curves that define a closed loop.  If the
            // entity isn't suitable for region creation, simply return null.
            var curves = new DBObjectCollection();
            try
            {
                // Attempt to explode the entity into its constituent curves.  If
                // Explode throws or produces no curves, the entity cannot form a region.
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
                // Take ownership of the first region and dispose the rest to avoid
                // memory leaks.  Region.CreateFromCurves returns a DBObjectCollection of
                // newly allocated regions, all of which need disposal.
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
                // Dispose the intermediate curves regardless of success or failure
                foreach (DBObject obj in curves)
                {
                    obj.Dispose();
                }
            }
        }

        /// <summary>
        /// Computes the area of the intersection between two regions.  The boundary
        /// region is not modified; a clone is created to perform the boolean operation.
        /// Returns zero if the intersection is empty or if an error occurs.
        /// </summary>
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

        /// <summary>
        /// Finds the index of a field in a collection of field definitions by name.  Returns
        /// -1 when the field is not found.  A case‑insensitive comparison is used.
        /// </summary>
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