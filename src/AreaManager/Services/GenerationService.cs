using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AreaManager.Models;

namespace AreaManager.Services
{
    /// <summary>
    /// Provides entry points for generating both the temporary areas table and the crown area
    /// usage summary.  This updated version applies the "Induction Bend" table style to
    /// all tables and delegates area calculations to WorkspaceAreaService.
    /// </summary>
    public static class GenerationService
    {
        public static void GenerateTemporaryAreasTable()
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var editor = document.Editor;
            var database = document.Database;

            var pairs = BlockAttributeService.GetUniqueAttributePairs(editor, database);
            if (pairs.Count == 0)
            {
                editor.WriteMessage("\nNo matching blocks found.");
                return;
            }

            var rows = BuildTempAreaRows(pairs);
            InsertTable(database, editor, rows);
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
                var identifier = pair.Item1?.Trim() ?? string.Empty;
                var enterText = pair.Item2?.Trim() ?? string.Empty;

                var description = MapIdentifierDescription(identifier);
                var row = BuildTempAreaRow(description, identifier, enterText);
                rows.Add(row);
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
                    var part1 = TextParsingService.ExtractLastNumber(parts[0]);
                    var part2 = TextParsingService.ExtractLastNumber(parts[1]);

                    if (IsNumeric(part1) && IsNumeric(part2))
                    {
                        var width = TextParsingService.ParseDoubleOrDefault(part1);
                        var length = TextParsingService.ParseDoubleOrDefault(part2);
                        var area = (width * length) / 10000.0;

                        row.Width = width.ToString("0.0");
                        row.Length = length.ToString("0.0");
                        row.AreaHa = area.ToString("0.000");
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
                if (string.IsNullOrWhiteSpace(description))
                {
                    // Default unknown identifiers to a generic bucket so they still contribute to totals
                    description = "";
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
            options.AddAllowedClass(typeof(Table), true);

            var result = editor.GetEntity(options);
            return result.Status == PromptStatus.OK ? result.ObjectId : ObjectId.Null;
        }

        private static void ApplyTableStyle(Transaction transaction, Database database, Editor editor, Table table, string styleName)
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
    }
}