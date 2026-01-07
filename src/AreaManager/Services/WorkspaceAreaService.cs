using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
using AreaManager.Models;
using AcadTable = Autodesk.AutoCAD.DatabaseServices.Table;
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;
using ObjectDataTable = Autodesk.Gis.Map.ObjectData.Table;

namespace AreaManager.Services
{
    /// <summary>
    /// Service responsible for reading workspace rows from a source table and computing
    /// aggregate area values for the crown area usage report.  This implementation has
    /// been updated to better handle tables without explicit header rows and to
    /// interpret "Yes"/"No" values in the "within existing dispositions" column.
    /// </summary>
    public static class WorkspaceAreaService
    {
        private const string WorkspaceTableName = "WORKSPACENUM";
        private const string WorkspaceFieldName = "WORKSPACENUM";

        public static List<WorkspaceAreaRow> CalculateWorkspaceAreasFromTable(Editor editor, Database database, ObjectId tableId)
        {
            var results = new List<WorkspaceAreaRow>();

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = transaction.GetObject(tableId, OpenMode.ForRead) as AcadTable;
                if (table == null)
                {
                    return results;
                }

                WarnOnDuplicateWorkspaceShapes(editor, transaction, database);
                var entries = ReadWorkspaceEntriesFromTable(editor, table);
                results.AddRange(entries.Select(BuildWorkspaceAreaRow));

                transaction.Commit();
            }

            return results
                .OrderBy(row => ParseIdentifierPrefix(row.WorkspaceId))
                .ThenBy(row => ParseIdentifierNumber(row.WorkspaceId))
                .ThenBy(row => row.WorkspaceId, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void WarnOnDuplicateWorkspaceShapes(Editor editor, Transaction transaction, Database database)
        {
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var mapApp = HostMapApplicationServices.Application;
            var tables = mapApp.ActiveProject.ODTables;

            if (!tables.IsTableDefined(WorkspaceTableName))
            {
                return;
            }

            ObjectDataTable odTable = tables[WorkspaceTableName];

            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId objectId in modelSpace)
            {
                var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                if (entity == null)
                {
                    continue;
                }

                var records = odTable.GetObjectTableRecords(0, entity, MapOpenMode.OpenForRead, false);
                if (records.Count == 0)
                {
                    continue;
                }

                var record = records[0];
                var fieldIndex = FindFieldIndex(odTable.FieldDefinitions, WorkspaceFieldName);
                if (fieldIndex < 0)
                {
                    continue;
                }

                var value = record[fieldIndex].StrValue;
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                var key = value.Trim();
                if (seen.ContainsKey(key))
                {
                    seen[key] += 1;
                }
                else
                {
                    seen[key] = 1;
                }
            }

            var duplicates = seen.Where(item => item.Value > 1).Select(item => item.Key).ToList();
            if (duplicates.Count > 0)
            {
                editor.WriteMessage($"\nWarning: Duplicate WORKSPACENUM values found in object data: {string.Join(", ", duplicates)}.");
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

        /// <summary>
        /// Reads each row from the selected workspace source table and produces a list of
        /// WorkspaceTableEntry objects.  The logic here interprets the "within existing dispositions"
        /// column: when the cell contains "Yes" (case insensitive) and no numeric disposition value
        /// is present, the row's entire area is considered part of the existing disposition; when the
        /// cell contains "No" and no numeric cut value is present, the area is treated as existing
        /// cut.  Numeric values in either column override this behaviour.
        /// </summary>
        private static List<WorkspaceTableEntry> ReadWorkspaceEntriesFromTable(Editor editor, AcadTable table)
        {
            var entries = new List<WorkspaceTableEntry>();
            var headerRow = FindHeaderRow(table);
            var columnMap = MapWorkspaceColumns(table, headerRow);
            var startRow = headerRow >= 0 ? headerRow + 1 : 0;

            var duplicates = new Dictionary<string, WorkspaceTableEntry>(StringComparer.OrdinalIgnoreCase);

            for (var rowIndex = startRow; rowIndex < table.Rows.Count; rowIndex++)
            {
                var workspaceId = ReadCellText(table, rowIndex, columnMap.WorkspaceIdColumn);
                if (string.IsNullOrWhiteSpace(workspaceId))
                {
                    continue;
                }

                var entry = new WorkspaceTableEntry
                {
                    WorkspaceId = workspaceId.Trim(),
                    Width = ReadCellNumber(table, rowIndex, columnMap.WidthColumn),
                    Length = ReadCellNumber(table, rowIndex, columnMap.LengthColumn),
                    AreaHa = ReadCellNumber(table, rowIndex, columnMap.AreaColumn),
                    ExistingDispositionHa = ReadCellNumber(table, rowIndex, columnMap.ExistingDispositionColumn),
                    ExistingCutHa = ReadCellNumber(table, rowIndex, columnMap.ExistingCutColumn),
                    TotalHa = ReadCellNumber(table, rowIndex, columnMap.TotalColumn)
                };

                // Calculate the total area using width/length or area fields as appropriate
                entry.TotalHa = ResolveTotalArea(entry);

                // Interpret the 'within existing dispositions' column.  If the cell contains
                // 'Yes' or 'No' (case insensitive) and the existing disposition field isn't numeric,
                // use the entire area as the disposition or cut accordingly.
                if (columnMap.ExistingDispositionColumn >= 0)
                {
                    if (!entry.ExistingDispositionHa.HasValue)
                    {
                        var dispText = ReadCellText(table, rowIndex, columnMap.ExistingDispositionColumn);
                        if (!string.IsNullOrWhiteSpace(dispText))
                        {
                            var normalized = dispText.Trim().ToUpperInvariant();
                            if (normalized == "YES")
                            {
                                if (entry.TotalHa.HasValue)
                                {
                                    entry.ExistingDispositionHa = entry.TotalHa;
                                    entry.ExistingCutHa = 0.0;
                                }
                            }
                            else if (normalized == "NO")
                            {
                                if (!entry.ExistingCutHa.HasValue && entry.TotalHa.HasValue)
                                {
                                    entry.ExistingCutHa = entry.TotalHa;
                                    entry.ExistingDispositionHa = 0.0;
                                }
                            }
                        }
                    }
                }

                if (!duplicates.TryGetValue(entry.WorkspaceId, out var existing))
                {
                    duplicates[entry.WorkspaceId] = entry;
                    continue;
                }

                if (EntriesMatch(existing, entry))
                {
                    continue;
                }

                var resolved = ResolveDuplicateEntry(existing, entry);
                duplicates[entry.WorkspaceId] = resolved;

                editor.WriteMessage($"\nWarning: Duplicate workspace '{entry.WorkspaceId}' with different sizes/areas detected. Using the larger area.");
            }

            entries.AddRange(duplicates.Values);
            return entries;
        }

        /// <summary>
        /// Maps column indices to their respective roles.  When a header row isn't found,
        /// this method assumes a fixed column layout corresponding to the default
        /// temporary areas table: description, workspace ID, width, length, area,
        /// within dispositions, existing cut disturbance, new cut disturbance.
        /// </summary>
        private static WorkspaceColumnMap MapWorkspaceColumns(AcadTable table, int headerRow)
        {
            var map = new WorkspaceColumnMap();
            if (headerRow < 0)
            {
                // Default mapping for a table without a header.  Column 1 holds the workspace ID (identifier),
                // columns 2 and 3 hold width and length, column 4 is area, column 5 is the within dispositions
                // flag (Yes/No), column 6 is the existing cut area, and column 7 (if present) is new cut (ignored).
                map.WorkspaceIdColumn = 1;
                map.WidthColumn = 2;
                map.LengthColumn = 3;
                map.AreaColumn = 4;
                map.ExistingDispositionColumn = 5;
                map.ExistingCutColumn = 6;
                map.TotalColumn = -1;
                return map;
            }

            for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
            {
                var header = ReadCellText(table, headerRow, colIndex);
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                var normalized = header.Trim().ToUpperInvariant();
                if (map.WorkspaceIdColumn < 0 && (normalized.Contains("WORKSPACE") || normalized.Contains("W#") || normalized == "ID" || normalized.Contains("DESCRIPTION")))
                {
                    map.WorkspaceIdColumn = colIndex;
                    continue;
                }

                if (map.WidthColumn < 0 && normalized.Contains("WIDTH"))
                {
                    map.WidthColumn = colIndex;
                    continue;
                }

                if (map.LengthColumn < 0 && normalized.Contains("LENGTH"))
                {
                    map.LengthColumn = colIndex;
                    continue;
                }

                if (map.ExistingDispositionColumn < 0 && normalized.Contains("DISPOSITION"))
                {
                    map.ExistingDispositionColumn = colIndex;
                    continue;
                }

                if (map.ExistingCutColumn < 0 && normalized.Contains("EXISTING CUT"))
                {
                    map.ExistingCutColumn = colIndex;
                    continue;
                }

                if (map.TotalColumn < 0 && normalized.Contains("TOTAL"))
                {
                    map.TotalColumn = colIndex;
                    continue;
                }

                if (map.AreaColumn < 0 && normalized.Contains("AREA"))
                {
                    map.AreaColumn = colIndex;
                }
            }

            if (map.WorkspaceIdColumn < 0)
            {
                // default to identifier column if not mapped
                map.WorkspaceIdColumn = 1;
            }

            return map;
        }

        private static string ReadCellText(AcadTable table, int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex >= table.Rows.Count || columnIndex < 0 || columnIndex >= table.Columns.Count)
            {
                return string.Empty;
            }

            return table.Cells[rowIndex, columnIndex].TextString ?? string.Empty;
        }

        private static double? ReadCellNumber(AcadTable table, int rowIndex, int columnIndex)
        {
            var text = ReadCellText(table, rowIndex, columnIndex);
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var numberText = TextParsingService.ExtractLastNumber(text);
            if (string.IsNullOrWhiteSpace(numberText))
            {
                return null;
            }

            return TextParsingService.ParseDoubleOrDefault(numberText);
        }

        private static double? ResolveTotalArea(WorkspaceTableEntry entry)
        {
            if (entry.TotalHa.HasValue)
            {
                return entry.TotalHa.Value;
            }

            if (entry.AreaHa.HasValue)
            {
                return entry.AreaHa.Value;
            }

            if (entry.Width.HasValue && entry.Length.HasValue)
            {
                return (entry.Width.Value * entry.Length.Value) / 10000.0;
            }

            if (entry.ExistingCutHa.HasValue || entry.ExistingDispositionHa.HasValue)
            {
                return (entry.ExistingCutHa ?? 0.0) + (entry.ExistingDispositionHa ?? 0.0);
            }

            return null;
        }

        /// <summary>
        /// Constructs a WorkspaceAreaRow from a raw table entry by populating totals and disturbances.
        /// Existing disposition and cut values are passed through directly, and the total is computed
        /// if not explicitly provided.  New cut disturbance is always zero in this context.
        /// </summary>
        private static WorkspaceAreaRow BuildWorkspaceAreaRow(WorkspaceTableEntry entry)
        {
            var existingCut = entry.ExistingCutHa ?? 0.0;
            var existingDisposition = entry.ExistingDispositionHa ?? 0.0;
            var total = entry.TotalHa ?? (existingCut + existingDisposition);

            return new WorkspaceAreaRow
            {
                WorkspaceId = entry.WorkspaceId,
                ExistingCutHa = existingCut,
                ExistingDispositionHa = existingDisposition,
                TotalHa = total,
                ExistingCutDisturbanceHa = existingCut,
                NewCutDisturbanceHa = 0.0
            };
        }

        /// <summary>
        /// Attempts to locate a header row by scanning the first two rows of the table for cells
        /// containing typical header text (e.g., "WORKSPACE" or "DESCRIPTION").  Returns the index
        /// of the header row or -1 if none is found.
        /// </summary>
        private static int FindHeaderRow(AcadTable table)
        {
            var maxScan = Math.Min(table.Rows.Count, 2);
            for (var rowIndex = 0; rowIndex < maxScan; rowIndex++)
            {
                for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                {
                    var text = ReadCellText(table, rowIndex, colIndex);
                    if (IsHeaderText(text))
                    {
                        return rowIndex;
                    }
                }
            }
            return -1;
        }

        /// <summary>
        /// Determines whether a given cell value likely belongs to a header row based on keywords.
        /// </summary>
        private static bool IsHeaderText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            var upper = text.Trim().ToUpperInvariant();
            return upper.Contains("DESCRIPTION") || upper.Contains("WORKSPACE") || upper.Contains("W#") || upper.Contains("ID");
        }

        private static bool EntriesMatch(WorkspaceTableEntry first, WorkspaceTableEntry second)
        {
            return ValuesMatch(first.Width, second.Width)
                && ValuesMatch(first.Length, second.Length)
                && ValuesMatch(first.AreaHa, second.AreaHa)
                && ValuesMatch(first.TotalHa, second.TotalHa)
                && ValuesMatch(first.ExistingDispositionHa, second.ExistingDispositionHa)
                && ValuesMatch(first.ExistingCutHa, second.ExistingCutHa);
        }

        private static bool ValuesMatch(double? first, double? second)
        {
            if (!first.HasValue && !second.HasValue)
            {
                return true;
            }

            if (!first.HasValue || !second.HasValue)
            {
                return false;
            }

            return Math.Abs(first.Value - second.Value) < 0.0001;
        }

        private static WorkspaceTableEntry ResolveDuplicateEntry(WorkspaceTableEntry existing, WorkspaceTableEntry candidate)
        {
            var existingArea = existing.TotalHa ?? existing.AreaHa ?? 0.0;
            var candidateArea = candidate.TotalHa ?? candidate.AreaHa ?? 0.0;

            if (!existing.TotalHa.HasValue && existing.Width.HasValue && existing.Length.HasValue)
            {
                existingArea = (existing.Width.Value * existing.Length.Value) / 10000.0;
            }

            if (!candidate.TotalHa.HasValue && candidate.Width.HasValue && candidate.Length.HasValue)
            {
                candidateArea = (candidate.Width.Value * candidate.Length.Value) / 10000.0;
            }

            return candidateArea > existingArea ? candidate : existing;
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

        private class WorkspaceColumnMap
        {
            public int WorkspaceIdColumn { get; set; } = -1;
            public int WidthColumn { get; set; } = -1;
            public int LengthColumn { get; set; } = -1;
            public int AreaColumn { get; set; } = -1;
            public int ExistingDispositionColumn { get; set; } = -1;
            public int ExistingCutColumn { get; set; } = -1;
            public int TotalColumn { get; set; } = -1;
        }

        private class WorkspaceTableEntry
        {
            public string WorkspaceId { get; set; }
            public double? Width { get; set; }
            public double? Length { get; set; }
            public double? AreaHa { get; set; }
            public double? ExistingDispositionHa { get; set; }
            public double? ExistingCutHa { get; set; }
            public double? TotalHa { get; set; }
        }
    }
}