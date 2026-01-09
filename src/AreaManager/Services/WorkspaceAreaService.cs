using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map;
using AreaManager.Models;

// Alias the AutoCAD table class to avoid ambiguity with Autodesk.Gis.Map.ObjectData.Table.
using AcadTable = Autodesk.AutoCAD.DatabaseServices.Table;
using MapOpenMode = Autodesk.Gis.Map.Constants.OpenMode;
using ObjectDataTable = Autodesk.Gis.Map.ObjectData.Table;

namespace AreaManager.Services
{
    /// <summary>
    /// Reads workspace rows from a source AutoCAD table and computes aggregate area values
    /// for the crown area usage report.
    ///
    /// Notes:
    /// - The source table is expected to be the "Temporary Areas" table produced by this tool.
    /// - The "Within Existing Dispositions" column can contain:
    ///     * "No"  (treat within-dispo as 0)
    ///     * "Yes" (treat within-dispo as TOTAL area)
    ///     * a numeric value (treat within-dispo as that number)
    /// - The "Existing Cut Disturbance" column is interpreted as TOTAL existing disturbance
    ///   (existing cut + existing disposition). We derive cut-only by subtracting within-dispo.
    ///
    /// This file is written to compile and run on AutoCAD 2015 and 2025.
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

        /// <summary>
        /// Warn if the drawing contains multiple OD-boundary shapes with the same WORKSPACENUM value.
        /// This does not stop processing, but it is a strong indicator that area calculations may be wrong.
        ///
        /// IMPORTANT: Map ObjectData objects must be disposed. Failing to dispose these over repeated
        /// runs is a common cause of instability / access violations.
        /// </summary>
        private static void WarnOnDuplicateWorkspaceShapes(Editor editor, Transaction transaction, Database database)
        {
            var mapApp = HostMapApplicationServices.Application;
            var tables = mapApp.ActiveProject.ODTables;

            if (!tables.IsTableDefined(WorkspaceTableName))
            {
                return;
            }

            using (ObjectDataTable odTable = tables[WorkspaceTableName])
            {
                var fieldIndex = FindFieldIndex(odTable.FieldDefinitions, WorkspaceFieldName);
                if (fieldIndex < 0)
                {
                    return;
                }

                var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId objectId in modelSpace)
                {
                    var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                    if (entity == null)
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
                            using (var mapValue = record[fieldIndex])
                            {
                                var value = mapValue.StrValue;
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
                        }
                    }
                    catch
                    {
                        // MapException / interop failures can occur on proxy entities etc.
                        // Ignore and keep scanning.
                    }
                }

                var duplicates = seen
                    .Where(item => item.Value > 1)
                    .Select(item => item.Key)
                    .ToList();

                if (duplicates.Count > 0)
                {
                    editor.WriteMessage(
                        $"\nWarning: Duplicate WORKSPACENUM values found in object data: {string.Join(", ", duplicates)}.");
                }
            }
        }

        private static int FindFieldIndex(Autodesk.Gis.Map.ObjectData.FieldDefinitions fieldDefinitions, string fieldName)
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
        /// Reads each row from the selected workspace source table and produces a list of entries.
        /// Duplicate workspace IDs are resolved by choosing the entry with the larger total area.
        /// </summary>
        private static List<WorkspaceTableEntry> ReadWorkspaceEntriesFromTable(Editor editor, AcadTable table)
        {
            var headerRow = FindHeaderRow(table);
            var columnMap = MapWorkspaceColumns(table, headerRow);
            var startRow = headerRow >= 0 ? headerRow + 1 : 0;

            var unique = new Dictionary<string, WorkspaceTableEntry>(StringComparer.OrdinalIgnoreCase);

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
                    TotalHa = ReadCellNumber(table, rowIndex, columnMap.TotalColumn),

                    ExistingDispositionHa = ReadCellNumber(table, rowIndex, columnMap.ExistingDispositionColumn),
                    ExistingDispositionText = ReadCellText(table, rowIndex, columnMap.ExistingDispositionColumn),

                    ExistingCutDisturbanceHa = ReadCellNumber(table, rowIndex, columnMap.ExistingCutDisturbanceColumn),
                    NewCutDisturbanceHa = ReadCellNumber(table, rowIndex, columnMap.NewCutDisturbanceColumn)
                };

                entry.TotalHa = ResolveTotalArea(entry);

                if (!unique.TryGetValue(entry.WorkspaceId, out var existing))
                {
                    unique[entry.WorkspaceId] = entry;
                    continue;
                }

                if (EntriesMatch(existing, entry))
                {
                    continue;
                }

                var resolved = ResolveDuplicateEntry(existing, entry);
                unique[entry.WorkspaceId] = resolved;

                editor.WriteMessage(
                    $"\nWarning: Duplicate workspace '{entry.WorkspaceId}' with different sizes/areas detected. Using the larger area.");
            }

            return unique.Values.ToList();
        }

        /// <summary>
        /// Map column indices to their roles.
        /// If no header row exists, assume the default Temporary Areas table layout:
        ///   0 Description
        ///   1 Identifier
        ///   2 Width
        ///   3 Length
        ///   4 Area (ha)
        ///   5 Within Existing Dispositions (Yes/No/number)
        ///   6 Existing Cut Disturbance (ha)
        ///   7 New Cut Disturbance (ha)
        /// </summary>
        private static WorkspaceColumnMap MapWorkspaceColumns(AcadTable table, int headerRow)
        {
            var map = new WorkspaceColumnMap();

            if (headerRow < 0)
            {
                map.WorkspaceIdColumn = 1;
                map.WidthColumn = 2;
                map.LengthColumn = 3;
                map.AreaColumn = 4;
                map.ExistingDispositionColumn = 5;
                map.ExistingCutDisturbanceColumn = 6;
                map.NewCutDisturbanceColumn = 7;
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

                if (map.WorkspaceIdColumn < 0 &&
                    (normalized.Contains("WORKSPACE") || normalized.Contains("W#") || normalized == "ID" || normalized.Contains("DESCRIPTION")))
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

                if (map.ExistingDispositionColumn < 0 && (normalized.Contains("DISPOSITION") || normalized.Contains("WITHIN")))
                {
                    map.ExistingDispositionColumn = colIndex;
                    continue;
                }

                if (map.ExistingCutDisturbanceColumn < 0 && (normalized.Contains("EXISTING CUT") || normalized.Contains("CUT DIST")))
                {
                    map.ExistingCutDisturbanceColumn = colIndex;
                    continue;
                }

                if (map.NewCutDisturbanceColumn < 0 && normalized.Contains("NEW CUT"))
                {
                    map.NewCutDisturbanceColumn = colIndex;
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

            // Reasonable defaults if header mapping fails.
            if (map.WorkspaceIdColumn < 0)
            {
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

            // If the table only provides disturbance values, derive total from them.
            if (entry.ExistingCutDisturbanceHa.HasValue || entry.NewCutDisturbanceHa.HasValue)
            {
                return (entry.ExistingCutDisturbanceHa ?? 0.0) + (entry.NewCutDisturbanceHa ?? 0.0);
            }

            return null;
        }

        private static WorkspaceAreaRow BuildWorkspaceAreaRow(WorkspaceTableEntry entry)
        {
            var total = entry.TotalHa ?? 0.0;

            // Within (Existing Disposition): numeric overrides; otherwise interpret Yes/No.
            var existingDisposition = entry.ExistingDispositionHa ?? 0.0;
            if (!entry.ExistingDispositionHa.HasValue)
            {
                var dispText = (entry.ExistingDispositionText ?? string.Empty).Trim();
                if (dispText.Equals("YES", StringComparison.OrdinalIgnoreCase))
                {
                    existingDisposition = total;
                }
                else
                {
                    // "NO", blank, or anything else
                    existingDisposition = 0.0;
                }
            }

            // Existing Cut Disturbance: already includes disposition.
            var existingCutDisturbance = entry.ExistingCutDisturbanceHa;
            if (!existingCutDisturbance.HasValue)
            {
                // If missing, try to derive from total and new-cut.
                if (entry.NewCutDisturbanceHa.HasValue)
                {
                    existingCutDisturbance = Math.Max(0.0, total - entry.NewCutDisturbanceHa.Value);
                }
                else
                {
                    existingCutDisturbance = 0.0;
                }
            }

            var newCutDisturbance = entry.NewCutDisturbanceHa;
            if (!newCutDisturbance.HasValue)
            {
                newCutDisturbance = Math.Max(0.0, total - existingCutDisturbance.Value);
            }

            var existingCutOnly = Math.Max(0.0, existingCutDisturbance.Value - existingDisposition);

            return new WorkspaceAreaRow
            {
                WorkspaceId = entry.WorkspaceId,
                ExistingCutHa = existingCutOnly,
                ExistingDispositionHa = existingDisposition,
                TotalHa = total,
                ExistingCutDisturbanceHa = existingCutDisturbance.Value,
                NewCutDisturbanceHa = newCutDisturbance.Value
            };
        }

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
                && ValuesMatch(first.TotalHa, second.TotalHa)
                && ValuesMatch(first.AreaHa, second.AreaHa)
                && ValuesMatch(first.ExistingDispositionHa, second.ExistingDispositionHa)
                && string.Equals((first.ExistingDispositionText ?? string.Empty).Trim(), (second.ExistingDispositionText ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase)
                && ValuesMatch(first.ExistingCutDisturbanceHa, second.ExistingCutDisturbanceHa)
                && ValuesMatch(first.NewCutDisturbanceHa, second.NewCutDisturbanceHa);
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
            public int ExistingCutDisturbanceColumn { get; set; } = -1;
            public int NewCutDisturbanceColumn { get; set; } = -1;
            public int TotalColumn { get; set; } = -1;
        }

        private class WorkspaceTableEntry
        {
            public string WorkspaceId { get; set; }

            public double? Width { get; set; }
            public double? Length { get; set; }

            public double? AreaHa { get; set; }
            public double? TotalHa { get; set; }

            public double? ExistingDispositionHa { get; set; }
            public string ExistingDispositionText { get; set; }

            public double? ExistingCutDisturbanceHa { get; set; }
            public double? NewCutDisturbanceHa { get; set; }
        }
    }
}
