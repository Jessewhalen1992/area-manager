using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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

        public static List<WorkspaceAreaRow> CalculateWorkspaceAreasFromTable(Editor editor, Database database, ObjectId tableId)
        {
            var results = new List<WorkspaceAreaRow>();

            using (var transaction = database.TransactionManager.StartTransaction())
            {
                var table = transaction.GetObject(tableId, OpenMode.ForRead) as Table;
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

        private static List<WorkspaceTableEntry> ReadWorkspaceEntriesFromTable(Editor editor, Table table)
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

                entry.TotalHa = ResolveTotalArea(entry);

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

        private static int FindHeaderRow(Table table)
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

        private static WorkspaceColumnMap MapWorkspaceColumns(Table table, int headerRow)
        {
            var map = new WorkspaceColumnMap();
            if (headerRow < 0)
            {
                map.WorkspaceIdColumn = 0;
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
                map.WorkspaceIdColumn = 0;
            }

            return map;
        }

        private static string ReadCellText(Table table, int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex >= table.Rows.Count || columnIndex < 0 || columnIndex >= table.Columns.Count)
            {
                return string.Empty;
            }

            return table.Cells[rowIndex, columnIndex].TextString ?? string.Empty;
        }

        private static double? ReadCellNumber(Table table, int rowIndex, int columnIndex)
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
