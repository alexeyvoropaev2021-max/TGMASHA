using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

[assembly: CommandClass(typeof(TL_Plugin.TL_Optimized))]

namespace TL_Plugin
{
    public class TL_Optimized
    {
        private const string SingleLineBlock = "Блок однолинейной схемы";
        private const string TerminalsBlock = "Клеммники";
        private const string CircuitBreakersBlock = "TL1 - Изображения автоматов";
        private const string SpecEOMAddress = "TL1 - Адрес для макроса";

        static TL_Optimized()
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private static string GetExcelFilePath(Document doc, string layerName)
        {
            var db = doc.Database;
            using var tr = db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in ms)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is DBText text && text.Layer == layerName)
                    return text.TextString;
            }
            return string.Empty;
        }

        private static List<ExcelRow> ReadExcelSheet(string filePath, string sheetName)
        {
            var rows = new List<ExcelRow>();
            if (!File.Exists(filePath)) return rows;

            using var pkg = new ExcelPackage(new FileInfo(filePath));
            var ws = pkg.Workbook.Worksheets[sheetName];
            if (ws == null) return rows;

            int r = 4;
            while (true)
            {
                var b = ws.Cells[$"B{r}"].Text;
                var ab = ws.Cells[$"AB{r}"].Text;
                var ac = ws.Cells[$"AC{r}"].Text;

                if (string.IsNullOrEmpty(b) && string.IsNullOrEmpty(ab) && string.IsNullOrEmpty(ac))
                    break;

                rows.Add(new ExcelRow(r, ws));
                r++;
            }
            return rows;
        }

        private static double GetBlockWidth(string widthValue, double defaultWidth = 50.0)
        {
            if (string.IsNullOrEmpty(widthValue))
                return defaultWidth;

            if (double.TryParse(widthValue.Trim(), out var width) && width > 0)
                return width;

            return defaultWidth;
        }

        private static int GetBlockCountFromAF(string colAF)
        {
            if (string.IsNullOrEmpty(colAF))
                return 1;

            if (double.TryParse(colAF, out var value))
            {
                if (value <= 2)
                    return 1;
                else if (value <= 4)
                    return 3;
                else
                    return (int)Math.Ceiling(value / 2.0);
            }

            return 1;
        }

        private static double CalculateInterval(ExcelRow row, int blockType)
        {
            switch (blockType)
            {
                case 0:
                    return 144.36;
                case 1:
                    return GetTerminalInterval(row.ColU);
                case 2:
                    double multiplier = 1.0;
                    if (!string.IsNullOrEmpty(row.ColAF) &&
                        double.TryParse(row.ColAF, out double afValue) && afValue > 0)
                    {
                        multiplier = afValue;
                    }
                    return 180.0 * multiplier;
                default:
                    return 0;
            }
        }

        private static double GetTerminalInterval(string colU)
        {
            if (string.IsNullOrEmpty(colU))
                return 42.5;

            if (!double.TryParse(colU, out var val))
                return 42.5;

            return val switch
            {
                1.5 => 42.5,
                2.5 => 52.5,
                4.0 => 62.5,
                6.0 => 75.0,
                10.0 => 100.0,
                _ => 42.5
            };
        }

        // ИСПРАВЛЕННЫЙ метод: ищет ТОЛЬКО свойство "Vis" для всех блоков
        private static void InsertBlockWithVisualization(
            BlockTableRecord space,
            Transaction tr,
            BlockTableRecord def,
            ExcelRow row,
            Point3d pt,
            string visualization,
            int blockType)
        {
            var br = new BlockReference(pt, def.ObjectId);
            space.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);

            // Заполнение атрибутов
            foreach (ObjectId id in def)
            {
                var obj = tr.GetObject(id, OpenMode.ForRead);
                if (obj is not AttributeDefinition ad || ad.Constant)
                    continue;

                var ar = new AttributeReference();
                ar.SetAttributeFromBlock(ad, br.BlockTransform);

                switch (blockType)
                {
                    case 0:
                        FillSingleLineAttributes(ar, ad, row);
                        break;
                    case 1:
                        FillTerminalAttributes(ar, ad, row);
                        break;
                    case 2:
                        FillCircuitBreakerAttributes(ar, ad, row);
                        break;
                }

                br.AttributeCollection.AppendAttribute(ar);
                tr.AddNewlyCreatedDBObject(ar, true);
            }

            // Установка Dynamic Block свойства "Vis"
            if (br.IsDynamicBlock && !string.IsNullOrEmpty(visualization) && visualization != "#N/A" && visualization != "0")
            {
                try
                {
                    var props = br.DynamicBlockReferencePropertyCollection;

                    foreach (DynamicBlockReferenceProperty prop in props)
                    {
                        // ВСЕ БЛОКИ ИСПОЛЬЗУЮТ "Vis"
                        if (prop.PropertyName.Equals("Vis", StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                prop.Value = visualization;
                                br.RecordGraphicsModified(true);
                            }
                            catch { }
                            break;
                        }
                    }
                }
                catch { }
            }
        }

        private static void FillSingleLineAttributes(AttributeReference ar, AttributeDefinition ad, ExcelRow row)
        {
            switch (ad.Tag.ToUpper())
            {
                case "ПОЗ-ОБОЗНАЧЕНИЕ": ar.TextString = row.ColAC; break;
                case "ОПИСАНИЕ": ar.TextString = row.ColB; break;
                case "ФАЗА": ar.TextString = row.GetPhaseString(); break;
                case "КЛЕММА1": ar.TextString = row.ColAD; break;
                case "КЛЕММА2": ar.TextString = row.ColAE; break;
                case "АКТУАТОР": ar.TextString = $"[{row.ColR}] {row.ColS}"; break;
                case "ГРУППА": ar.TextString = row.ColA; break;
                case "IРАС": ar.TextString = row.ColN; break;
                case "PУСТ": ar.TextString = row.ColM; break;
                case "КАБЕЛЬ": ar.TextString = $"{row.ColT}{row.ColU}"; break;
                case "ПРОЛОЖЕНИЕ": ar.TextString = row.GetProlozhenie(); break;
                case "АВТОМАТ":
                case "ПОЗ_ОБОЗН1":
                    ar.TextString = FormatCircuitBreakerInfo(row);
                    break;
                case "НОМЕР": ar.TextString = row.ColH + row.ColJ; break;
                case "ТИП": ar.TextString = row.ColK; break;
                case "ПОЯСНЕНИЕ": ar.TextString = row.ColB; break;
                case "АРТИКУЛ": ar.TextString = row.ColAA; break;
                case "VENDOR": ar.TextString = row.ColZ; break;
            }
        }

        private static void FillTerminalAttributes(AttributeReference ar, AttributeDefinition ad, ExcelRow row)
        {
            switch (ad.Tag.ToUpper())
            {
                case "ПОЗ-ОБОЗНАЧЕНИЕ": ar.TextString = row.ColAC; break;
                case "ОПИСАНИЕ": ar.TextString = row.ColB; break;
                case "ФАЗА": ar.TextString = row.GetPhaseString(); break;
                case "КЛЕММА1": ar.TextString = row.ColAD; break;
                case "КЛЕММА2": ar.TextString = row.ColAE; break;
                case "АКТУАТОР": ar.TextString = $"{row.ColR}{row.ColS}"; break;
                case "ГРУППА": ar.TextString = row.ColA; break;
                case "IРАС": ar.TextString = row.ColN; break;
                case "PУСТ": ar.TextString = row.ColM; break;
                case "КАБЕЛЬ": ar.TextString = $"{row.ColT}{row.ColU}"; break;
                case "ПРОЛОЖЕНИЕ": ar.TextString = row.GetProlozhenie(); break;
                case "НОМЕР": ar.TextString = row.ColH + row.ColJ; break;
                case "ТИП": ar.TextString = $"{row.ColK} {row.ColI}"; break;
                case "ПОЯСНЕНИЕ": ar.TextString = row.ColB; break;
                case "АРТИКУЛ": ar.TextString = row.ColAA; break;
                case "VENDOR": ar.TextString = row.ColZ; break;
            }
        }

        private static void FillCircuitBreakerAttributes(AttributeReference ar, AttributeDefinition ad, ExcelRow row)
        {
            switch (ad.Tag.ToUpper())
            {
                case "НОМЕР": ar.TextString = row.ColH + row.ColJ; break;
                case "ТИП": ar.TextString = $"{row.ColK} {row.ColI}"; break;
                case "ПОЯСНЕНИЕ": ar.TextString = row.ColB; break;
                case "АРТИКУЛ": ar.TextString = row.ColAA; break;
                case "VENDOR": ar.TextString = row.ColZ; break;
            }
        }

        private static string FormatCircuitBreakerInfo(ExcelRow row)
        {
            var breaker = row.ColH == "QF+N" ? "QF" : row.ColH;
            return $"{breaker}{row.ColJ}\n{row.ColK} {row.ColI}\n{row.ColZ}\n{row.ColAA}";
        }

        [CommandMethod("TL_Однолинейная_схема")]
        public static void DrawSingleLineDiagram()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var path = GetExcelFilePath(doc, SpecEOMAddress);

            if (string.IsNullOrEmpty(path))
            {
                ed.WriteMessage("\nФайл не найден");
                return;
            }

            var rows = ReadExcelSheet(path, "Расчет нагрузок");
            if (rows.Count == 0) return;

            var ppr = ed.GetPoint("\nУкажите точку вставки: ");
            if (ppr.Status != PromptStatus.OK) return;

            double currentX = ppr.Value.X;
            double y = ppr.Value.Y;
            double z = ppr.Value.Z;

            using var tr = db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            if (!bt.Has(SingleLineBlock))
            {
                ed.WriteMessage($"\nОшибка: блок '{SingleLineBlock}' не найден");
                return;
            }

            var def = (BlockTableRecord)tr.GetObject(bt[SingleLineBlock], OpenMode.ForRead);

            foreach (var row in rows)
            {
                if (string.IsNullOrEmpty(row.ColAB)) continue;

                var pt = new Point3d(currentX, y, z);
                InsertBlockWithVisualization(space, tr, def, row, pt, row.ColAB, 0);

                double interval = CalculateInterval(row, 0);
                currentX += interval;
            }
            tr.Commit();
            ed.WriteMessage($"\n✓ Однолинейная схема вставлена");
        }

        [CommandMethod("TL_Клеммники")]
        public static void DrawTerminals()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var path = GetExcelFilePath(doc, SpecEOMAddress);

            if (string.IsNullOrEmpty(path))
            {
                ed.WriteMessage("\nФайл не найден");
                return;
            }

            var rows = ReadExcelSheet(path, "Расчет нагрузок");
            if (rows.Count == 0) return;

            var ppr = ed.GetPoint("\nУкажите точку вставки: ");
            if (ppr.Status != PromptStatus.OK) return;

            double currentX = ppr.Value.X;
            double y = ppr.Value.Y;
            double z = ppr.Value.Z;

            using var tr = db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            if (!bt.Has(TerminalsBlock))
            {
                ed.WriteMessage($"\nОшибка: блок '{TerminalsBlock}' не найден");
                return;
            }

            var def = (BlockTableRecord)tr.GetObject(bt[TerminalsBlock], OpenMode.ForRead);

            int count = 0;
            foreach (var row in rows)
            {
                string visualization = row.ColAM?.Trim();

                if (string.IsNullOrEmpty(visualization) || visualization == "#N/A" || visualization == "0")
                    continue;

                int blockCount = GetBlockCountFromAF(row.ColAF);
                double blockWidth = GetBlockWidth(row.ColAO, 50.0);

                for (int i = 0; i < blockCount; i++)
                {
                    var pt = new Point3d(currentX, y, z);
                    InsertBlockWithVisualization(space, tr, def, row, pt, visualization, 1);

                    currentX += blockWidth;
                    count++;
                }
            }

            ed.WriteMessage($"\n✓ Вставлено клемм: {count}");
            tr.Commit();
        }

        [CommandMethod("TL_Изображения_автоматов")]
        public static void DrawCircuitBreakers()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var path = GetExcelFilePath(doc, SpecEOMAddress);

            if (string.IsNullOrEmpty(path))
            {
                ed.WriteMessage("\nФайл не найден");
                return;
            }

            var dataRows = ReadExcelSheet(path, "Расчет нагрузок");
            if (dataRows.Count == 0) return;

            var ppr = ed.GetPoint("\nУкажите точку вставки: ");
            if (ppr.Status != PromptStatus.OK) return;

            double currentX = ppr.Value.X;
            double y = ppr.Value.Y;
            double z = ppr.Value.Z;

            using var tr = db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            if (!bt.Has(CircuitBreakersBlock))
            {
                ed.WriteMessage($"\nОшибка: блок '{CircuitBreakersBlock}' не найден");
                return;
            }

            var def = (BlockTableRecord)tr.GetObject(bt[CircuitBreakersBlock], OpenMode.ForRead);

            int count = 0;
            foreach (var row in dataRows)
            {
                string visualization = row.ColAN?.Trim();

                if (string.IsNullOrEmpty(visualization) || visualization == "#N/A" || visualization == "0")
                    continue;

                double blockWidth = GetBlockWidth(row.ColAP, 50.0);

                var pt = new Point3d(currentX, y, z);
                InsertBlockWithVisualization(space, tr, def, row, pt, visualization, 2);

                currentX += blockWidth;
                count++;
            }

            ed.WriteMessage($"\n✓ Вставлено автоматов: {count}");
            tr.Commit();
        }

        private class ExcelRow(int r, ExcelWorksheet ws)
        {
            public int Row = r;
            public string ColA = ws.Cells[$"A{r}"].Text, ColB = ws.Cells[$"B{r}"].Text, ColC = ws.Cells[$"C{r}"].Text, ColH = ws.Cells[$"H{r}"].Text, ColI = ws.Cells[$"I{r}"].Text, ColJ = ws.Cells[$"J{r}"].Text, ColK = ws.Cells[$"K{r}"].Text, ColM = GetNumeric(ws, "M", r), ColN = GetNumeric(ws, "N", r),
                          ColO = GetNumeric(ws, "O", r), ColP = GetNumeric(ws, "P", r), ColQ = GetNumeric(ws, "Q", r), ColR = ws.Cells[$"R{r}"].Text, ColS = ws.Cells[$"S{r}"].Text, ColT = ws.Cells[$"T{r}"].Text, ColU = ws.Cells[$"U{r}"].Text,
                          ColV = ws.Cells[$"V{r}"].Text, ColW = GetNumeric(ws, "W", r), ColX = GetNumeric(ws, "X", r), ColZ = ws.Cells[$"Z{r}"].Text, ColAA = ws.Cells[$"AA{r}"].Text, ColAB = ws.Cells[$"AB{r}"].Text,
                          ColAC = ws.Cells[$"AC{r}"].Text, ColAD = ws.Cells[$"AD{r}"].Text, ColAE = ws.Cells[$"AE{r}"].Text, ColAF = ws.Cells[$"AF{r}"].Text, ColAG = ws.Cells[$"AG{r}"].Text,
                          ColAM = ws.Cells[$"AM{r}"].Text,
                          ColAN = ws.Cells[$"AN{r}"].Text,
                          ColAO = ws.Cells[$"AO{r}"].Text,
                          ColAP = ws.Cells[$"AP{r}"].Text;

            private static string GetNumeric(ExcelWorksheet ws, string col, int row)
            {
                var text = ws.Cells[$"{col}{row}"].Text;
                return double.TryParse(text, out var v) ? Math.Round(v, 2).ToString() : "";
            }

            public string GetPhaseString()
            {
                var phases = new List<string>();
                if (!string.IsNullOrEmpty(ColO)) phases.Add("A");
                if (!string.IsNullOrEmpty(ColP)) phases.Add("B");
                if (!string.IsNullOrEmpty(ColQ)) phases.Add("C");
                return string.Join(",", phases);
            }

            public string GetProlozhenie()
            {
                return string.IsNullOrEmpty(ColA)
                    ? ""
                    : $"{ColV} L={ColW}м dU={ColX}";
            }
        }
    }
}
