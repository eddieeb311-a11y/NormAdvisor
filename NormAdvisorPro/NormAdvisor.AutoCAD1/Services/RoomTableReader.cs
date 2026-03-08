using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace NormAdvisor.AutoCAD1.Services
{
    public class RoomTableReader
    {
        /// <summary>
        /// Хэрэглэгчээс горим сонгуулж өрөө таних
        /// </summary>
        public List<Models.RoomInfo> ReadRooms()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // Горим сонгох
            var options = new PromptKeywordOptions(
                "\nӨрөөний хүснэгт таних горим сонгоно уу [Table/Region]: ", "Table Region");
            options.AllowNone = false;

            var result = ed.GetKeywords(options);
            if (result.Status != PromptStatus.OK)
                return new List<Models.RoomInfo>();

            if (result.StringResult == "Table")
                return ReadFromTable();
            else
                return ReadFromRegion();
        }

        /// <summary>
        /// Горим 1: AutoCAD Table объектоос уншина
        /// </summary>
        public List<Models.RoomInfo> ReadFromTableDirect() => ReadFromTable();
        private List<Models.RoomInfo> ReadFromTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var rooms = new List<Models.RoomInfo>();

            // Table сонгуулах
            var promptOpts = new PromptEntityOptions("\nХүснэгт (Table) сонгоно уу: ");
            promptOpts.SetRejectMessage("\nTable объект сонгоно уу.");
            promptOpts.AddAllowedClass(typeof(Table), true);

            var entResult = ed.GetEntity(promptOpts);
            if (entResult.Status != PromptStatus.OK)
                return rooms;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var table = tr.GetObject(entResult.ObjectId, OpenMode.ForRead) as Table;
                if (table == null) return rooms;

                // Баганы индекс олох (№, Нэр, Талбай)
                int colNum = -1, colName = -1, colArea = -1;

                // Header мөрөөс баганы нэр олох (эхний 3 мөрийг шалгана)
                for (int row = 0; row < Math.Min(3, table.Rows.Count); row++)
                {
                    // Merge хийсэн гарчиг мөрийг алгасах (нэг cell бүх баганыг хамарсан)
                    // Жишээ нь: "ӨРӨӨНИЙ ТОДОРХОЙЛОЛТ" гэсэн merge мөр
                    bool isMergedTitle = false;
                    try
                    {
                        // Хэрэв энэ мөрийн эхний cell merge хийсэн бол алгасна
                        var range = CellRange.Create(table, row, 0, row, table.Columns.Count - 1);
                        if (table.Cells[row, 0].IsMerged == true)
                        {
                            string firstCell = GetCellText(table, row, 0).ToLower().Trim();
                            if (firstCell.Contains("тодорхойлолт") || firstCell.Contains("жагсаалт") ||
                                firstCell.Length > 15)
                            {
                                isMergedTitle = true;
                            }
                        }
                    }
                    catch { }

                    if (isMergedTitle) continue;

                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        string cellText = GetCellText(table, row, col).ToLower().Trim();

                        if (string.IsNullOrEmpty(cellText)) continue;

                        if (cellText.Contains("№") || cellText.Contains("д/д") || cellText == "no")
                            colNum = col;
                        else if (cellText.Contains("нэр") || cellText.Contains("name"))
                            colName = col;
                        else if (cellText.Contains("талбай") || cellText.Contains("м²") || cellText.Contains("area") || cellText.Contains("хэмжээ"))
                            colArea = col;
                    }
                    if (colNum >= 0 || colName >= 0) break;
                }

                // Хэрэв "нэр" олдсон ч "талбай" олдоогүй бол сүүлийн баганыг талбай гэж тооцно
                if (colName >= 0 && colArea < 0)
                {
                    colArea = table.Columns.Count - 1;
                    if (colArea == colName) colArea = -1; // нэр = сүүлийн багана бол талбай байхгүй
                }

                // Хэрэв баганы нэр олдохгүй бол автоматаар таамаглана
                // (ихэвчлэн 1-р багана = №, 2-р = Нэр, 3-р = Талбай)
                if (colName < 0)
                {
                    if (table.Columns.Count >= 3)
                    {
                        colNum = 0;
                        colName = 1;
                        colArea = 2;
                    }
                    else if (table.Columns.Count >= 2)
                    {
                        colName = 0;
                        colArea = 1;
                    }
                    else
                    {
                        colName = 0;
                    }
                }

                // Data мөрүүдийг уншина (header-ийн дараахаас)
                int startRow = FindDataStartRow(table);
                int roomNumber = 1;
                string currentSection = ""; // Одоогийн секц/бүлгийн нэр

                for (int row = startRow; row < table.Rows.Count; row++)
                {
                    // Merge хийсэн секц гарчиг мөр шалгах (жишээ: "А-Кофешоп", "Б-Хүнсний дэлгүүр")
                    bool isSectionHeader = false;
                    try
                    {
                        if (table.Cells[row, 0].IsMerged == true)
                        {
                            string mergedText = GetCellText(table, row, 0).Trim();
                            if (!string.IsNullOrEmpty(mergedText) &&
                                !mergedText.ToLower().Contains("нийт") &&
                                !mergedText.ToLower().Contains("total") &&
                                !mergedText.ToLower().Contains("бүгд"))
                            {
                                currentSection = mergedText;
                                isSectionHeader = true;
                            }
                        }
                    }
                    catch { }
                    if (isSectionHeader) continue;

                    string name = colName >= 0 ? GetCellText(table, row, colName).Trim() : "";
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    // "Нийт" мөр бол алгасах
                    if (name.ToLower().Contains("нийт") || name.ToLower().Contains("total") || name.ToLower().Contains("бүгд"))
                        continue;

                    // Секц гарчиг шиг мөр — дугааргүй, талбайгүй, нэр нь урт
                    // (merge хийгдээгүй ч гэсэн)
                    string numCheck = colNum >= 0 ? GetCellText(table, row, colNum).Trim() : "";
                    string areaCheck = colArea >= 0 ? GetCellText(table, row, colArea).Trim() : "";
                    if (string.IsNullOrEmpty(numCheck) && string.IsNullOrEmpty(areaCheck) && name.Length > 1)
                    {
                        currentSection = name;
                        continue;
                    }

                    var room = new Models.RoomInfo
                    {
                        Name = name,
                        RawText = name,
                        SectionName = currentSection
                    };

                    // Дугаар уншина (жишээ: "a1", "b2", "1", "15")
                    if (colNum >= 0)
                    {
                        string numText = GetCellText(table, row, colNum).Trim();
                        room.RoomId = numText; // Анхны текст хадгалах
                        if (int.TryParse(numText, out int num))
                            room.Number = num;
                        else
                            room.Number = roomNumber;
                    }
                    else
                    {
                        room.Number = roomNumber;
                        room.RoomId = roomNumber.ToString();
                    }

                    // Талбай уншина
                    if (colArea >= 0)
                    {
                        string areaText = GetCellText(table, row, colArea).Trim();
                        room.Area = ParseArea(areaText);
                    }

                    rooms.Add(room);
                    roomNumber++;
                }

                tr.Commit();
            }

            return rooms;
        }

        /// <summary>
        /// Горим 2: Region сонгох (2 цэг) → Text/MText уншина
        /// </summary>
        public List<Models.RoomInfo> ReadFromRegionDirect() => ReadFromRegion();
        private List<Models.RoomInfo> ReadFromRegion()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var rooms = new List<Models.RoomInfo>();

            // 1-р цэг: дээд зүүн
            var pt1Result = ed.GetPoint("\nДээд зүүн булангийн цэг сонгоно уу: ");
            if (pt1Result.Status != PromptStatus.OK) return rooms;

            // 2-р цэг: доод баруун
            var pt2Options = new PromptCornerOptions("\nДоод баруун булангийн цэг сонгоно уу: ", pt1Result.Value);
            var pt2Result = ed.GetCorner(pt2Options);
            if (pt2Result.Status != PromptStatus.OK) return rooms;

            Point3d pt1 = pt1Result.Value;
            Point3d pt2 = pt2Result.Value;

            // Бүсийн хязгаар тодорхойлох
            double minX = Math.Min(pt1.X, pt2.X);
            double maxX = Math.Max(pt1.X, pt2.X);
            double minY = Math.Min(pt1.Y, pt2.Y);
            double maxY = Math.Max(pt1.Y, pt2.Y);

            // Бүс дотор байгаа бүх Text/MText олох
            var textEntities = new List<TextEntity>();

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId objId in btr)
                {
                    var ent = tr.GetObject(objId, OpenMode.ForRead);

                    if (ent is DBText dbText)
                    {
                        var pos = dbText.Position;
                        if (IsInBounds(pos, minX, maxX, minY, maxY))
                        {
                            textEntities.Add(new TextEntity
                            {
                                Text = dbText.TextString.Trim(),
                                X = pos.X,
                                Y = pos.Y,
                                Height = dbText.Height
                            });
                        }
                    }
                    else if (ent is MText mText)
                    {
                        var pos = mText.Location;
                        if (IsInBounds(pos, minX, maxX, minY, maxY))
                        {
                            // MText-ийн formatting арилгах
                            string cleanText = CleanMText(mText.Contents);
                            textEntities.Add(new TextEntity
                            {
                                Text = cleanText.Trim(),
                                X = pos.X,
                                Y = pos.Y,
                                Height = mText.TextHeight
                            });
                        }
                    }
                    // BlockReference дотор текст хайх
                    else if (ent is BlockReference blkRef)
                    {
                        ScanBlockForText(blkRef, tr, minX, maxX, minY, maxY, textEntities);
                    }
                }

                tr.Commit();
            }

            if (textEntities.Count == 0)
            {
                ed.WriteMessage("\nСонгосон бүсэд текст олдсонгүй.");
                return rooms;
            }

            // Text-үүдийг мөр (row) болгон бүлэглэнэ
            // Y координатаар ойролцоо (tolerance) байвал нэг мөр гэж тооцно
            var rows = GroupIntoRows(textEntities);

            // Мөр бүрийг parse хийнэ
            int roomNumber = 1;
            foreach (var row in rows)
            {
                // Мөр доторхи text-үүдийг X-ээр эрэмбэлнэ (зүүнээс баруун)
                var sortedTexts = row.OrderBy(t => t.X).ToList();

                var room = ParseRowTexts(sortedTexts, roomNumber);
                if (room != null)
                {
                    rooms.Add(room);
                    roomNumber++;
                }
            }

            return rooms;
        }

        #region Helper methods

        private string GetCellText(Table table, int row, int col)
        {
            try
            {
                // Value нь цэвэр утга (formatting-гүй) өгдөг
                var cellValue = table.Cells[row, col].Value;
                if (cellValue != null)
                {
                    string val = cellValue.ToString();
                    if (!string.IsNullOrEmpty(val) && !val.Contains(@"{\"))
                        return val;
                }

                // TextString нь MText formatting кодтой байж болно
                var textString = table.Cells[row, col].TextString;
                if (string.IsNullOrEmpty(textString)) return "";

                // MText formatting арилгах
                return CleanMText(textString);
            }
            catch
            {
                return "";
            }
        }

        private int FindDataStartRow(Table table)
        {
            // Header мөрийг алгасах (гарчиг + баганы нэр = 2-3 мөр байж болно)
            for (int row = 0; row < Math.Min(5, table.Rows.Count); row++)
            {
                string firstCell = GetCellText(table, row, 0).ToLower().Trim();

                // Header мөрүүд: гарчиг, баганы нэр гэх мэт
                if (firstCell.Contains("№") || firstCell.Contains("нэр") ||
                    firstCell.Contains("д/д") || firstCell.Contains("name") ||
                    firstCell.Contains("өрөө") || firstCell.Contains("тодорхойлолт") ||
                    firstCell.Contains("жагсаалт") || string.IsNullOrEmpty(firstCell))
                    continue;

                // Тоо эхэлсэн бол data мөр олдлоо
                if (int.TryParse(firstCell, out _))
                    return row;

                continue;
            }
            return 2; // Default: 3 дахь мөрөөс эхлэх (гарчиг + header)
        }

        private double ParseArea(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;

            // "16.94 м²" → 16.94
            // "16,94"    → 16.94
            text = text.Replace("м²", "").Replace("m²", "").Replace("м2", "").Trim();
            text = text.Replace(',', '.');

            // Зөвхөн тоо хэсгийг олох
            var match = Regex.Match(text, @"[\d]+\.?[\d]*");
            if (match.Success && double.TryParse(match.Value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double area))
            {
                // AutoCAD Table Value нь мм² өгч болно → м² руу хөрвүүлэх
                // Хэрэв 10000-аас их бол мм² гэж тооцно (10000 мм² = 0.01 м²)
                if (area > 10000)
                    area /= 1000000.0;

                return Math.Round(area, 2);
            }
            return 0;
        }

        private bool IsInBounds(Point3d pos, double minX, double maxX, double minY, double maxY)
        {
            return pos.X >= minX && pos.X <= maxX && pos.Y >= minY && pos.Y <= maxY;
        }

        private string CleanMText(string mTextContent)
        {
            if (string.IsNullOrEmpty(mTextContent)) return "";

            // AutoCAD MText formatting codes арилгах
            // \A1; = alignment, \f...; = font, \H...; = height, etc.
            string cleaned = Regex.Replace(mTextContent, @"\\[AaCcFfHhLlOoQqSsTtWw][^;]*;", "");
            cleaned = Regex.Replace(cleaned, @"\\[Pp]", "\n"); // \P = newline
            cleaned = Regex.Replace(cleaned, @"\{|\}", ""); // { } хаалт арилгах
            cleaned = Regex.Replace(cleaned, @"\\~", " "); // \~ = space
            return cleaned.Trim();
        }

        private List<List<TextEntity>> GroupIntoRows(List<TextEntity> texts)
        {
            // Y координатаар эрэмбэлнэ (дээрээс доош = Y буурах)
            var sorted = texts.OrderByDescending(t => t.Y).ToList();

            var rows = new List<List<TextEntity>>();
            var currentRow = new List<TextEntity> { sorted[0] };
            double currentY = sorted[0].Y;

            // Tolerance: текстийн өндрийн 0.7 дахин
            double tolerance = sorted[0].Height * 0.7;
            if (tolerance < 1.0) tolerance = 1.0; // min 1 unit

            for (int i = 1; i < sorted.Count; i++)
            {
                if (Math.Abs(sorted[i].Y - currentY) <= tolerance)
                {
                    // Нэг мөрт нэмнэ
                    currentRow.Add(sorted[i]);
                }
                else
                {
                    // Шинэ мөр эхэлнэ
                    rows.Add(currentRow);
                    currentRow = new List<TextEntity> { sorted[i] };
                    currentY = sorted[i].Y;
                    tolerance = sorted[i].Height * 0.7;
                    if (tolerance < 1.0) tolerance = 1.0;
                }
            }
            rows.Add(currentRow);

            return rows;
        }

        private Models.RoomInfo ParseRowTexts(List<TextEntity> rowTexts, int defaultNumber)
        {
            if (rowTexts.Count == 0) return null;

            // Header / гарчиг мөр шалгах — алгасна
            string combined = string.Join(" ", rowTexts.Select(t => t.Text)).ToLower();
            if (combined.Contains("№") && (combined.Contains("нэр") || combined.Contains("өрөө")))
                return null;
            if (combined.Contains("нийт") || combined.Contains("total") || combined.Contains("бүгд"))
                return null;
            // "ӨРӨӨНИЙ ТОДОРХОЙЛОЛТ", "ӨРӨӨНИЙ ЖАГСААЛТ" гэсэн гарчиг алгасах
            if (combined.Contains("тодорхойлолт") || combined.Contains("жагсаалт"))
                return null;
            // "өрөөний нэр" гэсэн header мөр (№ байхгүй ч)
            if (combined.Contains("өрөөний нэр"))
                return null;

            var room = new Models.RoomInfo { Number = defaultNumber };

            if (rowTexts.Count >= 3)
            {
                // 3+ текст: [№] [Нэр] [Талбай]
                string firstText = rowTexts[0].Text;
                room.RoomId = firstText; // Анхны дугаар текст хадгалах

                if (int.TryParse(firstText, out int num))
                {
                    room.Number = num;
                    room.Name = rowTexts[1].Text;
                    room.Area = ParseArea(rowTexts[2].Text);
                }
                else
                {
                    // "a1", "b2" гэсэн формат — нэрийг 2 дахь текстээс авна
                    room.Name = rowTexts[1].Text;
                    room.Area = ParseArea(rowTexts[2].Text);
                }
            }
            else if (rowTexts.Count == 2)
            {
                // 2 текст: [Нэр] [Талбай] эсвэл [№] [Нэр]
                if (double.TryParse(rowTexts[1].Text.Replace(',', '.').Replace("м²", "").Trim(),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out _))
                {
                    room.Name = rowTexts[0].Text;
                    room.Area = ParseArea(rowTexts[1].Text);
                }
                else
                {
                    room.Name = string.Join(" ", rowTexts.Select(t => t.Text));
                }
            }
            else
            {
                // 1 текст: зөвхөн нэр
                room.Name = rowTexts[0].Text;
            }

            room.RawText = string.Join(" | ", rowTexts.Select(t => t.Text));

            // Хоосон нэртэй бол алгасна
            if (string.IsNullOrWhiteSpace(room.Name)) return null;

            return room;
        }

        #endregion

        /// <summary>
        /// BlockReference доторхи текст (AttributeReference, DBText, MText) хайж цуглуулах.
        /// Координатуудыг BlockTransform ашиглан world координат руу хөрвүүлнэ.
        /// </summary>
        private void ScanBlockForText(BlockReference blkRef, Transaction tr,
            double minX, double maxX, double minY, double maxY, List<TextEntity> textEntities)
        {
            Matrix3d transform = blkRef.BlockTransform;

            // AttributeReference (блокийн attribute утгууд — ихэвчлэн өрөөний дугаар)
            try
            {
                foreach (ObjectId attId in blkRef.AttributeCollection)
                {
                    try
                    {
                        var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                        if (att == null) continue;

                        string content = att.TextString?.Trim();
                        if (string.IsNullOrEmpty(content)) continue;

                        var pos = att.Position;
                        if (!IsInBounds(pos, minX, maxX, minY, maxY)) continue;

                        textEntities.Add(new TextEntity
                        {
                            Text = content,
                            X = pos.X,
                            Y = pos.Y,
                            Height = att.Height > 0 ? att.Height : 2.5
                        });
                    }
                    catch { }
                }
            }
            catch { }

            // Block definition дотор текст хайх
            try
            {
                var blockDef = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (blockDef == null) return;

                foreach (ObjectId entId in blockDef)
                {
                    Entity blockEnt;
                    try { blockEnt = tr.GetObject(entId, OpenMode.ForRead) as Entity; }
                    catch { continue; }
                    if (blockEnt == null) continue;

                    if (blockEnt is DBText txt)
                    {
                        string content = txt.TextString?.Trim();
                        if (string.IsNullOrEmpty(content)) continue;

                        var worldPos = txt.Position.TransformBy(transform);
                        if (!IsInBounds(worldPos, minX, maxX, minY, maxY)) continue;

                        textEntities.Add(new TextEntity
                        {
                            Text = content,
                            X = worldPos.X,
                            Y = worldPos.Y,
                            Height = txt.Height > 0 ? txt.Height : 2.5
                        });
                    }
                    else if (blockEnt is MText mtxt)
                    {
                        string content = CleanMText(mtxt.Contents ?? mtxt.Text ?? "").Trim();
                        if (string.IsNullOrEmpty(content)) continue;

                        var worldPos = mtxt.Location.TransformBy(transform);
                        if (!IsInBounds(worldPos, minX, maxX, minY, maxY)) continue;

                        textEntities.Add(new TextEntity
                        {
                            Text = content,
                            X = worldPos.X,
                            Y = worldPos.Y,
                            Height = mtxt.TextHeight > 0 ? mtxt.TextHeight : 2.5
                        });
                    }
                    else if (blockEnt is AttributeDefinition attDef && attDef.Constant)
                    {
                        string content = attDef.TextString?.Trim();
                        if (string.IsNullOrEmpty(content)) continue;

                        var worldPos = attDef.Position.TransformBy(transform);
                        if (!IsInBounds(worldPos, minX, maxX, minY, maxY)) continue;

                        textEntities.Add(new TextEntity
                        {
                            Text = content,
                            X = worldPos.X,
                            Y = worldPos.Y,
                            Height = attDef.Height > 0 ? attDef.Height : 2.5
                        });
                    }
                    // Nested block
                    else if (blockEnt is BlockReference nestedRef)
                    {
                        try { ScanBlockForText(nestedRef, tr, minX, maxX, minY, maxY, textEntities); }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private class TextEntity
        {
            public string Text { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Height { get; set; }
        }
    }
}
