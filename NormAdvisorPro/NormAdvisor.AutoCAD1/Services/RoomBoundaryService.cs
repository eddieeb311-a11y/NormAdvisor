using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using GI = Autodesk.AutoCAD.GraphicsInterface;
using NormAdvisor.AutoCAD1.Models;

namespace NormAdvisor.AutoCAD1.Services
{
    /// <summary>
    /// Ó¨Ñ€Ó©Ó©Ð½Ð¸Ð¹ polyline Ñ…Ò¯Ñ€ÑÑÑ‚ÑÐ¹ Ð°Ð¶Ð¸Ð»Ð»Ð°Ñ… ÑÐµÑ€Ð²Ð¸Ñ
    /// XData Ñ…Ð°Ð´Ð³Ð°Ð»Ð°Ñ…, highlight, scan
    /// </summary>
    public class RoomBoundaryService
    {
        public const string RegAppName = "NORMADVISOR";
        public const string LayerName = "NORM_ROOM_BOUNDARY";
        public const string RoomAreaLayerName = "NORM_ROOM_AREA";
        public const string RoomNumLayerName = "NORM_ROOM_NUM";
        private static readonly object VirtualLock = new object();
        private static readonly Dictionary<string, Extents3d> VirtualBoundsByRoomId =
            new Dictionary<string, Extents3d>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<int, Extents3d> VirtualBoundsByRoomNumber =
            new Dictionary<int, Extents3d>();

        public struct LinkResult
        {
            public ObjectId PolylineId;
            public double DrawnArea;
            public int RoomNumber;
            public string RoomId;
        }

        private string BuildRoomMatchKey(RoomInfo room)
        {
            if (!string.IsNullOrWhiteSpace(room.RoomId) && Regex.IsMatch(room.RoomId, "[A-Za-zА-Яа-я]"))
                return room.RoomId.Trim();

            return $"#N{room.Number}:{(room.Name ?? string.Empty).Trim()}";
        }

        /// <summary>
        /// âœŽ Ñ‚Ð¾Ð²Ñ‡: [Ð—ÑƒÑ€Ð°Ñ…/Ð¡Ð¾Ð½Ð³Ð¾Ñ…] ÑÐ¾Ð½Ð³Ð¾Ð»Ñ‚ â†’ Ð—ÑƒÑ€Ð°Ñ… = AREA ÑˆÐ¸Ð³, Ð¡Ð¾Ð½Ð³Ð¾Ñ… = Ð±Ð°Ð¹Ð³Ð°Ð° polyline
        /// </summary>
        public LinkResult DrawOrSelectBoundary(RoomInfo room)
        {
            var result = new LinkResult { PolylineId = ObjectId.Null, DrawnArea = 0 };

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return result;
            var ed = doc.Editor;

            var kwOpts = new PromptKeywordOptions(
                $"\n[{room.Number}. {room.Name}] [Ð—ÑƒÑ€Ð°Ñ…(Z)/Ð¡Ð¾Ð½Ð³Ð¾Ñ…(S)] <Ð—ÑƒÑ€Ð°Ñ…>: ");
            kwOpts.Keywords.Add("Ð—ÑƒÑ€Ð°Ñ…", "Z", "Ð—ÑƒÑ€Ð°Ñ…(Z)");
            kwOpts.Keywords.Add("Ð¡Ð¾Ð½Ð³Ð¾Ñ…", "S", "Ð¡Ð¾Ð½Ð³Ð¾Ñ…(S)");
            kwOpts.Keywords.Default = "Ð—ÑƒÑ€Ð°Ñ…";
            kwOpts.AllowNone = true;

            var kwResult = ed.GetKeywords(kwOpts);
            if (kwResult.Status == PromptStatus.Cancel) return result;

            string choice = kwResult.Status == PromptStatus.OK ? kwResult.StringResult : "Ð—ÑƒÑ€Ð°Ñ…";

            return choice == "Ð¡Ð¾Ð½Ð³Ð¾Ñ…" ? SelectExistingPolyline(room) : DrawNewPolyline(room);
        }

        /// <summary>
        /// Ð¨ÑƒÑƒÐ´ Ð±Ð°Ð¹Ð³Ð°Ð° polyline ÑÐ¾Ð½Ð³Ð¾Ñ… (NORMSELECTROOM-Ð´ Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°)
        /// </summary>
        public LinkResult SelectBoundary(RoomInfo room)
        {
            return SelectExistingPolyline(room);
        }

        /// <summary>
        /// AREA command ÑˆÐ¸Ð³ Ð·ÑƒÑ€Ð°Ñ… Ð³Ð¾Ñ€Ð¸Ð¼:
        /// - GetPoint() + rubber-band ÑˆÑƒÐ³Ð°Ð¼
        /// - TransientGraphics Ð½Ð¾Ð³Ð¾Ð¾Ð½ polygon preview (Ð±Ð¾Ð´Ð¸Ñ‚ Ñ†Ð°Ð³Ð°Ð°Ñ€)
        /// - Ð”ÑƒÑƒÑÐ°Ñ…Ð°Ð´ visualization Ð°Ð»Ð³Ð° Ð±Ð¾Ð»Ð½Ð¾
        /// - Polyline Ð½ÑŒ Ð½Ò¯Ð´ÑÐ½Ð´ Ñ…Ð°Ñ€Ð°Ð³Ð´Ð°Ñ…Ð³Ò¯Ð¹ Ð´Ð°Ð²Ñ…Ð°Ñ€Ð³Ð° Ð´ÑÑÑ€ Ñ…Ð°Ð´Ð³Ð°Ð»Ð°Ð³Ð´Ð°Ð½Ð°
        /// </summary>
        private LinkResult DrawNewPolyline(RoomInfo room)
        {
            var result = new LinkResult { PolylineId = ObjectId.Null, DrawnArea = 0 };

            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            ed.WriteMessage($"\n  [{room.Number}. {room.Name}] Ð¦ÑÐ³ ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ. Enter = Ð´ÑƒÑƒÑÐ³Ð°Ñ…");

            var points = new List<Point2d>();
            Point3d? lastPoint = null;

            // TransientGraphics preview
            Polyline transPoly = null;
            var tm = GI.TransientManager.CurrentTransientManager;
            var intCol = new IntegerCollection();

            try
            {
                while (true)
                {
                    PromptPointOptions ptOpts;
                    if (lastPoint == null)
                    {
                        ptOpts = new PromptPointOptions("\n  Ð­Ñ…Ð½Ð¸Ð¹ Ñ†ÑÐ³: ");
                    }
                    else
                    {
                        string areaInfo = points.Count >= 3 ? $" ~{CalcTempArea(points):F1}Ð¼Â²" : "";
                        ptOpts = new PromptPointOptions($"\n  Ð”Ð°Ñ€Ð°Ð°Ð³Ð¸Ð¹Ð½ Ñ†ÑÐ³ [{points.Count}Ñ†{areaInfo}]: ");
                        ptOpts.UseBasePoint = true;
                        ptOpts.BasePoint = lastPoint.Value;
                        ptOpts.AllowNone = true;
                    }

                    var ptResult = ed.GetPoint(ptOpts);

                    if (ptResult.Status == PromptStatus.None) // Enter = Ð´ÑƒÑƒÑÐ³Ð°Ñ…
                        break;
                    if (ptResult.Status == PromptStatus.Cancel) // Esc = Ñ†ÑƒÑ†Ð»Ð°Ñ…
                    {
                        EraseTransient(tm, ref transPoly, intCol);
                        ed.WriteMessage("\n  Ð¦ÑƒÑ†Ð°Ð»Ð»Ð°Ð°.");
                        return result;
                    }
                    if (ptResult.Status == PromptStatus.OK)
                    {
                        points.Add(new Point2d(ptResult.Value.X, ptResult.Value.Y));
                        lastPoint = ptResult.Value;

                        // ÐÐ¾Ð³Ð¾Ð¾Ð½ polygon preview ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
                        UpdateTransientPoly(tm, ref transPoly, points, intCol);
                    }
                }

                // Transient preview Ð°Ñ€Ð¸Ð»Ð³Ð°Ñ… (AREA ÑˆÐ¸Ð³ Ð°Ð»Ð³Ð° Ð±Ð¾Ð»Ð½Ð¾)
                EraseTransient(tm, ref transPoly, intCol);

                if (points.Count < 3)
                {
                    ed.WriteMessage("\n  3-Ð°Ð°Ñ Ð´ÑÑÑˆ Ñ†ÑÐ³ ÑˆÐ°Ð°Ñ€Ð´Ð»Ð°Ð³Ð°Ñ‚Ð°Ð¹.");
                    return result;
                }

                // Polyline Ò¯Ò¯ÑÐ³ÑÐ¶ XData Ñ…Ð°Ð´Ð³Ð°Ð»Ð°Ñ… (Ñ…Ð°Ñ€Ð°Ð³Ð´Ð°Ñ…Ð³Ò¯Ð¹ Ð´Ð°Ð²Ñ…Ð°Ñ€Ð³Ð° Ð´ÑÑÑ€)
                using (var lockDoc = doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    EnsureRegApp(db, tr);
                    EnsureLayer(db, tr);

                    var pl = new Polyline();
                    for (int i = 0; i < points.Count; i++)
                        pl.AddVertexAt(i, points[i], 0, 0, 0);
                    pl.Closed = true;
                    pl.Layer = LayerName;
                    pl.Visible = false;

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    ms.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);

                    AttachXData(pl, tr, db, room);

                    result.PolylineId = pl.ObjectId;
                    result.DrawnArea = NormalizeAreaToM2(pl.Area, room.Area);

                    // Layer-Ð¸Ð¹Ð³ Ð±Ð°Ñ€Ð°Ð³ Ñ…Ð°Ñ€Ð°Ð³Ð´Ð°Ñ…Ð³Ò¯Ð¹ Ð±Ð¾Ð»Ð³Ð¾Ñ… (color 251 = Ð¼Ð°Ñˆ Ñ†Ð°Ð¹Ð²Ð°Ñ€ ÑÐ°Ð°Ñ€Ð°Ð»)
                    var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (layerTable.Has(LayerName))
                    {
                        var layer = (LayerTableRecord)tr.GetObject(layerTable[LayerName], OpenMode.ForWrite);
                        layer.Color = Color.FromColorIndex(ColorMethod.ByAci, 251);
                        layer.IsPlottable = false;
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n  {points.Count} Ñ†ÑÐ³. Ð¢Ð°Ð»Ð±Ð°Ð¹: {result.DrawnArea:F2} Ð¼Â²");
                }
            }
            catch (System.Exception ex)
            {
                EraseTransient(tm, ref transPoly, intCol);
                ed.WriteMessage($"\n  ÐÐ»Ð´Ð°Ð°: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// TransientGraphics â€” Ð½Ð¾Ð³Ð¾Ð¾Ð½ polygon preview ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
        /// </summary>
        private void UpdateTransientPoly(GI.TransientManager tm, ref Polyline poly, List<Point2d> points, IntegerCollection intCol)
        {
            // Ð¥ÑƒÑƒÑ‡Ð½Ñ‹Ð³ ÑƒÑÑ‚Ð³Ð°Ñ…
            EraseTransient(tm, ref poly, intCol);

            if (points.Count < 2) return;

            poly = new Polyline();
            for (int i = 0; i < points.Count; i++)
                poly.AddVertexAt(i, points[i], 0, 0, 0);
            if (points.Count >= 3) poly.Closed = true;
            poly.Color = Color.FromColorIndex(ColorMethod.ByAci, 3); // ÐÐ¾Ð³Ð¾Ð¾Ð½
            poly.LineWeight = LineWeight.LineWeight030;

            tm.AddTransient(poly, GI.TransientDrawingMode.DirectTopmost, 128, intCol);
        }

        /// <summary>
        /// Transient entity Ð°Ñ€Ð¸Ð»Ð³Ð°Ñ…
        /// </summary>
        private void EraseTransient(GI.TransientManager tm, ref Polyline poly, IntegerCollection intCol)
        {
            if (poly != null)
            {
                try { tm.EraseTransient(poly, intCol); } catch { }
                poly.Dispose();
                poly = null;
            }
        }

        /// <summary>
        /// Shoelace formula â€” Ñ‚Ò¯Ñ€ Ñ‚Ð°Ð»Ð±Ð°Ð¹ Ñ‚Ð¾Ð¾Ñ†Ð¾Ð¾Ð»Ð¾Ñ…
        /// </summary>
        private double CalcTempArea(List<Point2d> pts)
        {
            double area = 0;
            int n = pts.Count;
            for (int i = 0; i < n; i++)
            {
                int j = (i + 1) % n;
                area += pts[i].X * pts[j].Y;
                area -= pts[j].X * pts[i].Y;
            }
            return Math.Abs(area) / 2.0;
        }

        /// <summary>
        /// Ð‘Ð°Ð¹Ð³Ð°Ð° polyline ÑÐ¾Ð½Ð³Ð¾Ð¶ Ñ…Ð¾Ð»Ð±Ð¾Ñ…
        /// </summary>
        private LinkResult SelectExistingPolyline(RoomInfo room)
        {
            var result = new LinkResult { PolylineId = ObjectId.Null, DrawnArea = 0 };

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return result;

            var ed = doc.Editor;
            var db = doc.Database;

            var opts = new PromptEntityOptions($"\n  Polyline ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ: ");
            opts.SetRejectMessage("\n  Ð—Ó©Ð²Ñ…Ó©Ð½ Polyline ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ.");
            opts.AddAllowedClass(typeof(Polyline), true);
            opts.AddAllowedClass(typeof(Polyline2d), true);
            opts.AddAllowedClass(typeof(Polyline3d), true);

            var entResult = ed.GetEntity(opts);
            if (entResult.Status != PromptStatus.OK) return result;

            using (var lockDoc = doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    EnsureRegApp(db, tr);
                    EnsureLayer(db, tr);

                    var entity = (Entity)tr.GetObject(entResult.ObjectId, OpenMode.ForWrite);

                    // ÐÐ»ÑŒ Ñ…ÑÐ´Ð¸Ð¹Ð½ XData Ð±Ð°Ð¹Ð³Ð°Ð° ÑÑÑÑ…
                    var existingXData = entity.GetXDataForApplication(RegAppName);
                    if (existingXData != null)
                    {
                        ed.WriteMessage("\n  ÐÐ½Ñ…Ð°Ð°Ñ€ÑƒÑƒÐ»Ð³Ð°: Ð­Ð½Ñ polyline Ó©Ó©Ñ€ Ó©Ñ€Ó©Ó©Ñ‚ÑÐ¹ Ñ…Ð¾Ð»Ð±Ð¾Ð³Ð´ÑÐ¾Ð½ Ð±Ð°Ð¹ÑÐ°Ð½. Ð”Ð°Ñ€Ð¶ Ð±Ð¸Ñ‡Ð½Ñ.");
                    }

                    entity.Layer = LayerName;
                    AttachXData(entity, tr, db, room);

                    double area = 0;
                    if (entity is Polyline pl)
                    {
                        if (!pl.Closed)
                        {
                            ed.WriteMessage("\n  Polyline Ñ…Ð°Ð°Ð³Ð´Ð°Ð°Ð³Ò¯Ð¹ Ð±Ð°Ð¹ÑÐ°Ð½. Ð¥Ð°Ð°Ð¶ Ð±Ð°Ð¹Ð½Ð°...");
                            pl.Closed = true;
                        }
                        area = pl.Area;
                    }
                    else if (entity is Polyline2d pl2d)
                        area = pl2d.Area;
                    else if (entity is Polyline3d pl3d)
                        area = pl3d.Area;

                    tr.Commit();

                    result.PolylineId = entResult.ObjectId;
                    result.DrawnArea = NormalizeAreaToM2(area, room.Area);
                }
            }

            return result;
        }

        /// <summary>
        /// RegApp Ð±Ò¯Ñ€Ñ‚Ð³ÑÑ…
        /// </summary>
        private void EnsureRegApp(Database db, Transaction tr)
        {
            var regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!regTable.Has(RegAppName))
            {
                regTable.UpgradeOpen();
                var regApp = new RegAppTableRecord { Name = RegAppName };
                regTable.Add(regApp);
                tr.AddNewlyCreatedDBObject(regApp, true);
            }
        }

        /// <summary>
        /// NORM_ROOM_BOUNDARY layer Ò¯Ò¯ÑÐ³ÑÑ… (locked Ð±Ð¾Ð» unlock Ñ…Ð¸Ð¹Ð½Ñ)
        /// </summary>
        private void EnsureLayer(Database db, Transaction tr)
        {
            var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!layerTable.Has(LayerName))
            {
                layerTable.UpgradeOpen();
                var layer = new LayerTableRecord
                {
                    Name = LayerName,
                    Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                };
                layerTable.Add(layer);
                tr.AddNewlyCreatedDBObject(layer, true);
            }
            else
            {
                // Layer Ð±Ð°Ð¹Ð³Ð°Ð° Ð±Ð¾Ð» locked ÑÑÑÑ…Ð¸Ð¹Ð³ ÑˆÐ°Ð»Ð³Ð°Ð¶ unlock Ñ…Ð¸Ð¹Ð½Ñ
                var existing = (LayerTableRecord)tr.GetObject(layerTable[LayerName], OpenMode.ForRead);
                if (existing.IsLocked)
                {
                    existing.UpgradeOpen();
                    existing.IsLocked = false;
                }
            }
        }

        /// <summary>
        /// Entity Ð´ÑÑÑ€ Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ XData Ñ…Ð°Ð´Ð³Ð°Ð»Ð°Ñ…
        /// </summary>
        private void AttachXData(Entity entity, Transaction tr, Database db, RoomInfo room)
        {
            entity.XData = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, RegAppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, "ROOM"),
                new TypedValue((int)DxfCode.ExtendedDataInteger32, room.Number),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, BuildRoomMatchKey(room)),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, room.Name),
                new TypedValue((int)DxfCode.ExtendedDataReal, room.Area)
            );
        }

        /// <summary>
        /// Entity-Ð¸Ð¹Ð½ XData-Ð°Ð°Ñ Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ð¼ÑÐ´ÑÑÐ»ÑÐ» ÑƒÐ½ÑˆÐ¸Ñ…
        /// </summary>
        public RoomInfo ReadXData(Entity entity)
        {
            var xdata = entity.GetXDataForApplication(RegAppName);
            if (xdata == null) return null;

            var values = xdata.AsArray();
            if (values.Length < 5) return null;
            if (values[1].Value.ToString() != "ROOM") return null;

            string roomId = string.Empty;
            string name;
            double area;

            if (values.Length >= 6)
            {
                roomId = values[3].Value?.ToString() ?? string.Empty;
                name = values[4].Value?.ToString() ?? string.Empty;
                area = Convert.ToDouble(values[5].Value);
            }
            else
            {
                name = values[3].Value?.ToString() ?? string.Empty;
                area = Convert.ToDouble(values[4].Value);
            }

            return new RoomInfo
            {
                Number = Convert.ToInt32(values[2].Value),
                RoomId = roomId,
                Name = name,
                Area = area
            };
        }

        /// <summary>
        /// Ó¨Ñ€Ó©Ó©Ð½Ð¸Ð¹ polyline Ñ€ÑƒÑƒ zoom Ñ…Ð¸Ð¹Ð½Ñ.
        /// Native _.ZOOM _W ÐºÐ¾Ð¼Ð°Ð½Ð´ Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð° (ed.SetCurrentView-ÑÑÑ Ð¸Ð»Ò¯Ò¯ Ð½Ð°Ð¹Ð´Ð²Ð°Ñ€Ñ‚Ð°Ð¹).
        /// roomNumber Ó©Ð³Ð²Ó©Ð» XData scan Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°.
        /// </summary>
        public void ZoomToRoom(ObjectId polylineId, int roomNumber = 0, string roomId = null)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                if (TryGetVirtualBoundary(roomId, roomNumber, out Extents3d virtualExt))
                {
                    ZoomByExtents(doc, virtualExt);
                    return;
                }

                ObjectId targetId = polylineId;
                if (!string.IsNullOrWhiteSpace(roomId))
                {
                    var foundById = FindExistingRoomBoundariesByRoomId(db);
                    if (foundById.TryGetValue(roomId, out ObjectId idByRoomId) && idByRoomId != ObjectId.Null)
                        targetId = idByRoomId;

                    // Compatibility fallback: "#N{number}:{name}" key.
                    if ((targetId == ObjectId.Null || !targetId.IsValid) &&
                        TryParseCompositeRoomKey(roomId, out int n, out string namePart))
                    {
                        var foundByName = FindExistingRoomBoundaryByNumberAndName(db, n, namePart);
                        if (foundByName != ObjectId.Null) targetId = foundByName;
                    }
                }
                else if (roomNumber > 0)
                {
                    var found = FindExistingRoomBoundaries(db);
                    if (found.TryGetValue(roomNumber, out ObjectId xId) && xId != ObjectId.Null)
                        targetId = xId;
                }

                if (targetId == ObjectId.Null || !targetId.IsValid)
                {
                    ed.WriteMessage("\n  Polyline not found.");
                    return;
                }

                Extents3d ext;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var entity = (Entity)tr.GetObject(targetId, OpenMode.ForRead);
                    ext = entity.GeometricExtents;
                    tr.Commit();
                }

                try { ed.SetImpliedSelection(new ObjectId[] { targetId }); } catch { }
                ZoomByExtents(doc, ext);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nZoom error: {ex.Message}");
            }
        }

        private void ZoomByExtents(Document doc, Extents3d ext)
        {
            var ed = doc.Editor;

            double padX = (ext.MaxPoint.X - ext.MinPoint.X) * 0.25;
            double padY = (ext.MaxPoint.Y - ext.MinPoint.Y) * 0.25;
            if (padX < 1.0) padX = 1.0;
            if (padY < 1.0) padY = 1.0;

            double x1 = ext.MinPoint.X - padX;
            double y1 = ext.MinPoint.Y - padY;
            double x2 = ext.MaxPoint.X + padX;
            double y2 = ext.MaxPoint.Y + padY;

            double cx = (x1 + x2) / 2.0;
            double cy = (y1 + y2) / 2.0;
            double w = Math.Max(1.0, x2 - x1);
            double h = Math.Max(1.0, y2 - y1);

            try
            {
                using (var view = ed.GetCurrentView())
                {
                    view.CenterPoint = new Point2d(cx, cy);
                    view.Width = w;
                    view.Height = h;
                    ed.SetCurrentView(view);
                }
                return;
            }
            catch
            {
            }

            var ci = System.Globalization.CultureInfo.InvariantCulture;
            string center = string.Format(ci, "{0:F2},{1:F2}", cx, cy);
            string height = string.Format(ci, "{0:F2}", Math.Max(w, h));
            doc.SendStringToExecute($"_.ZOOM _C {center} {height}\n", false, false, false);
        }

        private void SaveVirtualBoundary(RoomInfo room, List<Point2d> points)
        {
            if (points == null || points.Count == 0) return;

            double minX = points.Min(p => p.X);
            double minY = points.Min(p => p.Y);
            double maxX = points.Max(p => p.X);
            double maxY = points.Max(p => p.Y);
            var ext = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));

            lock (VirtualLock)
            {
                var key = BuildRoomMatchKey(room);
                VirtualBoundsByRoomId[key] = ext;
                VirtualBoundsByRoomNumber[room.Number] = ext;
            }
        }

        private bool TryGetVirtualBoundary(string roomId, int roomNumber, out Extents3d ext)
        {
            lock (VirtualLock)
            {
                if (!string.IsNullOrWhiteSpace(roomId) && VirtualBoundsByRoomId.TryGetValue(roomId, out ext))
                    return true;
                if (roomNumber > 0 && VirtualBoundsByRoomNumber.TryGetValue(roomNumber, out ext))
                    return true;
            }

            ext = default(Extents3d);
            return false;
        }

        private void ClearVirtualBoundaries()
        {
            lock (VirtualLock)
            {
                VirtualBoundsByRoomId.Clear();
                VirtualBoundsByRoomNumber.Clear();
            }
        }

        public void DiagnoseRoomLookup(int roomNumber, string roomId = null)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            ed.WriteMessage("\n=== ZOOM DIAG ===");
            ed.WriteMessage($"\n  Input number: {roomNumber}, key: \"{roomId}\"");

            if (TryGetVirtualBoundary(roomId, roomNumber, out Extents3d vext))
            {
                ed.WriteMessage($"\n  Virtual: FOUND ext=({vext.MinPoint.X:F1},{vext.MinPoint.Y:F1})-({vext.MaxPoint.X:F1},{vext.MaxPoint.Y:F1})");
            }
            else
            {
                ed.WriteMessage("\n  Virtual: not found");
            }

            var byId = FindExistingRoomBoundariesByRoomId(db);
            if (!string.IsNullOrWhiteSpace(roomId) && byId.TryGetValue(roomId, out ObjectId idByKey))
                ed.WriteMessage($"\n  XData key: FOUND id={idByKey}");
            else
                ed.WriteMessage("\n  XData key: not found");

            var byNum = FindExistingRoomBoundaries(db);
            if (roomNumber > 0 && byNum.TryGetValue(roomNumber, out ObjectId idByNum))
                ed.WriteMessage($"\n  XData number: FOUND id={idByNum}");
            else
                ed.WriteMessage("\n  XData number: not found");

            if (!string.IsNullOrWhiteSpace(roomId) && TryParseCompositeRoomKey(roomId, out int n, out string nm))
            {
                var idByName = FindExistingRoomBoundaryByNumberAndName(db, n, nm);
                if (idByName != ObjectId.Null)
                    ed.WriteMessage($"\n  Composite fallback: FOUND id={idByName}");
                else
                    ed.WriteMessage("\n  Composite fallback: not found");
            }

            ed.WriteMessage("\n=== END DIAG ===");
        }

        public void ClearHighlight()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            try
            {
                doc.Editor.SetImpliedSelection(new ObjectId[0]);
            }
            catch { }
        }

        /// <summary>
        /// ModelSpace-Ð°Ð°Ñ NORMADVISOR XData-Ñ‚Ð°Ð¹ Ð±Ò¯Ñ… entity Ð¾Ð»Ð¾Ñ…
        /// </summary>
        public Dictionary<int, ObjectId> FindExistingRoomBoundaries(Database db)
        {
            var result = new Dictionary<int, ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    var xdata = entity.GetXDataForApplication(RegAppName);
                    if (xdata == null) continue;

                    var values = xdata.AsArray();
                    if (values.Length >= 5 && values[1].Value.ToString() == "ROOM")
                    {
                        int roomNumber = Convert.ToInt32(values[2].Value);
                        result[roomNumber] = id;
                    }
                }

                tr.Commit();
            }

            return result;
        }

        public Dictionary<string, ObjectId> FindExistingRoomBoundariesByRoomId(Database db)
        {
            var result = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    var room = ReadXData(entity);
                    if (room == null) continue;
                    if (string.IsNullOrWhiteSpace(room.RoomId)) continue;

                    result[room.RoomId] = id;
                }

                tr.Commit();
            }

            return result;
        }

        /// <summary>
        /// ÐÐ²Ñ‚Ð¾Ð¼Ð°Ñ‚ Ñ‚Ð°Ð½Ð¸Ñ…: ÑÐ¾Ð½Ð³Ð¾ÑÐ¾Ð½ Ð±Ò¯Ñ Ð´Ð¾Ñ‚Ð¾Ñ€ Ñ…Ð°Ð°Ð»Ñ‚Ñ‚Ð°Ð¹ polyline + Ñ‚ÐµÐºÑÑ‚ Ñ…Ð°Ð¹Ð¶, Ó©Ñ€Ó©Ó©Ñ‚ÑÐ¹ match Ñ…Ð¸Ð¹Ð½Ñ.
        /// boundsMin/boundsMax = null Ð±Ð¾Ð» Ð±Ò¯Ñ… ModelSpace-Ð°Ð°Ñ Ñ…Ð°Ð¹Ð½Ð°.
        /// </summary>

        private ObjectId FindExistingRoomBoundaryByNumberAndName(Database db, int roomNumber, string roomName)
        {
            if (roomNumber <= 0 || string.IsNullOrWhiteSpace(roomName))
                return ObjectId.Null;

            string nameNorm = NormalizeForMatch(roomName);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    var room = ReadXData(entity);
                    if (room == null) continue;
                    if (room.Number != roomNumber) continue;

                    string rn = NormalizeForMatch(room.Name ?? string.Empty);
                    if (rn == nameNorm)
                        return id;
                }

                tr.Commit();
            }

            return ObjectId.Null;
        }

        private bool TryParseCompositeRoomKey(string key, out int number, out string name)
        {
            number = 0;
            name = string.Empty;
            if (string.IsNullOrWhiteSpace(key)) return false;
            if (!key.StartsWith("#N", StringComparison.OrdinalIgnoreCase)) return false;

            int colon = key.IndexOf(':');
            if (colon < 0) return false;

            string numText = key.Substring(2, colon - 2);
            if (!int.TryParse(numText, out number)) return false;

            name = key.Substring(colon + 1).Trim();
            return !string.IsNullOrWhiteSpace(name);
        }

        public List<LinkResult> AutoMatchBoundaries(List<RoomInfo> rooms, Point2d? boundsMin = null, Point2d? boundsMax = null, string layerFilter = null)
        {
            var results = new List<LinkResult>();

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return results;
            var ed = doc.Editor;
            var db = doc.Database;
            ClearVirtualBoundaries();
            bool hasBounds = boundsMin.HasValue && boundsMax.HasValue;
            double bMinX = hasBounds ? boundsMin.Value.X : 0;
            double bMinY = hasBounds ? boundsMin.Value.Y : 0;
            double bMaxX = hasBounds ? boundsMax.Value.X : 0;
            double bMaxY = hasBounds ? boundsMax.Value.Y : 0;
            var allowedLayers = ParseLayerFilter(layerFilter);

            using (var lockDoc = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                EnsureRegApp(db, tr);
                EnsureLayer(db, tr);

                // Field Ñ‚Ð¾Ð¾Ð»ÑƒÑƒÑ€ Ñ†ÑÐ²ÑÑ€Ð»ÑÑ…
                _fieldTextCount = 0;
                _fieldResolvedCount = 0;

                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // 1. Ð‘Ò¯Ñ… Ñ…Ð°Ð°Ð»Ñ‚Ñ‚Ð°Ð¹ Polyline Ñ†ÑƒÐ³Ð»ÑƒÑƒÐ»Ð°Ñ… (NORMADVISOR XData-Ð³Ò¯Ð¹)
                var closedPolylines = new List<PolylineData>();
                // 2. Ð‘Ò¯Ñ… Ñ‚ÐµÐºÑÑ‚ Ñ†ÑƒÐ³Ð»ÑƒÑƒÐ»Ð°Ñ…
                var textEntities = new List<TextData>();

                foreach (ObjectId id in ms)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    // Polyline
                    if (entity is Polyline pl && pl.Closed && pl.NumberOfVertices >= 3)
                    {
                        // Layer ÑˆÒ¯Ò¯Ð»Ñ‚
                        if (!IsLayerAllowed(pl.Layer, allowedLayers))
                            continue;

                        // ÐÐ»ÑŒ Ñ…ÑÐ´Ð¸Ð¹Ð½ NORMADVISOR XData-Ñ‚Ð°Ð¹ Ð±Ð¾Ð» Ð°Ð»Ð³Ð°ÑÐ°Ñ…

                        var pts = new List<Point2d>();
                        for (int i = 0; i < pl.NumberOfVertices; i++)
                            pts.Add(pl.GetPoint2dAt(i));

                        // Ð‘Ò¯Ñ ÑˆÒ¯Ò¯Ð»Ñ‚: polyline-Ð¸Ð¹Ð½ Ð´ÑƒÐ½Ð´Ð°Ð¶ Ñ†ÑÐ³ Ð±Ò¯Ñ Ð´Ð¾Ñ‚Ð¾Ñ€ Ð±Ð°Ð¹Ñ… Ñ‘ÑÑ‚Ð¾Ð¹
                        if (hasBounds)
                        {
                            double cx = pts.Average(p => p.X);
                            double cy = pts.Average(p => p.Y);
                            if (cx < bMinX || cx > bMaxX || cy < bMinY || cy > bMaxY)
                                continue;
                        }

                        closedPolylines.Add(new PolylineData
                        {
                            ObjectId = id,
                            Points = pts,
                            Area = pl.Area
                        });
                    }
                    // DBText
                    else if (entity is DBText txt)
                    {
                        var pos = new Point2d(txt.Position.X, txt.Position.Y);
                        if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                        string content = GetTextContent(txt, tr);
                        if (string.IsNullOrEmpty(content)) continue;

                        textEntities.Add(new TextData
                        {
                            Content = content,
                            Position = pos
                        });
                    }
                    // MText
                    else if (entity is MText mtxt)
                    {
                        var pos = new Point2d(mtxt.Location.X, mtxt.Location.Y);
                        if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                        string content = GetTextContent(mtxt, tr);
                        if (string.IsNullOrEmpty(content)) continue;

                        textEntities.Add(new TextData
                        {
                            Content = content,
                            Position = pos
                        });
                    }
                    // BlockReference â€” Ð´Ð¾Ñ‚Ð¾Ñ€ Polyline + Ñ‚ÐµÐºÑÑ‚ Ñ…Ð°Ð¹Ñ…
                    else if (entity is BlockReference blkRef)
                    {
                        ScanBlockReference(blkRef, tr, hasBounds, bMinX, bMinY, bMaxX, bMaxY,
                            closedPolylines, textEntities, layerFilter);
                    }
                }

                string layerInfo = !string.IsNullOrEmpty(layerFilter) ? $", layers=\"{layerFilter}\"" : "";
                string boundsInfo = hasBounds ? $" (selected bounds{layerInfo})" : $" (all drawing{layerInfo})";
                ed.WriteMessage($"\n  Scan: {closedPolylines.Count} polylines, {textEntities.Count} texts{boundsInfo}");

                // DEBUG: Ñ‚ÐµÐºÑÑ‚ Ð¶Ð¸ÑˆÑÑ
                if (textEntities.Count > 0)
                {
                    int showCount = System.Math.Min(textEntities.Count, 10);
                    ed.WriteMessage($"\n  Text samples ({textEntities.Count} total):");
                    for (int i = 0; i < showCount; i++)
                        ed.WriteMessage($"\n    \"{textEntities[i].Content}\"");
                }
                else
                {
                    ed.WriteMessage("\n  WARNING: No text found.");
                }

                if (_fieldTextCount > 0)
                    ed.WriteMessage($"\n  Field text: found={_fieldTextCount}, resolved={_fieldResolvedCount}");

                // â”€â”€ ÐÑÐ³Ð¶ Ñ‚Ð¾Ð´Ð¾Ñ€Ñ…Ð¾Ð¹Ð»Ð¾Ñ… â”€â”€
                double polyAvgArea = closedPolylines.Any() ? closedPolylines.Average(p => p.Area) : 1;
                double roomAvgArea = rooms.Where(r => r.Area > 0).Select(r => r.Area).DefaultIfEmpty(1).Average();
                double areaScale = 1.0;
                if (roomAvgArea > 0 && polyAvgArea / roomAvgArea > 500)
                    areaScale = 1000000.0;
                ed.WriteMessage($"\n  Area scale: {areaScale:F0} (poly avg={polyAvgArea:F1}, room avg={roomAvgArea:F1})");

                // â”€â”€ Matching Ð±ÑÐ»Ñ‚Ð³ÑÐ» â”€â”€
                int matched = 0;
                int layerMatched = 0;
                bool msUpgraded = false;
                // Always rematch all rooms on each run.
                var unmatchedRooms = rooms.ToList();
                ed.WriteMessage($"\n  Rooms pending: {unmatchedRooms.Count}/{rooms.Count}");
                var availablePolys = new List<PolylineData>(closedPolylines);

                // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
                // PASS 1: Ð¢ÐµÐºÑÑ‚-Ð´Ð¾Ñ‚Ð¾Ñ€-polyline matching
                // Polyline Ð±Ò¯Ñ€Ð¸Ð¹Ð½ Ð´Ð¾Ñ‚Ð¾Ñ€ Ð±Ð°Ð¹Ð³Ð°Ð° Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ Ð¾Ð»Ð¶, Ó©Ñ€Ó©Ó©Ñ‚ÑÐ¹ Ñ‚Ð¾Ñ…Ð¸Ñ€ÑƒÑƒÐ»Ð½Ð°
                // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
                // ═══════════════════════════════════════════════
                // PASS 0: Layer-based matching (100% accurate)
                // NORM_ROOM_AREA layer -> polyline
                // NORM_ROOM_NUM layer -> text (дугаар)
                // ═══════════════════════════════════════════════
                {
                    var roomAreaPolys = new List<PolylineData>();
                    var roomNumTexts = new List<TextData>();

                    // NORM_ROOM_AREA layer дээрх polyline цуглуулах
                    foreach (var poly in closedPolylines)
                    {
                        var ent = tr.GetObject(poly.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent != null && string.Equals(ent.Layer, RoomAreaLayerName, StringComparison.OrdinalIgnoreCase))
                            roomAreaPolys.Add(poly);
                    }

                    // NORM_ROOM_NUM layer дээрх текст DB-с скан хийх
                    foreach (ObjectId id in ms)
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity == null) continue;
                        if (!string.Equals(entity.Layer, RoomNumLayerName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        string content = null;
                        Point2d pos = default;

                        if (entity is DBText dtxt)
                        {
                            content = GetTextContent(dtxt, tr);
                            pos = new Point2d(dtxt.Position.X, dtxt.Position.Y);
                        }
                        else if (entity is MText mtxt2)
                        {
                            content = GetTextContent(mtxt2, tr);
                            pos = new Point2d(mtxt2.Location.X, mtxt2.Location.Y);
                        }

                        if (!string.IsNullOrEmpty(content))
                        {
                            if (!hasBounds || IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY))
                                roomNumTexts.Add(new TextData { Content = content, Position = pos });
                        }
                    }

                    if (roomAreaPolys.Count > 0 && roomNumTexts.Count > 0)
                    {
                        ed.WriteMessage($"\n\n  -- PASS 0: layer-based matching ({RoomAreaLayerName}: {roomAreaPolys.Count} polys, {RoomNumLayerName}: {roomNumTexts.Count} texts) --");

                        var matchedRoomsP0 = new HashSet<RoomInfo>();
                        var matchedPolysP0 = new HashSet<PolylineData>();

                        foreach (var poly in roomAreaPolys)
                        {
                            var insideTexts = roomNumTexts.Where(t => IsPointInsidePolygon(t.Position, poly.Points)).ToList();
                            if (insideTexts.Count == 0) continue;

                            string numContent = insideTexts[0].Content.Trim();

                            // 1) RoomId-р хайх (a1, b2)
                            RoomInfo matchedRoom = unmatchedRooms.FirstOrDefault(r =>
                                !matchedRoomsP0.Contains(r) &&
                                !string.IsNullOrWhiteSpace(r.RoomId) &&
                                string.Equals(r.RoomId.Trim(), numContent, StringComparison.OrdinalIgnoreCase));

                            // 2) Дугаараар хайх
                            if (matchedRoom == null && int.TryParse(numContent, out int parsedNum))
                            {
                                matchedRoom = unmatchedRooms.FirstOrDefault(r =>
                                    !matchedRoomsP0.Contains(r) && r.Number == parsedNum);
                            }

                            // 3) Нэрээр хайх
                            if (matchedRoom == null)
                            {
                                matchedRoom = unmatchedRooms.FirstOrDefault(r =>
                                    !matchedRoomsP0.Contains(r) &&
                                    !string.IsNullOrWhiteSpace(r.Name) &&
                                    string.Equals(r.Name.Trim(), numContent, StringComparison.OrdinalIgnoreCase));
                            }

                            if (matchedRoom == null) continue;

                            var linkResult = LinkPolyToRoom(poly, matchedRoom, tr, db, ms, ref msUpgraded, areaScale);
                            results.Add(linkResult);
                            matchedRoomsP0.Add(matchedRoom);
                            matchedPolysP0.Add(poly);
                            availablePolys.Remove(poly);
                            matched++;
                            layerMatched++;

                            ed.WriteMessage($"\n    LAYER \"{numContent}\" -> {matchedRoom.Number}. {matchedRoom.Name} ({linkResult.DrawnArea:F2} m2)");
                        }

                        unmatchedRooms.RemoveAll(r => matchedRoomsP0.Contains(r));
                        ed.WriteMessage($"\n    Layer matched: {layerMatched}/{rooms.Count}");

                        if (unmatchedRooms.Count == 0)
                        {
                            var layerTable2 = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                            if (layerTable2.Has(LayerName))
                            {
                                var layer2 = (LayerTableRecord)tr.GetObject(layerTable2[LayerName], OpenMode.ForWrite);
                                layer2.Color = Color.FromColorIndex(ColorMethod.ByAci, 251);
                                layer2.IsPlottable = false;
                            }
                            tr.Commit();
                            ed.WriteMessage($"\n\n  === Result: {matched}/{rooms.Count} rooms (layer:{layerMatched}) - ALL MATCHED ===");
                            return results;
                        }
                    }
                    else if (roomAreaPolys.Count > 0 || roomNumTexts.Count > 0)
                    {
                        ed.WriteMessage($"\n\n  -- PASS 0: skipped ({RoomAreaLayerName}: {roomAreaPolys.Count}, {RoomNumLayerName}: {roomNumTexts.Count} - both needed) --");
                    }
                }

                // ═══════════════════════════════════════════════
                // PASS 1: text-inside-polyline matching (fallback)
                // ═══════════════════════════════════════════════
                ed.WriteMessage($"\n\n  -- PASS 1: text matching --");
                int textMatched = 0;

                if (textEntities.Count > 0 && availablePolys.Count > 0)
                {
                    // Polyline Ð±Ò¯Ñ€Ð´ ÑÐ¼Ð°Ñ€ Ñ‚ÐµÐºÑÑ‚ Ð´Ð¾Ñ‚Ð¾Ñ€ Ð½ÑŒ Ð±Ð°Ð¹Ð³Ð°Ð°Ð³ Ð¾Ð»Ð¾Ñ…
                    var polyTexts = new System.Collections.Generic.Dictionary<PolylineData, List<TextData>>();
                    foreach (var poly in availablePolys)
                    {
                        var insideTexts = new List<TextData>();
                        foreach (var txt in textEntities)
                        {
                            if (IsPointInsidePolygon(txt.Position, poly.Points))
                                insideTexts.Add(txt);
                        }
                        if (insideTexts.Count > 0)
                            polyTexts[poly] = insideTexts;
                    }

                    ed.WriteMessage($"\n    Text-inside-polyline: {polyTexts.Count}/{availablePolys.Count}");

                    // DEBUG: ÑÑ…Ð½Ð¸Ð¹ Ñ…ÑÐ´ÑÐ½ polyline-Ð¸Ð¹Ð½ Ð´Ð¾Ñ‚Ð¾Ñ€ Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ Ñ…Ð°Ñ€ÑƒÑƒÐ»Ð°Ñ…
                    int dbgCount = 0;
                    foreach (var kv in polyTexts)
                    {
                        if (dbgCount >= 5) break;
                        string txts = string.Join(", ", kv.Value.Select(t => $"\"{t.Content}\""));
                        ed.WriteMessage($"\n      PL(area={kv.Key.Area / areaScale:F2}): {txts}");
                        dbgCount++;
                    }

                    // Ó¨Ñ€Ó©Ó© Ð±Ò¯Ñ€Ñ‚ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… polyline Ð¾Ð»Ð¾Ñ…
                    var matchedRooms = new HashSet<RoomInfo>();
                    var matchedPolySet = new HashSet<PolylineData>();

                    // PASS 1A: RoomId anchor match (strongest rule)
                    foreach (var room in unmatchedRooms.ToList())
                    {
                        if (string.IsNullOrWhiteSpace(room.RoomId)) continue;

                        PolylineData bestAnchorPoly = null;
                        double bestDiff = double.MaxValue;

                        foreach (var kv in polyTexts)
                        {
                            if (matchedPolySet.Contains(kv.Key)) continue;

                            bool hasAnchor = kv.Value.Any(t =>
                                string.Equals((t.Content ?? string.Empty).Trim(), room.RoomId, StringComparison.OrdinalIgnoreCase) ||
                                (t.Content ?? string.Empty).Trim().StartsWith(room.RoomId + ".", StringComparison.OrdinalIgnoreCase) ||
                                (t.Content ?? string.Empty).Trim().StartsWith(room.RoomId + " ", StringComparison.OrdinalIgnoreCase));
                            if (!hasAnchor) continue;

                            double diff = room.Area > 0
                                ? Math.Abs((kv.Key.Area / areaScale) - room.Area) / room.Area
                                : 0;

                            if (diff < bestDiff)
                            {
                                bestDiff = diff;
                                bestAnchorPoly = kv.Key;
                            }
                        }

                        if (bestAnchorPoly == null) continue;

                        var anchorResult = LinkPolyToRoom(bestAnchorPoly, room, tr, db, ms, ref msUpgraded, areaScale);
                        results.Add(anchorResult);
                        matchedRooms.Add(room);
                        matchedPolySet.Add(bestAnchorPoly);
                        availablePolys.Remove(bestAnchorPoly);
                        matched++;
                        textMatched++;

                        string src = bestAnchorPoly.IsFromBlock ? " [block]" : "";
                        ed.WriteMessage($"\n    ANCHOR {room.RoomId} {room.Name} -> {anchorResult.DrawnArea:F2} m2{src}");
                    }

                    // PASS 1B: Exact numeric anchor for rooms without RoomId (e.g. 4, 5, 11).
                    foreach (var room in unmatchedRooms.ToList())
                    {
                        if (matchedRooms.Contains(room)) continue;
                        if (!string.IsNullOrWhiteSpace(room.RoomId)) continue;

                        string numText = room.Number.ToString();
                        PolylineData bestNumPoly = null;
                        int bestExactCount = -1;
                        int bestPureNumberCount = int.MaxValue;
                        double bestDiff = double.MaxValue;

                        foreach (var kv in polyTexts)
                        {
                            if (matchedPolySet.Contains(kv.Key)) continue;

                            int exactCount = kv.Value.Count(t =>
                                string.Equals((t.Content ?? string.Empty).Trim(), numText, StringComparison.OrdinalIgnoreCase));
                            if (exactCount <= 0) continue;

                            int pureNumberCount = kv.Value.Count(t =>
                                Regex.IsMatch((t.Content ?? string.Empty).Trim(), @"^\d+$"));

                            double diff = room.Area > 0
                                ? Math.Abs((kv.Key.Area / areaScale) - room.Area) / room.Area
                                : 0;

                            bool better = false;
                            if (exactCount > bestExactCount) better = true;
                            else if (exactCount == bestExactCount && pureNumberCount < bestPureNumberCount) better = true;
                            else if (exactCount == bestExactCount && pureNumberCount == bestPureNumberCount && diff < bestDiff) better = true;

                            if (better)
                            {
                                bestNumPoly = kv.Key;
                                bestExactCount = exactCount;
                                bestPureNumberCount = pureNumberCount;
                                bestDiff = diff;
                            }
                        }

                        if (bestNumPoly == null) continue;

                        var numResult = LinkPolyToRoom(bestNumPoly, room, tr, db, ms, ref msUpgraded, areaScale);
                        results.Add(numResult);
                        matchedRooms.Add(room);
                        matchedPolySet.Add(bestNumPoly);
                        availablePolys.Remove(bestNumPoly);
                        matched++;
                        textMatched++;

                        string srcN = bestNumPoly.IsFromBlock ? " [block]" : "";
                        ed.WriteMessage($"\n    ANCHOR-N {room.Number} {room.Name} -> {numResult.DrawnArea:F2} m2{srcN}");
                    }

                    // PASS 1C: generic text scoring fallback.
                    foreach (var room in unmatchedRooms.ToList())
                    {
                        if (matchedRooms.Contains(room)) continue;
                        PolylineData bestPoly = null;
                        int bestScore = 0; // Ð˜Ð»Ò¯Ò¯ Ð¾Ð»Ð¾Ð½ Ñ‚ÐµÐºÑÑ‚ Ñ‚Ð¾Ñ…Ð¸Ñ€ÑÐ¾Ð½ Ð½ÑŒ Ð´ÑÑÑ€

                        foreach (var kv in polyTexts)
                        {
                            if (matchedPolySet.Contains(kv.Key)) continue;

                            int score = ScoreRoomToPolyline(room, kv.Value);
                            if (score > bestScore)
                            {
                                bestScore = score;
                                bestPoly = kv.Key;
                            }
                            else if (score > 0 && score == bestScore && bestPoly != null)
                            {
                                double curDiff = room.Area > 0
                                    ? Math.Abs((kv.Key.Area / areaScale) - room.Area) / room.Area
                                    : double.MaxValue;
                                double bestDiff = room.Area > 0
                                    ? Math.Abs((bestPoly.Area / areaScale) - room.Area) / room.Area
                                    : double.MaxValue;
                                if (curDiff < bestDiff)
                                    bestPoly = kv.Key;
                            }
                        }

                        if (bestPoly != null && bestScore >= 120)
                        {
                            var linkResult = LinkPolyToRoom(bestPoly, room, tr, db, ms, ref msUpgraded, areaScale);
                            results.Add(linkResult);
                            matchedRooms.Add(room);
                            matchedPolySet.Add(bestPoly);
                            availablePolys.Remove(bestPoly);
                            matched++;
                            textMatched++;

                            string src = bestPoly.IsFromBlock ? " [Ð±Ð»Ð¾Ðº]" : "";
                            ed.WriteMessage($"\n    OK {room.RoomId} {room.Name} -> {linkResult.DrawnArea:F2} m2{src}");
                        }
                    }

                    // unmatchedRooms ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
                    unmatchedRooms.RemoveAll(r => matchedRooms.Contains(r));
                }

                ed.WriteMessage($"\n    Text matched: {textMatched}");

                // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
                // PASS 2: Area-Ð°Ð°Ñ€ greedy matching (Ò¯Ð»Ð´ÑÑÐ½ Ó©Ñ€Ó©Ó©Ð½Ò¯Ò¯Ð´ÑÐ´)
                // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
                int areaMatched = 0;
                if (unmatchedRooms.Count > 0 && availablePolys.Count > 0)
                {
                    ed.WriteMessage($"\n\n  -- PASS 2: area fallback ({unmatchedRooms.Count} rooms <-> {availablePolys.Count} polylines) --");

                    var candidates = new List<Tuple<RoomInfo, PolylineData, double>>();
                    foreach (var room in unmatchedRooms)
                    {
                        if (room.Area <= 0) continue;
                        foreach (var poly in availablePolys)
                        {
                            double polyAreaM2 = poly.Area / areaScale;
                            double diff = Math.Abs(polyAreaM2 - room.Area) / room.Area;
                            candidates.Add(Tuple.Create(room, poly, diff));
                        }
                    }

                    candidates.Sort((a, b) => a.Item3.CompareTo(b.Item3));

                    var usedRooms = new HashSet<RoomInfo>();
                    var usedPolys = new HashSet<PolylineData>();

                    foreach (var cand in candidates)
                    {
                        if (usedRooms.Contains(cand.Item1)) continue;
                        if (usedPolys.Contains(cand.Item2)) continue;
                        // 200% Ð·Ó©Ñ€Ò¯Ò¯Ñ‚ÑÐ¹ Ð±Ð¾Ð» Ð°Ð»Ð³Ð°ÑÐ½Ð° (Ð´ÑÐ½Ð´Ò¯Ò¯ Ð¸Ñ… Ð·Ó©Ñ€Ò¯Ò¯Ñ‚ÑÐ¹ match Ñ…Ð¸Ð¹Ñ…Ð³Ò¯Ð¹)
                        if (cand.Item3 > 0.35) continue;

                        var linkResult = LinkPolyToRoom(cand.Item2, cand.Item1, tr, db, ms, ref msUpgraded, areaScale);
                        results.Add(linkResult);
                        usedRooms.Add(cand.Item1);
                        usedPolys.Add(cand.Item2);
                        matched++;
                        areaMatched++;

                        double pct = cand.Item3 * 100;
                        string src = cand.Item2.IsFromBlock ? " [Ð±Ð»Ð¾Ðº]" : "";
                        ed.WriteMessage($"\n    OK {cand.Item1.RoomId} {cand.Item1.Name} ({cand.Item1.Area:F2}) -> {linkResult.DrawnArea:F2} m2 (diff:{pct:F0}%){src}");
                    }

                    ed.WriteMessage($"\n    Area matched: {areaMatched}");
                }

                // Layer-Ð¸Ð¹Ð³ Ñ…Ð°Ñ€Ð°Ð³Ð´Ð°Ñ…Ð³Ò¯Ð¹ Ð±Ð¾Ð»Ð³Ð¾Ñ…
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (layerTable.Has(LayerName))
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerTable[LayerName], OpenMode.ForWrite);
                    layer.Color = Color.FromColorIndex(ColorMethod.ByAci, 251);
                    layer.IsPlottable = false;
                }

                tr.Commit();
                ed.WriteMessage($"\n\n  === Result: {matched}/{rooms.Count} rooms (layer:{layerMatched}, text:{textMatched}, area:{areaMatched}) ===");
                int notMatched = rooms.Count(r => !r.HasBoundary) - matched;
                if (notMatched > 0)
                    ed.WriteMessage($"\n  Unmatched rooms: {notMatched} (link manually)");
            }

            return results;
        }

        /// <summary>
        /// Polyline-Ð³ Ó©Ñ€Ó©Ó©Ð½Ð´ Ñ…Ð¾Ð»Ð±Ð¾Ð¶, XData Ñ…Ð°Ð²ÑÐ°Ñ€Ð³Ð°Ñ… helper
        /// </summary>
        private LinkResult LinkPolyToRoom(PolylineData poly, RoomInfo room,
            Transaction tr, Database db, BlockTableRecord ms, ref bool msUpgraded, double areaScale)
        {
            ObjectId resultPolyId;
            double area;

            if (poly.IsFromBlock)
            {
                // Virtual boundary mode: no extra polyline creation.
                SaveVirtualBoundary(room, poly.Points);
                resultPolyId = poly.ObjectId;
                area = poly.Area / areaScale;
            }
            else
            {
                var polyEntity = (Entity)tr.GetObject(poly.ObjectId, OpenMode.ForWrite);
                polyEntity.Layer = LayerName;
                AttachXData(polyEntity, tr, db, room);

                resultPolyId = poly.ObjectId;
                area = 0;
                if (polyEntity is Polyline matchedPl) area = matchedPl.Area / areaScale;
            }

            return new LinkResult
            {
                PolylineId = resultPolyId,
                DrawnArea = area,
                RoomNumber = room.Number,
                RoomId = BuildRoomMatchKey(room)
            };
        }

        private bool TryFindExistingBoundaryId(Transaction tr, BlockTableRecord ms, RoomInfo room, out ObjectId boundaryId)
        {
            boundaryId = ObjectId.Null;

            foreach (ObjectId id in ms)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                var x = ReadXData(ent);
                if (x == null) continue;

                if (!string.IsNullOrWhiteSpace(room.RoomId) &&
                    string.Equals(x.RoomId, room.RoomId, StringComparison.OrdinalIgnoreCase))
                {
                    boundaryId = id;
                    return true;
                }

                if (string.IsNullOrWhiteSpace(room.RoomId) && x.Number == room.Number)
                {
                    boundaryId = id;
                    return true;
                }
            }

            return false;
        }

        private void ReplacePolylineGeometry(Polyline pl, List<Point2d> points)
        {
            while (pl.NumberOfVertices > 0)
                pl.RemoveVertexAt(pl.NumberOfVertices - 1);

            for (int i = 0; i < points.Count; i++)
                pl.AddVertexAt(i, points[i], 0, 0, 0);

            pl.Closed = true;
        }

        /// <summary>
        /// Ó¨Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ð´ÑƒÐ³Ð°Ð°Ñ€/Ð½ÑÑ€ÑÐ½Ð´ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… Ñ‚ÐµÐºÑÑ‚ Ð¾Ð»Ð¾Ñ…
        /// Ð”ÑƒÐ³Ð°Ð°Ñ€ (RoomId, Number) â†’ ÐÑÑ€ÑÑÑ€ (ÑÐ³ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… â†’ Ð°Ð³ÑƒÑƒÐ»ÑÐ°Ð½ â†’ normalize) Ð³ÑÑÑÐ½ Ð´Ð°Ñ€Ð°Ð°Ð»Ð»Ð°Ð°Ñ€
        /// </summary>
        private List<TextData> FindMatchingTexts(RoomInfo room, List<TextData> texts)
        {
            var result = new List<TextData>();
            string numStr = room.Number.ToString();
            string roomId = room.RoomId ?? "";

            // ÐÑÑ€Ð¸Ð¹Ð³ normalize Ñ…Ð¸Ð¹Ñ… (Ð·Ð°Ð¹, Ð·ÑƒÑ€Ð°Ð°Ñ, Ñ‚Ð¾Ð¼/Ð¶Ð¸Ð¶Ð¸Ð³)
            string roomNameNorm = NormalizeForMatch(room.Name);
            string sectionNorm = NormalizeForMatch(room.SectionName);

            foreach (var txt in texts)
            {
                string content = txt.Content;

                // RoomId-ÑÑÑ€ ÑÐ³ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… (Ð¶Ð¸ÑˆÑÑ: "a1", "b2", "d3")
                if (!string.IsNullOrEmpty(roomId) &&
                    string.Equals(content, roomId, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(txt);
                    continue;
                }

                // Ð¯Ð³ Ð´ÑƒÐ³Ð°Ð°Ñ€ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… (Ð¶Ð¸ÑˆÑÑ: "1", "2", "15")
                if (content == numStr)
                {
                    result.Add(txt);
                    continue;
                }

                // Ð”ÑƒÐ³Ð°Ð°Ñ€ + Ñ†ÑÐ³/Ð·Ð°Ð¹ (Ð¶Ð¸ÑˆÑÑ: "1.", "1. Ð¢Ð°Ð¼Ð±ÑƒÑ€")
                if (content.StartsWith(numStr + ".") || content.StartsWith(numStr + " "))
                {
                    result.Add(txt);
                    continue;
                }

                // RoomId + Ñ†ÑÐ³/Ð·Ð°Ð¹ (Ð¶Ð¸ÑˆÑÑ: "a1.", "a1 Ð¢Ð°Ð¼Ð±ÑƒÑ€")
                if (!string.IsNullOrEmpty(roomId) &&
                    (content.StartsWith(roomId + ".", StringComparison.OrdinalIgnoreCase) ||
                     content.StartsWith(roomId + " ", StringComparison.OrdinalIgnoreCase)))
                {
                    result.Add(txt);
                    continue;
                }

                // ÐÑÑ€ÑÑÑ€ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… (Ð¶Ð¸ÑˆÑÑ: "Ð¢Ð°Ð¼Ð±ÑƒÑ€", "ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿", "Ð-ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿")
                if (!string.IsNullOrEmpty(room.Name) && room.Name.Length > 2)
                {
                    // 1. Ð¯Ð³ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… (case-insensitive)
                    if (string.Equals(content, room.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(txt);
                        continue;
                    }

                    // 2. Ð¢ÐµÐºÑÑ‚ Ð½ÑŒ Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ð½ÑÑ€Ð¸Ð¹Ð³ Ð°Ð³ÑƒÑƒÐ»ÑÐ°Ð½ (Ð¶Ð¸ÑˆÑÑ: "Ð-ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿ Ð·Ð°Ð»" Ð°Ð³ÑƒÑƒÐ»ÑÐ°Ð½ "ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿")
                    if (content.IndexOf(room.Name, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        result.Add(txt);
                        continue;
                    }

                    // 3. Ó¨Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ð½ÑÑ€ Ð½ÑŒ Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ Ð°Ð³ÑƒÑƒÐ»ÑÐ°Ð½ (Ð¶Ð¸ÑˆÑÑ: Ð½ÑÑ€="ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿ Ð·Ð°Ð»", Ñ‚ÐµÐºÑÑ‚="ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿")
                    if (content.Length > 2 &&
                        room.Name.IndexOf(content, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        result.Add(txt);
                        continue;
                    }

                    // 4. Normalize Ñ…Ð¸Ð¹ÑÑÐ½ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… (Ð·ÑƒÑ€Ð°Ð°Ñ, Ð·Ð°Ð¹ Ð·ÑÑ€Ð³Ð¸Ð¹Ð³ Ð°Ñ€Ð¸Ð»Ð³Ð°Ð¶ Ñ…Ð°Ñ€ÑŒÑ†ÑƒÑƒÐ»Ð°Ñ…)
                    //    "Ð-ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿" vs "Ð ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿" vs "ÐÐºÐ¾Ñ„ÐµÑˆÐ¾Ð¿"
                    if (!string.IsNullOrEmpty(roomNameNorm) && roomNameNorm.Length > 2)
                    {
                        string contentNorm = NormalizeForMatch(content);
                        if (!string.IsNullOrEmpty(contentNorm) && contentNorm.Length > 2)
                        {
                            if (contentNorm.Contains(roomNameNorm) || roomNameNorm.Contains(contentNorm))
                            {
                                result.Add(txt);
                                continue;
                            }
                        }
                    }
                }

                // 5. Ð¡ÐµÐºÑ†/Ð±Ò¯Ð»Ð³Ð¸Ð¹Ð½ Ð½ÑÑ€ÑÑÑ€ Ñ‚Ð¾Ñ…Ð¸Ñ€Ð¾Ñ… (Ð¶Ð¸ÑˆÑÑ: "Ð-ÐšÐ¾Ñ„ÐµÑˆÐ¾Ð¿" ÑÐµÐºÑ†Ð¸Ð¹Ð½ Ð½ÑÑ€)
                if (!string.IsNullOrEmpty(room.SectionName) && room.SectionName.Length > 2)
                {
                    if (string.Equals(content, room.SectionName, StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(txt);
                        continue;
                    }
                    if (content.IndexOf(room.SectionName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        room.SectionName.IndexOf(content, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        result.Add(txt);
                        continue;
                    }
                    if (!string.IsNullOrEmpty(sectionNorm) && sectionNorm.Length > 2)
                    {
                        string contentNorm = NormalizeForMatch(content);
                        if (!string.IsNullOrEmpty(contentNorm) && contentNorm.Length > 2 &&
                            (contentNorm.Contains(sectionNorm) || sectionNorm.Contains(contentNorm)))
                        {
                            result.Add(txt);
                            continue;
                        }
                    }
                }
            }

            return result;
        }
        private int ScoreRoomToPolyline(RoomInfo room, List<TextData> texts)
        {
            int best = 0;
            foreach (var txt in texts)
            {
                int score = ScoreTextMatch(room, txt.Content);
                if (score > best) best = score;
            }
            return best;
        }

        private int ScoreTextMatch(RoomInfo room, string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return 0;

            string text = content.Trim();
            string numStr = room.Number.ToString();
            string roomId = room.RoomId ?? string.Empty;
            string roomName = room.Name ?? string.Empty;
            string sectionName = room.SectionName ?? string.Empty;
            bool roomIdHasLetter = !string.IsNullOrEmpty(roomId) && Regex.IsMatch(roomId, "[A-Za-zА-Яа-я]");
            bool textIsPureNumber = Regex.IsMatch(text, @"^\d+$");

            if (!string.IsNullOrEmpty(roomId))
            {
                if (string.Equals(text, roomId, StringComparison.OrdinalIgnoreCase))
                    return 1000;

                if (text.StartsWith(roomId + ".", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith(roomId + " ", StringComparison.OrdinalIgnoreCase))
                    return 900;
            }

            if (!string.IsNullOrEmpty(roomName) && roomName.Length > 2)
            {
                if (string.Equals(text, roomName, StringComparison.OrdinalIgnoreCase))
                    return 700;

                if (text.IndexOf(roomName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    roomName.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                    return 550;

                string n1 = NormalizeForMatch(text);
                string n2 = NormalizeForMatch(roomName);
                if (!string.IsNullOrEmpty(n1) && !string.IsNullOrEmpty(n2) &&
                    (n1.Contains(n2) || n2.Contains(n1)))
                    return 500;
            }

            if (!string.IsNullOrEmpty(sectionName) && sectionName.Length > 2)
            {
                if (string.Equals(text, sectionName, StringComparison.OrdinalIgnoreCase))
                    return 300;

                if (text.IndexOf(sectionName, StringComparison.OrdinalIgnoreCase) >= 0)
                    return 250;
            }

            if (!roomIdHasLetter)
            {
                if (text == numStr)
                    return 80;

                if (text.StartsWith(numStr + ".") || text.StartsWith(numStr + " "))
                    return 70;

                if (Regex.IsMatch(text, $@"\b{Regex.Escape(numStr)}\b"))
                    return 60;
            }
            else if (textIsPureNumber)
            {
                return 0;
            }

            return 0;
        }


        /// <summary>
        /// Ð¢ÐµÐºÑÑ‚Ð¸Ð¹Ð³ match Ñ…Ð¸Ð¹Ñ…ÑÐ´ Ð½Ð¾Ñ€Ð¼Ð°Ð»Ð¸Ð·Ð°Ñ†Ð¸ Ñ…Ð¸Ð¹Ñ…
        /// Ð—Ð°Ð¹, Ð·ÑƒÑ€Ð°Ð°Ñ, Ñ†ÑÐ³ Ð·ÑÑ€ÑÐ³ Ñ‚ÑÐ¼Ð´ÑÐ³Ñ‚Ò¯Ò¯Ð´Ð¸Ð¹Ð³ Ð°Ñ€Ð¸Ð»Ð³Ð°Ð¶, Ð¶Ð¸Ð¶Ð¸Ð³ Ò¯ÑÑÐ³ Ð±Ð¾Ð»Ð³Ð¾Ð½Ð¾
        /// </summary>
        private string NormalizeForMatch(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("-", "")
                .Replace("â€“", "")
                .Replace(" ", "")
                .Replace(".", "")
                .Replace(",", "")
                .ToLowerInvariant();
        }

        private HashSet<string> ParseLayerFilter(string layerFilter)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(layerFilter)) return result;

            var parts = layerFilter.Split(new[] { '|', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var name = p.Trim();
                if (!string.IsNullOrEmpty(name)) result.Add(name);
            }
            return result;
        }

        private bool IsLayerAllowed(string layerName, HashSet<string> allowedLayers)
        {
            if (allowedLayers == null || allowedLayers.Count == 0) return true;
            if (string.IsNullOrWhiteSpace(layerName)) return false;
            return allowedLayers.Contains(layerName.Trim());
        }


        /// <summary>
        /// Ð¦ÑÐ³ Ñ‚ÑÐ³Ñˆ Ó©Ð½Ñ†Ó©Ð³Ñ‚ Ð±Ò¯Ñ Ð´Ð¾Ñ‚Ð¾Ñ€ Ð±Ð°Ð¹Ð³Ð°Ð° ÑÑÑÑ… (Point2d)
        /// </summary>
        private bool IsInBounds(Point2d pos, double minX, double maxX, double minY, double maxY)
        {
            return pos.X >= minX && pos.X <= maxX && pos.Y >= minY && pos.Y <= maxY;
        }

        /// <summary>
        /// Ray casting - Ñ†ÑÐ³ polygon Ð´Ð¾Ñ‚Ð¾Ñ€ Ð±Ð°Ð¹Ð³Ð°Ð° ÑÑÑÑ…
        /// </summary>
        private bool IsPointInsidePolygon(Point2d point, List<Point2d> polygon)
        {
            bool inside = false;
            int n = polygon.Count;

            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                double xi = polygon[i].X, yi = polygon[i].Y;
                double xj = polygon[j].X, yj = polygon[j].Y;

                if (((yi > point.Y) != (yj > point.Y)) &&
                    (point.X < (xj - xi) * (point.Y - yi) / (yj - yi) + xi))
                {
                    inside = !inside;
                }
            }

            return inside;
        }

        /// <summary>
        /// BlockReference Ð´Ð¾Ñ‚Ð¾Ñ€ polyline + Ñ‚ÐµÐºÑÑ‚ Ñ…Ð°Ð¹Ð¶ Ñ†ÑƒÐ³Ð»ÑƒÑƒÐ»Ð°Ñ….
        /// Ð‘Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€Ñ…Ð¸ entity-Ð¸Ð¹Ð½ ÐºÐ¾Ð¾Ñ€Ð´Ð¸Ð½Ð°Ñ‚ÑƒÑƒÐ´Ñ‹Ð³ BlockTransform Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½ world ÐºÐ¾Ð¾Ñ€Ð´Ð¸Ð½Ð°Ñ‚ Ñ€ÑƒÑƒ Ñ…Ó©Ñ€Ð²Ò¯Ò¯Ð»Ð½Ñ.
        /// </summary>
        private void ScanBlockReference(BlockReference blkRef, Transaction tr,
            bool hasBounds, double bMinX, double bMinY, double bMaxX, double bMaxY,
            List<PolylineData> closedPolylines, List<TextData> textEntities, string layerFilter = null)
        {
            Matrix3d transform = blkRef.BlockTransform;
            string blkRefLayer = blkRef.Layer; // Block ref-Ð¸Ð¹Ð½ Ó©Ó©Ñ€Ð¸Ð¹Ð½Ñ… Ð½ÑŒ layer (layer "0" entity-Ð´ Ó©Ð²Ð»Ó©Ð³Ð´Ó©Ð½Ó©)
            var allowedLayers = ParseLayerFilter(layerFilter);

            // AttributeReference Ñ‚ÐµÐºÑÑ‚ (Ð±Ð»Ð¾ÐºÐ¸Ð¹Ð½ attribute ÑƒÑ‚Ð³ÑƒÑƒÐ´)
            foreach (ObjectId attId in blkRef.AttributeCollection)
            {
                try
                {
                    var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (att == null) continue;

                    string content = GetTextContent(att, tr);
                    if (string.IsNullOrEmpty(content)) continue;

                    var pos = new Point2d(att.Position.X, att.Position.Y);
                    if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                    textEntities.Add(new TextData { Content = content, Position = pos });
                }
                catch { }
            }

            // Block definition Ð´Ð¾Ñ‚Ð¾Ñ€ entity Ñ…Ð°Ð¹Ñ…
            var btr = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return;

            foreach (ObjectId entId in btr)
            {
                Entity ent;
                try { ent = tr.GetObject(entId, OpenMode.ForRead) as Entity; }
                catch { continue; }
                if (ent == null) continue;

                // Polyline (Ð±Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€)
                if (ent is Polyline pl && pl.Closed && pl.NumberOfVertices >= 3)
                {
                    string effectiveLayer = string.Equals(pl.Layer, "0", StringComparison.OrdinalIgnoreCase)
                        ? blkRefLayer : pl.Layer;
                    if (!IsLayerAllowed(effectiveLayer, allowedLayers))
                        continue;


                    // Vertex-Ò¯Ò¯Ð´Ð¸Ð¹Ð³ world ÐºÐ¾Ð¾Ñ€Ð´Ð¸Ð½Ð°Ñ‚ Ñ€ÑƒÑƒ Ñ…Ó©Ñ€Ð²Ò¯Ò¯Ð»ÑÑ…
                    var pts = new List<Point2d>();
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        var localPt = pl.GetPoint3dAt(i);
                        var worldPt = localPt.TransformBy(transform);
                        pts.Add(new Point2d(worldPt.X, worldPt.Y));
                    }

                    if (hasBounds)
                    {
                        double cx = pts.Average(p => p.X);
                        double cy = pts.Average(p => p.Y);
                        if (cx < bMinX || cx > bMaxX || cy < bMinY || cy > bMaxY)
                            continue;
                    }

                    double area = CalcTempArea(pts);
                    closedPolylines.Add(new PolylineData
                    {
                        ObjectId = entId,
                        Points = pts,
                        Area = area,
                        IsFromBlock = true
                    });
                }
                // DBText (Ð±Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€)
                else if (ent is DBText txt)
                {
                    string content = GetTextContent(txt, tr);
                    if (string.IsNullOrEmpty(content)) continue;

                    var worldPos = txt.Position.TransformBy(transform);
                    var pos = new Point2d(worldPos.X, worldPos.Y);
                    if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                    textEntities.Add(new TextData { Content = content, Position = pos });
                }
                // MText (Ð±Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€)
                else if (ent is MText mtxt)
                {
                    string content = GetTextContent(mtxt, tr);
                    if (string.IsNullOrEmpty(content)) continue;

                    var worldPos = mtxt.Location.TransformBy(transform);
                    var pos = new Point2d(worldPos.X, worldPos.Y);
                    if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                    textEntities.Add(new TextData { Content = content, Position = pos });
                }
                // AttributeDefinition (constant attribute)
                else if (ent is AttributeDefinition attDef && attDef.Constant)
                {
                    string content = GetTextContent(attDef, tr);
                    if (string.IsNullOrEmpty(content)) continue;

                    var worldPos = attDef.Position.TransformBy(transform);
                    var pos = new Point2d(worldPos.X, worldPos.Y);
                    if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                    textEntities.Add(new TextData { Content = content, Position = pos });
                }
                // Nested BlockReference (Ð½ÑÐ³ Ñ‚Ò¯Ð²ÑˆÐ¸Ð½ Ð³Ò¯Ð½Ð·Ð³Ð¸Ð¹)
                else if (ent is BlockReference nestedRef)
                {
                    // Nested Ð±Ð»Ð¾ÐºÐ¸Ð¹Ð½ Ñ…ÑƒÐ²ÑŒÐ´ Ñ…Ð¾ÑÐ¾Ð»ÑÐ¾Ð½ transform Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°
                    try
                    {
                        ScanNestedBlock(nestedRef, tr, transform,
                            hasBounds, bMinX, bMinY, bMaxX, bMaxY,
                            closedPolylines, textEntities, layerFilter, blkRefLayer);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Nested BlockReference Ð´Ð¾Ñ‚Ð¾Ñ€ polyline + Ñ‚ÐµÐºÑÑ‚ Ñ…Ð°Ð¹Ñ… (Ñ…Ð¾ÑÐ¾Ð»ÑÐ¾Ð½ transform Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°)
        /// </summary>
        private void ScanNestedBlock(BlockReference nestedRef, Transaction tr, Matrix3d parentTransform,
            bool hasBounds, double bMinX, double bMinY, double bMaxX, double bMaxY,
            List<PolylineData> closedPolylines, List<TextData> textEntities,
            string layerFilter = null, string parentBlockLayer = null)
        {
            Matrix3d combinedTransform = nestedRef.BlockTransform * parentTransform;
            // Nested ref-Ð¸Ð¹Ð½ effective layer: Ñ…ÑÑ€ÑÐ² "0" Ð±Ð¾Ð» ÑÑ†ÑÐ³ Ð±Ð»Ð¾ÐºÐ¸Ð¹Ð½ layer-Ð³ Ð°Ð²Ð½Ð°
            string nestedRefLayer = string.Equals(nestedRef.Layer, "0", StringComparison.OrdinalIgnoreCase)
                ? (parentBlockLayer ?? nestedRef.Layer) : nestedRef.Layer;

            var btr = tr.GetObject(nestedRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return;

            var allowedLayers = ParseLayerFilter(layerFilter);
            foreach (ObjectId entId in btr)
            {
                Entity ent;
                try { ent = tr.GetObject(entId, OpenMode.ForRead) as Entity; }
                catch { continue; }
                if (ent == null) continue;

                if (ent is Polyline pl && pl.Closed && pl.NumberOfVertices >= 3)
                {
                    string effectiveLayer = string.Equals(pl.Layer, "0", StringComparison.OrdinalIgnoreCase)
                        ? nestedRefLayer : pl.Layer;
                    if (!IsLayerAllowed(effectiveLayer, allowedLayers))
                        continue;

                    var pts = new List<Point2d>();
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        var worldPt = pl.GetPoint3dAt(i).TransformBy(combinedTransform);
                        pts.Add(new Point2d(worldPt.X, worldPt.Y));
                    }

                    if (hasBounds)
                    {
                        double cx = pts.Average(p => p.X);
                        double cy = pts.Average(p => p.Y);
                        if (cx < bMinX || cx > bMaxX || cy < bMinY || cy > bMaxY)
                            continue;
                    }

                    closedPolylines.Add(new PolylineData
                    {
                        ObjectId = entId,
                        Points = pts,
                        Area = CalcTempArea(pts),
                        IsFromBlock = true
                    });
                }
                else if (ent is DBText txt)
                {
                    string content = GetTextContent(txt, tr);
                    if (string.IsNullOrEmpty(content)) continue;

                    var pos = new Point2d(
                        txt.Position.TransformBy(combinedTransform).X,
                        txt.Position.TransformBy(combinedTransform).Y);
                    if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                    textEntities.Add(new TextData { Content = content, Position = pos });
                }
                else if (ent is MText mtxt)
                {
                    string content = GetTextContent(mtxt, tr);
                    if (string.IsNullOrEmpty(content)) continue;

                    var pos = new Point2d(
                        mtxt.Location.TransformBy(combinedTransform).X,
                        mtxt.Location.TransformBy(combinedTransform).Y);
                    if (hasBounds && !IsInBounds(pos, bMinX, bMaxX, bMinY, bMaxY)) continue;

                    textEntities.Add(new TextData { Content = content, Position = pos });
                }
            }
        }

        // Field Ñ‚ÐµÐºÑÑ‚ Ð¸Ð»Ñ€Ò¯Ò¯Ð»ÑÑ… Ñ‚Ð¾Ð¾Ð»ÑƒÑƒÑ€ (debug)
        private int _fieldTextCount;
        private int _fieldResolvedCount;

        /// <summary>
        /// AutoCAD Field Ð¾Ð±ÑŠÐµÐºÑ‚Ð¾Ð¾Ñ evaluated (Ñ…Ð°Ñ€ÑƒÑƒÐ»ÑÐ°Ð½) ÑƒÑ‚Ð³Ñ‹Ð³ Ð°Ð²Ð°Ñ….
        /// %<\AcField...>% Ð¼Ð°Ñ€ÐºÐµÑ€Ñ‚Ð°Ð¹ Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ ÑˆÐ¸Ð¹Ð´Ð²ÑÑ€Ð»ÑÑ….
        /// </summary>
        private string ResolveFieldText(string rawText, Entity entity, Transaction tr)
        {
            if (string.IsNullOrEmpty(rawText)) return rawText;

            // Field Ð¼Ð°Ñ€ÐºÐµÑ€ Ð±Ð°Ð¹Ð³Ð°Ð° ÑÑÑÑ… ÑˆÐ°Ð»Ð³Ð°Ñ…
            bool hasFieldMarker = rawText.Contains("%<") && rawText.Contains(">%");

            if (!hasFieldMarker) return rawText;

            _fieldTextCount++;

            // 1-Ñ€ Ð°Ñ€Ð³Ð°: Entity-Ð¸Ð¹Ð½ extension dictionary-Ð°Ð°Ñ Field Ð¾Ð±ÑŠÐµÐºÑ‚ Ð°Ð²Ð°Ñ…
            try
            {
                if (entity.ExtensionDictionary != ObjectId.Null)
                {
                    var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    if (extDict != null && extDict.Contains("ACAD_FIELD"))
                    {
                        var fieldDictId = extDict.GetAt("ACAD_FIELD");
                        var fieldDict = tr.GetObject(fieldDictId, OpenMode.ForRead) as DBDictionary;
                        if (fieldDict != null)
                        {
                            foreach (var entry in fieldDict)
                            {
                                try
                                {
                                    var field = tr.GetObject(entry.Value, OpenMode.ForRead) as Field;
                                    if (field != null)
                                    {
                                        // Field-Ð¸Ð¹Ð½ cached (ÑÒ¯Ò¯Ð»Ð´ evaluate Ñ…Ð¸Ð¹ÑÑÐ½) ÑƒÑ‚Ð³Ð° Ð°Ð²Ð°Ñ…
                                        string fieldValue = field.Value?.ToString()?.Trim();
                                        if (!string.IsNullOrEmpty(fieldValue) && !fieldValue.Contains("%<"))
                                        {
                                            _fieldResolvedCount++;
                                            return fieldValue;
                                        }

                                        // GetFieldCode-Ð¾Ð¾Ñ€ field code Ð°Ð²Ð°Ñ…
                                        try
                                        {
                                            string code = field.GetFieldCode();
                                            // Field code Ð´Ð¾Ñ‚Ð¾Ñ€ evaluated Ñ‚ÐµÐºÑÑ‚ Ð±Ð°Ð¹Ð¶ Ð±Ð¾Ð»Ð½Ð¾
                                            if (!string.IsNullOrEmpty(code) && !code.Contains("%<"))
                                            {
                                                _fieldResolvedCount++;
                                                return code.Trim();
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            catch { }

            // 2-Ñ€ Ð°Ñ€Ð³Ð°: MText Ð±Ð¾Ð» .Text property Ð°ÑˆÐ¸Ð³Ð»Ð°Ñ… (evaluated ÑƒÑ‚Ð³Ð° Ð±ÑƒÑ†Ð°Ð°Ð´Ð°Ð³)
            if (entity is MText mtext)
            {
                try
                {
                    string textProp = mtext.Text;
                    if (!string.IsNullOrEmpty(textProp) && !textProp.Contains("%<"))
                    {
                        _fieldResolvedCount++;
                        return textProp.Trim();
                    }
                }
                catch { }
            }

            // 3-Ñ€ Ð°Ñ€Ð³Ð°: DBText-Ð¸Ð¹Ð½ TextString (Ð¸Ñ…ÑÐ²Ñ‡Ð»ÑÐ½ evaluated Ð±Ð°Ð¹Ð´Ð°Ð³)
            if (entity is DBText dbtext)
            {
                try
                {
                    string ts = dbtext.TextString;
                    if (!string.IsNullOrEmpty(ts) && !ts.Contains("%<"))
                    {
                        _fieldResolvedCount++;
                        return ts.Trim();
                    }
                }
                catch { }
            }

            // 4-Ñ€ Ð°Ñ€Ð³Ð°: Field Ð¼Ð°Ñ€ÐºÐµÑ€ Ñ…Ð°ÑÐ°Ñ… â€” "Ñ‚ÐµÐºÑÑ‚ %<\AcField...>% Ñ‚ÐµÐºÑÑ‚" â†’ "Ñ‚ÐµÐºÑÑ‚ Ñ‚ÐµÐºÑÑ‚"
            string cleaned = Regex.Replace(rawText, @"%<[^>]*>%", "").Trim();
            if (!string.IsNullOrEmpty(cleaned))
            {
                _fieldResolvedCount++;
                return cleaned;
            }

            // Ð¨Ð¸Ð¹Ð´Ð²ÑÑ€Ð»ÑÑ… Ð±Ð¾Ð»Ð¾Ð¼Ð¶Ð³Ò¯Ð¹ â€” Ð°Ð½Ñ…Ð½Ñ‹ Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ Ð±ÑƒÑ†Ð°Ð°Ð½Ð°
            return rawText;
        }

        /// <summary>
        /// Ð¯Ð¼Ð°Ñ€ Ñ‡ text entity-Ð¸Ð¹Ð½ Ð°Ð³ÑƒÑƒÐ»Ð³Ñ‹Ð³ Ð°Ð²Ð°Ñ… (Field-Ð¸Ð¹Ð³ ÑˆÐ¸Ð¹Ð´Ð²ÑÑ€Ð»ÑÑ… Ð¾Ñ€Ð¾Ð»Ð´Ð»Ð¾Ð³Ð¾Ñ‚Ð¾Ð¹)
        /// </summary>
        private string GetTextContent(Entity entity, Transaction tr)
        {
            string raw = "";

            if (entity is DBText txt)
            {
                raw = txt.TextString ?? "";
            }
            else if (entity is MText mtxt)
            {
                // .Text = evaluated text, .Contents = raw with format codes
                raw = mtxt.Text ?? mtxt.Contents ?? "";
            }
            else if (entity is AttributeReference att)
            {
                raw = att.TextString ?? "";
            }
            else if (entity is AttributeDefinition attDef)
            {
                raw = attDef.TextString ?? "";
            }

            raw = raw.Trim();
            if (string.IsNullOrEmpty(raw)) return "";

            // Field Ð¼Ð°Ñ€ÐºÐµÑ€ Ð±Ð°Ð¹Ð²Ð°Ð» resolve Ñ…Ð¸Ð¹Ñ…
            if (raw.Contains("%<") && raw.Contains(">%"))
            {
                return ResolveFieldText(raw, entity, tr);
            }

            return raw;
        }

        // Ð¢Ò¯Ñ€ Ð±Ò¯Ñ‚Ñ†Ò¯Ò¯Ð´
        private class PolylineData
        {
            public ObjectId ObjectId;
            public List<Point2d> Points;
            public double Area;
            public bool IsFromBlock; // true = Ð±Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€, ÑˆÐ¸Ð½Ñ polyline Ò¯Ò¯ÑÐ³ÑÑ… ÑˆÐ°Ð°Ñ€Ð´Ð»Ð°Ð³Ð°Ñ‚Ð°Ð¹
        }

        private class TextData
        {
            public string Content;
            public Point2d Position;
        }

        /// <summary>
        /// Polyline-Ð¸Ð¹Ð½ Ñ‚Ð°Ð»Ð±Ð°Ð¹ Ð°Ð²Ð°Ñ…
        /// </summary>
        private double NormalizeAreaToM2(double area, double expectedAreaM2)
        {
            if (area <= 0) return 0;
            if (expectedAreaM2 > 0 && area / expectedAreaM2 > 500) return area / 1000000.0;
            if (area > 1000000) return area / 1000000.0;
            return area;
        }

        public double GetPolylineArea(ObjectId polylineId, double expectedAreaM2 = 0)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return 0;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var entity = tr.GetObject(polylineId, OpenMode.ForRead);
                double area = 0;

                if (entity is Polyline pl)
                    area = pl.Area;
                else if (entity is Polyline2d pl2d)
                    area = pl2d.Area;
                else if (entity is Polyline3d pl3d)
                    area = pl3d.Area;

                tr.Commit();
                return NormalizeAreaToM2(area, expectedAreaM2);
            }
        }
    }
}



































