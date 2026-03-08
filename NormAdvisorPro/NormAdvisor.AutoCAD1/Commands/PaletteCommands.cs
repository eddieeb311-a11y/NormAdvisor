using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using NormAdvisor.AutoCAD1.Models;
using NormAdvisor.AutoCAD1.Services;
using NormAdvisor.AutoCAD1.ViewModels;

namespace NormAdvisor.AutoCAD1.Commands
{
    public class PaletteCommands
    {
        /// <summary>
        /// NORMADVISOR â€” ÐŸÐ°Ð»ÐµÑ‚Ñ‚ Ñ†Ð¾Ð½Ñ… Ð½ÑÑÑ…/Ñ…Ð°Ð°Ñ…
        /// </summary>
        [CommandMethod("NORMADVISOR")]
        public void TogglePalette()
        {
            try
            {
                NormPaletteSet.Instance.Toggle();
            }
            catch (System.Exception ex)
            {
                var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage($"\nÐŸÐ°Ð»ÐµÑ‚Ñ‚ Ð½ÑÑÑ…ÑÐ´ Ð°Ð»Ð´Ð°Ð°: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMROOMS â€” Ó¨Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ñ…Ò¯ÑÐ½ÑÐ³Ñ‚ Ñ‚Ð°Ð½Ð¸Ñ… (Table ÑÑÐ²ÑÐ» Region)
        /// </summary>
        [CommandMethod("NORMROOMS")]
        public void ReadRooms()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var reader = new RoomTableReader();
            var rooms = reader.ReadRooms();

            if (rooms.Count == 0)
            {
                ed.WriteMessage("\nÓ¨Ñ€Ó©Ó© Ð¾Ð»Ð´ÑÐ¾Ð½Ð³Ò¯Ð¹.");
                return;
            }

            ed.WriteMessage($"\n\n===== Ó¨Ð Ó¨Ó¨ÐÐ˜Ð™ Ð–ÐÐ“Ð¡ÐÐÐ›Ð¢ ({rooms.Count} Ó©Ñ€Ó©Ó©) =====\n");
            foreach (var room in rooms)
            {
                ed.WriteMessage($"\n  {room}");
            }
            ed.WriteMessage($"\n\nÐÐ¸Ð¹Ñ‚: {rooms.Count} Ó©Ñ€Ó©Ó©");
            ed.WriteMessage("\n==========================================\n");
        }

        /// <summary>
        /// NORMPLACE â€” ÐŸÐ°Ð»ÐµÑ‚Ñ‚-ÑÑÑ Ñ‚Ó©Ñ…Ó©Ó©Ñ€Ó©Ð¼Ð¶ Ð±Ð°Ð¹Ñ€ÑˆÑƒÑƒÐ»Ð°Ñ…Ð°Ð´ Ð´ÑƒÑƒÐ´Ð°Ð³Ð´Ð°Ð½Ð°
        /// PlacementContext-Ð¾Ð¾Ñ Ó©Ð³Ó©Ð³Ð´Ó©Ð» Ð°Ð²Ð½Ð° (| Ñ‚ÑÐ¼Ð´ÑÐ³, Ð·Ð°Ð¹ Ð°ÑˆÐ¸Ð³Ð»Ð°Ñ…Ð³Ò¯Ð¹)
        /// </summary>
        [CommandMethod("NORMPLACE")]
        public void PlaceDevice()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                if (!PlacementContext.HasPending)
                {
                    ed.WriteMessage("\nÐ‘Ð°Ð¹Ñ€ÑˆÑƒÑƒÐ»Ð°Ñ… Ñ‚Ó©Ñ…Ó©Ó©Ñ€Ó©Ð¼Ð¶ ÑÐ¾Ð½Ð³Ð¾Ð³Ð´Ð¾Ð¾Ð³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°.");
                    return;
                }

                var category = PlacementContext.PendingCategory;
                var device = PlacementContext.PendingDevice;
                PlacementContext.Clear();

                var placement = new DevicePlacementService();
                placement.PlaceDevice(category, device);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nÐÐ»Ð´Ð°Ð°: {ex.Message}");
            }
        }
        /// <summary>
        /// NORMDRAWROOM â€” Palette-ÑÑÑ Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ polyline Ñ…Ò¯Ñ€ÑÑ Ñ…Ð¾Ð»Ð±Ð¾Ñ…Ð¾Ð´ Ð´ÑƒÑƒÐ´Ð°Ð³Ð´Ð°Ð½Ð°
        /// RoomDrawingContext-Ð¾Ð¾Ñ Ó©Ð³Ó©Ð³Ð´Ó©Ð» Ð°Ð²Ð½Ð°
        /// </summary>
        [CommandMethod("NORMDRAWROOM")]
        public void DrawRoomBoundary()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                if (!RoomDrawingContext.HasPending)
                {
                    ed.WriteMessage("\nÓ¨Ñ€Ó©Ó© ÑÐ¾Ð½Ð³Ð¾Ð³Ð´Ð¾Ð¾Ð³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°. ÐŸÐ°Ð»ÐµÑ‚Ñ‚Ð¾Ð¾Ñ Ó©Ñ€Ó©Ó© ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ.");
                    return;
                }

                var room = RoomDrawingContext.PendingRoom;
                RoomDrawingContext.Clear();

                ed.WriteMessage($"\n--- {room.Number}. {room.Name} ({room.Area:F2} Ð¼Â²) Ñ…Ò¯Ñ€ÑÑ Ð·ÑƒÑ€Ð½Ð° ---");

                var service = new RoomBoundaryService();
                var result = service.DrawOrSelectBoundary(room);

                if (result.PolylineId != ObjectId.Null)
                {
                    ed.WriteMessage($"\n  Polyline Ñ…Ð¾Ð»Ð±Ð¾Ð³Ð´Ð»Ð¾Ð¾. Ð‘Ð¾Ð´Ð¸Ñ‚ Ñ‚Ð°Ð»Ð±Ð°Ð¹: {result.DrawnArea:F2} Ð¼Â²");

                    double diff = result.DrawnArea - room.Area;
                    if (Math.Abs(diff) > 0.1)
                    {
                        double pct = room.Area > 0 ? (diff / room.Area) * 100 : 0;
                        ed.WriteMessage($"\n  Ð—Ó©Ñ€Ò¯Ò¯: {diff:+0.00;-0.00} Ð¼Â² ({pct:+0.0;-0.0}%)");
                    }

                    // ViewModel ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
                    RoomListViewModel.Current?.OnBoundaryDrawn(room, result.PolylineId, result.DrawnArea);

                    ed.WriteMessage($"\n  ÐÐ¼Ð¶Ð¸Ð»Ñ‚Ñ‚Ð°Ð¹!");
                }
                else
                {
                    ed.WriteMessage("\n  Ð¦ÑƒÑ†Ð°Ð»Ð»Ð°Ð°.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nÐÐ»Ð´Ð°Ð°: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMZOOMROOM â€” Palette-ÑÑÑ Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ polyline Ñ€ÑƒÑƒ zoom Ñ…Ð¸Ð¹Ñ…ÑÐ´ Ð´ÑƒÑƒÐ´Ð°Ð³Ð´Ð°Ð½Ð°
        /// RoomDrawingContext-Ð¾Ð¾Ñ Ó©Ð³Ó©Ð³Ð´Ó©Ð» Ð°Ð²Ð½Ð°; XData scan Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð° (stale ObjectId-Ð¸Ð¹Ð³ Ð·Ð°ÑÐ°Ñ…)
        /// </summary>
        [CommandMethod("NORMZOOMROOM")]
        public void ZoomToRoom()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                if (!RoomDrawingContext.HasPending)
                {
                    ed.WriteMessage("\nÓ¨Ñ€Ó©Ó© ÑÐ¾Ð½Ð³Ð¾Ð³Ð´Ð¾Ð¾Ð³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°.");
                    return;
                }

                var room = RoomDrawingContext.PendingRoom;
                RoomDrawingContext.Clear();

                var service = new RoomBoundaryService();
                // room.Number-Ð³ Ð´Ð°Ð²ÑƒÑƒÐ»Ð¶ Ó©Ð³ÑÐ½Ó©Ó©Ñ€ XData scan Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°
                string zoomKey = !string.IsNullOrWhiteSpace(room.RoomId) && System.Text.RegularExpressions.Regex.IsMatch(room.RoomId, "[A-Za-zА-Яа-я]")
                    ? room.RoomId.Trim()
                    : $"#N{room.Number}:{(room.Name ?? string.Empty).Trim()}";
                service.ZoomToRoom(room.BoundaryId, room.Number, zoomKey);
                ed.WriteMessage($"\nâ†’ {room.Number}. {room.Name}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nÐÐ»Ð´Ð°Ð°: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMSELECTROOM â€” Palette-Ð¸Ð¹Ð½ âŠ• Ñ‚Ð¾Ð²Ñ‡Ð¾Ð¾Ñ€ Ð±Ð°Ð¹Ð³Ð°Ð° polyline ÑÐ¾Ð½Ð³Ð¾Ð½ Ñ…Ð¾Ð»Ð±Ð¾Ñ…Ð¾Ð´ Ð´ÑƒÑƒÐ´Ð°Ð³Ð´Ð°Ð½Ð°
        /// </summary>
        /// <summary>
        /// NORMDIAGZOOM - zoom lookup diagnostic
        /// </summary>
        [CommandMethod("NORMDIAGZOOM")]
        public void DiagZoomRoom()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            try
            {
                if (!RoomDrawingContext.HasPending)
                {
                    ed?.WriteMessage("\nӨрөө сонгогдоогүй байна. Жагсаалтаас өрөөг сонгоод дахин ажиллуулна уу.");
                    return;
                }

                var room = RoomDrawingContext.PendingRoom;
                string zoomKey = !string.IsNullOrWhiteSpace(room.RoomId) && System.Text.RegularExpressions.Regex.IsMatch(room.RoomId, "[A-Za-zА-Яа-я]")
                    ? room.RoomId.Trim()
                    : $"#N{room.Number}:{(room.Name ?? string.Empty).Trim()}";

                var service = new RoomBoundaryService();
                service.DiagnoseRoomLookup(room.Number, zoomKey);
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\nDiag error: {ex.Message}");
            }
        }

        [CommandMethod("NORMSELECTROOM")]
        public void SelectRoomBoundary()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                if (!RoomDrawingContext.HasPending)
                {
                    ed.WriteMessage("\nÓ¨Ñ€Ó©Ó© ÑÐ¾Ð½Ð³Ð¾Ð³Ð´Ð¾Ð¾Ð³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°. ÐŸÐ°Ð»ÐµÑ‚Ñ‚Ð¾Ð¾Ñ Ó©Ñ€Ó©Ó© ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ.");
                    return;
                }

                var room = RoomDrawingContext.PendingRoom;
                RoomDrawingContext.Clear();

                ed.WriteMessage($"\n--- {room.Number}. {room.Name} â€” Ð±Ð°Ð¹Ð³Ð°Ð° polyline ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ ---");

                var service = new RoomBoundaryService();
                var result = service.SelectBoundary(room);

                if (result.PolylineId != ObjectId.Null)
                {
                    ed.WriteMessage($"\n  Polyline Ñ…Ð¾Ð»Ð±Ð¾Ð³Ð´Ð»Ð¾Ð¾. Ð‘Ð¾Ð´Ð¸Ñ‚ Ñ‚Ð°Ð»Ð±Ð°Ð¹: {result.DrawnArea:F2} Ð¼Â²");

                    double diff = result.DrawnArea - room.Area;
                    if (Math.Abs(diff) > 0.1)
                    {
                        double pct = room.Area > 0 ? (diff / room.Area) * 100 : 0;
                        ed.WriteMessage($"\n  Ð—Ó©Ñ€Ò¯Ò¯: {diff:+0.00;-0.00} Ð¼Â² ({pct:+0.0;-0.0}%)");
                    }

                    RoomListViewModel.Current?.OnBoundaryDrawn(room, result.PolylineId, result.DrawnArea);
                    ed.WriteMessage($"\n  ÐÐ¼Ð¶Ð¸Ð»Ñ‚Ñ‚Ð°Ð¹!");
                }
                else
                {
                    ed.WriteMessage("\n  Ð¦ÑƒÑ†Ð°Ð»Ð»Ð°Ð°.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nÐÐ»Ð´Ð°Ð°: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMAUTOMATCH â€” Ð‘Ð»Ð¾Ðº ÑÐ¾Ð½Ð³Ð¾Ð¾Ð´ Ð´Ð¾Ñ‚Ð¾Ñ€+Ð³Ð°Ð´Ð½Ð° polyline/Ñ‚ÐµÐºÑÑ‚ Ñ…Ð°Ð¹Ð¶, Ó©Ñ€Ó©Ó©Ñ‚ÑÐ¹ Ð°Ð²Ñ‚Ð¾Ð¼Ð°Ñ‚ Ñ…Ð¾Ð»Ð±Ð¾Ñ…
        /// </summary>
        [CommandMethod("NORMAUTOMATCH")]
        public void AutoMatchBoundaries()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                var vm = RoomListViewModel.Current;
                if (vm == null)
                {
                    ed.WriteMessage("\nÐŸÐ°Ð»ÐµÑ‚Ñ‚ Ð½ÑÑÐ³Ð´ÑÑÐ³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°.");
                    return;
                }

                var rooms = vm.GetAllRooms();
                if (rooms.Count == 0)
                {
                    ed.WriteMessage("\nÓ¨Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ð¶Ð°Ð³ÑÐ°Ð°Ð»Ñ‚ Ñ…Ð¾Ð¾ÑÐ¾Ð½ Ð±Ð°Ð¹Ð½Ð°.");
                    return;
                }

                // Ð‘Ð»Ð¾Ðº ÑÐ¾Ð½Ð³ÑƒÑƒÐ»Ð°Ñ…
                var blkOpts = new PromptEntityOptions("\nБайгуулалтын блок дээр дарна уу: ");
                blkOpts.SetRejectMessage("\nBlock сонгоно уу.");
                blkOpts.AddAllowedClass(typeof(BlockReference), true);

                var blkResult = ed.GetEntity(blkOpts);
                if (blkResult.Status != PromptStatus.OK) return;

                Autodesk.AutoCAD.Geometry.Point2d boundsMin, boundsMax;
                string layerFilter = null;
                var layerCounts = new System.Collections.Generic.Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                using (var trBlk = doc.Database.TransactionManager.StartTransaction())
                {
                    var blkRef = trBlk.GetObject(blkResult.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blkRef == null) { trBlk.Commit(); return; }

                    // Ð‘Ð»Ð¾ÐºÐ¸Ð¹Ð½ extents-ÑÑÑ bounds Ð°Ð²Ð°Ñ… (5% padding)
                    var ext = blkRef.GeometricExtents;
                    double padX = (ext.MaxPoint.X - ext.MinPoint.X) * 0.05;
                    double padY = (ext.MaxPoint.Y - ext.MinPoint.Y) * 0.05;
                    boundsMin = new Autodesk.AutoCAD.Geometry.Point2d(
                        ext.MinPoint.X - padX, ext.MinPoint.Y - padY);
                    boundsMax = new Autodesk.AutoCAD.Geometry.Point2d(
                        ext.MaxPoint.X + padX, ext.MaxPoint.Y + padY);

                    ed.WriteMessage($"\n  Блок: \"{blkRef.Name}\"");

                    // Layer Ñ‚Ð¾Ð¾Ð»Ð¾Ð»: Ð±Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€ closed polyline-ÑƒÑƒÐ´Ñ‹Ð³ layer-Ð°Ð°Ñ€ Ñ‚Ð¾Ð¾Ð»Ð¾Ñ…
                    string blkRefLayerName = blkRef.Layer;

                    var btr = trBlk.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null)
                    {
                        foreach (ObjectId subId in btr)
                        {
                            Entity subEnt;
                            try { subEnt = trBlk.GetObject(subId, OpenMode.ForRead) as Entity; }
                            catch { continue; }
                            if (subEnt == null) continue;

                            if (subEnt is Polyline pl && pl.Closed && pl.NumberOfVertices >= 3)
                            {
                                string effLayer = string.Equals(pl.Layer, "0", StringComparison.OrdinalIgnoreCase)
                                    ? blkRefLayerName : pl.Layer;
                                if (!layerCounts.ContainsKey(effLayer))
                                    layerCounts[effLayer] = 0;
                                layerCounts[effLayer]++;
                            }
                            else if (subEnt is BlockReference nestedRef)
                            {
                                try
                                {
                                    string nestedLayer = string.Equals(nestedRef.Layer, "0", StringComparison.OrdinalIgnoreCase)
                                        ? blkRefLayerName : nestedRef.Layer;
                                    var nestedBtr = trBlk.GetObject(nestedRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                    if (nestedBtr != null)
                                    {
                                        foreach (ObjectId nId in nestedBtr)
                                        {
                                            var nEnt = trBlk.GetObject(nId, OpenMode.ForRead) as Polyline;
                                            if (nEnt != null && nEnt.Closed && nEnt.NumberOfVertices >= 3)
                                            {
                                                string nEffLayer = string.Equals(nEnt.Layer, "0", StringComparison.OrdinalIgnoreCase)
                                                    ? nestedLayer : nEnt.Layer;
                                                if (!layerCounts.ContainsKey(nEffLayer))
                                                    layerCounts[nEffLayer] = 0;
                                                layerCounts[nEffLayer]++;
                                            }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                    }

                    trBlk.Commit();
                }

                // Layer ÑÐ¾Ð½Ð³Ð¾Ð»Ñ‚: Ñ…ÑÑ€ÑÐ³Ð»ÑÐ³Ñ‡ÑÐ´ layer-Ð¸Ð¹Ð½ Ð¶Ð°Ð³ÑÐ°Ð°Ð»Ñ‚ Ñ…Ð°Ñ€ÑƒÑƒÐ»Ð¶ ÑÐ¾Ð½Ð³ÑƒÑƒÐ»Ð°Ñ…
                if (!layerCounts.Any())
                {
                    ed.WriteMessage("\n  Блок дотор closed polyline олдсонгүй.");
                    return;
                }

                // Layer-ÑƒÑƒÐ´Ñ‹Ð³ polyline Ñ‚Ð¾Ð¾Ð³Ð¾Ð¾Ñ€ ÑÑ€ÑÐ¼Ð±ÑÐ»Ð¶ Ð¶Ð°Ð³ÑÐ°Ð°Ñ…
                var sortedLayers = layerCounts.OrderByDescending(kv => kv.Value).ToList();
                ed.WriteMessage($"\n\n  Polyline-тай layer-ууд ({sortedLayers.Count}):");
                for (int li = 0; li < sortedLayers.Count; li++)
                {
                    string star = sortedLayers[li].Value == rooms.Count ? " â˜…" : "";
                    ed.WriteMessage($"\n    {li + 1}. \"{sortedLayers[li].Key}\" - {sortedLayers[li].Value} polyline{star}");
                }

                // Ð¥ÑÑ€ÑÐ² Ð³Ð°Ð½Ñ† layer Ð±Ð°Ð¹Ð²Ð°Ð» ÑˆÑƒÑƒÐ´ Ð°Ð²Ð½Ð°
                if (sortedLayers.Count == 1)
                {
                    layerFilter = sortedLayers[0].Key;
                    ed.WriteMessage($"\n  -> Ганц layer: \"{layerFilter}\"");
                }
                else
                {
                    // Keyword prompt: layer ÑÐ¾Ð½Ð³ÑƒÑƒÐ»Ð°Ñ…
                    var layerOpts = new PromptKeywordOptions($"\n  Layer сонгоно уу ({rooms.Count} өрөөтэй ойр тоотой нь *):");
                    // A..Z keyword-ÑƒÑƒÐ´ (Ñ…Ð°Ð¼Ð³Ð¸Ð¹Ð½ Ð¸Ñ…Ð´ÑÑ 10 layer)
                    string[] kwLetters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };
                    int maxLayers = System.Math.Min(sortedLayers.Count, kwLetters.Length);

                    for (int li = 0; li < maxLayers; li++)
                    {
                        string kw = kwLetters[li];
                        string label = $"{kw}: {sortedLayers[li].Key} ({sortedLayers[li].Value})";
                        layerOpts.Keywords.Add(kw, kw, label);
                    }

                    // Default: Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ñ‚Ð¾Ð¾Ñ‚Ð¾Ð¹ Ð¾Ð¹Ñ€ Ñ‚Ð¾Ð¾Ñ‚Ð¾Ð¹
                    var best = sortedLayers
                        .Select((kv, idx) => new { kv, idx })
                        .OrderBy(x => System.Math.Abs(x.kv.Value - rooms.Count))
                        .First();
                    layerOpts.Keywords.Default = kwLetters[best.idx];
                    layerOpts.AllowNone = true;

                    var kwRes = ed.GetKeywords(layerOpts);
                    if (kwRes.Status != PromptStatus.OK && kwRes.Status != PromptStatus.None)
                        return;

                    string chosen = kwRes.Status == PromptStatus.None
                        ? layerOpts.Keywords.Default
                        : kwRes.StringResult;
                    int chosenIdx = Array.IndexOf(kwLetters, chosen);
                    if (chosenIdx < 0 || chosenIdx >= maxLayers)
                    {
                        ed.WriteMessage("\n  Layer сонголт буруу.");
                        return;
                    }
                    int primaryCount = sortedLayers[chosenIdx].Value;
                    layerFilter = sortedLayers[chosenIdx].Key;

                    if (primaryCount < rooms.Count && sortedLayers.Count > 1)
                    {
                        var fallbackNorm = sortedLayers
                            .Select((kv, idx) => new { kv, idx })
                            .FirstOrDefault(x => x.idx != chosenIdx && string.Equals(x.kv.Key, "NORM_ROOM_BOUNDARY", StringComparison.OrdinalIgnoreCase));

                        var fallback = fallbackNorm ?? sortedLayers
                            .Select((kv, idx) => new { kv, idx })
                            .Where(x => x.idx != chosenIdx && x.kv.Value >= Math.Max(rooms.Count / 2, 10))
                            .OrderBy(x => Math.Abs((primaryCount + x.kv.Value) - rooms.Count))
                            .FirstOrDefault();

                        if (fallback != null)
                        {
                            layerFilter = layerFilter + "|" + fallback.kv.Key;
                            ed.WriteMessage($"\n  -> Layers: \"{layerFilter}\" (primary={primaryCount}, fallback={fallback.kv.Value})");
                        }
                        else
                        {
                            ed.WriteMessage($"\n  -> Layer: \"{layerFilter}\" ({primaryCount} polyline)");
                        }
                    }
                }

                ed.WriteMessage($"\n=== АВТО ТАНИХ ({rooms.Count} өрөө) ===");

                var service = new RoomBoundaryService();
                var results = service.AutoMatchBoundaries(rooms, boundsMin, boundsMax, layerFilter);

                ed.WriteMessage($"\n=== Дүн: {results.Count} өрөө холбогдлоо ===");

                // ViewModel ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
                vm.OnAutoMatchCompleted(results);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nАвто таних алдаа: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMDIAGTEXT â€” Ð‘Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€Ñ…Ð¸ Ð±Ò¯Ñ… Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ Ð½Ð°Ñ€Ð¸Ð¹Ð²Ñ‡Ð»Ð°Ð½ Ñ…Ð°Ñ€ÑƒÑƒÐ»Ð°Ñ… (debug)
        /// </summary>
        [CommandMethod("NORMDIAGTEXT")]
        public void DiagBlockText()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                var blkOpts = new PromptEntityOptions("\nÐ‘Ð»Ð¾Ðº ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ: ");
                blkOpts.SetRejectMessage("\nBlock ÑÐ¾Ð½Ð³Ð¾Ð½Ð¾ ÑƒÑƒ.");
                blkOpts.AddAllowedClass(typeof(BlockReference), true);

                var blkResult = ed.GetEntity(blkOpts);
                if (blkResult.Status != PromptStatus.OK) return;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var blkRef = tr.GetObject(blkResult.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blkRef == null) { tr.Commit(); return; }

                    ed.WriteMessage($"\n\n=== Ð‘Ð›ÐžÐš Ð¢Ð•ÐšÐ¡Ð¢ Ð”Ð˜ÐÐ“ÐÐžÐ¡Ð¢Ð˜Ðš: \"{blkRef.Name}\" ===");
                    ed.WriteMessage($"\n  Layer: {blkRef.Layer}");
                    int textCount = 0;
                    int fieldCount = 0;

                    // 1. AttributeReference-ÑƒÑƒÐ´
                    ed.WriteMessage($"\n\n--- AttributeReference ({blkRef.AttributeCollection.Count}) ---");
                    foreach (ObjectId attId in blkRef.AttributeCollection)
                    {
                        try
                        {
                            var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (att == null) continue;
                            textCount++;
                            string raw = att.TextString ?? "(null)";
                            bool hasField = raw.Contains("%<");
                            if (hasField) fieldCount++;
                            string fieldTag = hasField ? " [FIELD]" : "";

                            // Extension dictionary ÑˆÐ°Ð»Ð³Ð°Ñ…
                            string extInfo = "";
                            if (att.ExtensionDictionary != ObjectId.Null)
                            {
                                var extDict = tr.GetObject(att.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                                if (extDict != null && extDict.Contains("ACAD_FIELD"))
                                {
                                    var fDict = tr.GetObject(extDict.GetAt("ACAD_FIELD"), OpenMode.ForRead) as DBDictionary;
                                    if (fDict != null)
                                    {
                                        foreach (var entry in fDict)
                                        {
                                            var field = tr.GetObject(entry.Value, OpenMode.ForRead) as Field;
                                            if (field != null)
                                            {
                                                string fVal = field.Value?.ToString() ?? "(null)";
                                                extInfo = $" | Field.Value=\"{fVal}\"";
                                            }
                                        }
                                    }
                                }
                            }

                            ed.WriteMessage($"\n  AttRef: Tag=\"{att.Tag}\" Text=\"{raw}\"{fieldTag}{extInfo}");
                            ed.WriteMessage($"    Pos=({att.Position.X:F0},{att.Position.Y:F0})");
                        }
                        catch (System.Exception ex) { ed.WriteMessage($"\n  AttRef err: {ex.Message}"); }
                    }

                    // 2. Block definition entities
                    var btr = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null)
                    {
                        int entCount = 0;
                        int plCount = 0;
                        int nestedBlkCount = 0;

                        ed.WriteMessage($"\n\n--- Block Definition \"{btr.Name}\" entities ---");

                        foreach (ObjectId entId in btr)
                        {
                            Entity ent;
                            try { ent = tr.GetObject(entId, OpenMode.ForRead) as Entity; }
                            catch { continue; }
                            if (ent == null) continue;
                            entCount++;

                            if (ent is DBText txt)
                            {
                                textCount++;
                                string raw = txt.TextString ?? "(null)";
                                bool hasField = raw.Contains("%<");
                                if (hasField) fieldCount++;
                                string fieldTag = hasField ? " [FIELD]" : "";

                                string extInfo = "";
                                if (txt.ExtensionDictionary != ObjectId.Null)
                                {
                                    try
                                    {
                                        var extDict = tr.GetObject(txt.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                                        if (extDict != null && extDict.Contains("ACAD_FIELD"))
                                        {
                                            var fDict = tr.GetObject(extDict.GetAt("ACAD_FIELD"), OpenMode.ForRead) as DBDictionary;
                                            if (fDict != null)
                                            {
                                                foreach (var entry in fDict)
                                                {
                                                    var field = tr.GetObject(entry.Value, OpenMode.ForRead) as Field;
                                                    if (field != null)
                                                    {
                                                        string fVal = field.Value?.ToString() ?? "(null)";
                                                        extInfo = $" | Field.Value=\"{fVal}\"";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }

                                ed.WriteMessage($"\n  DBText: \"{raw}\"{fieldTag}{extInfo} Layer={txt.Layer} Pos=({txt.Position.X:F0},{txt.Position.Y:F0})");
                            }
                            else if (ent is MText mtxt)
                            {
                                textCount++;
                                string rawText = mtxt.Text ?? "(null)";
                                string rawContents = mtxt.Contents ?? "(null)";
                                bool hasField = rawText.Contains("%<") || rawContents.Contains("%<");
                                if (hasField) fieldCount++;
                                string fieldTag = hasField ? " [FIELD]" : "";

                                string extInfo = "";
                                if (mtxt.ExtensionDictionary != ObjectId.Null)
                                {
                                    try
                                    {
                                        var extDict = tr.GetObject(mtxt.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                                        if (extDict != null && extDict.Contains("ACAD_FIELD"))
                                        {
                                            var fDict = tr.GetObject(extDict.GetAt("ACAD_FIELD"), OpenMode.ForRead) as DBDictionary;
                                            if (fDict != null)
                                            {
                                                foreach (var entry in fDict)
                                                {
                                                    var field = tr.GetObject(entry.Value, OpenMode.ForRead) as Field;
                                                    if (field != null)
                                                    {
                                                        string fVal = field.Value?.ToString() ?? "(null)";
                                                        extInfo = $" | Field.Value=\"{fVal}\"";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }

                                // Ð£Ñ€Ñ‚ Ñ‚ÐµÐºÑÑ‚Ð¸Ð¹Ð³ Ñ‚Ð¾Ð²Ñ‡Ð»Ð¾Ñ…
                                string showText = rawText.Length > 60 ? rawText.Substring(0, 60) + "..." : rawText;
                                string showCont = rawContents.Length > 60 ? rawContents.Substring(0, 60) + "..." : rawContents;
                                ed.WriteMessage($"\n  MText: Text=\"{showText}\"{fieldTag}{extInfo} Layer={mtxt.Layer}");
                                if (rawText != rawContents)
                                    ed.WriteMessage($"\n         Contents=\"{showCont}\"");
                            }
                            else if (ent is AttributeDefinition attDef)
                            {
                                textCount++;
                                string raw = attDef.TextString ?? "(null)";
                                string constStr = attDef.Constant ? " [CONST]" : "";
                                ed.WriteMessage($"\n  AttDef: Tag=\"{attDef.Tag}\" Text=\"{raw}\"{constStr} Layer={attDef.Layer}");
                            }
                            else if (ent is Polyline pl)
                            {
                                if (pl.Closed) plCount++;
                            }
                            else if (ent is BlockReference nested)
                            {
                                nestedBlkCount++;
                                // Nested Ð±Ð»Ð¾Ðº Ð´Ð¾Ñ‚Ð¾Ñ€ Ñ‚ÐµÐºÑÑ‚ Ñ…Ð°Ð¹Ñ…
                                ed.WriteMessage($"\n  NestedBlock: \"{nested.Name}\" Layer={nested.Layer}");

                                // Nested Ð±Ð»Ð¾ÐºÐ¸Ð¹Ð½ attribute-ÑƒÑƒÐ´
                                foreach (ObjectId nAttId in nested.AttributeCollection)
                                {
                                    try
                                    {
                                        var nAtt = tr.GetObject(nAttId, OpenMode.ForRead) as AttributeReference;
                                        if (nAtt == null) continue;
                                        textCount++;
                                        string nRaw = nAtt.TextString ?? "(null)";
                                        bool nHasField = nRaw.Contains("%<");
                                        if (nHasField) fieldCount++;
                                        ed.WriteMessage($"\n    AttRef: Tag=\"{nAtt.Tag}\" Text=\"{nRaw}\"{(nHasField ? " [FIELD]" : "")}");
                                    }
                                    catch { }
                                }

                                // Nested definition Ð´Ð¾Ñ‚Ð¾Ñ€ Ñ‚ÐµÐºÑÑ‚
                                var nBtr = tr.GetObject(nested.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (nBtr != null)
                                {
                                    foreach (ObjectId nEntId in nBtr)
                                    {
                                        Entity nEnt;
                                        try { nEnt = tr.GetObject(nEntId, OpenMode.ForRead) as Entity; }
                                        catch { continue; }
                                        if (nEnt == null) continue;

                                        if (nEnt is DBText nTxt)
                                        {
                                            textCount++;
                                            string nRaw = nTxt.TextString ?? "(null)";
                                            bool nHasField = nRaw.Contains("%<");
                                            if (nHasField) fieldCount++;
                                            ed.WriteMessage($"\n    DBText: \"{nRaw}\"{(nHasField ? " [FIELD]" : "")} Layer={nTxt.Layer}");
                                        }
                                        else if (nEnt is MText nMtxt)
                                        {
                                            textCount++;
                                            string nRaw = nMtxt.Text ?? "(null)";
                                            bool nHasField = nRaw.Contains("%<") || (nMtxt.Contents ?? "").Contains("%<");
                                            if (nHasField) fieldCount++;
                                            string nShow = nRaw.Length > 60 ? nRaw.Substring(0, 60) + "..." : nRaw;
                                            ed.WriteMessage($"\n    MText: \"{nShow}\"{(nHasField ? " [FIELD]" : "")} Layer={nMtxt.Layer}");
                                        }
                                    }
                                }
                            }
                        }

                        ed.WriteMessage($"\n\n--- ÐÐ¸Ð¹Ñ‚ ---");
                        ed.WriteMessage($"\n  Entity: {entCount}, ClosedPL: {plCount}, NestedBlk: {nestedBlkCount}");
                        ed.WriteMessage($"\n  Text: {textCount}, Field: {fieldCount}");
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n=== Ð”Ð˜ÐÐ“ÐÐžÐ¡Ð¢Ð˜Ðš Ð”Ð£Ð£Ð¡Ð›ÐÐ ===\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nÐ”Ð¸Ð°Ð³Ð½Ð¾ÑÑ‚Ð¸Ðº Ð°Ð»Ð´Ð°Ð°: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMBLOCKS â€” DWG Ñ„Ð°Ð¹Ð» Ð±Ò¯Ñ€Ð¸Ð¹Ð½ Ð±Ð»Ð¾Ðº Ð½ÑÑ€Ð¸Ð¹Ð³ Ð¶Ð°Ð³ÑÐ°Ð°Ñ… (debug)
        /// </summary>
        /// <summary>
        /// NORM5X10LAN - LAN ugsraltyn buduuvchiin ehnii prototype (floor x LAN count)
        /// </summary>
        [CommandMethod("NORM5X10LAN")]
        public void DrawLanStackPrototype()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                var ptRes = ed.GetPoint("\nБүдүүвч эхлэх цэг сонгоно уу: ");
                if (ptRes.Status != PromptStatus.OK) return;

                var floorOpts = new PromptIntegerOptions("\nДавхрын тоо <5>: ")
                {
                    DefaultValue = 5,
                    AllowNone = true,
                    LowerLimit = 1,
                    UpperLimit = 50
                };
                floorOpts.UseDefaultValue = true;
                var floorRes = ed.GetInteger(floorOpts);
                if (floorRes.Status == PromptStatus.Cancel) return;
                int floors = floorRes.Status == PromptStatus.None ? 5 : floorRes.Value;

                var lanOpts = new PromptIntegerOptions("\nНэг давхарт LAN тоо <10>: ")
                {
                    DefaultValue = 10,
                    AllowNone = true,
                    LowerLimit = 1,
                    UpperLimit = 128
                };
                lanOpts.UseDefaultValue = true;
                var lanRes = ed.GetInteger(lanOpts);
                if (lanRes.Status == PromptStatus.Cancel) return;
                int lanPerFloor = lanRes.Status == PromptStatus.None ? 10 : lanRes.Value;

                using (var lockDoc = doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var existing = new HashSet<ObjectId>();
                    foreach (ObjectId id in ms) existing.Add(id);

                    var ubBlockId = EnsureUbMedeelelBlock(db, ed);
                    DrawLanStackSchematic(ms, tr, ptRes.Value, floors, lanPerFloor, ubBlockId);

                    const double scaleFactor = 140.0;
                    var scaleMx = Matrix3d.Scaling(scaleFactor, ptRes.Value);
                    foreach (ObjectId id in ms)
                    {
                        if (existing.Contains(id)) continue;
                        var ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        ent?.TransformBy(scaleMx);
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nУгсралтын бүдүүвч зурлаа: {floors} давхар x {lanPerFloor} LAN (scale x140).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nNORM5X10LAN алдаа: {ex.Message}");
            }
        }

        private void DrawLanStackSchematic(
            BlockTableRecord ms,
            Transaction tr,
            Point3d origin,
            int floors,
            int lanPerFloor,
            ObjectId ubBlockId)
        {
            double left = origin.X;
            double top = origin.Y;

            int safeFloors = Math.Max(1, floors);
            int cols = safeFloors >= 5 ? 2 : 1;
            int leftCount = (int)Math.Ceiling(safeFloors / 2.0);
            int rightCount = safeFloors - leftCount;

            double panelGap = 16.0;
            double panelW = 178.0;
            double rowH = 36.0;
            double headerH = 16.0;
            double panelTopOffset = 18.0;
            double panelBottomPad = 16.0;
            double maxRows = Math.Max(leftCount, Math.Max(1, rightCount));
            double panelH = headerH + 10.0 + maxRows * rowH + panelBottomPad;

            double sheetW = cols * panelW + (cols - 1) * panelGap + 20.0;
            double sheetH = panelTopOffset + panelH + 42.0;

            double sheetL = left;
            double sheetR = left + sheetW;
            double sheetT = top;
            double sheetB = top - sheetH;

            AddRect(ms, tr, sheetL, sheetT, sheetR, sheetB);
            AddRect(ms, tr, sheetL + 2, sheetT - 2, sheetR - 2, sheetB + 2);
            AddTextCentered(ms, tr, sheetL + 4, sheetR - 4, sheetT - 12, sheetT - 2.5, "ГУРВАЛСАН ҮЙЛЧИЛГЭЭНИЙ УГСРАЛТЫН БҮДҮҮВЧ", 3.0);

            double panel1X = sheetL + 8;
            double panelYTop = sheetT - panelTopOffset;

            DrawLanPanel(ms, tr, panel1X, panelYTop, panelW, rowH, leftCount, safeFloors, leftCount, lanPerFloor, ubBlockId, "K-1");

            if (cols == 2)
            {
                double panel2X = panel1X + panelW + panelGap;
                DrawLanPanel(ms, tr, panel2X, panelYTop, panelW, rowH, rightCount, leftCount, 1, lanPerFloor, ubBlockId, "K-2");
            }

            DrawBasementRack(ms, tr, sheetL + 8, sheetB + 12, ubBlockId);
        }

        private void DrawLanPanel(
            BlockTableRecord ms,
            Transaction tr,
            double x,
            double yTop,
            double width,
            double rowH,
            int rowCount,
            int floorFrom,
            int floorTo,
            int lanPerFloor,
            ObjectId ubBlockId,
            string riserPrefix)
        {
            if (rowCount <= 0) return;

            double h = 26.0 + rowCount * rowH;
            double yBottom = yTop - h;

            double xNo = x + 12.0;
            double xRoom = x + 58.0;
            double xBoard = x + 96.0;
            double xRiser = x + 140.0;
            double xRight = x + width;

            AddRect(ms, tr, x, yTop, xRight, yBottom);
            AddLine(ms, tr, x, yTop - 10, xRight, yTop - 10);
            AddLine(ms, tr, x, yTop - 16, xRight, yTop - 16);

            AddLine(ms, tr, xNo, yTop - 10, xNo, yBottom);
            AddLine(ms, tr, xRoom, yTop - 10, xRoom, yBottom);
            AddLine(ms, tr, xBoard, yTop - 10, xBoard, yBottom);
            AddLine(ms, tr, xRiser, yTop - 10, xRiser, yBottom);

            AddTextCentered(ms, tr, x + 0.5, xRight - 0.5, yTop - 10, yTop, "ГУРВАЛСАН ҮЙЛЧИЛГЭЭ", 2.0);
            AddTextCentered(ms, tr, x, xNo, yTop - 16, yTop - 10, "№", 1.8);
            AddTextCentered(ms, tr, xNo, xRoom, yTop - 16, yTop - 10, "Өрөө", 1.8);
            AddTextCentered(ms, tr, xRoom, xBoard, yTop - 16, yTop - 10, "Айлын самбар", 1.8);
            AddTextCentered(ms, tr, xBoard, xRiser, yTop - 16, yTop - 10, "Босоо сувагчлал", 1.8);

            for (int i = 0; i < rowCount; i++)
            {
                int floor = floorFrom - i;
                if (floor < floorTo) break;

                double rowTop = yTop - 16.0 - i * rowH;
                double rowBottom = rowTop - rowH;
                double yMid = (rowTop + rowBottom) / 2.0;

                AddLine(ms, tr, x, rowBottom, xRight, rowBottom);
                AddText(ms, tr, new Point3d(x + 2.0, yMid - 1.0, 0), $"{floor}-р", 1.8);
                AddText(ms, tr, new Point3d(x + 2.0, yMid - 4.0, 0), "давхар", 1.6);

                DrawFloorRoomStack(ms, tr, xNo + 2, rowTop - 4, xRoom - 2, rowBottom + 3, floor, lanPerFloor);

                double ontX = xRoom + 12;
                double riserX = xRiser + 14;
                double yA = rowTop - 8;
                double yB = rowTop - 17;
                double yC = rowTop - 26;

                DrawOntRun(ms, tr, ubBlockId, ontX, yA, riserX, $"A-{floor:D2}-1");
                DrawOntRun(ms, tr, ubBlockId, ontX, yB, riserX, $"A-{floor:D2}-2");
                DrawOntRun(ms, tr, ubBlockId, ontX, yC, riserX, $"A-{floor:D2}-3");

                AddLine(ms, tr, riserX + 8, yA, riserX + 8, yC);
                AddText(ms, tr, new Point3d(xRiser + 1, yMid - 1.0, 0), $"{riserPrefix}-{floor}", 1.8);

                if (!TryInsertUbSymbol(ms, tr, ubBlockId, new Point3d(riserX + 5.5, yMid - 0.8, 0), "FDF UB", 0.5))
                {
                    DrawFdbFallback(ms, tr, riserX + 3, yMid + 4, 10, 8);
                }
            }
        }

        private void DrawFloorRoomStack(
            BlockTableRecord ms,
            Transaction tr,
            double xLeft,
            double yTop,
            double xRight,
            double yBottom,
            int floor,
            int lanPerFloor)
        {
            double g1 = yTop;
            double g2 = yTop - 9;
            double g3 = yTop - 18;

            DrawRoomGroup(ms, tr, xLeft, xRight, g1, $"A,C,F,H  сууц", floor, 1, lanPerFloor);
            DrawRoomGroup(ms, tr, xLeft, xRight, g2, $"B,G  сууц", floor, 2, lanPerFloor);
            DrawRoomGroup(ms, tr, xLeft, xRight, g3, $"E,D  сууц", floor, 3, lanPerFloor);
        }

        private void DrawRoomGroup(
            BlockTableRecord ms,
            Transaction tr,
            double xLeft,
            double xRight,
            double y,
            string label,
            int floor,
            int group,
            int lanPerFloor)
        {
            AddText(ms, tr, new Point3d(xLeft + 0.5, y, 0), label, 1.4);

            double symX = xLeft + 2;
            double joinX = xRight - 6;
            double y1 = y - 2.2;
            double y2 = y - 4.2;
            double y3 = y - 6.2;

            AddCircle(ms, tr, symX + 2.0, y1, 0.9);
            AddText(ms, tr, new Point3d(symX + 4.0, y1 - 0.8, 0), "IPTEL", 1.1);

            AddSolidTriangle(ms, tr, symX + 2.0, y2, 1.5);
            AddText(ms, tr, new Point3d(symX + 4.0, y2 - 0.8, 0), "PTV", 1.1);

            AddRect(ms, tr, symX + 1.0, y3 + 0.8, symX + 3.2, y3 - 0.8);
            AddText(ms, tr, new Point3d(symX + 4.0, y3 - 0.8, 0), $"LAN-{Math.Max(1, lanPerFloor / 3)}", 1.1);

            AddLine(ms, tr, symX + 8.8, y1, joinX, y1);
            AddLine(ms, tr, symX + 8.0, y2, joinX, y2);
            AddLine(ms, tr, symX + 9.0, y3, joinX, y3);

            AddLine(ms, tr, joinX, y1, joinX, y3);
            AddCircle(ms, tr, joinX + 1.6, (y1 + y3) / 2.0, 0.7);
            AddText(ms, tr, new Point3d(joinX + 0.8, y3 - 2.2, 0), $"A-{floor},{group}", 1.0);
        }

        private void DrawOntRun(
            BlockTableRecord ms,
            Transaction tr,
            ObjectId ubBlockId,
            double ontX,
            double y,
            double riserX,
            string cableLabel)
        {
            bool inserted = TryInsertUbSymbol(ms, tr, ubBlockId, new Point3d(ontX + 5, y - 1.4, 0), "ONT UB", 0.45);
            if (!inserted)
            {
                AddRect(ms, tr, ontX, y + 2.2, ontX + 10, y - 2.2);
                AddText(ms, tr, new Point3d(ontX + 2.0, y - 0.9, 0), "ONT", 1.4);
            }

            AddLine(ms, tr, ontX + 10, y, riserX, y);
            AddCircle(ms, tr, riserX + 1.2, y, 0.7);
            AddText(ms, tr, new Point3d(ontX - 3, y + 0.6, 0), cableLabel, 1.0);
        }

        private void DrawFdbFallback(BlockTableRecord ms, Transaction tr, double x, double yTop, double w, double h)
        {
            AddRect(ms, tr, x, yTop, x + w, yTop - h);
            AddSolidTriangle(ms, tr, x + w * 0.55, yTop - h * 0.5, h * 0.7);
            AddText(ms, tr, new Point3d(x + 1.2, yTop - 2.8, 0), "FDB", 1.2);
        }

        private void DrawBasementRack(BlockTableRecord ms, Transaction tr, double x, double y, ObjectId ubBlockId)
        {
            double w = 110;
            double h = 34;
            AddRect(ms, tr, x, y + h, x + w, y);
            AddText(ms, tr, new Point3d(x + 2, y + h - 4, 0), "Зоорийн давхар", 1.8);
            AddRect(ms, tr, x + 46, y + h - 8, x + w - 4, y + 4);
            AddText(ms, tr, new Point3d(x + 59, y + 11, 0), "Rack-d", 1.8);

            double baseX = x + 52;
            double baseY = y + 20;
            for (int i = 0; i < 3; i++)
            {
                double yy = baseY - i * 7;
                if (!TryInsertUbSymbol(ms, tr, ubBlockId, new Point3d(baseX + 12, yy, 0), "FDF UB", 0.45))
                {
                    DrawFdbFallback(ms, tr, baseX + 8, yy + 3, 9, 6);
                }
                AddLine(ms, tr, baseX - 2, yy, baseX + 8, yy);
            }

            AddRect(ms, tr, x + 22, y + 16, x + 34, y + 10);
            AddText(ms, tr, new Point3d(x + 24, y + 11.5, 0), "ODF", 1.2);
            AddLine(ms, tr, x + 34, y + 13, baseX - 2, baseY);
            AddLine(ms, tr, x + 34, y + 13, baseX - 2, baseY - 7);
            AddLine(ms, tr, x + 34, y + 13, baseX - 2, baseY - 14);
        }        private ObjectId EnsureUbMedeelelBlock(Database targetDb, Editor ed)
        {
            try
            {
                var config = BlocksConfigService.Instance.LoadConfig();
                var ub = config.Categories.FirstOrDefault(c =>
                    string.Equals(c.BlockName, "UB Medeelel Holboo", StringComparison.OrdinalIgnoreCase));
                if (ub == null) return ObjectId.Null;

                string dwgPath = BlocksConfigService.Instance.GetDwgFullPath(ub.DwgFile);
                if (!File.Exists(dwgPath)) return ObjectId.Null;

                using (var tr = targetDb.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(targetDb.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(ub.BlockName))
                    {
                        var id = bt[ub.BlockName];
                        tr.Commit();
                        return id;
                    }
                    tr.Commit();
                }

                using (var sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, "");

                    var blockIds = new ObjectIdCollection();
                    using (var tr = sourceDb.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                        foreach (ObjectId id in bt)
                        {
                            var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            if (!btr.IsLayout) blockIds.Add(id);
                        }
                        tr.Commit();
                    }

                    if (blockIds.Count > 0)
                    {
                        var mapping = new IdMapping();
                        targetDb.WblockCloneObjects(blockIds, targetDb.BlockTableId, mapping, DuplicateRecordCloning.Ignore, false);
                    }
                }

                using (var tr = targetDb.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(targetDb.BlockTableId, OpenMode.ForRead);
                    if (bt.Has("UB Medeelel Holboo"))
                    {
                        var id = bt["UB Medeelel Holboo"];
                        tr.Commit();
                        return id;
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nUB block load aldaa: {ex.Message}");
            }

            return ObjectId.Null;
        }

        private bool TryInsertUbSymbol(BlockTableRecord ms, Transaction tr, ObjectId blockId, Point3d position, string visibilityState, double scale)
        {
            if (blockId == ObjectId.Null) return false;

            try
            {
                var br = new BlockReference(position, blockId)
                {
                    ScaleFactors = new Scale3d(scale, scale, scale)
                };
                ms.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                if (!string.IsNullOrWhiteSpace(visibilityState) && br.DynamicBlockTableRecord != ObjectId.Null)
                {
                    foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
                    {
                        if (!prop.PropertyName.StartsWith("Visibility", StringComparison.OrdinalIgnoreCase)) continue;
                        var allowed = prop.GetAllowedValues();
                        foreach (var v in allowed)
                        {
                            var s = v?.ToString() ?? string.Empty;
                            if (string.Equals(s, visibilityState, StringComparison.OrdinalIgnoreCase))
                            {
                                prop.Value = s;
                                br.RecordGraphicsModified(true);
                                break;
                            }
                        }
                        break;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void AddCircle(BlockTableRecord ms, Transaction tr, double x, double y, double r)
        {
            var c = new Circle(new Point3d(x, y, 0), Vector3d.ZAxis, r);
            ms.AppendEntity(c);
            tr.AddNewlyCreatedDBObject(c, true);
        }

        private static void AddSolidTriangle(BlockTableRecord ms, Transaction tr, double x, double y, double size)
        {
            var pl = new Polyline();
            pl.AddVertexAt(0, new Point2d(x - size * 0.5, y - size * 0.45), 0, 0, 0);
            pl.AddVertexAt(1, new Point2d(x - size * 0.5, y + size * 0.45), 0, 0, 0);
            pl.AddVertexAt(2, new Point2d(x + size * 0.6, y), 0, 0, 0);
            pl.Closed = true;
            ms.AppendEntity(pl);
            tr.AddNewlyCreatedDBObject(pl, true);
            try { pl.UpgradeOpen(); pl.ColorIndex = 256; } catch { }
        }
        private static void AddLine(BlockTableRecord ms, Transaction tr, double x1, double y1, double x2, double y2)
        {
            var ln = new Line(new Point3d(x1, y1, 0), new Point3d(x2, y2, 0));
            ms.AppendEntity(ln);
            tr.AddNewlyCreatedDBObject(ln, true);
        }

        private static void AddRect(BlockTableRecord ms, Transaction tr, double x1, double y1, double x2, double y2)
        {
            var pl = new Polyline();
            pl.AddVertexAt(0, new Point2d(x1, y1), 0, 0, 0);
            pl.AddVertexAt(1, new Point2d(x2, y1), 0, 0, 0);
            pl.AddVertexAt(2, new Point2d(x2, y2), 0, 0, 0);
            pl.AddVertexAt(3, new Point2d(x1, y2), 0, 0, 0);
            pl.Closed = true;
            ms.AppendEntity(pl);
            tr.AddNewlyCreatedDBObject(pl, true);
        }

        private static void AddTextCentered(
            BlockTableRecord ms,
            Transaction tr,
            double xLeft,
            double xRight,
            double yBottom,
            double yTop,
            string text,
            double height)
        {
            var dbt = new DBText
            {
                Height = height,
                TextString = text,
                HorizontalMode = TextHorizontalMode.TextCenter,
                VerticalMode = TextVerticalMode.TextVerticalMid,
                AlignmentPoint = new Point3d((xLeft + xRight) * 0.5, (yBottom + yTop) * 0.5, 0)
            };
            dbt.AdjustAlignment(ms.Database);
            ms.AppendEntity(dbt);
            tr.AddNewlyCreatedDBObject(dbt, true);
        }
        private static void AddText(BlockTableRecord ms, Transaction tr, Point3d pos, string text, double height)
        {
            var dbt = new DBText
            {
                Position = pos,
                Height = height,
                TextString = text
            };
            ms.AppendEntity(dbt);
            tr.AddNewlyCreatedDBObject(dbt, true);
        }
        [CommandMethod("NORMTEMPLATEINFO")]
        public void ExtractTemplateInfo()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            if (ed == null) return;

            try
            {
                var pathOpt = new PromptStringOptions("\nTemplate DWG зам <C:\\Users\\Byamba Erdene\\Desktop\\Template.dwg>: ")
                {
                    AllowSpaces = true,
                    DefaultValue = @"C:\Users\Byamba Erdene\Desktop\Template.dwg",
                    UseDefaultValue = true
                };
                var pathRes = ed.GetString(pathOpt);
                if (pathRes.Status == PromptStatus.Cancel) return;

                string dwgPath = string.IsNullOrWhiteSpace(pathRes.StringResult)
                    ? pathOpt.DefaultValue
                    : pathRes.StringResult.Trim();

                if (!File.Exists(dwgPath))
                {
                    ed.WriteMessage($"\nФайл олдсонгүй: {dwgPath}");
                    return;
                }

                var rects = new List<RectInfo>();
                var textHeights = new List<double>();
                var textStyles = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                using (var db = new Database(false, true))
                {
                    db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, "");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        foreach (ObjectId id in ms)
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent == null) continue;

                            if (ent is Polyline pl && IsAxisAlignedRect(pl, out var ex))
                            {
                                rects.Add(RectInfo.From(ex, ent.Layer, ent.ColorIndex));
                            }

                            if (ent is DBText dt)
                            {
                                if (dt.Height > 0) textHeights.Add(dt.Height);
                                string style = ResolveTextStyleName(tr, dt.TextStyleId);
                                if (!string.IsNullOrWhiteSpace(style))
                                {
                                    if (!textStyles.ContainsKey(style)) textStyles[style] = 0;
                                    textStyles[style]++;
                                }
                            }
                            else if (ent is MText mt)
                            {
                                if (mt.TextHeight > 0) textHeights.Add(mt.TextHeight);
                                string style = ResolveTextStyleName(tr, mt.TextStyleId);
                                if (!string.IsNullOrWhiteSpace(style))
                                {
                                    if (!textStyles.ContainsKey(style)) textStyles[style] = 0;
                                    textStyles[style]++;
                                }
                            }
                        }

                        tr.Commit();
                    }
                }

                if (rects.Count == 0)
                {
                    ed.WriteMessage("\nТэгш өнцөгт polyline олдсонгүй.");
                    return;
                }

                var sorted = rects.OrderByDescending(r => r.Area).ToList();
                var outer = sorted[0];
                var inner = sorted.Skip(1).FirstOrDefault(r => Contains(outer.Ext, r.Ext));
                var scope = inner ?? outer;

                var insideScope = sorted.Where(r => !ReferenceEquals(r, outer) && !ReferenceEquals(r, inner) && Contains(scope.Ext, r.Ext)).ToList();
                var workArea = insideScope.OrderByDescending(r => r.Area).FirstOrDefault();
                var titleBlock = insideScope
                    .Where(r => r.Center.X > (scope.Ext.MinPoint.X + scope.Ext.MaxPoint.X) / 2.0 && r.Center.Y < (scope.Ext.MinPoint.Y + scope.Ext.MaxPoint.Y) / 2.0)
                    .OrderByDescending(r => r.Area)
                    .FirstOrDefault();

                var panelCandidates = insideScope
                    .Where(r => !ReferenceEquals(r, workArea) && !ReferenceEquals(r, titleBlock))
                    .Where(r => r.H > scope.H * 0.25 && r.W > scope.W * 0.15)
                    .OrderByDescending(r => r.Area)
                    .Take(4)
                    .ToList();

                double commonHeight = textHeights.Count > 0
                    ? textHeights.GroupBy(h => Math.Round(h, 2)).OrderByDescending(g => g.Count()).First().Key
                    : 0;
                var topStyle = textStyles.OrderByDescending(kv => kv.Value).FirstOrDefault();

                var sb = new StringBuilder();
                sb.AppendLine("=== TEMPLATE INFO ===");
                sb.AppendLine($"DWG: {dwgPath}");
                sb.AppendLine($"Outer frame: W={outer.W:F0}, H={outer.H:F0}, layer={outer.Layer}");
                if (inner != null)
                {
                    sb.AppendLine($"Inner frame: W={inner.W:F0}, H={inner.H:F0}, layer={inner.Layer}");
                    sb.AppendLine($"Offsets (L,R,T,B): {(inner.Ext.MinPoint.X - outer.Ext.MinPoint.X):F0}, {(outer.Ext.MaxPoint.X - inner.Ext.MaxPoint.X):F0}, {(outer.Ext.MaxPoint.Y - inner.Ext.MaxPoint.Y):F0}, {(inner.Ext.MinPoint.Y - outer.Ext.MinPoint.Y):F0}");
                }
                if (workArea != null)
                {
                    sb.AppendLine($"Work area candidate: W={workArea.W:F0}, H={workArea.H:F0}, layer={workArea.Layer}");
                }
                if (titleBlock != null)
                {
                    sb.AppendLine($"Title block candidate: W={titleBlock.W:F0}, H={titleBlock.H:F0}, layer={titleBlock.Layer}");
                }
                if (panelCandidates.Count > 0)
                {
                    sb.AppendLine($"Panel candidates ({panelCandidates.Count}):");
                    for (int i = 0; i < panelCandidates.Count; i++)
                    {
                        var p = panelCandidates[i];
                        sb.AppendLine($"  {i + 1}) W={p.W:F0}, H={p.H:F0}, center=({p.Center.X:F0},{p.Center.Y:F0}), layer={p.Layer}");
                    }
                }

                if (commonHeight > 0)
                    sb.AppendLine($"Common text height: {commonHeight:F2}");
                if (!string.IsNullOrWhiteSpace(topStyle.Key))
                    sb.AppendLine($"Top text style: {topStyle.Key} ({topStyle.Value})");

                ed.WriteMessage("\n" + sb.ToString().Replace("\r\n", "\n"));

                string outPath = Path.Combine(@"C:\Users\Byamba Erdene\Desktop\Norm", "template_info.txt");
                File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8);
                ed.WriteMessage($"\nТайлан хадгаллаа: {outPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nNORMTEMPLATEINFO алдаа: {ex.Message}");
            }
        }

        private static string ResolveTextStyleName(Transaction tr, ObjectId styleId)
        {
            try
            {
                if (styleId == ObjectId.Null) return string.Empty;
                var rec = (TextStyleTableRecord)tr.GetObject(styleId, OpenMode.ForRead);
                return rec?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private sealed class RectInfo
        {
            public Extents2d Ext { get; set; }
            public string Layer { get; set; }
            public int ColorIndex { get; set; }
            public double W => Ext.MaxPoint.X - Ext.MinPoint.X;
            public double H => Ext.MaxPoint.Y - Ext.MinPoint.Y;
            public double Area => W * H;
            public Point2d Center => new Point2d((Ext.MinPoint.X + Ext.MaxPoint.X) * 0.5, (Ext.MinPoint.Y + Ext.MaxPoint.Y) * 0.5);

            public static RectInfo From(Extents2d ex, string layer, int color)
            {
                return new RectInfo { Ext = ex, Layer = layer ?? string.Empty, ColorIndex = color };
            }
        }

        private static bool IsAxisAlignedRect(Polyline pl, out Extents2d ex)
        {
            ex = default;
            if (pl == null || !pl.Closed || pl.NumberOfVertices != 4) return false;

            var pts = new Point2d[4];
            for (int i = 0; i < 4; i++)
            {
                if (Math.Abs(pl.GetBulgeAt(i)) > 1e-6) return false;
                pts[i] = pl.GetPoint2dAt(i);
            }

            for (int i = 0; i < 4; i++)
            {
                var a = pts[i];
                var b = pts[(i + 1) % 4];
                bool vertical = Math.Abs(a.X - b.X) < 1e-6;
                bool horizontal = Math.Abs(a.Y - b.Y) < 1e-6;
                if (!vertical && !horizontal) return false;
            }

            double minX = pts.Min(p => p.X);
            double maxX = pts.Max(p => p.X);
            double minY = pts.Min(p => p.Y);
            double maxY = pts.Max(p => p.Y);
            if (maxX - minX < 1 || maxY - minY < 1) return false;

            ex = new Extents2d(new Point2d(minX, minY), new Point2d(maxX, maxY));
            return true;
        }

        private static bool Contains(Extents2d outer, Extents2d inner)
        {
            return inner.MinPoint.X >= outer.MinPoint.X - 1e-6
                && inner.MaxPoint.X <= outer.MaxPoint.X + 1e-6
                && inner.MinPoint.Y >= outer.MinPoint.Y - 1e-6
                && inner.MaxPoint.Y <= outer.MaxPoint.Y + 1e-6;
        }
        [CommandMethod("NORMBLOCKS")]
        public void ListBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                var config = BlocksConfigService.Instance.LoadConfig();

                foreach (var cat in config.Categories)
                {
                    string dwgPath = BlocksConfigService.Instance.GetDwgFullPath(cat.DwgFile);
                    ed.WriteMessage($"\n\n=== {cat.Name} ({cat.DwgFile}) ===");

                    if (!File.Exists(dwgPath))
                    {
                        ed.WriteMessage($"\n  DWG Ð¾Ð»Ð´ÑÐ¾Ð½Ð³Ò¯Ð¹: {dwgPath}");
                        continue;
                    }

                    using (var db = new Database(false, true))
                    {
                        db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, "");
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            foreach (ObjectId id in bt)
                            {
                                var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                                if (!btr.IsLayout)
                                {
                                    string type = btr.IsAnonymous ? "anon" : (btr.IsDynamicBlock ? "DYNAMIC" : "static");
                                    ed.WriteMessage($"\n  [{type}] '{btr.Name}'");
                                }
                            }

                            string match = bt.Has(cat.BlockName) ? "OK" : "ÐžÐ›Ð”Ð¡ÐžÐÐ“Ò®Ð™!";
                            ed.WriteMessage($"\n  Ð¢Ð¾Ñ…Ð¸Ñ€Ð³Ð¾Ð¾: '{cat.BlockName}' -> {match}");
                            tr.Commit();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nÐÐ»Ð´Ð°Ð°: {ex.Message}");
            }
        }

        /// <summary>
        /// NORMPREPLAYERS — Зураг дээрх area polyline болон дугаар текстийг
        /// NORM_ROOM_AREA / NORM_ROOM_NUM layer руу шилжүүлэх
        /// Хэрэглэгч: area polyline-н layer сонгоно -> NORM_ROOM_AREA руу шилжүүлнэ
        ///            дугаар текстийн layer сонгоно -> NORM_ROOM_NUM руу шилжүүлнэ
        /// </summary>
        [CommandMethod("NORMPREPLAYERS")]
        public void PrepareNormLayers()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                using (var lockDoc = doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                    // NORM_ROOM_AREA layer үүсгэх
                    if (!layerTable.Has(RoomBoundaryService.RoomAreaLayerName))
                    {
                        layerTable.UpgradeOpen();
                        var newLayer = new LayerTableRecord();
                        newLayer.Name = RoomBoundaryService.RoomAreaLayerName;
                        newLayer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 30); // orange
                        layerTable.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    // NORM_ROOM_NUM layer үүсгэх
                    if (!layerTable.Has(RoomBoundaryService.RoomNumLayerName))
                    {
                        if (!layerTable.IsWriteEnabled) layerTable.UpgradeOpen();
                        var newLayer = new LayerTableRecord();
                        newLayer.Name = RoomBoundaryService.RoomNumLayerName;
                        newLayer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1); // red
                        layerTable.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    tr.Commit();
                }

                // === 1. Area polyline layer сонгох ===
                ed.WriteMessage($"\n\n=== NORM Layer Prep ===");
                ed.WriteMessage($"\n  Өрөөний хилийн polyline-г агуулсан layer-н нэрийг оруулна уу.");
                ed.WriteMessage($"\n  (Эсвэл polyline-г шууд сонгож болно)");

                var kwOpts = new PromptKeywordOptions(
                    "\n[Layer нэрээр(L)/Polyline сонгох(S)] <Layer>: ");
                kwOpts.Keywords.Add("Layer", "L", "Layer(L)");
                kwOpts.Keywords.Add("Select", "S", "Select(S)");
                kwOpts.Keywords.Default = "Layer";
                kwOpts.AllowNone = true;

                var kwResult = ed.GetKeywords(kwOpts);
                string mode = kwResult.Status == PromptStatus.OK ? kwResult.StringResult : "Layer";

                int polyMoved = 0;
                int textMoved = 0;

                if (mode == "Layer")
                {
                    // Layer нэрээр шилжүүлэх
                    var strOpts = new PromptStringOptions("\n  Area polyline-н layer нэр: ");
                    strOpts.AllowSpaces = true;
                    var strResult = ed.GetString(strOpts);
                    if (strResult.Status != PromptStatus.OK) return;
                    string areaLayerName = strResult.StringResult.Trim();

                    // Дугаар текстийн layer
                    var strOpts2 = new PromptStringOptions("\n  Дугаар текстийн layer нэр: ");
                    strOpts2.AllowSpaces = true;
                    var strResult2 = ed.GetString(strOpts2);
                    if (strResult2.Status != PromptStatus.OK) return;
                    string numLayerName = strResult2.StringResult.Trim();

                    using (var lockDoc = doc.LockDocument())
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        foreach (ObjectId id in ms)
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity == null) continue;

                            // Area polyline шилжүүлэх
                            if (entity is Polyline pl && pl.Closed && pl.NumberOfVertices >= 3 &&
                                string.Equals(pl.Layer, areaLayerName, StringComparison.OrdinalIgnoreCase))
                            {
                                entity.UpgradeOpen();
                                entity.Layer = RoomBoundaryService.RoomAreaLayerName;
                                polyMoved++;
                            }
                            // Дугаар текст шилжүүлэх
                            else if ((entity is DBText || entity is MText) &&
                                string.Equals(entity.Layer, numLayerName, StringComparison.OrdinalIgnoreCase))
                            {
                                entity.UpgradeOpen();
                                entity.Layer = RoomBoundaryService.RoomNumLayerName;
                                textMoved++;
                            }
                        }

                        tr.Commit();
                    }
                }
                else
                {
                    // Polyline-г сонгож шилжүүлэх
                    ed.WriteMessage("\n  Өрөөний хилийн closed polyline-г сонгоно уу (олноор сонгож болно):");
                    var selOpts = new PromptSelectionOptions();
                    selOpts.MessageForAdding = "\n  Polyline сонго: ";
                    var selResult = ed.GetSelection(selOpts);
                    if (selResult.Status == PromptStatus.OK)
                    {
                        using (var lockDoc = doc.LockDocument())
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            foreach (var selObj in selResult.Value)
                            {
                                var ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                                if (ent is Polyline pl && pl.Closed)
                                {
                                    ent.UpgradeOpen();
                                    ent.Layer = RoomBoundaryService.RoomAreaLayerName;
                                    polyMoved++;
                                }
                            }
                            tr.Commit();
                        }
                    }

                    // Дугаар текст сонгох
                    ed.WriteMessage("\n  Одоо дугаар текстүүдийг сонгоно уу:");
                    var selResult2 = ed.GetSelection(selOpts);
                    if (selResult2.Status == PromptStatus.OK)
                    {
                        using (var lockDoc = doc.LockDocument())
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            foreach (var selObj in selResult2.Value)
                            {
                                var ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                                if (ent is DBText || ent is MText)
                                {
                                    ent.UpgradeOpen();
                                    ent.Layer = RoomBoundaryService.RoomNumLayerName;
                                    textMoved++;
                                }
                            }
                            tr.Commit();
                        }
                    }
                }

                ed.WriteMessage($"\n\n  Done: {polyMoved} polylines -> {RoomBoundaryService.RoomAreaLayerName}");
                ed.WriteMessage($"\n        {textMoved} texts -> {RoomBoundaryService.RoomNumLayerName}");
                ed.WriteMessage($"\n  AutoMatch ажиллуулахад PASS 0 100% зөв тааруулна.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nАлдаа: {ex.Message}");
            }
        }
    }
}














