using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using NormAdvisor.AutoCAD1.Models;

namespace NormAdvisor.AutoCAD1.Services
{
    /// <summary>
    /// Мастер байгуулалт болон хуулбарыг удирдах сервис.
    /// - NOD (Named Objects Dictionary) дээр мастер lock төлөв хадгална
    /// - XData-тай polyline-уудыг скан хийж хуулбар илрүүлнэ
    /// </summary>
    public class MasterCopyService
    {
        private const string DictName = "NORMADVISOR_MASTER";
        private const string KeyLocked = "IS_LOCKED";
        private const string KeyTimestamp = "LOCK_TIME";

        /// <summary>
        /// Мастер lock хийгдсэн эсэх
        /// </summary>
        public bool IsMasterLocked(Database db)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var result = GetLockState(db, tr);
                tr.Commit();
                return result;
            }
        }

        /// <summary>
        /// Мастер lock хийх — бүх өрөөний boundary-д MASTER тэмдэг тавина
        /// </summary>
        public void LockMaster(Database db)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // NOD дээр lock state хадгалах
                var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

                DBDictionary normDict;
                if (nod.Contains(DictName))
                {
                    normDict = (DBDictionary)tr.GetObject(nod.GetAt(DictName), OpenMode.ForWrite);
                }
                else
                {
                    normDict = new DBDictionary();
                    nod.SetAt(DictName, normDict);
                    tr.AddNewlyCreatedDBObject(normDict, true);
                }

                // IS_LOCKED = 1
                SetXrecord(normDict, tr, KeyLocked, 1);
                SetXrecord(normDict, tr, KeyTimestamp, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // Бүх NORMADVISOR polyline-д MASTER тэмдэг нэмэх
                MarkAllAsMaster(db, tr);

                tr.Commit();
            }
        }

        /// <summary>
        /// Мастер unlock хийх
        /// </summary>
        public void UnlockMaster(Database db)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

                if (nod.Contains(DictName))
                {
                    var normDict = (DBDictionary)tr.GetObject(nod.GetAt(DictName), OpenMode.ForWrite);
                    SetXrecord(normDict, tr, KeyLocked, 0);
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Хуулбарын тоо тодорхойлох.
        /// Ижил room number-тэй polyline-ууд олон удаа давтагдаж байвал → хуулбар байна.
        /// Хуулбарын тоо = (нийт group тоо) - 1 (мастер)
        /// </summary>
        public CopyDetectionResult DetectCopies(Database db)
        {
            var result = new CopyDetectionResult();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // Бүх NORMADVISOR polyline олж, room number-ээр нь бүлэглэх
                var roomGroups = new Dictionary<int, List<BoundaryInfo>>();

                foreach (ObjectId id in ms)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    var xdata = entity.GetXDataForApplication(RoomBoundaryService.RegAppName);
                    if (xdata == null) continue;

                    var values = xdata.AsArray();
                    if (values.Length < 5 || values[1].Value.ToString() != "ROOM") continue;

                    int roomNum = Convert.ToInt32(values[2].Value);
                    string roomName = values[3].Value.ToString();
                    bool isMaster = values.Length >= 7 && values[5].Value.ToString() == "MASTER";

                    // Polyline-ийн төвийг олох
                    Point3d center = Point3d.Origin;
                    try
                    {
                        var ext = entity.GeometricExtents;
                        center = new Point3d(
                            (ext.MinPoint.X + ext.MaxPoint.X) / 2,
                            (ext.MinPoint.Y + ext.MaxPoint.Y) / 2,
                            0);
                    }
                    catch { }

                    if (!roomGroups.ContainsKey(roomNum))
                        roomGroups[roomNum] = new List<BoundaryInfo>();

                    roomGroups[roomNum].Add(new BoundaryInfo
                    {
                        ObjectId = id,
                        RoomNumber = roomNum,
                        RoomName = roomName,
                        Center = center,
                        IsMaster = isMaster
                    });
                }

                // Хуулбар тоолох: room number бүрд хэдэн polyline байгааг шалгах
                if (roomGroups.Count > 0)
                {
                    // Хамгийн олон давтагдсан room number-ийн тоог авна
                    int maxCount = roomGroups.Values.Max(g => g.Count);
                    result.CopyCount = maxCount > 1 ? maxCount - 1 : 0;
                    result.TotalBoundaries = roomGroups.Values.Sum(g => g.Count);
                    result.MasterBoundaries = roomGroups.Values.Sum(g => g.Count(b => b.IsMaster));
                    result.RoomGroups = roomGroups;

                    // Хуулбар бүлгүүдийг offset-ээр тодорхойлох
                    if (maxCount > 1)
                    {
                        result.LayoutGroups = DetectLayoutGroups(roomGroups);
                    }
                }

                tr.Commit();
            }

            return result;
        }

        /// <summary>
        /// Polyline-уудыг байршлаар нь бүлэглэх (ижил offset-тэй бол нэг layout)
        /// </summary>
        private List<LayoutGroup> DetectLayoutGroups(Dictionary<int, List<BoundaryInfo>> roomGroups)
        {
            var groups = new List<LayoutGroup>();

            // Хамгийн олон entry-тэй room-ийг reference болгон авах
            var refRoom = roomGroups.OrderByDescending(kv => kv.Value.Count).First();
            int groupCount = refRoom.Value.Count;

            // Reference room-ийн polyline бүр нэг layout бүлэг
            for (int i = 0; i < groupCount; i++)
            {
                var refBoundary = refRoom.Value[i];
                var group = new LayoutGroup
                {
                    Index = i,
                    IsMaster = refBoundary.IsMaster,
                    ReferenceCenter = refBoundary.Center,
                    Boundaries = new List<BoundaryInfo> { refBoundary }
                };

                // Бусад room-уудаас ижил offset-тэй polyline олох
                if (i == 0)
                {
                    // Эхний бүлэг = master (эсвэл MASTER тэмдэгтэй)
                    foreach (var kv in roomGroups)
                    {
                        if (kv.Key == refRoom.Key) continue;
                        var masterBoundary = kv.Value.FirstOrDefault(b => b.IsMaster) ?? kv.Value.FirstOrDefault();
                        if (masterBoundary != null)
                            group.Boundaries.Add(masterBoundary);
                    }
                }

                groups.Add(group);
            }

            // Master бүлгийг эхэнд тавих
            var masterGroup = groups.FirstOrDefault(g => g.IsMaster || g.Index == 0);
            if (masterGroup != null)
                masterGroup.IsMaster = true;

            return groups;
        }

        /// <summary>
        /// Бүх NORMADVISOR polyline-д MASTER тэмдэг нэмэх (XData-д нэмэлт field)
        /// </summary>
        private void MarkAllAsMaster(Database db, Transaction tr)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            var regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!regTable.Has(RoomBoundaryService.RegAppName)) return;

            string groupId = Guid.NewGuid().ToString("N").Substring(0, 8);

            foreach (ObjectId id in ms)
            {
                var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (entity == null) continue;

                var xdata = entity.GetXDataForApplication(RoomBoundaryService.RegAppName);
                if (xdata == null) continue;

                var values = xdata.AsArray();
                if (values.Length < 5 || values[1].Value.ToString() != "ROOM") continue;

                // MASTER тэмдэг + GroupId нэмэх
                entity.UpgradeOpen();
                entity.XData = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, RoomBoundaryService.RegAppName),
                    values[1], // "ROOM"
                    values[2], // RoomNumber
                    values[3], // RoomName
                    values[4], // TableArea
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, "MASTER"),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, groupId)
                );
            }
        }

        private bool GetLockState(Database db, Transaction tr)
        {
            var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
            if (!nod.Contains(DictName)) return false;

            var normDict = (DBDictionary)tr.GetObject(nod.GetAt(DictName), OpenMode.ForRead);
            if (!normDict.Contains(KeyLocked)) return false;

            var xrec = (Xrecord)tr.GetObject(normDict.GetAt(KeyLocked), OpenMode.ForRead);
            var data = xrec.Data;
            if (data == null) return false;

            var vals = data.AsArray();
            return vals.Length > 0 && Convert.ToInt32(vals[0].Value) == 1;
        }

        private void SetXrecord(DBDictionary dict, Transaction tr, string key, object value)
        {
            Xrecord xrec;
            if (dict.Contains(key))
            {
                xrec = (Xrecord)tr.GetObject(dict.GetAt(key), OpenMode.ForWrite);
            }
            else
            {
                xrec = new Xrecord();
                dict.SetAt(key, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }

            short dxfCode = value is int ? (short)DxfCode.Int32 : (short)DxfCode.Text;
            xrec.Data = new ResultBuffer(new TypedValue(dxfCode, value));
        }
    }

    // ========== Models ==========

    public class BoundaryInfo
    {
        public ObjectId ObjectId { get; set; }
        public int RoomNumber { get; set; }
        public string RoomName { get; set; }
        public Point3d Center { get; set; }
        public bool IsMaster { get; set; }
    }

    public class LayoutGroup
    {
        public int Index { get; set; }
        public bool IsMaster { get; set; }
        public Point3d ReferenceCenter { get; set; }
        public List<BoundaryInfo> Boundaries { get; set; } = new List<BoundaryInfo>();
    }

    public class CopyDetectionResult
    {
        public int CopyCount { get; set; }
        public int TotalBoundaries { get; set; }
        public int MasterBoundaries { get; set; }
        public Dictionary<int, List<BoundaryInfo>> RoomGroups { get; set; } = new Dictionary<int, List<BoundaryInfo>>();
        public List<LayoutGroup> LayoutGroups { get; set; } = new List<LayoutGroup>();
    }
}
