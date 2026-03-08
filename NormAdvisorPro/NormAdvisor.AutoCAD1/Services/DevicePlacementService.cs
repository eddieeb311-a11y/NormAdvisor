using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using NormAdvisor.AutoCAD1.Models;

namespace NormAdvisor.AutoCAD1.Services
{
    /// <summary>
    /// AutoCAD дээр төхөөрөмжийн блок байршуулах сервис
    /// </summary>
    public class DevicePlacementService
    {
        private readonly BlocksConfigService _configService;

        public DevicePlacementService()
        {
            _configService = BlocksConfigService.Instance;
        }

        /// <summary>
        /// Төхөөрөмжийг AutoCAD зураг дээр байршуулах
        /// Хэрэглэгчээс цэг сонгуулж, блок оруулна
        /// </summary>
        public bool PlaceDevice(DeviceCategory category, DeviceInfo device)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return false;

            var ed = doc.Editor;
            var db = doc.Database;

            // DWG файлаас блок ачаалах
            string dwgPath = _configService.GetDwgFullPath(category.DwgFile);
            if (!File.Exists(dwgPath))
            {
                ed.WriteMessage($"\nDWG файл олдсонгүй: {dwgPath}");
                return false;
            }

            // Цэг сонгуулах
            var ptResult = ed.GetPoint($"\n{device.Name} байршуулах цэг сонгоно уу: ");
            if (ptResult.Status != PromptStatus.OK) return false;

            using (var lockDoc = doc.LockDocument())
            {
                ObjectId blockRefId = ObjectId.Null;

                try
                {
                    // Transaction 1: Блок оруулах
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        ObjectId blockId = LoadBlockFromDwg(db, dwgPath, category.BlockName);
                        if (blockId == ObjectId.Null)
                        {
                            ed.WriteMessage($"\nБлок олдсонгүй: {category.BlockName}");
                            tr.Commit();
                            return false;
                        }

                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        var blockRef = new BlockReference(ptResult.Value, blockId);
                        btr.AppendEntity(blockRef);
                        tr.AddNewlyCreatedDBObject(blockRef, true);

                        // Attribute нэмэх
                        AddAttributes(tr, blockRef, blockId, device.TagPrefix);

                        blockRefId = blockRef.ObjectId;
                        tr.Commit();
                    }

                    // Transaction 2: Visibility тохируулах (блок commit болсны дараа)
                    if (blockRefId != ObjectId.Null && !string.IsNullOrEmpty(device.VisibilityState))
                    {
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var blockRef = (BlockReference)tr.GetObject(blockRefId, OpenMode.ForWrite);
                            SetVisibilityState(blockRef, device.VisibilityState, ed);
                            tr.Commit();
                        }
                    }

                    ed.WriteMessage($"\n✓ {device.Name} амжилттай байршууллаа.");
                    return true;
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\nАлдаа: {ex.Message}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Гадаад DWG файлаас блок тодорхойлолт ачаалах
        /// 1) WblockClone - named block байвал бүх блокийг clone хийнэ (dynamic block support)
        /// 2) Database.Insert fallback - named block байхгүй бол model space-ийг блок болгон оруулна
        /// </summary>
        private ObjectId LoadBlockFromDwg(Database targetDb, string dwgPath, string blockName)
        {
            // Блок аль хэдийн байгаа эсэхийг шалгах
            using (var tr = targetDb.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(targetDb.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    var id = bt[blockName];
                    tr.Commit();
                    return id;
                }
                tr.Commit();
            }

            // Гадаад DWG-ээс ачаалах
            bool cloned = false;
            using (var sourceDb = new Database(false, true))
            {
                sourceDb.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, "");

                // Нэрлэсэн блок байгаа эсэх шалгах
                bool hasNamedBlock = false;
                using (var tr = sourceDb.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                    hasNamedBlock = bt.Has(blockName);
                    tr.Commit();
                }

                if (hasNamedBlock)
                {
                    // Named block олдлоо → бүх блокийг clone хийх (anonymous blocks оруулах)
                    var blockIds = new ObjectIdCollection();
                    using (var tr = sourceDb.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                        foreach (ObjectId id in bt)
                        {
                            var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            if (!btr.IsLayout)
                                blockIds.Add(id);
                        }
                        tr.Commit();
                    }

                    if (blockIds.Count > 0)
                    {
                        var mapping = new IdMapping();
                        targetDb.WblockCloneObjects(blockIds, targetDb.BlockTableId,
                            mapping, DuplicateRecordCloning.Ignore, false);
                        cloned = true;
                    }
                }
            }

            // Fallback: WblockClone ашиглаж model space-ийг блок болгох
            if (!cloned)
            {
                // Named block олдоогүй тохиолдолд алгасна
                return ObjectId.Null;
            }

            // Ачаалсан блокын ID олох
            using (var tr = targetDb.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(targetDb.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    var id = bt[blockName];
                    tr.Commit();
                    return id;
                }
                tr.Commit();
            }

            return ObjectId.Null;
        }

        /// <summary>
        /// Dynamic block-ийн visibility state тохируулах
        /// IsDynamicBlock нь WblockClone-ийн дараа false буцааж болно
        /// тиймээс DynamicBlockTableRecord шалгана
        /// </summary>
        private void SetVisibilityState(BlockReference blockRef, string stateName, Editor ed)
        {
            if (string.IsNullOrEmpty(stateName)) return;

            if (blockRef.DynamicBlockTableRecord == ObjectId.Null)
            {
                ed.WriteMessage($"\n[DEBUG] DynamicBlockTableRecord = Null, dynamic block биш");
                return;
            }

            // Боломжит утгуудыг жагсааж, яг тохирохыг олох
            foreach (DynamicBlockReferenceProperty prop in blockRef.DynamicBlockReferencePropertyCollection)
            {
                if (!prop.PropertyName.StartsWith("Visibility")) continue;

                ed.WriteMessage($"\n[DEBUG] Visibility property: '{prop.PropertyName}', current='{prop.Value}', readOnly={prop.ReadOnly}");

                // Боломжит утгууд
                var allowedValues = prop.GetAllowedValues();
                ed.WriteMessage($"\n[DEBUG] Боломжит утгууд ({allowedValues.Length}):");
                string exactMatch = null;
                foreach (var val in allowedValues)
                {
                    string valStr = val.ToString();
                    ed.WriteMessage($"\n[DEBUG]   '{valStr}'");
                    if (string.Equals(valStr, stateName, StringComparison.OrdinalIgnoreCase))
                        exactMatch = valStr;
                }

                if (exactMatch != null)
                {
                    prop.Value = exactMatch;
                    blockRef.RecordGraphicsModified(true);
                    ed.WriteMessage($"\n[DEBUG] SET -> '{exactMatch}', verify='{prop.Value}'");
                }
                else
                {
                    ed.WriteMessage($"\n[DEBUG] '{stateName}' утга боломжит жагсаалтад ОЛДСОНГҮЙ!");
                }
                break;
            }
        }

        /// <summary>
        /// Блокын attribute-уудыг нэмэх
        /// </summary>
        private void AddAttributes(Transaction tr, BlockReference blockRef, ObjectId blockDefId, string tagPrefix)
        {
            var blockDef = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);

            foreach (ObjectId attDefId in blockDef)
            {
                var obj = tr.GetObject(attDefId, OpenMode.ForRead);
                if (obj is AttributeDefinition attDef && !attDef.Constant)
                {
                    var attRef = new AttributeReference();
                    attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);

                    if (attDef.Tag.ToUpper().Contains("TAG") ||
                        attDef.Tag.ToUpper().Contains("NAME") ||
                        attDef.Tag.ToUpper().Contains("NUMBER"))
                    {
                        attRef.TextString = tagPrefix;
                    }

                    blockRef.AttributeCollection.AppendAttribute(attRef);
                    tr.AddNewlyCreatedDBObject(attRef, true);
                }
            }
        }
    }
}
