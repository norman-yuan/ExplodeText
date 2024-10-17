using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using CadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using static System.Net.Mime.MediaTypeNames;
using Autodesk.AutoCAD.Colors;

namespace ExplodeText
{
    public class TextExploder
    {
        private readonly Dictionary<ObjectId, string> _errDict = new Dictionary<ObjectId, string>();

        private Document _dwg = null;
        private Database _db = null;
        private Editor _ed = null;

        private ObjectId _textId = ObjectId.Null;
        private Extents3d _txtExtents = new Extents3d();
        private List<string> _existingWmfBlockNames;
        private string _layer = "0";
        private double _rotation = 0;
        //private string _textValue = "";

        private ObjectIdCollection _textCurveIds = new ObjectIdCollection();

        public bool ExplodeText(
            Document dwg, ObjectId txtEntId, 
            string explodeToLayer=null, bool eraseTextEnt = true)
        {
            //if (txtEntId.ObjectClass.DxfName.ToUpper() != "TEXT")
            //{
            //    _errDict.Add(txtEntId, "Not a DBText.");
            //}

            _textId= txtEntId;
            _dwg = dwg;
            _ed = _dwg.Editor;
            _db = _dwg.Database;

            string wmfFileName = null;
            string wmfFilePath = null;

            // generate WMF image file
            if (!ExportWmfImage(_textId, out wmfFileName, out wmfFilePath)) return false;

            // import the WMF image as block
            var importFile = $"{wmfFilePath}{wmfFileName}.wmf";
            try
            {
                if (!ImportWmfImage(importFile, _txtExtents.MinPoint))
                {
                    return false;
                }
            }
            finally
            {
                if (File.Exists(importFile)) File.Delete(importFile);
            }

            // process WMF image block
            if (!string.IsNullOrEmpty(explodeToLayer) &&
                LayerExists(explodeToLayer))
            {
                _layer = explodeToLayer;
            }

            _existingWmfBlockNames = GetExistingWmfBlockNames();
            if (!PostWmfImportWork()) return false;

            // erase original text
            if (eraseTextEnt)
            {
                using (var tran = _db.TransactionManager.StartTransaction())
                {
                    var txt = (DBText)tran.GetObject(_textId, OpenMode.ForWrite);
                    txt.Erase();
                    tran.Commit();
                }
            }

            return true;
        }

        #region private methods: epxort DBText as WMF image

        private bool PrepareTextForExport(
            ObjectId txtEntId, 
            out Extents3d origExt,
            out Extents3d zoomExt,
            out string layer, 
            out double rotation)
        {
            origExt = new Extents3d();
            zoomExt = new Extents3d();
            layer = "";
            rotation = 0.0;

            var hasText = true;
            using (var tran=_db.TransactionManager.StartOpenCloseTransaction())
            {
                var txt = (DBText)tran.GetObject(txtEntId, OpenMode.ForWrite);
                if (txt.TextString.Trim().Length > 0)
                {
                    rotation = txt.Rotation;
                    origExt = txt.GeometricExtents;
                    layer = txt.Layer;

                    // rotate the text to horizontal
                    txt.TransformBy(
                        Matrix3d.Rotation(-txt.Rotation, Vector3d.ZAxis, txt.Position));

                    // mirror the text in order for TrueFont text to be converted to curves
                    using (var mirrorLine = new Line3d())
                    {
                        mirrorLine.Set(txt.Position, new Point3d(txt.Position.X, txt.Position.Y+1000, txt.Position.Z));
                        txt.TransformBy(Matrix3d.Mirroring(mirrorLine));
                    }

                    zoomExt = txt.GeometricExtents; 
                }
                else
                {
                    _ed.WriteMessage("\nText entity has empty text string!");
                    hasText = false;
                }
                
                tran.Commit();
            }
            return hasText;
        }

        private bool LayerExists(string layer)
        {
            var exists = false;
            using (var tran = _db.TransactionManager.StartOpenCloseTransaction())
            {
                var layerTable = (LayerTable)tran.GetObject(
                    _db.LayerTableId, OpenMode.ForRead);
                exists = layerTable.Has(layer);
                tran.Commit();
            }
            return exists;
        }

        private bool ExportWmfImage(
            ObjectId txtId, out string wmfFileName, out string wmfFilePath)
        {
            var exportOk = true;

            wmfFileName = string.Empty;
            wmfFilePath = string.Empty;

            var mirrText = (short)CadApp.GetSystemVariable("MIRRTEXT");
            try
            {
                CadApp.SetSystemVariable("MIRRTEXT", 1);

                exportOk = PrepareTextForExport(
                    txtId, out _txtExtents, out Extents3d zoomExt, out _layer, out _rotation);
                if (exportOk)
                {
                    using (var tempView = new WmfZoomedView(_ed, zoomExt))
                    {
                        CadApp.UpdateScreen();

                        dynamic preferences = CadApp.Preferences;
                        wmfFilePath = preferences.Files.TempFilePath;
                        wmfFileName = GetWmfFileName(txtId);
                        exportOk = ExportTextEntity(txtId, zoomExt, wmfFileName, wmfFilePath);
                    }
                }
            }
            finally
            {
                CadApp.SetSystemVariable("MIRRTEXT", mirrText);
            }
            return exportOk;
        }

        private string GetWmfFileName(ObjectId id)
        {
            var name = id.Handle.ToString();
            var i = 1;
            var fileName = "";
            while(true)
            {
                fileName = $"{name}_{i}";
                if (System.IO.File.Exists(fileName))
                {
                    i++;
                }
                else
                {
                    return fileName;
                }
            }
        }

        private bool ExportTextEntity(
            ObjectId txtId, Extents3d zoomExtents, string wmfFile, string wmfFilePath)
        {
            var ok = true;
            dynamic comDoc = _dwg.GetAcadDocument();
            dynamic ss = comDoc.SelectionSets.Add(txtId.Handle.ToString());
            var hiddenEnts = new List<ObjectId>();
            try
            {
                var pt1 = new double[]
                { 
                    zoomExtents.MinPoint.X, 
                    zoomExtents.MinPoint.Y, 
                    zoomExtents.MinPoint.Z 
                };
                var pt2 = new double[] 
                { 
                    zoomExtents.MaxPoint.X, 
                    zoomExtents.MaxPoint.Y, 
                    zoomExtents.MaxPoint.Z 
                };
                ss.Select(0, pt1, pt2);
                
                var count = ss.Count;
                if (count>1)
                {
                    hiddenEnts = HideExtraEntitiesInSelectionSet(ss);
                }

                var fName = $"{wmfFilePath}{wmfFile}.wmf";
                if (File.Exists(fName)) File.Delete(fName);
                comDoc.Export($"{wmfFilePath}{wmfFile}", "wmf", ss);
            }
            catch(System.Exception ex)
            {
                _ed.WriteMessage(
                    $"\nExporting WMF file error:\n{ex.Message}");
                ok = false;
            }
            finally
            {
                ss.Delete();
                if (hiddenEnts.Count>0)
                {

                }
            }

            return ok;
        }

        private List<ObjectId> HideExtraEntitiesInSelectionSet(dynamic selectionSet)
        {
            var ids=new List<ObjectId>();


            return ids;
        }

        #endregion

        #region private methods: import WMF image

        private List<string> GetExistingWmfBlockNames()
        {
            List<string> names;
            using (var tran = _db.TransactionManager.StartTransaction())
            {
                var blkTable = (BlockTable)tran.GetObject(
                    _db.BlockTableId, OpenMode.ForRead);
                names = blkTable.Cast<ObjectId>()
                    .Select(id => ((BlockTableRecord)tran.GetObject(id, OpenMode.ForRead)).Name)
                    .Where(name => name.ToUpper().StartsWith("WMF")).ToList();
                tran.Commit();
            }
            return names;
        }

        private bool ImportWmfImage(string wmfFile, Point3d insPt)
        {
            try
            {
                var pt = new double[] { insPt.X, insPt.Y, insPt.Z };
                dynamic comDoc = _dwg.GetAcadDocument();
                comDoc.Import(wmfFile, pt, 1.0);
                return true;
            }
            catch(System.Exception ex)
            {
                _ed.WriteMessage($"\nImportiing WMF image error:\n{ex.Message}");
                return false;
            }
        }

        #endregion

        #region private methods: Convert imported WMF block into curves

        private bool PostWmfImportWork()
        {
            var ok = true;
            using (var tran = _db.TransactionManager.StartTransaction())
            {
                var space=(BlockTableRecord)tran.GetObject(
                    _db.CurrentSpaceId, OpenMode.ForWrite);
                var wmfBlkRef = FindWmfBlock(space, tran);
                if (wmfBlkRef!=null)
                {
                    wmfBlkRef.UpgradeOpen();        
                    ok = ConvertWmfBlockReferenceToCurves(wmfBlkRef, space, tran);
                    if (ok)
                    {
                        try
                        {
                            wmfBlkRef.Erase();

                            var wmfBlkDef = (BlockTableRecord)tran.GetObject(
                                wmfBlkRef.BlockTableRecord, OpenMode.ForWrite);
                            wmfBlkDef.Erase();
                        }
                        catch { }
                    }
                }
                else
                {
                    _ed.WriteMessage(
                        $"\nCannot find imported WMF image block reference.");
                    ok = false;
                }

                tran.Commit();
            }

            return ok;
        }

        private BlockReference FindWmfBlock(BlockTableRecord space, Transaction tran)
        {
            var blks = new List<BlockReference>();
            foreach (ObjectId id in space)
            {
                if (id.ObjectClass.DxfName.ToUpper() != "INSERT") continue;
                var blk = (BlockReference)tran.GetObject(id, OpenMode.ForRead);
                if (blk.Name.ToUpper().StartsWith("WMF"))
                {
                    blks.Add(blk);
                }
            }

            if (blks.Count>0)
            {
                return (from b in blks orderby b.ObjectId.Handle.Value descending select b).First();
            }
            else
            {
                return null;
            }
        }

        private BlockTableRecord FindImportedWmfBlock(Transaction tran)
        {
            var blkTable = (BlockTable)tran.GetObject(
                _db.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId blkId in blkTable)
            {
                var blk = (BlockTableRecord)tran.GetObject(
                    blkId, OpenMode.ForRead);
                if (blk.Name.ToUpper().StartsWith("WMF"))
                {
                    if (IsImportedWmfBlock(blk.Name)) return blk;
                }
            }
            return null;
        }

        private bool IsImportedWmfBlock(string blkName)
        {
            foreach (var name in _existingWmfBlockNames)
            {
                if (name.ToUpper() == blkName.ToUpper()) return false;
            }
            return true;
        }

        private BlockReference FindImportedWmfBlockReference(
            string blkName, Transaction tran)
        {
            var blkId = ObjectId.Null;

            var space = (BlockTableRecord)tran.GetObject(
                _db.CurrentSpaceId, OpenMode.ForRead);
            foreach (ObjectId entId in space)
            {
                if (entId.ObjectClass.DxfName.ToUpper()=="INSERT")
                {
                    var blk=(BlockReference)tran.GetObject(entId, OpenMode.ForRead);
                    if (blk.Name.ToUpper() == blkName.ToUpper()) return blk;
                }
            }

            return null;
        }

        private bool ConvertWmfBlockReferenceToCurves(
            BlockReference blkRef, BlockTableRecord space, Transaction tran)
        {
            // mirror and rotate the block reference before exploding it
            using (var mirrorLine = new Line3d())
            {
                mirrorLine.Set(blkRef.Position, new Point3d(blkRef.Position.X, blkRef.Position.Y + 1000, blkRef.Position.Z));
                blkRef.TransformBy(Matrix3d.Mirroring(mirrorLine));
            }
            blkRef.TransformBy(Matrix3d.Rotation(_rotation, Vector3d.ZAxis, blkRef.Position));

            // explode the block reference
            var ents = ExplodeWmfBlock(blkRef, tran);

            var wmfExtents = GetWmfExtents(ents);
            var scale = 
                (_txtExtents.MaxPoint.X - _txtExtents.MinPoint.X) / 
                (wmfExtents.MaxPoint.X - wmfExtents.MinPoint.X);
            var mtScale = Matrix3d.Scaling(scale, wmfExtents.MinPoint);
            var mtMove = Matrix3d.Displacement(
                wmfExtents.MinPoint.GetVectorTo(_txtExtents.MinPoint));
            //var mtRotation = Matrix3d.Rotation(
            //    _rotation, Vector3d.ZAxis, _txtExtents.MinPoint);

            foreach (var ent in ents)
            {
                if (!string.IsNullOrEmpty(_layer))
                {
                    ent.Layer = _layer;
                    ent.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);
                }
                ent.TransformBy(mtScale);
                ent.TransformBy(mtMove);
                //ent.TransformBy(mtRotation);

                space.AppendEntity(ent);
                tran.AddNewlyCreatedDBObject(ent, true);
            }

            return true;
        }

        private List<Entity> ExplodeWmfBlock(BlockReference blkRef, Transaction tran)
        {
            var ents=new List<Entity>();
            using (var objs = new DBObjectCollection())
            {
                blkRef.Explode(objs);
                foreach (DBObject obj in objs)
                {
                    var ent = (Entity)obj;
                    
                    ents.Add(ent);
                }
            }
            return ents;
        }

        private Extents3d GetWmfExtents(List<Entity> ents)
        {
            var ext = new Extents3d();
            foreach (var ent in ents)
            {
                ext.AddExtents(ent.GeometricExtents);
            }
            return ext;
        }
        #endregion
    }
}
