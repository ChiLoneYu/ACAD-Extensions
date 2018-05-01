using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System.Drawing;
using System.Text.RegularExpressions;

namespace AcadExts
{
    [rtn("Inserts XREF DWG into host DWG and converts host DWG to new format")]
    internal sealed class XRefInserter : DwgProcessor
    {
        private readonly String _excelpath = String.Empty;
        public String ExcelPath { get { return _excelpath; } }

        List<ExcelAccessor.Row> ExcelList = new List<ExcelAccessor.Row>();
        List<String> extraTMObjects = new List<String>();

        Regex letterNo = new Regex(@"^(\s)*[A-Z]\d{6}(\S)*(\s)*$");
        Regex msRegex = new Regex(@"^(\s)*MS\d{6}(\S)*(\s)*$");
        Regex viewLetterRegex = new Regex(@"^([A-Z]){1,2}$");
        Regex calloutRegex = new Regex(@"^(\d)+");
        Regex refDesRegex = new Regex(@"^[A-Z]+(\S)*(\d)+(\s)*(\S)*$");

        public XRefInserter(String inPath, BackgroundWorker inBw, String inXmlPath)
            : base(inPath, inBw)
        {
            this._excelpath = inXmlPath;
        }

        public override String Process()
        {
            try
            {
                BeforeProcessing();
            }
            catch (System.Exception se)
            {
                //_Logger.Dispose();
                return "XREF Insertion processing exception: " + se.Message;
            }

            // Validate paths
            //if (!CheckDirPath()) { return "Invalid path: " + _Path; }
            if (!ExcelPath.isFilePathOK()) { return "Invalid Excel file: " + ExcelPath; }

            //StartTimer();

            // Open error logger
            //try
            //{
            //    _Logger = new Logger(_Path + "\\InsertXRefErrorLog.txt");
            //}
            //catch (System.Exception se)
            //{
            //    return "Could not create log file in: " + _Path + " because: " + se.Message;
            //}

            // Get dwgs
            try
            {
                GetDwgList(SearchOption.TopDirectoryOnly,
                           delegate(String inFile)
                           {
                               return true;
                               //return (System.IO.Path.GetFileNameWithoutExtension(inFile).Length >= 13);
                           });
            }
            catch (System.Exception se)
            {
                _Logger.Dispose();
                return "Could not access DWG files in: " + _Path + " because: " + se.Message;
            }

            // Get data from excel file
            try
            {
                ExcelList = ExcelAccessor.GetExcelData(ExcelPath);
            }
            catch (System.Exception se)
            {
                _Logger.Dispose();
                return "Could not read Excel file because: " + se.Message;
            }

            try
            {
                foreach (String currentDWGstring in DwgList)
                {
                    String ms = String.Empty;
                    String dwgSavePath = String.Empty;

                    Database oldDb = HostApplicationServices.WorkingDatabase;

                    // Host DWG
                    using (Database db = new Database(false, true))
                    {
                        // Read in DWG file
                        try
                        {
                            db.ReadDwgFile(currentDWGstring, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception e)
                        {
                            _Logger.Log(String.Concat(currentDWGstring + ": Could not read DWG because: ", e.Message));
                            continue;
                        }

                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            ObjectId idROMAND = new ObjectId();

                            #region Get object id for ROMAND font
                            SymbolTable symTable = (SymbolTable)acTrans.GetObject(db.TextStyleTableId, OpenMode.ForRead);

                            foreach (ObjectId id in symTable)
                            {
                                TextStyleTableRecord symbol = (TextStyleTableRecord)acTrans.GetObject(id, OpenMode.ForRead);

                                if (symbol.Name.ToUpper().Contains("ROMAND")) { idROMAND = symbol.ObjectId; }
                            }
                            #endregion

                            ObjectId oidBlockI = new ObjectId();
                            double prefixYVAL = 0.0d;
                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                #region Change font, height and width factor of "Prefix all ref des" and "designations with" block definitions
                                if (bt.Has("PRENOTE2"))
                                {
                                    oidBlockI = bt["PRENOTE2"];
                                    BlockTableRecord btr = acTrans.GetObject(oidBlockI, OpenMode.ForRead) as BlockTableRecord;

                                    foreach (ObjectId idkOid in btr)
                                    {
                                        AttributeDefinition ad = acTrans.GetObject(idkOid, OpenMode.ForWrite) as AttributeDefinition;
                                        if (ad.TextString.ToUpper().Contains("PREFIX ALL"))
                                        {
                                            ad.WidthFactor = .75d;
                                            prefixYVAL = ad.Position.Y;
                                        }
                                        if (ad.TextString.ToUpper().Contains("DESIGNATIONS"))
                                        {
                                            ad.WidthFactor = .75d;
                                            ad.Position = new Point3d(ad.Position.X, prefixYVAL - .1133, 0);
                                        }
                                        ad.TextStyleId = idROMAND;
                                        ad.Height = .07;
                                    }
                                }
                                #endregion

                                #region Change font of value of prenote
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord)
                                {
                                    foreach (ObjectId oid in btr)
                                    {
                                        Entity ent = acTrans.GetObject(oid, OpenMode.ForRead) as Entity;

                                        if (ent.GetType() == typeof(BlockReference))
                                        {
                                            BlockReference br = ent as BlockReference;

                                            if (!br.Name.ToUpper().Contains("PRENOTE2")) { continue; }

                                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection ac = br.AttributeCollection;

                                            foreach (ObjectId oidac in ac)
                                            {
                                                DBObject dbObj = acTrans.GetObject(oidac, OpenMode.ForWrite) as DBObject;
                                                AttributeReference acAttRef = dbObj as AttributeReference;
                                                acAttRef.TextStyleId = idROMAND;
                                                acAttRef.Height = .07;
                                                acAttRef.WidthFactor = .75d;
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }

                            try
                            {
                                // Name new file after DWG MS#
                                ms = Utilities.getDwgMsNum(db, acTrans);
                                dwgSavePath = System.IO.Path.GetDirectoryName(currentDWGstring) + "\\" + ms.Replace("S", "").ToLower() + ".dwg";
                            }
                            catch { _Logger.Log("Skipping file: Could not find MS# for DWG: " + currentDWGstring); acTrans.Commit(); continue; }

                            #region Make 2 new custom properties in new dwg for original file name and time stamp
                            DatabaseSummaryInfoBuilder dsib = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
                            IDictionary customPropTable = dsib.CustomPropertyTable;

                            customPropTable.Add("RTN_FILENAME", currentDWGstring);
                            customPropTable.Add("RTN_CONVERSION_DATE", DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"));

                            db.SummaryInfo = dsib.ToDatabaseSummaryInfo();

                            #endregion

                            Point3d ExtMinBefore = db.Extmin;
                            Point3d ExtMaxBefore = db.Extmax;

                            #region Get all extra objects that are not any of the specified types, and add them all to extraTMObjects list
                            //foreach (Entity ent in ToEach.toEachEntity(db, acTrans, delegate(Entity inEnt)
                            //{
                            //    if (inEnt.Layer.isFramePageLayerName())
                            //    {
                            //        if (inEnt.GetType() == typeof(DBText))
                            //        {
                            //            DBText dbt = inEnt as DBText;
                            //            String ut = dbt.TextString.Trim().ToUpper();
                            //            return !((ut.Contains("TM") && (ut.Contains("-")) && ut.HasNums()) ||
                            //                     (ut.StartsWith("IV") && ut.Contains("-")) ||
                            //                     (ut.Contains("PREFIX ALL") && ut.Contains("DESIGNATIONS WITH")));
                            //        }
                            //        if (inEnt.GetType() == typeof(BlockReference))
                            //        {
                            //            BlockReference br = inEnt as BlockReference;
                            //            return !(br.Name.Trim().ToUpper().StartsWith("FCL1") ||
                            //                     br.Name.Trim().ToUpper().Contains("PRENOTE"));
                            //        }
                            //        return true;
                            //    }
                            //    return false;
                            //}))
                            //{
                            //    if (ent.GetType() == typeof(DBText)) { String type = ent.ToString(); DBText dbt = ent as DBText; extraTMObjects.Add(currentDWGstring + " (" + ms + "): Frame Page Layer: " + ent.Layer + " contains extra entity: " + type.Substring(type.LastIndexOf('.') + 1) + " : " + dbt.TextString); }
                            //    else
                            //    {
                            //        String type = ent.ToString(); extraTMObjects.Add(currentDWGstring + " (" + ms + "): Frame Page Layer: " + ent.Layer + " contains extra entity: " + type.Substring(type.LastIndexOf('.') + 1));
                            //    }
                            //}
                            #endregion

                            #region add GTYPE layer if it doesnt already exist and (0,0,0) point
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable)
                            {
                                if (lt.Has("_GTYPE_RPSTL")) { _Logger.Log(currentDWGstring + " (" + ms + "): DWG already has a _GTYPE_RPSTL layer"); }

                                // Add GTYPE layer if dwg doesnt have it already
                                if (!lt.Has("_GTYPE_RPSTL"))
                                {
                                    using (LayerTableRecord newGTypeRec = new LayerTableRecord())
                                    {
                                        newGTypeRec.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.White);
                                        newGTypeRec.Name = "_GTYPE_RPSTL";

                                        // Upgrade lt layertable from forread to forwrite
                                        // lt.UpgradeOpen();

                                        // Add layertable record to layer table and add to transaction
                                        lt.Add(newGTypeRec);
                                        acTrans.AddNewlyCreatedDBObject(newGTypeRec, true);

                                        // Add (0,0) point to the new GTYPE layer
                                        using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                                        {
                                            using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                            {
                                                DBPoint point = new DBPoint(new Point3d(0, 0, 0));

                                                try
                                                {
                                                    db.Clayer = lt["_GTYPE_RPSTL"];
                                                    btr.AppendEntity(point);
                                                    acTrans.AddNewlyCreatedDBObject(point, true);
                                                    point.Layer = "_GTYPE_RPSTL";
                                                }
                                                catch (System.Exception se) { _Logger.Log(currentDWGstring + " (" + ms + "): Could not add GTYPE layer: " + se.Message); }
                                            }
                                        }
                                        newGTypeRec.IsFrozen = true;
                                        newGTypeRec.IsLocked = true;
                                    }
                                }
                            }
                            #endregion

                            #region delete fig cap and IV# and book code
                            using (ObjectIdCollection oidc = new ObjectIdCollection())
                            {
                                // TODO getalldwgobjects is throwing exception for objects on locked layers
                                foreach (DBObject dbo in Utilities.getAllDwgObjs(db, acTrans, OpenMode.ForWrite, _Logger))
                                {
                                    //Entity currentEnt = acTrans.GetObject(dbo.Id, OpenMode.ForWrite) as Entity;
                                    using (Entity currentEnt = dbo as Entity)
                                    {
                                        if (currentEnt.GetType() == typeof(BlockReference))
                                        {
                                            //try{
                                            using (BlockReference br = acTrans.GetObject(dbo.Id, OpenMode.ForWrite) as BlockReference)
                                            {
                                                if (br.Name.Trim().ToUpper().StartsWith("FCL1"))
                                                {
                                                    // NOTE erase must be passed true
                                                    br.Erase(true);
                                                    oidc.Add(br.Id);
                                                }
                                            }
                                            //}catch { }
                                        }

                                        if (currentEnt.GetType() == typeof(DBText))
                                        {
                                            using (DBText dbt = acTrans.GetObject(currentEnt.Id, OpenMode.ForWrite) as DBText)
                                            {
                                                if (dbt.Position.Y <= 1.125 || dbt.Position.Y >= 10.375)
                                                {
                                                    // NOTE erase must be passed true
                                                    dbt.Erase(true);
                                                    oidc.Add(dbt.Id);
                                                }
                                            }
                                        }
                                    }
                                }
                                db.Purge(oidc);
                            }
                            #endregion

                            Boolean FPsShouldHavePrefixes = false;

                            #region Is there any prefix found on any TM layer
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        if (ltr.Name.isFramePageLayerName())
                                        {
                                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                                            {
                                                foreach (ObjectId oidBT in bt)
                                                {
                                                    using (BlockTableRecord btr = acTrans.GetObject(oidBT, OpenMode.ForRead) as BlockTableRecord)
                                                    {
                                                        foreach (ObjectId btrOID in btr)
                                                        {
                                                            using (Entity ent = acTrans.GetObject(btrOID, OpenMode.ForRead) as Entity)
                                                            {
                                                                if (ent.GetType() == typeof(DBText) && (ent.Layer == ltr.Name))
                                                                {
                                                                    DBText entDBText = acTrans.GetObject(btrOID, OpenMode.ForRead) as DBText;
                                                                    String entString = entDBText.TextString.Trim().ToUpper();

                                                                    if ((entString.Contains("PREFIX")) || (entString.Contains("DESIGNATIONS")))
                                                                    {
                                                                        FPsShouldHavePrefixes = true;
                                                                    }
                                                                }
                                                                if (ent.GetType() == typeof(BlockReference) && (ent.Layer == ltr.Name))
                                                                {
                                                                    BlockReference br = ent as BlockReference;

                                                                    if (br.Name.Trim().ToUpper().Contains("PRENOTE"))
                                                                    {
                                                                        FPsShouldHavePrefixes = true;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            Boolean dwgHasFramePageLayersWithoutPrefixes = false;

                            #region do all frame pages have prefix notes if they should?

                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        if (ltr.Name.isFramePageLayerName())
                                        {
                                            // Layer is a frame page, look for a prefix
                                            Boolean prefixFoundOnThisLayer = false;

                                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                                            {
                                                foreach (ObjectId oidBT in bt)
                                                {
                                                    using (BlockTableRecord btr = acTrans.GetObject(oidBT, OpenMode.ForRead) as BlockTableRecord)
                                                    {
                                                        foreach (ObjectId btrOID in btr)
                                                        {
                                                            using (Entity ent = acTrans.GetObject(btrOID, OpenMode.ForRead) as Entity)
                                                            {
                                                                if (ent.GetType() == typeof(DBText) && (ltr.Name == ent.Layer))
                                                                {
                                                                    DBText entDBText = acTrans.GetObject(btrOID, OpenMode.ForRead) as DBText;
                                                                    String entString = entDBText.TextString.Trim().ToUpper();

                                                                    if ((entString.Contains("PREFIX")) || (entString.Contains("DESIGNATIONS")))
                                                                    {
                                                                        prefixFoundOnThisLayer = true;
                                                                    }
                                                                }
                                                                if (ent.GetType() == typeof(BlockReference) && (ltr.Name == ent.Layer))
                                                                {
                                                                    BlockReference br = ent as BlockReference;
                                                                    String entsLayerName = br.Layer.Trim().ToUpper();

                                                                    if (br.Name.Trim().ToUpper().Contains("PRENOTE"))
                                                                    {
                                                                        prefixFoundOnThisLayer = true;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // If a frame page is found that doesnt have a prefix but other frame pages in this dwg do have prefixes, log error
                                            if (!prefixFoundOnThisLayer && FPsShouldHavePrefixes) { dwgHasFramePageLayersWithoutPrefixes = true; /*myLogger.Log("Frame page without prefix detected: " + currentDWG.FileNameNoExt + " : " + ltr.Name);*/ }
                                        }
                                    }
                                }
                            }

                            #endregion

                            if (dwgHasFramePageLayersWithoutPrefixes) { _Logger.Log(currentDWGstring + " (" + ms + "): has 1 or more TM layers that are missing prefixes"); }

                            #region if no prefixes on fps, delete them and make ms-0 layer
                            ObjectIdCollection framePagesToDelete = new ObjectIdCollection();
                            List<String> UsedMLayers = new List<String>();

                            if (!FPsShouldHavePrefixes)
                            {
                                using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable)
                                {
                                    // delete frame page layers
                                    foreach (ObjectId oid in lt)
                                    {
                                        LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForWrite) as LayerTableRecord;
                                        if (ltr.Name.isFramePageLayerName())
                                        {
                                            ltr.Erase(true);
                                            framePagesToDelete.Add(oid);
                                        }
                                    }
                                    // make new layer with name ms#-0
                                    using (LayerTableRecord LayerNew = new LayerTableRecord())
                                    {
                                        String newLayerName = LayerNew.Name = ms.Replace("S", "").ToLower() + "-0";
                                        LayerNew.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                        lt.Add(LayerNew);
                                        acTrans.AddNewlyCreatedDBObject(LayerNew, true);
                                        UsedMLayers.Add(newLayerName);
                                    }
                                }
                                db.Purge(framePagesToDelete);
                            }

                            #endregion

                            #region change tm layers to ms-x layers
                            if (FPsShouldHavePrefixes)
                            {
                                //List<String> Legacy = ExcelList["LegacyLayer"].OrderByDescending( x => x.Length).ToList();

                                ObjectIdCollection oidsFordelete = new ObjectIdCollection();

                                Dictionary<String, String> PriorityChart = new Dictionary<String, String>();

                                using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                                {
                                    foreach (ObjectId layerId in lt)
                                    {
                                        using (LayerTableRecord layer = acTrans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord)
                                        {
                                            try
                                            {
                                                if (layer.Name.Trim().isFramePageLayerName())
                                                {
                                                    String newLayerCode;

                                                    foreach (ExcelAccessor.Row currentRow in ExcelList)
                                                    {
                                                        if (layer.Name.Trim().StartsWith(currentRow.LegacyKey))
                                                        {
                                                            newLayerCode = currentRow.NewLayer;
                                                            int PriorityCode = currentRow.Priority;

                                                            if (PriorityChart.ContainsKey(newLayerCode))
                                                            {
                                                                if (Convert.ToInt32(PriorityChart[newLayerCode]) < PriorityCode)
                                                                {
                                                                    try { Utilities.delLayer(acTrans, db, layer.Name); }
                                                                    catch { }
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    PriorityChart.Remove(newLayerCode);
                                                                }
                                                            }

                                                            PriorityChart.Add(newLayerCode, PriorityCode.ToString());

                                                            String newLayerName = ms.Replace("S", "").ToLower() + "-" + newLayerCode;

                                                            ObjectIdCollection oidc = new ObjectIdCollection();
                                                            if (lt.Has(newLayerName))
                                                            {
                                                                UsedMLayers.Remove(newLayerName);
                                                                db.Clayer = lt["0"];
                                                                Utilities.delLayer(acTrans, db, newLayerName);
                                                                //layer.Erase();
                                                                //oidc.Add(layer.Id);
                                                            }
                                                            db.Purge(oidc);

                                                            try { layer.Name = newLayerName; layer.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red); }
                                                            catch
                                                            { _Logger.Log(currentDWGstring + " (" + ms + "): " + layer.Name + " can't be replaced with " + newLayerName); }

                                                            layer.IsFrozen = true;
                                                            // if (!UsedMLayers.Contains(newLayerName)) { UsedMLayers.Add(newLayerName); }
                                                            // need to add layer to layertable like in ms-0 funcito nwhe nthere are ne prefixes ??
                                                            UsedMLayers.Add(newLayerName);

                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            catch
                                            {
                                                _Logger.Log(currentDWGstring + " (" + ms + "): DWG Processing error occured converting TM layers to m#-x layers ");
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region insert xref
                            // Find XRef dwg
                            XrefGraph graph = db.GetHostDwgXrefGraph(true);
                            for (int i = 0; i < graph.NumNodes; i++)
                            {
                                ObjectIdCollection oic = new ObjectIdCollection();
                                XrefGraphNode node = graph.GetXrefNode(i);

                                //if (node.IsNested) { /*ed.WriteMessage(nl + "XREF nested node found: " + node.Name + ". continuing to next node..." + nl);*/ }
                                //else
                                if (!node.IsNested)
                                {
                                    // ed.WriteMessage(nl +"XREF node found: " + node.Name + nl);
                                    String xRefName = node.Name;
                                    String pathName = "";

                                    using (BlockTableRecord btr = (BlockTableRecord)acTrans.GetObject(node.BlockTableRecordId, OpenMode.ForWrite))
                                    {
                                        xRefName = node.Name;
                                        pathName = btr.PathName;
                                    }

                                    // Open the host Block table for read
                                    //BlockTable acBlkTbl;
                                    using (BlockTable acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                                    {
                                        // Open the host Block table record Model space for write
                                        //BlockTableRecord acBlkTblRec;
                                        using (BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                        {
                                            // XRef DWG
                                            using (Database dbXRef = new Database(false, true))
                                            {
                                                // Read in DWG file
                                                try { dbXRef.ReadDwgFile(pathName, System.IO.FileShare.Read, true, String.Empty); }
                                                catch (System.Exception e)
                                                {
                                                    _Logger.Log(String.Concat(currentDWGstring + " (" + ms + "): Could not read XREF: ", pathName, " because: ", e.Message));
                                                    continue;
                                                }

                                                using (Transaction transXRef = dbXRef.TransactionManager.StartTransaction())
                                                {
                                                    BlockTable btDB = transXRef.GetObject(dbXRef.BlockTableId, OpenMode.ForRead, false) as BlockTable;
                                                    // Only want model space 
                                                    //foreach (ObjectId btrId in btDB)
                                                    //{
                                                    BlockTableRecord btrDB = transXRef.GetObject(btDB[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                                    // NOTE Using this line (V) instead will cause a random white line to appear on the drawing instead of the "For" text
                                                    //BlockTableRecord btrDB = transXRef.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                                                    //{
                                                    foreach (ObjectId id in btrDB)
                                                    {
                                                        oic.Add(id);
                                                    }
                                                    //}
                                                    // }

                                                    IdMapping acIdMap = new IdMapping();

                                                    // clone xref dwg objects to host dwg
                                                    db.WblockCloneObjects(oic, acBlkTblRec.ObjectId, acIdMap, DuplicateRecordCloning.Replace, false);
                                                }
                                            }
                                        }
                                    }
                                    //System.Windows.Forms.MessageBox.Show(xRefName + "\t" + pathName);
                                    try { db.DetachXref(node.BlockTableRecordId); }
                                    catch { _Logger.Log(currentDWGstring + " (" + ms + "): Could not detach XREF: " + node.Name /*pathName*/); continue; }
                                }
                            }
                            #endregion

                            Point3d ExtMinAfter = new Point3d();
                            Point3d ExtMaxAfter = new Point3d();

                            List<Entity> entArray = new List<Entity>();

                            ObjectIdCollection oidsDelete = new ObjectIdCollection();

                            #region delete stuff to calculate extents

                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                {
                                    foreach (ObjectId oid in btr)
                                    {
                                        Entity ent = acTrans.GetObject(oid, OpenMode.ForWrite) as Entity;

                                        if (ent.GetType() == typeof(DBText))
                                        {
                                            DBText dbt = ent as DBText;

                                            if (dbt.TextString.ToUpper().Contains("PREFIX") ||
                                                dbt.TextString.ToUpper().Contains("DESIGNATION") ||
                                                dbt.TextString.ToUpper().Contains("MS"))
                                            {
                                                entArray.Add(ent.Clone() as Entity);
                                                oidsDelete.Add(dbt.Id);
                                                dbt.Erase(true);
                                            }
                                        }
                                    }
                                }
                            }
                            db.Purge(oidsDelete);
                            #endregion

                            db.UpdateExt(true);

                            ExtMinAfter = db.Extmin;
                            ExtMaxAfter = db.Extmax;

                            // Host DWG

                            BlockTable blockTab = acTrans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                            BlockTableRecord blockTabRec = acTrans.GetObject(blockTab[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                            #region put back prefix des and ms
                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                {
                                    foreach (Entity ent in entArray)
                                    {
                                        btr.AppendEntity(ent);
                                        acTrans.AddNewlyCreatedDBObject(ent, true);
                                    }
                                }
                            }
                            #endregion

                            // Move callouts, leader lines, REFDESs, and view letters from COUT layer to each m#-x layer 
                            #region move view letters, REFDESs, callouts, REFDES prefix, and leader lines from COUT to ms-x layer
                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                foreach (ObjectId oid in bt)
                                {
                                    BlockTableRecord btr = acTrans.GetObject(oid, OpenMode.ForRead) as BlockTableRecord;

                                    foreach (ObjectId oidEnt in btr)
                                    {
                                        Entity ent = acTrans.GetObject(oidEnt, OpenMode.ForRead) as Entity;

                                        if (ent.GetType() == typeof(DBText))
                                        {
                                            DBText dbt = ent as DBText;
                                            String dbtString = dbt.TextString.Trim().ToUpper();

                                            if (ent.Layer == "COUT")
                                            {
                                                /*if (!((refDesRegex.IsMatch(dbtString)) ||
                                                   (viewLetterRegex.IsMatch(dbtString)) ||
                                                   ((calloutRegex.IsMatch(dbtString)) && (!dbtString.Contains("-")))))
                                                {
                                                    //  myLogger.Log(currentDWG.FileNameNoExt + ": Copying unknown text from COUT to m#-x layer: " + dbt.TextString);
                                                }*/

                                                foreach (String msLayer in UsedMLayers)
                                                {
                                                    Entity clone = ent.Clone() as Entity;
                                                    if (calloutRegex.IsMatch(dbtString)) { DBText dbt2 = ent as DBText; dbt2.Height = .07; dbt2.TextStyleId = idROMAND; }

                                                    // clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                    btr.AppendEntity(clone);
                                                    acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                    // NOTE this layer change statement must come after the append entity and addnewlycreateddbobject calls or key not found Exception will be thrown

                                                    try { clone.Layer = msLayer; }
                                                    catch
                                                    {
                                                        _Logger.Log(currentDWGstring + " (" + ms + "): Could not move text object ( " + dbt.TextString + " ) from COUT to new layer: " + msLayer);
                                                        #region for debugging
                                                        //// For debugging
                                                        //List<String> currentLayers = new List<string>();
                                                        //using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                                                        //{
                                                        //    foreach (ObjectId oidltr in lt)
                                                        //    {
                                                        //        using (LayerTableRecord ltr = acTrans.GetObject(oidltr, OpenMode.ForRead) as LayerTableRecord)
                                                        //        {
                                                        //            currentLayers.Add(ltr.Name);
                                                        //        }
                                                        //    }
                                                        //}
                                                        #endregion
                                                    }
                                                }
                                            }
                                            continue;
                                        }

                                        if ((ent.GetType() == typeof(Polyline)) && (ent.Layer == "COUT"))
                                        {
                                            // Polyline line = ent as Polyline;
                                            foreach (String msLayer in UsedMLayers)
                                            {
                                                Entity clone = ent.Clone() as Entity;
                                                //ccc
                                                //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                btr.AppendEntity(clone);
                                                acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                clone.Layer = msLayer;
                                            }
                                            continue;
                                        }

                                        if ((ent.GetType() == typeof(AttributeDefinition)) && (ent.Layer == "COUT"))
                                        {
                                            // Polyline line = ent as Polyline;
                                            foreach (String msLayer in UsedMLayers)
                                            {
                                                Entity clone = ent.Clone() as Entity;
                                                //ccc
                                                //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                btr.AppendEntity(clone);
                                                acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                clone.Layer = msLayer;
                                            }
                                            continue;
                                        }

                                        if ((ent.GetType() == typeof(Polyline2d)) && (ent.Layer == "COUT"))
                                        {
                                            //Polyline2d line = ent as Polyline2d;
                                            foreach (String msLayer in UsedMLayers)
                                            {
                                                Entity clone = ent.Clone() as Entity;
                                                //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                btr.AppendEntity(clone);
                                                acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                clone.Layer = msLayer;
                                            }
                                            continue;
                                        }

                                        if (ent.GetType() == typeof(BlockReference) && ent.Layer == "COUT")
                                        {
                                            //Entity br = ent as Entity;
                                            BlockReference br = ent as BlockReference;
                                            br.UpgradeOpen();

                                            //br.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0);
                                            //br.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);

                                            ent.UpgradeOpen();

                                            if (String.Equals(br.Name, "ar") || String.Equals(br.Name, "AR"))
                                            {
                                                foreach (String msLayer in UsedMLayers)
                                                {
                                                    Entity clone = ent.Clone() as Entity;
                                                    //clone.Color = Color.FromColorIndex(ColorMethod.ByBlock, 0);
                                                    //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                    //clone.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255,0,0);
                                                    btr.AppendEntity(clone);
                                                    acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                    clone.Layer = msLayer;
                                                }
                                            }
                                            continue;
                                        }
                                        //  ed.Regen();

                                        //if (ent.GetType() == typeof(DBText) && ent.Layer == "COUT")
                                        //{
                                        //    foreach (String msLayer in UsedMLayers)
                                        //    {
                                        //        Entity clone = ent.Clone() as Entity;
                                        //        clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                        //        btr.AppendEntity(clone);
                                        //        acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                        //        // NOTE this layer change statement must come after the append entity and addnewlycreateddbobject calls or key not found Exception will be thrown
                                        //        clone.Layer = msLayer;
                                        //    }
                                        //    continue;
                                        //}

                                        if (ent.GetType() == typeof(Line) && ent.Layer == "COUT")
                                        {
                                            foreach (String msLayer in UsedMLayers)
                                            {
                                                Entity clone = ent.Clone() as Entity;

                                                //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                btr.AppendEntity(clone);
                                                acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                // NOTE this layer change statement must come after the append entity and addnewlycreateddbobject calls or key not found Exception will be thrown
                                                clone.Layer = msLayer;
                                            }
                                            continue;
                                        }

                                        if (ent.GetType() == typeof(Solid) && ent.Layer == "COUT")
                                        {
                                            foreach (String msLayer in UsedMLayers)
                                            {
                                                Entity clone = ent.Clone() as Entity;

                                                //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                btr.AppendEntity(clone);
                                                acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                // NOTE this layer change statement must come after the append entity and addnewlycreateddbobject calls or key not found Exception will be thrown
                                                clone.Layer = msLayer;
                                            }
                                            continue;
                                        }

                                        if (ent.GetType() == typeof(Arc) && ent.Layer == "COUT")
                                        {
                                            foreach (String msLayer in UsedMLayers)
                                            {
                                                Entity clone = ent.Clone() as Entity;

                                                //clone.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                                                btr.AppendEntity(clone);
                                                acTrans.AddNewlyCreatedDBObject(clone as DBObject, true);
                                                // NOTE this layer change statement must come after the append entity and addnewlycreateddbobject calls or key not found Exception will be thrown
                                                clone.Layer = msLayer;
                                            }
                                            continue;
                                        }

                                        if (ent.Layer == "COUT")
                                        {
                                            _Logger.Log(currentDWGstring + " (" + ms + "): Attempting to copy unknown object to m#-x layer of type: " + ent.GetType().ToString() + " on layer: " + ent.Layer);
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region move and check prefix

                            bool prefixExists = false;
                            double textY = 0;

                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                {
                                    foreach (ObjectId id in btr)
                                    {
                                        Entity ent = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;

                                        if (ent.GetType() == typeof(DBText))
                                        {
                                            DBText textobj = ent as DBText;
                                            string text = textobj.TextString.ToUpper().Trim();

                                            if (text.Contains("PREFIX"))
                                            {
                                                prefixExists = true;
                                                //prefixAllYVal = textobj.Position.Y;
                                                textobj.WidthFactor = .75d;
                                                textobj.Position = new Point3d(ExtMinAfter.X, ExtMinAfter.Y - .17, 0);
                                                textobj.Height = .07;
                                                textobj.TextStyleId = idROMAND;
                                                textY = textobj.Position.Y - .2233;
                                                continue;
                                            }
                                            if (text.Contains("DESIGNATIONS"))
                                            {
                                                prefixExists = true;
                                                textobj.WidthFactor = .75d;
                                                textobj.Position = new Point3d(ExtMinAfter.X, ExtMinAfter.Y - .2833/* .4619 */, 0);
                                                textobj.Height = .07;
                                                textobj.TextStyleId = idROMAND;
                                                textY = textobj.Position.Y - .11;

                                                //List<DBText> ThirdLineTexts = textobj.getThirdLinePrefixT(db, acTrans);

                                                continue;
                                            }
                                        }
                                        if (ent.GetType() == typeof(BlockReference))
                                        {
                                            // NOTE BUG -- set textY in here
                                            BlockReference br = ent as BlockReference;

                                            //if (br.Name.ToUpper().Contains("PRENOTE")) { myLogger.Log("Block Reference dwg prefix found in dwg: " + currentDWG.FileNameNoExt); }

                                            if (br.Name.ToUpper().Contains("PRENOTE3"))
                                            {
                                                //dwgHasOneorMorePreNote3s = true;
                                                _Logger.Log(currentDWGstring + " (" + ms + "): Prenote 3 found in layer: " + br.Layer);
                                            }

                                            if (br.Name.Trim().ToUpper().Contains("PRENOTE2"))
                                            {
                                                //System.Windows.Forms.MessageBox.Show(currentDWG.FileNameNoExt);
                                                prefixExists = true;
                                                Point3d originalBRpos = br.Position;
                                                br.Position = new Point3d(ExtMinAfter.X, ExtMinAfter.Y - .17, 0);
                                                textY = br.Position.Y - .2233d;// -.2719;

                                                Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol = br.AttributeCollection;

                                                foreach (ObjectId objID in attCol)
                                                {
                                                    DBObject dbObj = acTrans.GetObject(objID, OpenMode.ForWrite) as DBObject;
                                                    AttributeReference acAttRef = dbObj as AttributeReference;
                                                    acAttRef.Position = new Point3d(br.Position.X + .845, br.Position.Y - .1133, 0);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region move MS

                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                //foreach (ObjectId oid in bt)
                                //{
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                {
                                    foreach (ObjectId id in btr)
                                    {
                                        Entity ent = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;
                                        if (ent.GetType() == typeof(DBText))
                                        {
                                            DBText MS = ent as DBText;
                                            String msString = MS.TextString.Trim().ToUpper();

                                            if (msString.StartsWith("MS") || letterNo.IsMatch(msString))
                                            {
                                                MS.TextStyleId = idROMAND;
                                                MS.Height = .06;
                                                MS.WidthFactor = .75d;

                                                // Get MS # width
                                                double width = MS.TextWidth();//db, acTrans);

                                                if (width <= 0)
                                                {
                                                    _Logger.Log(currentDWGstring + " (" + ms + "): MS# with unknown width detected: " + msString);
                                                    continue;
                                                }

                                                if (prefixExists)
                                                {
                                                    //MS.moveTextTo(acTrans, db, ExtMinAfter.X - width, textY);
                                                    //MS.Justify = AttachmentPoint.MiddleMid;
                                                    MS.Position = new Point3d(ExtMaxAfter.X - width, textY, 0);
                                                    //MS.Justify = AttachmentPoint.MiddleMid;
                                                }
                                                else
                                                {
                                                    //MS.moveTextTo(acTrans, db, ExtMinAfter.X - width, ExtMinAfter.Y - .11D);
                                                    //MS.Justify = AttachmentPoint.MiddleMid;
                                                    MS.Position = new Point3d(ExtMaxAfter.X - width, ExtMinAfter.Y - .11d, 0);
                                                    //MS.Justify = AttachmentPoint.MiddleMid;
                                                }

                                                //  MS.Justify = AttachmentPoint.BaseRight;
                                                //  MS.AlignmentPoint = new Point3d(ExtMaxAfter.X, MS.Position.Y , 0);
                                            }
                                        }
                                    }
                                }
                            }

                            #endregion

                            #region copy ms to m# layers and delete from 0 layer
                            ObjectIdCollection OriginalMsId = new ObjectIdCollection();
                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                                {
                                    foreach (ObjectId id in btr)
                                    {
                                        Entity ent = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;
                                        if (ent.GetType() == typeof(DBText))
                                        {
                                            DBText dbt = ent as DBText;
                                            if ((ent.Layer == "0") && (dbt.TextString.Trim().ToUpper().StartsWith("MS") || letterNo.IsMatch(dbt.TextString.ToUpper().Trim())))
                                            {
                                                // found ms num

                                                // Copy it to every m# layer
                                                foreach (String mLayer in UsedMLayers)
                                                {
                                                    DBText dbtClone = ent.Clone() as DBText;

                                                    //  if (mLayer.Substring(mLayer.LastIndexOf("-") + 1).Length > 1) { /* 2 #s after dash */ }

                                                    btr.AppendEntity(dbtClone as Entity);
                                                    acTrans.AddNewlyCreatedDBObject(dbtClone as DBObject, true);
                                                    // NOTE this layer change statement must come after the append entity and addnewlycreateddbobject calls or key not found Exception will be thrown
                                                    dbtClone.Layer = mLayer;
                                                    dbtClone.TextString = dbt.TextString + "-" + mLayer.Substring(mLayer.LastIndexOf('-') + 1);
                                                }
                                                ent.Erase(true);
                                                //   dbt.Erase(true);
                                                //OriginalMsId.Add(dbt.Id);
                                                OriginalMsId.Add(ent.Id);
                                                // break;
                                            }
                                        }
                                    }
                                    db.Purge(OriginalMsId);
                                }
                            }
                            #endregion

                            #region lock and unfreeze 0 layer and make it current

                            // Pass anonymous method to toeach method for it to use to lock COUT and 0 layers
                            try
                            {
                                ToEach.toEachLayerRecord(db, acTrans, OpenMode.ForWrite, delegate(LayerTableRecord inLTR)
                                {
                                    if (String.Equals(inLTR.Name.Trim(), "0") /* || String.Equals(inLTR.Name.ToUpper().Trim(), "COUT")*/)
                                    {
                                        inLTR.IsLocked = true;

                                        if (inLTR.IsFrozen == true) { inLTR.IsFrozen = false; }

                                        db.Clayer = inLTR.ObjectId;
                                    }
                                });
                            }
                            catch
                            {
                                _Logger.Log(currentDWGstring + " (" + ms + "): Error occured while trying to lock 0 layer in: " + currentDWGstring);
                            }

                            #endregion

                            ObjectIdCollection coutObjs = new ObjectIdCollection();

                            // Delete COUT layer
                            try { Utilities.delLayer(acTrans, db, "COUT"); }
                            catch { _Logger.Log(currentDWGstring + " (" + ms + "): COUT Layer could not be deleted"); }

                            #region Make sure only layers are: 0, gtype, and m#-x layer names
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        String layerName = ltr.Name.Trim().ToUpper();
                                        if (!String.Equals(layerName, "0") &&
                                             !String.Equals(layerName, "_GTYPE_RPSTL") &&
                                             !UsedMLayers.Contains(layerName.ToLower()))
                                        {
                                            _Logger.Log(currentDWGstring + " (" + ms + "): Layer: " + "\"" + layerName + "\"" + " is not \"0\", \"COUT\", or frame page layer. Deleting layer...");

                                            Utilities.delLayer(acTrans, db, ltr.Name);
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region make sure only (0,0) point is on gtype layer
                            ObjectIdCollection gtypeObjs = new ObjectIdCollection();
                            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            {
                                //foreach (ObjectId oid in bt)
                                // {
                                using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord)
                                {
                                    foreach (ObjectId oid2 in btr)
                                    {
                                        Entity ent = acTrans.GetObject(oid2, OpenMode.ForRead) as Entity;

                                        if ((ent.Layer.ToUpper().Trim() == "_GTYPE_RPSTL") && (ent.GetType() != typeof(DBPoint)))
                                        {
                                            try
                                            {
                                                ent.UpgradeOpen();
                                                btr.UpgradeOpen();
                                                gtypeObjs.Add(ent.Id);
                                                ent.Erase(true);
                                                _Logger.Log(currentDWGstring + " (" + ms + "): Deleting unknown object on GTYPE layer: " + ent.ToString() + " " + ent.GetType().ToString());
                                            }
                                            catch
                                            {
                                                _Logger.Log(currentDWGstring + " (" + ms + "): Unknown object on GTYPE layer: " + ent.ToString() + " " + ent.GetType().ToString());
                                            }
                                        }

                                        if (ent.GetType() == typeof(DBText))
                                        {
                                            DBText dbtxt = ent as DBText;
                                            if (dbtxt.TextString.ToUpper().Contains("FOR") && dbtxt.TextString.HasNums())
                                            {
                                                ent.UpgradeOpen();
                                                //System.Windows.Forms.MessageBox.Show(dbtxt.Layer);
                                                dbtxt.UpgradeOpen();
                                                dbtxt.Layer = "0";
                                                dbtxt.TextStyleId = idROMAND;
                                                dbtxt.Height = .07;
                                                dbtxt.WidthFactor = .75;
                                                _Logger.Log(currentDWGstring + " (" + ms + "): FOR # text found: \"" + dbtxt.TextString + "\" on layer: " + dbtxt.Layer);
                                            }
                                        }
                                    }
                                }
                                // }
                            }
                            db.Purge(gtypeObjs);
                            #endregion

                            acTrans.Commit();
                        }

                        try { db.SaveAs(dwgSavePath, DwgVersion.Current); }
                        catch (Autodesk.AutoCAD.Runtime.Exception e)
                        {
                            _Logger.Log(currentDWGstring + " (" + ms + "): DWG Could not be saved because: " + e.Message);
                        }
                        catch (System.Exception e)
                        {
                            _Logger.Log(currentDWGstring + " (" + ms + "): DWG Could not be saved because: " + e.Message);
                        }

                        HostApplicationServices.WorkingDatabase = oldDb;

                        DwgCounter++;

                        try
                        {
                            _Bw.ReportProgress( Utilities.GetPercentage(DwgCounter, NumDwgs)); 
                        }
                        catch { _Logger.Log("Progress bar report error"); }

                        if (_Bw.CancellationPending)
                        {
                            _Logger.Log("Processing cancelled by user at dwg " + DwgCounter.ToString() + " out of " + NumDwgs);
                            if (_Logger.ErrorCount > 0) { return "Processing cancelled. Error log file: " + _Logger.Path; }
                            else
                            {
                                return "Processing cancelled at dwg " + DwgCounter + " out of " + NumDwgs;
                            }
                        }
                    }
                }

                foreach (String error in extraTMObjects) { _Logger.Log(error); }

                StopTimer();

                return (String.Concat(DwgCounter,
                                     " of ",
                                     NumDwgs,
                                     " DWG files processed in ",
                                     TimePassed,
                                     ((_Logger.ErrorCount > 0) ? (". Log file: " + _Logger.Path) : (""))));
            }
            catch (System.Exception se)
            {
                _Logger.Log("Unhandled processing exception: " + se.Message);
                return "Unhandled processing Exception: " + se.Message;
            }
            finally
            {
                AfterProcessing();
            }
        }
    }
}
