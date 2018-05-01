using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Text.RegularExpressions;
using System.IO;
using System.ComponentModel;
using System.Diagnostics;

namespace AcadExts
{
    [rtn("Processes RPSTLs for U.S. Gov. Delivery")]
    internal sealed class DeliveryFormatter : DwgProcessor
    {
        private String Suffix { get; set; }

        public DeliveryFormatter(String inPath, BackgroundWorker inBw, String inSuffix) : base(inPath, inBw)
        {
            Suffix = inSuffix;
        }
        
        public override String Process()
        {
            Regex FileFormat = new System.Text.RegularExpressions.Regex(@"^m(\d){6}([a-z]{0,2})?$");
            Regex LayerFormat = new System.Text.RegularExpressions.Regex(@"^m(\d){6}([a-z]){0,2}-(\w){1,2}$");

            try
            {
                BeforeProcessing();
            }
            catch(System.Exception se)
            {
                return "RPSTL Gov. Delivery Exception: " + se.Message;
            }
            //StartTimer();

            //if (!CheckDirPath()) { return "Invalid path: " + _Path; }

            //// Open error logger
            //try { _Logger = new Logger(_Path + "\\RpstlDeliveryErrors.txt"); }
            //catch { return "Could not write log file in: " + _Path; }

            // Get all DWG files
            try { GetDwgList(SearchOption.TopDirectoryOnly); }
            catch (SystemException se) 
            {
                _Logger.Log(" DWG files could not be enumerated because: " + se.Message);
                _Logger.Dispose();
                return " DWG files could not be enumerated because: " + se.Message;
            }

            try
            {
                foreach (String currentDWG in DwgList)
                {
                    Boolean found0 = false;
                    Boolean foundGType = false;
                    Boolean foundmx0 = false;
                    ObjectId foundmx0Id = new ObjectId();
                    Boolean foundmxSuffix = false;
                    String newfi = "";
                    List<String> LayersForDelete = new List<String>();

                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Processing cancelled at dwg " + currentDWG + " out of " + NumDwgs);
                        return "Processing cancelled at dwg " + DwgCounter.ToString() + " out of " + NumDwgs;
                    }

                    DwgCounter++;

                    try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                    catch { _Logger.Log("Progress bar report error"); }

                    using (Database db = new Database(false, true))
                    {
                        try
                        {
                            db.ReadDwgFile(currentDWG, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception e)
                        {
                            _Logger.Log("Could not read DWG: " + currentDWG + " because: " + e.Message); continue;
                        }

                        if (!FileFormat.IsMatch(System.IO.Path.GetFileNameWithoutExtension(db.Filename)))
                        {
                            _Logger.Log("Skipping: " + db.Filename + " because of incorrect name format");
                            continue;
                        }

                        String dwgMsName = System.IO.Path.GetFileNameWithoutExtension(db.Filename).Trim();

                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForWrite) as LayerTableRecord)
                                    {
                                        String layerName = ltr.Name.Trim();
                                        ltr.IsLocked = false;

                                        if (String.Equals(layerName, "0")) { found0 = true; db.Clayer = ltr.Id; ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255); continue; }
                                        if (String.Equals(layerName, "_GTYPE_RPSTL")) { foundGType = true; LayersForDelete.Add(ltr.Name); continue; }
                                        if (Regex.IsMatch(layerName, @"^m(\d){6}([a-z]){0,2}-0") && String.Equals(layerName.Remove(layerName.LastIndexOf('-')), dwgMsName)) { foundmx0 = true; ltr.IsFrozen = false; foundmx0Id = ltr.Id; ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255); continue; }
                                        if (Regex.IsMatch(layerName, @"^m(\d){6}([a-z]){0,2}-" + Suffix + "$") && String.Equals(layerName.Remove(layerName.LastIndexOf('-')), dwgMsName)) { foundmxSuffix = true; ltr.IsFrozen = false; ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255); continue; }
                                        if (String.Equals(layerName, "FILENAME") || String.Equals(layerName, "SCALE")) { LayersForDelete.Add(ltr.Name); continue; }

                                        // myLogger.Log("Deleting Unknown layer: " + ltr.Name + " in: " + currentDWG);
                                        LayersForDelete.Add(ltr.Name);
                                    }
                                }
                            }
                            if (!foundmx0 && !foundmxSuffix) { _Logger.Log("Could not find mx-0 or mx-suffix layer in: " + currentDWG); }
                            if (!found0) { _Logger.Log("0 Layer not found"); }
                            if (!foundGType) { _Logger.Log("GTYPE Layer not found"); }

                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForWrite) as LayerTableRecord)
                                    {
                                        String layerName = ltr.Name.Trim();

                                        if (foundmx0 && LayerFormat.IsMatch(layerName) && ltr.Id != foundmx0Id) { LayersForDelete.Add(ltr.Name); _Logger.Log("Layer: " + layerName + " not allowed in: " + System.IO.Path.GetFileNameWithoutExtension(db.Filename)); continue; }
                                        if (!foundmx0 && foundmxSuffix && LayerFormat.IsMatch(layerName)) { ltr.IsFrozen = false; ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255); continue; }
                                    }
                                }
                            }

                            foreach (String LayerForDelete in LayersForDelete) { Utilities.delLayer(acTrans, db, LayerForDelete); }

                            if (foundmx0) { newfi = String.Concat(db.Filename.Remove(db.Filename.LastIndexOf('.')), "-", "0", ".dwg"); }
                            if (!foundmx0) { newfi = String.Concat(db.Filename.Remove(db.Filename.LastIndexOf('.')), "-", Suffix.Trim(), ".dwg"); }

                            acTrans.Commit();
                        }
                        try { db.SaveAs(newfi, DwgVersion.Current); }
                        catch (System.Exception se) { _Logger.Log("Couldn't save Dwg: " + System.IO.Path.GetFileNameWithoutExtension(db.Filename) + " because: " + se.Message); }
                    }
                }

                //if (_Logger.ErrorCount > 0) { return "Log file: " + _Logger.Path; } else { return "No errors found"; }
            
            }
            catch (System.Exception e) { _Logger.Log("Unhandled exception: " + e.Message); return "Unhandled exception: " + e.Message; }
            finally
            {
                AfterProcessing();
            }

            return String.Concat(DwgCounter,
                                 " out of ",
                                 NumDwgs,
                                 " processed in ",
                                 TimePassed,
                                 (_Logger.ErrorCount > 0) ? (". Log file: " + _Logger.Path) : ("."));
        }
    }
}
