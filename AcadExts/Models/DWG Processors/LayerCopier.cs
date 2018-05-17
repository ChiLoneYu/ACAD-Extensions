using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;

namespace AcadExts
{
    // Program 1
    [rtn("Copies old m layer into new m layer with new dash number")]
    internal sealed class LayerCopier : DwgProcessor
    {
        // Matches an M number
        //private readonly Regex MRegex = new Regex(@"^M(\d){5,7}-\S*$");
        // Matches everything after "DESIGNATIONS WITH"
        //private readonly Regex PrefixValueRegex = new Regex(@"(?<=^DESIGNATIONS WITH\s)\S*$");
        // Matches everything after MSxxxxxx-"
        //private readonly Regex MSValueRegex = new Regex(@"(?<=^MS(\d){5,7}-)\S*$");
        // Matches a prefix note
        private readonly Regex PrefixRegex = new Regex(@"^DESIGNATIONS WITH\s.+$");
        // Matches everything after xxxxxx-
        private readonly Regex DashValueRegex = new Regex(@"(?<=^m(\d){6}[A-Za-z]{0,1}-)\d{1,2}[A-Za-z]?$");
        // Matches a MS#
        private readonly Regex MSRegex = new Regex(@"^MS(\d){6}[A-Za-z]{0,1}-.*$");

        private class Row
        {
            // Columns from text file
            public String col1 { private set; get; }
            public String col2 { private set; get; }
            public String col3 { private set; get; }

            // Columns from text file prepended with m
            //public String col1m { private set; get; }
            //public String col2m { private set; get; }
            //public String col3m { private set; get; }

            public Row(String inCol1, String inCol2, String inCol3)
            {
                col1 = inCol1;
                col2 = inCol2;
                col3 = inCol3;

                //col1m = String.Concat("m", col1);
                //col2m = String.Concat("m", col2);
                //col3m = String.Concat("m", col3);
            }

            public override string ToString()
            {
                return (String.Concat(col1, "|", col2, "|", col3));
            }
        }

        private readonly String MapPath = String.Empty;

        public LayerCopier(String inPath, BackgroundWorker inBw, String inMapPath)
            : base(inPath, inBw)
        {
            MapPath = inMapPath;
        }

        public override string Process()
        {
            List<Row> mappings = new List<Row>();

            //  Check file path
            if (!MapPath.isFilePathOK()) { return "Invalid File: " + MapPath; }

            // Open error logger
            try
            {
                BeforeProcessing();
            }
            catch (System.Exception se)
            {
                return "Layer Copy processing exception: " + se.Message;
            }

            // Get Dwgs
            try { GetDwgList(SearchOption.TopDirectoryOnly, (s) => true); }
            catch (System.Exception se)
            {
                _Logger.Dispose();
                return "Could not access DWG files in: " + _Path + " because: " + se.Message;
            }

            // Get mappings from text file
            try
            {
                String line = String.Empty;

                using (StreamReader reader = new StreamReader(MapPath))
                {
                    while ((line = reader.ReadLine()) != null)
                    {
                        String[] substrings = line.Split(new[] { '|' }, 3);
                        mappings.Add(new Row(substrings[0], substrings[1], substrings[2]));
                    }
                }
            }
            catch
            {
                _Logger.Log("Error parsing file: " + MapPath);
                _Logger.Dispose();
                return "Error parsing file: " + MapPath;
            }

            // Process
            try
            {
                // Save current db
                Database oldDb = HostApplicationServices.WorkingDatabase;

                foreach (String dwg in DwgList)
                {
                    String justFileName = String.Empty;
                    String oldDashNumber = String.Empty;
                    String newDashNumber = String.Empty;
                    Row correspondingRow = null;

                    //Check for BackgroundWorker cancellation
                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Layer copying cancelled by user at dwg " + DwgCounter.ToString() + " out of " + NumDwgs);
                        break;
                    }

                    justFileName = System.IO.Path.GetFileNameWithoutExtension(dwg);

                    // Find corresponding row in mapping list
                    IEnumerable<Row> correspondingRows = mappings.Where<Row>(r => String.Equals(r.col1, justFileName));    

                    if (correspondingRows.Count() < 1)
                    {
                        _Logger.Log("Skipping file because no mapping was found for: " + justFileName);
                        continue;
                    }
                    if (correspondingRows.Count() > 1)
                    {
                        correspondingRow = correspondingRows.First<Row>();
                        _Logger.Log("More than one mapping entry was found for: " + justFileName + ", using: " + correspondingRow.ToString());
                    }
                    correspondingRow = correspondingRows.First<Row>();

                    // Get old and new dash numbers from corresponding row in text file
                    try
                    {
                        oldDashNumber = DashValueRegex.Match(correspondingRow.col2).Value;
                        newDashNumber = DashValueRegex.Match(correspondingRow.col3).Value;
                    }
                    catch (System.Exception se)
                    {
                        _Logger.Log(se.Message + ": Error parsing dash numbers for: " + correspondingRow.col2 + " and " + correspondingRow.col3);
                        continue;
                    }

                    if (String.IsNullOrWhiteSpace(oldDashNumber) || String.IsNullOrWhiteSpace(newDashNumber))
                    {
                        _Logger.Log("old and new dash numbers not found in: " + correspondingRow.col2 + " and " + correspondingRow.col3);
                        continue;
                    }

                    using (Database db = new Database(false, true))
                    {
                        try
                        {
                            db.ReadDwgFile(dwg, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception e)
                        {
                            _Logger.Log(String.Concat(dwg + ": Could not read DWG because: ", e.Message));
                            continue;
                        }

                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                if (lt.Has(correspondingRow.col2))
                                {
                                    using (LayerTableRecord layerToCopy = acTrans.GetObject(lt[correspondingRow.col2], OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        // clone layer
                                        LayerTableRecord newLayer = (LayerTableRecord)layerToCopy.Clone();

                                        // rename the new layer using third column in corresponding row
                                        newLayer.Name = correspondingRow.col3;

                                        // upgrade layertable for write and add the new layer to the layertable and db
                                        lt.UpgradeOpen();
                                        lt.Add(newLayer);
                                        acTrans.AddNewlyCreatedDBObject(newLayer, true);

                                        // Get the block table record for this drawing's model space
                                        BlockTableRecord btr = acTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite) as BlockTableRecord;

                                        // Go through every entity in the block table record, and check if it's on the old layer,
                                        // and if it is, clone the entity, and add the clone to the new layer
                                        foreach (ObjectId entOid in btr)
                                        {
                                            using (Entity ent = acTrans.GetObject(entOid, OpenMode.ForRead) as Entity)
                                            {
                                                // Check if entity is on col2 layer
                                                if (String.Equals(ent.Layer, correspondingRow.col2))
                                                {
                                                    // Clone entity
                                                    Entity entityClone = ent.Clone() as Entity;

                                                    // if ent used for clone is text and a prefix or ms#, change the value
                                                    // (original entity, not the clone is used here to check type cause dxfname
                                                    // is still empty for the clone at this point)
                                                    if (ent.Id.ObjectClass.DxfName == "TEXT")
                                                    {
                                                        DBText dbt = entityClone as DBText;

                                                        // Check if text is a prefix note
                                                        if (PrefixRegex.IsMatch(dbt.TextString))
                                                        {
                                                            // use new dash number for prefix note on new layer
                                                            dbt.TextString = String.Concat(dbt.TextString.Substring(0, dbt.TextString.IndexOf("WITH") + 4), " " + newDashNumber);
                                                        }

                                                        // Check if text is MS#
                                                        if (MSRegex.IsMatch(dbt.TextString))
                                                        {
                                                            // use new dash number for MS# on new layer
                                                            dbt.TextString = String.Concat(dbt.TextString.Substring(0, dbt.TextString.IndexOf('-') + 1), newDashNumber);
                                                        }

                                                        btr.AppendEntity(dbt as Entity);
                                                        acTrans.AddNewlyCreatedDBObject(dbt as Entity, true);
                                                        dbt.Layer = newLayer.Name;
                                                    }
                                                    else
                                                    {
                                                        // Entity is not text, so add cloned entity to dwg, and change layer to newly added layer
                                                        btr.AppendEntity(entityClone);
                                                        acTrans.AddNewlyCreatedDBObject(entityClone as Entity, true);
                                                        entityClone.Layer = newLayer.Name;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    _Logger.Log(correspondingRow.col2 + " layer not found in dwg: " + dwg);
                                }
                            }
                            acTrans.Commit();
                        }
                        
                        try 
                        { db.SaveAs(String.Concat(Path.GetDirectoryName(db.Filename),
                                                  "\\",
                                                  correspondingRow.col3,
                                                  ".dwg"), DwgVersion.Current);
                        }
                        catch (System.Exception se) { _Logger.Log("Could not save processed dwg because: " + se.Message); }
                    }

                    // Increment dwg counter
                    DwgCounter++;

                    // Report current progress to progress bar
                    _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs));
                }
                // Put old db back
                HostApplicationServices.WorkingDatabase = oldDb;
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

            return String.Concat(DwgCounter,
                     " out of ",
                     NumDwgs,
                     " dwgs processed in ",
                     TimePassed,
                     ((_Logger.ErrorCount > 0) ? (". Log file: " + _Logger.Path) : "."));
        }
    }
}
