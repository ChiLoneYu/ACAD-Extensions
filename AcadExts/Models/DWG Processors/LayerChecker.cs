using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using System.Text.RegularExpressions;

namespace AcadExts
{
    [rtn("Checks Layers")]
    internal sealed class LayerChecker : DwgProcessor
    {
        private System.Boolean MakeChanges { get; set; }
        private System.Boolean MultiDir { get; set; }

        System.IO.StreamWriter writer = null;

        public LayerChecker(String inPath, BackgroundWorker inBw, Boolean inMultiDir, Boolean inMakeChanges)
            : base(inPath, inBw)
        {
            MakeChanges = inMakeChanges;
            MultiDir = inMultiDir;
        }

        public override String Process()
        {
            Regex pattern1 = new Regex(@"^[ACDFIJKNTVWX][0-9][0-9][0-9][0-9][0-9][0-9][A-Z]?((-|–)[0-9][0-9]?)?$");
            Regex pattern2 = new Regex(@"^[ABCDEGHKLNPRUTZ][0-9][0-9][0-9][0-9]?((-|–)[0-9][0-9]?)?$");
            Regex pattern3 = new Regex(@"^[0-9][0-9][0-9][0-9][0-9][0-9][A-Z]?((-|–)([0-9]|[A-Z]))?$");

            Int32 numErrorsPerDWG = 0, numErrorsTotal = 0;

            FileInfo textReport;

            //if (!CheckDirPath()) { return "Invalid path: " + _Path; }


            try
            {
                BeforeProcessing();
            }
            catch(System.Exception se)
            {
                return "Layer Checker processing exception: " + se.Message;
            }

            try
            {
                textReport = new FileInfo(_Path + "\\dwgsource_check_" + DateTime.Now.ToString("ddHHmmss") + ".txt");
                writer = new StreamWriter(textReport.FullName);
            }
            catch
            {
                _Logger.Dispose();
                return "Could not open checker log file in: " + _Path;
            }
            //try { _Logger = new Logger(_Path + "\\LayerCheckerErrorLog.txt"); }
            //catch { return "Could not create error log file in: " + _Path; }

            //StartTimer();

            #region Get dwgs and create checked dir if multi dir box is checked

            if (MultiDir)
            {
                try
                {
                    IEnumerable<String> dirList = Directory.EnumerateDirectories(_Path);
                    Directory.CreateDirectory(_Path + "\\checked\\");

                    foreach (String dirToCopy in dirList)
                    {
                        if (dirToCopy.Contains("checked")) { continue; }
                        Directory.CreateDirectory(String.Concat(_Path, "\\checked\\", dirToCopy.Substring(dirToCopy.LastIndexOf("\\") + 1)));
                    }
                }
                catch (System.Exception se)
                {
                    _Logger.Log("Error re-creating directory structure in new folder in: " + _Path + " because: " + se.Message);
                    return "Error re-creating directory structure in new folder in: " + _Path;
                }

                try { GetDwgList(SearchOption.AllDirectories, (inFileStr) => !inFileStr.Contains("\\checked\\")); }
                catch (System.Exception se)
                {
                    _Logger.Log(" Not all .dwg files could be enumerated because: " + se.Message);
                    return "Could not get all DWG files";
                }
            }
            else
            {
                try { GetDwgList(SearchOption.TopDirectoryOnly); }
                catch (System.Exception se)
                {
                    _Logger.Log(" Not all DWG files could be enumerated because: " + se.Message);
                    return "Could not get all dwg files in: " + _Path;
                }
            }
            #endregion
            try
            {
                foreach (String currentDWG in DwgList)
                {
                    Database oldDb = HostApplicationServices.WorkingDatabase;
                    StringBuilder currentDwgErrors = new StringBuilder();
                    numErrorsPerDWG = 0;
                    String ms = "";

                    using (Database db = new Database(false, true))
                    {
                        try
                        {
                            db.ReadDwgFile(currentDWG, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (System.Exception se)
                        {
                            _Logger.Log("Could not read DWG: " + currentDWG + " because: " + se.Message);
                            continue;
                        }

                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable)
                            {
                                if (MakeChanges) { db.Clayer = lt["0"]; }

                                foreach (ObjectId layerId in lt)
                                {
                                    LayerTableRecord layer = acTrans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;

                                    String curLayerName = layer.Name.ToUpper().Trim();

                                    if (curLayerName.Equals("FILENAME")) { ms = Utilities.msText(db, "FILENAME"); }

                                    if (curLayerName.Equals("MSNUM")) { ms = Utilities.msText(db, "MSNUM"); }

                                    if (MakeChanges) { layer.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255); }

                                    if ((String.Equals(curLayerName, "0")) ||
                                         (String.Equals(curLayerName, "DEFPOINTS")) ||
                                         (String.Equals(curLayerName, "COLUMN")) ||
                                         (String.Equals(curLayerName, "ST_TABLE_VISIBLE")) ||
                                         (String.Equals(curLayerName, "ST_TABLE_INVISIBLE")) ||
                                         (String.Equals(curLayerName, System.IO.Path.GetFileNameWithoutExtension(currentDWG).ToUpper().Trim()))
                                        )
                                    {
                                        if (MakeChanges)
                                        {
                                            layer.IsLocked = false;

                                            if (layer.IsFrozen) { layer.IsFrozen = false; }

                                            if (String.Equals(curLayerName, "ST_TABLE_INVISIBLE")) { layer.IsOff = true; } else { layer.IsOff = false; }
                                        }
                                    }
                                    else
                                    {
                                        if ((String.Equals(curLayerName, "IADS_HOTSPOTS")) ||
                                             (String.Equals(curLayerName, "TEMPLATE")) ||
                                             (String.Equals(curLayerName, "ST_AUTOCONVERT_MARKERS")) ||
                                             (String.Equals(curLayerName, "ZONE")) ||
                                             (String.Equals(curLayerName, "FILENAME")) ||
                                             (String.Equals(curLayerName, "SCALE")) ||
                                             (String.Equals(curLayerName, "MSNUM")) ||
                                             (curLayerName.StartsWith("REF_")) ||
                                             (((pattern1.IsMatch(curLayerName)) ||
                                               (pattern2.IsMatch(curLayerName)) ||
                                               (pattern3.IsMatch(curLayerName))) &&
                                               (!System.IO.Path.GetFileNameWithoutExtension(currentDWG).ToUpper().Equals(curLayerName))
                                             ))
                                        {
                                            if (MakeChanges)
                                            {
                                                layer.IsLocked = false;

                                                Utilities.delLayer(db, layer.Name, layer);
                                            }
                                        }
                                        else
                                        {
                                            currentDwgErrors.Append(layer.Name + "\t\t\t" + ms + Environment.NewLine);
                                            numErrorsPerDWG++;
                                        }
                                    }
                                    layer.Dispose();
                                }
                            }
                            acTrans.Commit();
                        }
                        if (MultiDir)
                        {
                            try
                            {
                                // Dwg is in subdir
                                if (!String.Equals(Path.GetDirectoryName(currentDWG), _Path))
                                {
                                    //Get subdir path and subdir name
                                    String subdir = Path.GetDirectoryName(currentDWG);
                                    String subdirName = subdir.Substring(subdir.LastIndexOf("\\"));

                                    db.SaveAs(Path.GetDirectoryName(subdir) + "\\checked" + subdirName + "\\" + Path.GetFileName(currentDWG), DwgVersion.Current);
                                }

                                // Dwg is in top dir
                                else
                                {
                                    db.SaveAs(_Path + "\\" + Path.GetFileName(currentDWG), DwgVersion.Current);
                                }
                            }
                            catch (System.Exception se)
                            {
                                _Logger.Log(currentDWG + " could not be saved because: " + se.Message);
                            }
                        }
                        else
                        {
                            try
                            {
                                db.SaveAs(currentDWG, DwgVersion.Current);
                            }
                            catch (System.Exception se)
                            {
                                _Logger.Log(currentDWG + " could not be saved because: " + se.Message);
                            }
                        }
                        HostApplicationServices.WorkingDatabase = oldDb;
                    }
                    numErrorsTotal += numErrorsPerDWG;

                    if (numErrorsPerDWG > 0)
                    {
                        currentDwgErrors.Insert(0, Utilities.nl + currentDWG + Utilities.nl + "--------" + Utilities.nl);
                        writer.Write(currentDwgErrors);
                    }

                    DwgCounter++;

                    try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                    catch { }

                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Layer Checking cancelled by user at dwg " + DwgCounter + " out of " + NumDwgs);
                        break;
                    }
                }
            }
            catch(System.Exception se) 
            {
                _Logger.Log("Processing Exception: " + se.Message);
                return "Processing Exception: " + se.Message;
            }
            finally 
            {
                try { writer.Close(); }
                catch (System.Exception se) { _Logger.Log("Couldn't close layer check file because: " + se.Message); }

                if (numErrorsTotal < 1)
                {
                    try { textReport.Delete(); }
                    catch { _Logger.Log("Couldn't delete empty check file"); }
                }

                AfterProcessing();
            }

            return String.Concat(DwgCounter.ToString(),
                " out of ",
                NumDwgs.ToString(),
                " dwgs processed in ",
                TimePassed,
                ". ",
                (_Logger.ErrorCount > 0) ? ("Error Log: " + _Logger.Path) : ("No errors found."));
        }
    }
}
