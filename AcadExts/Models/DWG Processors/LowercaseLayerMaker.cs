using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using System.Timers;
using System.Diagnostics;

namespace AcadExts
{
    [rtn("Makes all layer names lowercase")]
    internal sealed class LowercaseLayerMaker : DwgProcessor
    {

        public LowercaseLayerMaker(String inPath, BackgroundWorker inBw) : base(inPath, inBw)
        {
        }

        // copnvert all layer names to lowercase
        public override String Process()
        {

            if (!CheckDirPath()) { return "Invalid path: " + _Path; }

            try
            {
                _Logger = new Logger(String.Concat(_Path, "\\LowercaseLayerErrorLog.txt"));
            }
            catch (System.Exception se)
            {
                return "Could not create log file in: " + _Path + " because: " + se.Message;
            }

            try { GetDwgList(SearchOption.TopDirectoryOnly); }
            catch (System.Exception se)
            {
                _Logger.Log("Could not get dwg files because: " + se.Message);
                _Logger.Dispose();
                return "Could not get all dwg files";
            }

            StartTimer();

            try
            {
                foreach (String currentDwg in DwgList)
                {
                    System.Boolean changeMade = false;

                    using (Database db = new Database(false, true))
                    {
                        try
                        {
                            db.ReadDwgFile(currentDwg, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (System.Exception se)
                        {
                            _Logger.Log(String.Concat("Can't read DWG: ", currentDwg, " because: ", se.Message));
                            continue;
                        }
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        String layerName = ltr.Name.Trim();
                                        String layerNameLower = layerName.ToLower();

                                        if (!String.Equals(layerName, layerNameLower))
                                        {
                                            changeMade = true;
                                            ltr.UpgradeOpen();
                                            ltr.Name = layerNameLower;

                                        }
                                    }
                                }
                            }
                            acTrans.Commit();
                        }
                        if (changeMade) { db.SaveAs(currentDwg, DwgVersion.Current); }
                    }

                    DwgCounter++;

                    try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                    catch { }

                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Processing cancelled by user at DWG: " + DwgCounter.ToString());
                        return "Processing cancelled at dwg " + DwgCounter.ToString() + " out of " + NumDwgs.ToString();
                    }
                }
            }
            catch (System.Exception se) { _Logger.Log(String.Concat("Error: ", se.Message)); return "Processing error: " + se.Message; }
            finally
            {
                StopTimer();
                _Logger.Dispose();
            }

            return String.Concat(" ", DwgCounter.ToString(), " out of ", NumDwgs, " dwgs processed in ", TimePassed);
        }
    }
}
