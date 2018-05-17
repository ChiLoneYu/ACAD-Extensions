using Autodesk.AutoCAD.ApplicationServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;

namespace AcadExts
{
    //[rtn(".NET rewrite of FP_A_P AutoCAD command")]
    internal sealed class FP_A_P : DwgProcessor
    {
        public FP_A_P(String inPath, BackgroundWorker inBw)
            : base(inPath, inBw)
        {

        }

        public override String Process()
        {
            if (!CheckDirPath()) { return "Invalid path: " + _Path; }

            try { _Logger = new Logger(String.Concat(_Path, "\\FP_A_P_ErrorLog.txt")); }
            catch (System.Exception se) { return "Could not create log file in: " + _Path + " because: " + se.Message; }

            StartTimer();

            // Get dwgs
            try
            {
                GetDwgList(SearchOption.TopDirectoryOnly, inFile => true);
            }
            catch (System.Exception se)
            {
                _Logger.Log("Could not access DWG files in: " + _Path + " because: " + se.Message);
                _Logger.Dispose();
                return "Could not access DWG files in: " + _Path + " because: " + se.Message;
            }

            try
            {
                Database oldDB = HostApplicationServices.WorkingDatabase;

                foreach (String dwg in DwgList)
                {
                    using (Database db = new Database(false, true))
                    {
                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable)
                            {
                                if (String.Equals(db.Clayer, "0"))
                                {
                                    _Logger.Log(Path.GetFileNameWithoutExtension(db.Filename) + " has 0 layer set as current layer"); 
                                }

                                foreach(ObjectId oid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        if (!String.Equals(ltr.Name, "0"))
                                        {
                                            db.Clayer = lt[ltr.Name];
                                        }
                                    } 
                                }
                            }

                            if ("N:\\eps".isFilePathOK() || "N:\\eps".isDirectoryPathOK())
                            {
                                // preoutputsave
                                db.SaveAs(db.Filename, DwgVersion.Current);
                                String calling_fc = "FP_A_P";
                            }
                            else
                            {
                                // 2nd progn
                                _Logger.Log("No connection to the N:\\ Drive");
                            }

                            acTrans.Commit();
                        }
                    }
                }

                HostApplicationServices.WorkingDatabase = oldDB;
            }
            catch(System.Exception se)
            {
                _Logger.Log(String.Concat("Exception: ", se.Message));
            }
            finally
            {
                StopTimer();
                _Logger.Dispose();
            }
            
            return "";
        }

        private void PreOutputSave()
        {

        }
    }
}
