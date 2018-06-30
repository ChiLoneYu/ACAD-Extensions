using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;

namespace AcadExts
{
    [rtn("Converts dwg to new layer name")]
    internal sealed class LayerNamer : DwgProcessor
    {
        const String NewFolderName = "\\ProcessingOutput\\";

        private readonly Regex SuffixRegex = new Regex(@"^(\s*)(\d{1,2})(\s*)$");
        private readonly Regex PrefixRegex = new Regex(@"(?<=^DESIGNATIONS\sWITH\s).*");
        private readonly Regex msnoFileRegex = new Regex(@"(?<=(\s|\t)MS)(\d){6}[A-Z]?-(\d){1,2}([A-Z])?(?=\s|\t|$)");
        private readonly Regex DashNumRegex = new Regex(@"(?<=m(\d){6}([A-Za-z])?-)(\d){1,2}([A-Za-z])?");

        //private readonly String _Suffix = String.Empty;
        //public String Suffix { get { return _Suffix; } }

        //private readonly String _ExcelPath = String.Empty;
        //public String ExcelPath { get { return _ExcelPath; } }

        private readonly String _MSNOFile = String.Empty;
        public String MSNOFile { get { return _MSNOFile; } }

        public LayerNamer(String inPath, BackgroundWorker inBw, String inMSNOFile)
            : base(inPath, inBw)
        {
            //this._Suffix = inSuffix.Trim();
            //this._ExcelPath = inExcelPath;
            this._MSNOFile = inMSNOFile;
        }

        public override String Process()
        {
            //List<ExcelAccessor.Row> XlRows = null;
            List<String> msList = null;

            try { BeforeProcessing(); }
            catch (System.Exception se) { return "Layer Naming processing exception: " + se.Message; }

            if (String.IsNullOrWhiteSpace(MSNOFile) || !MSNOFile.isFilePathOK())
            {
                return "Invalid MSNO File: " + MSNOFile;
            }

            //if (String.IsNullOrWhiteSpace(ExcelPath) || !ExcelPath.isFilePathOK())
            //{
            //    return "Invalid Excel File: " + ExcelPath;
            //}

            //if (String.IsNullOrWhiteSpace(Suffix) || !SuffixRegex.IsMatch(Suffix))
            //{
            //    return "Invalid Suffix. Suffix must be a 1 or 2 digit number.";
            //}

            try
            {
                GetFiles(SearchOption.TopDirectoryOnly, null, "*.dwg");
                //GetDwgList(SearchOption.TopDirectoryOnly, (s) => true);
            }
            catch (System.Exception se)
            {
                _Logger.Dispose();
                return "Could not access DWG files in: " + _Path + " because: " + se.Message;
            }

            //try { XlRows = ExcelAccessor.GetExcelData(ExcelPath); }
            //catch (System.Exception se)
            //{
            //    _Logger.Log("Excel file could be read because: " + se.Message);
            //    _Logger.Dispose();
            //    return "Error reading Excel file: " + se.Message;
            //}

            //if (!XlRows.Select<ExcelAccessor.Row, String>((row) => row.NewLayer.Trim()).Contains(Suffix))
            //{
            //    _Logger.Log("New layer suffix not found in 'New Layer' column Excel file");
            //    _Logger.Dispose();
            //    return "New layer suffix not found";
            //}

            try
            {
                // Create new folder for output
                Directory.CreateDirectory(String.Concat(_Path, NewFolderName));

                msList = new List<string>();

                using (StreamReader msnoFileReader = new StreamReader(MSNOFile))
                {
                    String line = String.Empty;

                    while ((line = msnoFileReader.ReadLine()) != null)
                    {
                        if (msnoFileRegex.IsMatch(line))
                        {
                            msList.Add(String.Concat("m", msnoFileRegex.Match(line).Value.Trim().ToLower()));
                        }
                    }
                }

                foreach (String currentDwg in DwgList)
                {
                    //Boolean prefixExists = false;
                    //String prefix = String.Empty;
                    String NewFileName = String.Empty;
                    String DashNumber = String.Empty;

                    if (Path.GetFileName(currentDwg).Contains("-"))
                    {
                        _Logger.Log(Path.GetFileName(currentDwg) + " contains a dash, skipping file.");
                        continue;
                    }

                    DwgCounter++;

                    using (Database db = new Database(false, true))
                    {
                        // Read in DWG file
                        try
                        {
                            db.ReadDwgFile(currentDwg, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception e)
                        {
                            _Logger.Log(String.Concat(currentDwg + ": Could not read DWG because: ", e.Message));
                            continue;
                        }

                        using (Transaction acTrans = db.TransactionManager.StartTransaction())
                        {
                            #region comments
                            // Look for prefix, and get it if it exists
                            //using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                            //{
                            //    foreach (ObjectId oidBT in bt)
                            //    {
                            //        using (BlockTableRecord btr = acTrans.GetObject(oidBT, OpenMode.ForRead) as BlockTableRecord)
                            //        {
                            //            foreach (ObjectId btrOID in btr)
                            //            {
                            //                using (Entity ent = acTrans.GetObject(btrOID, OpenMode.ForRead) as Entity)
                            //                {
                            //                    if (ent.GetType() == typeof(DBText))
                            //                    {
                            //                        DBText dbt = ent as DBText;

                            //                        if (dbt.TextString.ToUpper().Contains("DESIGNATIONS WITH"))
                            //                        {
                            //                            prefix = PrefixRegex.Match(dbt.TextString).Value;
                            //                            prefixExists = true;
                            //                        }
                            //                    }
                            //                }
                            //            }
                            //        }
                            //    }
                            //}
                            #endregion

                            ToEach.toEachLayer(db, acTrans, delegate(LayerTableRecord ltrIN)
                            {
                                if (msList.Contains(ltrIN.Name) || String.Equals("0", ltrIN.Name))
                                {

                                    if (!String.Equals("0", ltrIN.Name))
                                    {
                                        // Get dash number
                                        DashNumber = DashNumRegex.Match(ltrIN.Name).Value.ToLower();
                                    }

                                    // dont delete layer if its the 0 layer or its in the list
                                    // turn it on and unfreeze it
                                    ltrIN.IsFrozen = false;
                                    ltrIN.IsOff = false;
                                }
                                else
                                {
                                    Utilities.delLayer(acTrans, db, ltrIN.Name);
                                }
                                return true;
                            });

                            // Iterate through layers
                            //using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            //{
                            //    foreach (ObjectId oid in lt)
                            //    {
                            //        using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForWrite) as LayerTableRecord)
                            //        {
                            //            if (msList.Contains(ltr.Name) || String.Equals("0", ltr.Name))
                            //            {

                            //                if (!String.Equals("0", ltr.Name)) 
                            //                {
                            //                    // Get dash number
                            //                    DashNumber = DashNumRegex.Match(ltr.Name).Value.ToLower();
                            //                }

                            //                // dont delete layer if its the 0 layer or its in the list
                            //                // turn it on and unfreeze it
                            //                ltr.IsFrozen = false;
                            //                ltr.IsOff = false;
                            //                continue;
                            //            }
                            //            else
                            //            {
                            //                Utilities.delLayer(acTrans, db, ltr.Name);
                            //            }
                            //            #region comments
                            //            //Lowercase the layer name 

                            //            // Turn on and unfreeze layer that ends in suffix
                            //            //if (ltr.Name.EndsWith(Suffix))
                            //            //{
                            //            //    //ltr.UpgradeOpen();
                            //            //    ltr.IsFrozen = false;
                            //            //    ltr.IsOff = false;
                            //            //}

                            //            // delete all layers except 0 and mathcing layer
                            //            // turn on and unfreeze 0 and matching layer

                            //            // put new dwgs in sub folder

                            //            //strip out s

                            //            // If prefix exists, delete all layers except prefix layer
                            //            //if (prefixExists && !ltr.Name.EndsWith(Suffix))
                            //            //{
                            //            //    Utilities.delLayer(acTrans, db, ltr.Name);
                            //            //}

                            //            //// If prefix doesn't exist, delete all layers (except 0 layer)
                            //            //if (!prefixExists && !ltr.Name.EndsWith("-0"))
                            //            //{
                            //            //    Utilities.delLayer(acTrans, db, ltr.Name);
                            //            //}

                            //            //if (!prefixExists && ltr.Name.EndsWith("-0"))
                            //            //{
                            //            //    //ltr.UpgradeOpen();
                            //            //    ltr.IsOff = false;
                            //            //    ltr.IsFrozen = false;
                            //            //}
                            //            #endregion
                            //        }
                            //    }
                            //}
                            acTrans.Commit();
                        }
                        // Save dwg with dash and suffix if a prefix exists
                        //if (prefixExists) 
                        //{
                        //    NewFileName = String.Concat(currentDwg.Substring(0, currentDwg.IndexOf(".dwg")), "-", Suffix, ".dwg");
                        //    db.SaveAs(NewFileName, DwgVersion.Current); 
                        //}

                        if (!String.IsNullOrWhiteSpace(DashNumber))
                        {
                            NewFileName = String.Concat(Path.GetDirectoryName(currentDwg),
                                                        NewFolderName,
                                                        Path.GetFileNameWithoutExtension(currentDwg),
                                                        "-",
                                                        DashNumber,
                                                        ".dwg");
                        }
                        else
                        {
                            NewFileName = String.Concat(Path.GetDirectoryName(currentDwg),
                                                       NewFolderName,
                                                       Path.GetFileNameWithoutExtension(currentDwg),
                                                       "-0.dwg");                            
                        }
                        db.SaveAs(NewFileName, DwgVersion.Current);
                    }

                    _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs));

                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Processing cancelled by user at dwg " + DwgCounter + " out of " + NumDwgs);
                        break;
                    }
                }// foreach loop
            }
            catch (System.Exception se)
            {
                return "Uncaught LayerName processing Exception: " + se.Message;
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
