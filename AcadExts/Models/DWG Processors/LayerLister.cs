using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using System.Windows;
using System.ComponentModel;
using System.Diagnostics;

namespace AcadExts
{
    // Make lists of all layers in all dwgs in current directory and in all nested directories
    [rtn("Makes layer lists of all DWGs")]
    internal sealed class LayerLister : DwgProcessor
    {
        public LayerLister(String inPath, BackgroundWorker inBw)
            : base(inPath, inBw)
        {
        }

        //https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/file-system/how-to-iterate-through-a-directory-tree
        public override String Process()
        {
            Int32 totalDwgCounter = 0;

            if (!CheckDirPath())
            {
                //ed.WriteMessage(Environment.NewLine + "Invalid path: " + Path);
                return "Invalid path: " + _Path;
            }

            try { _Logger = new Logger(_Path + "\\LayerListErrors.txt"); }
            catch (System.Exception se)
            {
                //ed.WriteMessage(Environment.NewLine + "Could not write log file because: " + se.Message + Environment.NewLine);
                return "Could not create log file in: " + _Path + " because: " + se.Message;
            }

            StartTimer();

            Stack<String> dirs = new Stack<String>();

            dirs.Push(_Path);

            while (dirs.Count > 0)
            {
                String currentDir = dirs.Pop();
                String[] subDirs;

                try
                {
                    subDirs = System.IO.Directory.GetDirectories(currentDir);
                }
                catch (System.Exception se)
                {
                    _Logger.Log("Sub directories in folder: " + _Path + " could not be accessed because: " + se.Message);
                    continue;
                }

                try
                {
                    GetDwgList(SearchOption.TopDirectoryOnly);
                }
                catch (System.Exception se)
                {
                    _Logger.Log("Files could not be accessed in: " + currentDir + " because: " + se.Message);
                    continue;
                }

                totalDwgCounter += NumDwgs;

                foreach (String dwg in DwgList)
                {
                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Processing cancelled by user at dwg " + DwgCounter + " out of " + NumDwgs);
                        _Logger.Dispose();
                        return "Layer listing cancelled at DWG: " + DwgCounter;
                    }

                    DwgCounter++;

                    try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                    catch
                    {
                        _Logger.Log("Progress bar error");
                    }

                    String dwgName = System.IO.Path.GetFileNameWithoutExtension(dwg);
                    System.IO.StreamWriter sw = null;

                    try
                    {
                        sw = new StreamWriter(currentDir + "\\" + dwgName + ".txt");
                    }
                    catch (System.Exception se) { _Logger.Log("Could not write layer list file in: " + currentDir + " because: " + se.Message); }

                    using (Database db = new Database(false, true))
                    {
                        try
                        {
                            db.ReadDwgFile(dwg, FileOpenMode.OpenForReadAndReadShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch (System.Exception se)
                        {
                            _Logger.Log("Could not read DWG: " + dwg + " because : " + se.Message);
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
                                        sw.WriteLine(ltr.Name.Trim());
                                    }
                                }
                            }
                        }
                    }
                    sw.Flush();
                    sw.Close();
                }
                foreach (String str in subDirs)
                {
                    dirs.Push(str);
                }
            }

            StopTimer();

            _Logger.Dispose();

            if (_Logger.ErrorCount > 0) { return _Logger.ErrorCount.ToString() + " errors found. Log file: " + _Path; }
            else
            {
                return String.Concat(totalDwgCounter, " lists made in: ", TimePassed);
            }
        }
    }
}
