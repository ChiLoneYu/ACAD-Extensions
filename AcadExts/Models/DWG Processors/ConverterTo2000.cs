using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;

namespace AcadExts
{
    [rtn("Convert DWGs to 2000 dwg file format")]
    internal sealed class ConverterTo2000 : DwgProcessor
    {
        public ConverterTo2000(String inPath, BackgroundWorker inBw) : base(inPath, inBw)
        {
        }

        public override String Process()
        {

            String newFolder = _Path + "\\ConvertedTo2000\\";

            try
            {
                BeforeProcessing();
            }
            catch(System.Exception se)
            {
                //_Logger.Dispose();
                return "2000 Conversion processing exception: " + se.Message;
            }
            //if (!CheckDirPath()) { return "Invalid path: " + _Path; }

            //try { _Logger = new Logger(String.Concat(_Path, "\\ConvertTo2000Errors.txt")); }
            //catch (System.Exception se)
            //{
            //    return "Could not create error log file in: " + _Path + " because: " + se.Message;
            //}

            //StartTimer();

            try { GetDwgList(SearchOption.TopDirectoryOnly); }
            catch (System.Exception se)
            {
                _Logger.Log("DWG files could not be accessed in: " + _Path + " because: " + se.Message);
                _Logger.Dispose();
                return "DWG files could not be accessed in: " + _Path + " because: " + se.Message;
            }

            if (NumDwgs < 1) { return "No DWGs found in: " + _Path; }

            try
            {
                Directory.CreateDirectory(newFolder);
            }
            catch (System.Exception se)
            {
                _Logger.Log("Could not create directory in: " + _Path + " because: " + se.Message);
                return "Could not create directory in: " + _Path + " because: " + se.Message;
            }

            try
            {
                foreach (String currentDwg in DwgList)
                {
                    using (Database db = new Database(false, true))
                    {
                        if (_Bw .CancellationPending)
                        {
                            _Logger.Log("DWG processing cancelled at DWG " + DwgCounter + " out of " + NumDwgs);
                            return "DWG processing cancelled at DWG " + DwgCounter + " out of " + NumDwgs;
                        }

                        try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter + 1, NumDwgs)); }catch { }
                        
                        try
                        {
                            db.ReadDwgFile(currentDwg, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                            db.CloseInput(true);
                        }
                        catch(System.Exception e)
                        {
                            _Logger.Log(String.Concat("Could not read DWG: ", System.IO.Path.GetFileName(currentDwg), " because: ", e.Message));
                            continue;
                        }

                        try
                        {
                            db.SaveAs(String.Concat(newFolder, System.IO.Path.GetFileName(currentDwg)), DwgVersion.AC1015);
                            DwgCounter++;
                        }
                        catch
                        {
                            _Logger.Log("Error saving DWG to 2000 format: " + currentDwg);
                        }
                    }
                }
            }
            catch (System.Exception se)
            {
                _Logger.Log("Exception occured: " + se.Message);
                return "Exception occured. Error Log: " + _Logger.Path;
            }
            finally
            {
                AfterProcessing();
            }

            return String.Concat(DwgCounter, " DWGs out of ", NumDwgs, " converted in ", TimePassed);
        }
    }
}
