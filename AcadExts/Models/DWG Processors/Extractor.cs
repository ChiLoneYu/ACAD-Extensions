using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Diagnostics;

namespace AcadExts
{
    [rtn("Copies dwgs in all subdirectories to top directory")]
    internal sealed class Extractor : DwgProcessor
    {
        
        public Extractor(String inPath, BackgroundWorker inBw) : base(inPath, inBw)
        {
        }
        
        public override String Process()
        {
            if (!CheckDirPath())
            { return String.Concat("Invalid path: ", _Path); }

            try { _Logger = new Logger(_Path + "\\ExtractErrors.txt"); }
            catch { return "Could not initiate log file in: " + _Path; }

            try { GetDwgList(SearchOption.AllDirectories); }
            catch(System.Exception se)
            {
                return "Could not access all files in: " + _Path + " because: " + se.Message; 
            }

            StartTimer();

            foreach(String currentFile in DwgList)
            {
                String newFilePath = String.Concat(_Path, "\\", System.IO.Path.GetFileName(currentFile));

                try { File.Copy(currentFile, newFilePath, false); }
                
                catch(IOException ioe)
                { _Logger.Log(String.Concat("File already exists in top directory: ", currentFile, " : ", ioe.Message)); continue; }
                
                catch (System.Exception se)
                { _Logger.Log(String.Concat("Error copying file: ", currentFile, " : ", se.Message)); continue; }

                DwgCounter++;

                try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                catch { _Logger.Log("Progress bar error"); }

                if (_Bw.CancellationPending)
                {
                    _Logger.Log("Extractor cancelled by user at dwg " + DwgCounter.ToString() + " out of " + NumDwgs);
                    break;
                }
            }

            StopTimer();

            _Logger.Dispose();

            return String.Concat(DwgCounter.ToString(), " out of ",
                                 NumDwgs, " DWGs copied in ",
                                 TimePassed);
        }
    }
}
