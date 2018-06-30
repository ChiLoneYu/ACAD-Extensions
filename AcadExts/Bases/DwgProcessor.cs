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
    internal abstract class DwgProcessor : IProcessor
    {
        // DWG folder path
        protected String _Path { get; private set; } 

        // Processing BackgroundWorker
        //private readonly BackgroundWorker _bw = null;
        //protected BackgroundWorker _Bw { get { return _bw; } }
        protected BackgroundWorker _Bw { get; private set; }

        // Error logger
        protected Logger _Logger { get; set; }

        // Processing timer
        protected System.Diagnostics.Stopwatch timer { get; set; }

        // # of dwgs in list
        protected Int32 NumDwgs { get { return DwgList.Count; } }

        //private List<String> _DwgList;
        //protected List<String> DwgList { get { return _DwgList; } }
        // List of dwgs to be processed
        protected List<String> DwgList { get; private set; }

        // dwg count. increment this as each dwg is processed
        private Int32 _DwgCounter = 0;
        protected Int32 DwgCounter { get { return _DwgCounter; } set { _DwgCounter = value; } }

        //  Processing timer
        protected String TimePassed { get { return timer.Elapsed.TotalSeconds.PrintTimeFromSeconds() ?? " null "; } }

        // Ctor
        public DwgProcessor(String inPath, BackgroundWorker inBw)
        {
            if (!String.IsNullOrWhiteSpace(inPath)) 
            {
                // Cut trailing slashes
                _Path = inPath.TrimEnd(new[] { '\\' });
            }

            if (inBw == null)
            {
                throw new ArgumentNullException("inBw", "Base dwg processor needs a BackgroundWorker instance");
            }
            else 
            {
                _Bw = inBw;
            }
        }

        protected virtual void BeforeProcessing()
        {
            // Check folder path
            if (!_Path.isDirectoryPathOK())
            {
                throw new System.IO. DirectoryNotFoundException(String.Concat("Folder not found: ", _Path));
            }

            // Open Logger
            try { _Logger = new Logger(String.Concat(/*System.IO.Path.GetDirectoryName()*/_Path, "\\ErrorLog.txt")); }
            catch
            {
                try { _Logger.Dispose(); } catch { }
                // Re-throw exception
                throw;
            }

            // Start timer
            timer = Stopwatch.StartNew();
        }

        protected virtual void AfterProcessing()
        {
            // Stop timer
            if (timer.IsRunning && timer != null)
            {
                timer.Stop();
            }

            // Close logger
            try { _Logger.Dispose(); } catch { }
        }

        // Pass null for no filter
        // File type defaults to .*
        public void GetFiles(SearchOption so, Predicate<String> filter, String extension = "*.*")
        {
            DwgList = System.IO.Directory.EnumerateFiles(_Path, extension, so)
                                         .ToList<String>()
                                         .FindAll(filter ?? (s => true));
        }

        //public void GetFileList(SearchOption so) 
        //{
        //    DwgList = Directory.EnumerateFiles(_Path, "*.*", so).ToList<String>();
        //    return;
        //}

        //public void GetDwgList(SearchOption so)
        //{
        //    //GetDwgList(so, (inFileStr) => { return true; });
        //    GetDwgList(so, FileStr => true);
        //    return;
        //}

        //public void GetDwgList(SearchOption so, Predicate<String> filter)
        //{
        //    DwgList = System.IO.Directory.EnumerateFiles(_Path, "*.dwg", so).ToList<String>().FindAll(filter);
        //    return;
        //}

        //public Boolean CheckDirPath()
        //{
        //    return _Path.isDirectoryPathOK();
        //}

        protected void StartTimer()
        {
            timer = Stopwatch.StartNew();
            return;
        }

        protected void StopTimer()
        {
            if (timer.IsRunning && timer != null)
            {
                timer.Stop();
            }
            return;
        }

        //Main dwg processing method for each program
        public abstract String Process();

        #region Template Pattern
        // In each dervied processing class, override initialize, performProcessing, and cleanup methods
        // and call below Process method (the template method) from view model. This will ensure the same template is being
        // used for all dwg processing functions.

        //public abstract void initialize();
        //public abstract void performProcessing();
        //public abstract String cleanup();

        //public sealed String Process()
        //{
        //    initialize();
        //    performProcessing();
        //    String result = cleanup();
        //    return result;
        //}
        #endregion

    }
}
