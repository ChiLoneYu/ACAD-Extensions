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
        private readonly String _path = String.Empty;
        protected String _Path { get { return _path; } }

        // Processing BackgroundWorker
        private readonly BackgroundWorker _bw = null;
        protected BackgroundWorker _Bw { get { return _bw; } }

        // Error logger
        protected Logger _Logger { get; set; }

        // Processing timer
        private System.Diagnostics.Stopwatch timer { get; set; }

        // # of dwgs in list
        protected Int32 NumDwgs { get { return DwgList.Count; } }

        private List<String> _DwgList;
        protected List<String> DwgList { get { return _DwgList; } }

        private Int32 _DwgCounter = 0;
        protected Int32 DwgCounter { get { return _DwgCounter; } set { _DwgCounter = value; } }

        protected String TimePassed { get { return timer.Elapsed.TotalSeconds.PrintTimeFromSeconds() ?? " null "; } }

        // Ctor
        public DwgProcessor(String inPath, BackgroundWorker inBw)
        {
            if (!String.IsNullOrWhiteSpace(inPath)) 
            {
                _path = inPath.TrimEnd(new[] { '\\' });
            }

            _bw = inBw;

            if (_Bw == null)
            {
                if (_Logger != null)
                {
                    try { _Logger.Log("BackgroundWorker is null"); }
                    catch { }
                }
                throw new ArgumentNullException("inBw", "Base dwg processor needs a BackgroundWorker instance");
            }
        }

        public void GetDwgList(SearchOption so)
        {
            GetDwgList(so, (inFileStr) => { return true; });
            return;
        }

        public void GetDwgList(SearchOption so, Predicate<String> filter)
        {
            _DwgList = System.IO.Directory.EnumerateFiles(_Path, "*.dwg", so).ToList<String>().FindAll(filter);
            return;
        }

        public Boolean CheckDirPath()
        {
            return _Path.isDirectoryPathOK();
        }

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
        // In each dervied processing class, override start, finish, and perform processing methods
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
