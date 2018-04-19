using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;

namespace AcadExts
{
    public class Logger : IDisposable
    {
        bool disposed = false;

        // Path
        private readonly String _path = String.Empty;
        public String Path { get { return _path; } }

        // Error count
        private Int32 _errorCount = 0;
        public Int32 ErrorCount { get { return _errorCount; } }

        private StreamWriter writer = null;

        //Ctor
        public  Logger(String inPath)
        {
            _path = inPath;

            try
            {
                writer = new System.IO.StreamWriter(path: _path, encoding: Encoding.ASCII, append:false);
            }
            catch {throw;}
        }

        // Log error message
        public void Log(String inError)
        {
            try
            {
                writer.WriteLine(String.Concat(inError, "   ", DateTime.Now.ToString("[M/d/yyyy  hh:mm:ss tt]"), Environment.NewLine), CultureInfo.GetCultureInfo("en-US"));
            }
            catch { }
            finally
            {
                _errorCount++;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (disposed) { return; }

            if (disposing && writer.BaseStream != null)
            {
                try
                {
                    writer.Flush();
                    writer.Close();
                }
                catch { }

                if (_errorCount == 0)
                {
                    // Try to delete log file automatically if its empty
                    try { File.Delete(_path); } catch { }
                }
            }
            disposed = true;
        }

        ~Logger()
        {
            Dispose(false);
        }
    }
}
