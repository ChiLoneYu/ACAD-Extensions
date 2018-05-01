using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml;
using System.Xml.Linq;
using System.Diagnostics;
using System.IO;

namespace AcadExts
{
    [rtn("Use XML mapping file to update reference information text in FBD drawing files for Qatar")]
    internal sealed class FBDUpdater : DwgProcessor
    {
        Int32 numFiles;
        String xmlPath;
        XmlReader xmlR = null;
        readonly Tuple<double, double, double, double> Coordinates;
        private Boolean filesNotSpecified;

        public FBDUpdater(String inPath, BackgroundWorker inBw, String inXmlPath, Tuple<double, double, double, double> coordinates, Boolean inFilesNotSpecified)
            : base(inPath, inBw)
        {
            Coordinates = coordinates;
            this.xmlPath = inXmlPath;
            filesNotSpecified = inFilesNotSpecified;
        }

        public override String Process()
        {
            //if (!CheckDirPath()) { return "Invalid path: " + _Path; }
            try
            {
                BeforeProcessing();
            }
            catch(System.Exception se)
            {
                return "FBD Updater Exception: " + se.Message;
            }

            if (!xmlPath.isFilePathOK()) { return "Invalid XML file path: " + xmlPath; }

            if (!String.Equals(".xml", System.IO.Path.GetExtension(xmlPath).ToLower()))
            {
                return "XML file does not have '.xml' extension";
            }

            //try { _Logger = new Logger(_Path + "\\updatefbdsLog.txt"); }
            //catch (System.Exception se) { return "Could not create log file in: " + _Path + " because: " + se.Message; }


            XmlReaderSettings xmlRSettings = new XmlReaderSettings();
            xmlRSettings.IgnoreWhitespace = true;

            try { xmlR = XmlReader.Create(xmlPath, xmlRSettings); }
            catch (System.Exception se)
            {
                _Logger.Dispose();
                return "XML reader could not be created in: " + _Path + " because: " + se.Message;
            }

            //StartTimer();

            try
            {
                // if files are specified
                if (!filesNotSpecified)
                {
                    try
                    {
                        // Get number of file tags in xml file
                        numFiles = XDocument.Load(xmlPath).Root.Elements("file").Count();
                    }
                    catch { }

                    while (xmlR.Read())
                    {
                        if (_Bw.CancellationPending)
                        {
                            _Logger.Log("Processing cancelled by user at dwg " + DwgCounter + " out of " + numFiles);
                            break;
                        }

                        if (xmlR.NodeType == XmlNodeType.Element && String.Equals(xmlR.Name.Trim(), "file"))
                        {
                            String oldfname = xmlR.GetAttribute("oldfname");
                            String newfname = xmlR.GetAttribute("newfname");

                            Dictionary<String, Dictionary<String, String>> map = new Dictionary<String, Dictionary<String, String>>();

                            map.Add("appendix", new Dictionary<string, string>());
                            map.Add("chapter", new Dictionary<string, string>());
                            map.Add("para0", new Dictionary<string, string>());
                            map.Add("section", new Dictionary<string, string>());
                            map.Add("figure", new Dictionary<string, string>());
                            map.Add("figsheet", new Dictionary<string, string>());
                            map.Add("figzone", new Dictionary<string, string>());
                            map.Add("table", new Dictionary<string, string>());
                            map.Add("tm", new Dictionary<string, string>());
                            map.Add("wp", new Dictionary<string, string>());

                            while (xmlR.Read())
                            {
                                if (XmlNodeType.Element == xmlR.NodeType && String.Equals(xmlR.Name.Trim(), "map"))
                                {
                                    String reftype = xmlR.GetAttribute("reftype");
                                    String oldVal = xmlR.GetAttribute("old").Trim();
                                    String newVal = xmlR.GetAttribute("new").Trim();

                                    if (String.Equals(reftype, "wp")) { map["wp"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "appendix")) { map["appendix"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "chapter")) { map["chapter"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "para0")) { map["para0"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "section")) { map["section"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "figure")) { map["figure"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "figsheet")) { map["figsheet"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "figzone")) { map["figzone"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "table")) { map["table"].Add(oldVal, newVal); }
                                    if (String.Equals(reftype, "tm")) { map["tm"].Add(oldVal, newVal); }
                                }

                                if (XmlNodeType.EndElement == xmlR.NodeType && String.Equals(xmlR.Name.Trim(), "file"))
                                { break; }
                            }

                            DwgUpdater DwgUpdater = new DwgUpdater(String.Concat(_Path, "\\", oldfname), String.Concat(_Path, "\\", newfname), map, _Logger, Coordinates);

                            try
                            {
                                DwgUpdater.Convert();
                            }
                            catch (System.IO.IOException ioe)
                            {
                                _Logger.Log("Could not convert file: " + oldfname + " because: " + ioe.Message);
                                continue;
                            }
                            catch (System.Exception se)
                            {
                                _Logger.Log("Error processing file: " + oldfname + " because: " + se.Message);
                                continue;
                            }
                            DwgCounter++;
                            _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, numFiles));
                        }
                    }
                }
                // files are not specified
                else
                {
                    // Create Converted folder
                    try { Directory.CreateDirectory(_Path + "\\Converted"); }
                    catch { return "Unable to create new directory for converted files in: " + _Path; }

                    if (XDocument.Load(xmlPath).Root.Elements("file").Count() > 0)
                    {
                        _Logger.Log("File tags are being ignored because \"files not specified\" box was checked.");
                    }

                    //List<String> dwgs = new List<String>();

                    Dictionary<String, Dictionary<String, String>> map = new Dictionary<String, Dictionary<String, String>>();

                    map.Add("appendix", new Dictionary<string, string>());
                    map.Add("chapter", new Dictionary<string, string>());
                    map.Add("para0", new Dictionary<string, string>());
                    map.Add("section", new Dictionary<string, string>());
                    map.Add("figure", new Dictionary<string, string>());
                    map.Add("figsheet", new Dictionary<string, string>());
                    map.Add("figzone", new Dictionary<string, string>());
                    map.Add("table", new Dictionary<string, string>());
                    map.Add("tm", new Dictionary<string, string>());
                    map.Add("wp", new Dictionary<string, string>());

                    while (xmlR.Read())
                    {
                        if (XmlNodeType.Element == xmlR.NodeType && String.Equals(xmlR.Name.Trim(), "map"))
                        {
                            String reftype = xmlR.GetAttribute("reftype");
                            String oldVal = xmlR.GetAttribute("old").Trim();
                            String newVal = xmlR.GetAttribute("new").Trim();

                            // ignore repeat keys

                            if (String.Equals(reftype, "wp"))
                            {
                                try { map["wp"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "appendix"))
                            {
                                try { map["appendix"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "chapter"))
                            {
                                try { map["chapter"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "para0"))
                            {
                                try { map["para0"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "section"))
                            {
                                try { map["section"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "figure"))
                            {
                                try { map["figure"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "figsheet"))
                            {
                                try { map["figsheet"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "figzone"))
                            {
                                try { map["figzone"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "table"))
                            {
                                try { map["table"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                            if (String.Equals(reftype, "tm"))
                            {
                                try { map["tm"].Add(oldVal, newVal); }
                                catch (ArgumentException) { }
                            }
                        }
                    }

                    try
                    {
                        GetDwgList(SearchOption.TopDirectoryOnly);
                        //dwgs = System.IO.Directory.EnumerateFiles(_Path, "*.dwg", System.IO.SearchOption.TopDirectoryOnly)
                        //                          .Where(f => !f.Contains("_Converted"))
                        //                          .ToList<String>();
                        numFiles = NumDwgs;
                    }
                    catch (System.Exception se)
                    {
                        _Logger.Log("Could not enumerate files because: " + se.Message);
                        return "Could not access files in: " + _Path + " because: " + se.Message;
                    }

                    foreach (String currentDwg in DwgList)
                    {
                        if (_Bw.CancellationPending)
                        {
                            _Logger.Log("Processing cancelled by user at dwg " + DwgCounter + " out of " + numFiles);
                            break;
                        }

                        DwgUpdater dwgUpdater = new DwgUpdater(currentDwg,
                                                               String.Concat(_Path,
                                                                             "\\Converted\\",
                                                                             Path.GetFileName(currentDwg)),
                                                               map,
                                                               _Logger,
                                                               Coordinates);

                        try
                        {
                            dwgUpdater.Convert();
                        }
                        catch (System.IO.IOException ioe)
                        {
                            _Logger.Log("Could not convert file: " + currentDwg + " because: " + ioe.Message);
                            continue;
                        }
                        catch (System.Exception se)
                        {
                            _Logger.Log("Error processing file: " + currentDwg + " because: " + se.Message);
                            continue;
                        }

                        DwgCounter++;

                        try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                        catch { }
                    }
                }
            }
            catch (ArgumentException ae)
            {
                _Logger.Log("Argument Exception: " + ae.Message);
            }
            catch (System.Exception se)
            {
                _Logger.Log("Unhandled exception outside of dwg conversion: " + se.Message);
            }

            finally
            {
                AfterProcessing();

                // Close XML reader
                try { xmlR.Close(); }
                catch (System.Exception se) { _Logger.Log("Error closing XML reader: " + se.Message); }
            }

            return String.Concat(DwgCounter.ToString(),
                                 " out of ",
                                 numFiles.ToString(),
                                 " dwgs converted in ",
                                 TimePassed,
                                 ". ",
                                 (_Logger.ErrorCount > 0) ? ("Error Log: " + _Logger.Path) : ("No errors found."));
        }
    }
}
