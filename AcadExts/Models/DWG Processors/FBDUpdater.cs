using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml;
using System.Xml.Linq;
using System.Diagnostics;

namespace AcadExts
{
    [rtn("Use XML mapping file to update reference information values in FBD DWGs for Qatar")]
    internal sealed class FBDUpdater : DwgProcessor
    {
        Int32 numFileTags;
        String xmlPath;
        XmlReader xmlR = null;
        readonly Tuple<double, double, double, double> Coordinates;

        public FBDUpdater(String inPath, BackgroundWorker inBw, String inXmlPath, Tuple<double, double, double, double> coordinates)
            : base(inPath, inBw)
        {
            Coordinates = coordinates;
            this.xmlPath = inXmlPath;
        }

        public override String Process()
        {
            if (!CheckDirPath()) { return "Invalid path: " + _Path; }
            if (!xmlPath.isFilePathOK()) { return "Invalid XML file path: " + xmlPath; }
            if (!String.Equals(".xml", System.IO.Path.GetExtension(xmlPath).ToLower()))
            {
                return "XML file does not have '.xml' extension";
            }

            try { _Logger = new Logger(_Path + "\\updatefbdsLog.txt"); }
            catch (System.Exception se) { return "Could not create log file in: " + _Path + " because: " + se.Message; }

            //try { dwgFileArray = System.IO.Directory.GetDirectories(Path, "*.dwg"); }
            //catch { return "Could not get DWG files in: " + Path; }

            XmlReaderSettings xmlRSettings = new XmlReaderSettings();
            xmlRSettings.IgnoreWhitespace = true;

            try { xmlR = XmlReader.Create(xmlPath, xmlRSettings); }
            catch (System.Exception se)
            {
                _Logger.Dispose();
                return "XML reader could not be created in: " + _Path + " because: " + se.Message;
            }

            try
            {
                // Get number of file tags in xml file
                numFileTags = XDocument.Load(xmlPath).Root.Elements("file").Count();
            }
            catch { }

            StartTimer();

            try
            {
                while (xmlR.Read())
                {
                    if (_Bw.CancellationPending)
                    {
                        _Logger.Log("Processing cancelled by user at dwg " + DwgCounter + " out of " + numFileTags);
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

                        DwgUpdater dwgUpdater = new DwgUpdater(String.Concat(_Path, "\\", oldfname), String.Concat(_Path, "\\", newfname), map, _Logger, Coordinates);

                        try
                        {
                            dwgUpdater.Convert();
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
                        _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, numFileTags));
                    }
                }
            }
            catch (System.Exception se)
            {
                _Logger.Log("Unhandled exception outside of dwg conversion: " + se.Message);
            }

            finally
            {
                StopTimer();

                // Close logger
                _Logger.Dispose();

                // Close XML reader
                try { xmlR.Close(); }
                catch (System.Exception se) { _Logger.Log("Error closing XML reader: " + se.Message); }
            }

            return String.Concat(DwgCounter.ToString(),
                                 " out of ",
                                 numFileTags.ToString(),
                                 " dwgs converted in ",
                                 TimePassed);
        }
    }
}
