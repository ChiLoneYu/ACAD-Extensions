using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.IO;
using System.Windows;
using System.Xml;
using System.Text.RegularExpressions;
using System.ComponentModel;

namespace AcadExts
{
    [rtn("Generates XML keyfiles for RPSTL DWGs")]
    internal sealed class KeyfileGenerator : DwgProcessor
    {
        Regex letterRegex = new Regex(@"^(\s)*[A-Z]\d{6}(\S)*(\s)*$");
        Regex msRegex = new Regex(@"^(\s)*MS\d{6}(\S)*(\s)*$");
        Regex viewLetterRegex = new Regex(@"^([A-Z]){1,2}$");
        Regex calloutRegex = new Regex(@"^(\d){1,4}([A-Z])?$");
        Regex refDesRegex = new Regex(@"^[A-Z]{1,2}(\d)+((\s)*(\S)*)*$");

        //Autodesk.AutoCAD.EditorInput.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

        public KeyfileGenerator(String inPath, BackgroundWorker inBw)
            : base(inPath, inBw)
        {
        }

        public override String Process()
        {
            try
            {
                BeforeProcessing();
            }
            catch(System.Exception se)
            {
                return "Keyfile Generation processing exception: " + se.Message;
            }
            //if (!CheckDirPath()) { return "Invalid path: " + _Path; }

            //try
            //{
            //    _Logger = new Logger(String.Concat(_Path, "\\KeyfileErrors.txt"));
            //}
            //catch (System.Exception se)
            //{
            //    // FATAL ERROR
            //    return "Could not create log file in: " + _Path + " because: " + se.Message;
            //}

            //StartTimer();

            // Get all DWG files
            try
            {
                GetDwgList(SearchOption.TopDirectoryOnly, delegate(String inFile) { return ((System.IO.Path.GetFileNameWithoutExtension(inFile).Length < 15)); });
            }
            catch (SystemException se)
            {
                _Logger.Log(" DWG files could not be enumerated because: " + se.Message);
                _Logger.Dispose();
                return " DWG files could not be enumerated because: " + se.Message;
            }

            if (NumDwgs == 0) { return "No DWGs found in: " + _Path; }

            foreach (String currentDWG in DwgList)
            {
                DwgCounter++;

                try { _Bw.ReportProgress(Utilities.GetPercentage(DwgCounter, NumDwgs)); }
                catch { }

                if (_Bw.CancellationPending)
                {
                    _Logger.Log("Keyfile generation cancelled by user at dwg " + DwgCounter.ToString() + " out of " + DwgList.Count);
                    break;
                }

                String dwgNameNoExt = System.IO.Path.GetFileNameWithoutExtension(currentDWG);
                String dwgNameExt = System.IO.Path.GetFileName(currentDWG);
                String ms = String.Empty;

                using (Database db = new Database(false, true))
                {
                    try
                    {
                        db.ReadDwgFile(currentDWG, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                        db.Regenmode = true;
                        db.CloseInput(true);
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception e)
                    {
                        _Logger.Log(String.Concat("Could not read DWG: ", dwgNameNoExt, " because: ", e.Message, "...Continuing to next DWG"));
                        continue;
                    }

                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        // Unlock and thaw all layers
                        ToEach.toEachLayer(db, acTrans, delegate(LayerTableRecord ltr)
                        {
                            if (ltr.IsLocked) { ltr.IsLocked = false; }
                            if (ltr.IsFrozen) { ltr.IsFrozen = false; }
                            return false;
                        });

                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Encoding = Encoding.ASCII;
                        settings.Indent = true;
                        XmlWriter xmlW = XmlWriter.Create(_Path + "\\" + dwgNameNoExt + "_key.xml", settings);

                        if (xmlW == null)
                        {
                            _Logger.Log("XML writer could not be initialized.");
                            _Logger.Dispose();
                            return "XML file could not be initialized.";
                        }

                        try
                        {
                            xmlW.WriteStartDocument();
                            xmlW.WriteStartElement("rpstl_keyfile");
                            xmlW.WriteAttributeString("version", "2.0");
                            xmlW.WriteStartElement("filename");
                            xmlW.WriteAttributeString("text", dwgNameExt);
                            xmlW.WriteEndElement();

                            // Write dwg extents
                            xmlW.WriteStartElement("extents");
                            xmlW.WriteAttributeString("x1", db.Extmin.X.truncstring(3));
                            xmlW.WriteAttributeString("y1", db.Extmin.Y.truncstring(3));
                            xmlW.WriteAttributeString("x2", db.Extmax.X.truncstring(3));
                            xmlW.WriteAttributeString("y2", db.Extmax.Y.truncstring(3));

                            double height = db.Extmax.Y - db.Extmin.Y;
                            double width = db.Extmax.X - db.Extmin.X;

                            // Write document height and width
                            xmlW.WriteAttributeString("height", height.truncstring(3));
                            xmlW.WriteAttributeString("width", width.truncstring(3));

                            xmlW.WriteEndElement();

                            // Unlock all layers
                            ToEach.toEachLayer(db, acTrans, delegate(LayerTableRecord ltr)
                                                            {
                                                                if (ltr.IsLocked) { ltr.IsLocked = false; }
                                                                return false;
                                                            }
                            );

                            db.Regenmode = true;

                            ObjectIdCollection oidc = new ObjectIdCollection();
                            ToEach.toEachEntity(db, acTrans, delegate(Entity ent) { oidc.Add(ent.Id); return false; });
                            db.Purge(oidc);

                            //Make Dictionary to map layerIds to layer names
                            // Note neccessary because the DBText .Layer property isn't working
                            Dictionary<ObjectId, String> LayerIdToLayerName = new Dictionary<ObjectId, string>();

                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId layerOid in lt)
                                {
                                    using (LayerTableRecord ltr = acTrans.GetObject(layerOid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        LayerIdToLayerName.Add(ltr.Id, ltr.Name);
                                    }
                                }
                            }

                            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
                            {
                                foreach (ObjectId oid in lt)
                                {
                                    List<DBText> msNos = new List<DBText>();
                                    List<DBText> callouts = new List<DBText>();
                                    List<DBText> viewLetters = new List<DBText>();
                                    List<DBText> miscTexts = new List<DBText>();
                                    List<BlockReference> cards = new List<BlockReference>();
                                    //IEnumerable<DBText> msnos = new List<DBText>();
                                    StringBuilder prefixString = new StringBuilder();

                                    using (LayerTableRecord ltr = acTrans.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                                    {
                                        List<DBText> linesBelowDesignations = new List<DBText>();

                                        using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                                        {
                                            using (BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord)
                                            {
                                                foreach (ObjectId oidbtr in btr)
                                                {
                                                    Entity ent = acTrans.GetObject(oidbtr, OpenMode.ForRead) as Entity;

                                                    if (ent.LayerId == ltr.Id)
                                                    {
                                                        if (ent.GetType() == typeof(DBText))
                                                        {
                                                            DBText dbt = ent as DBText;

                                                            String layerName = LayerIdToLayerName[dbt.LayerId];
                                                            // Skip everything on "FILENAME" and "SCALE" layers
                                                            if (String.Equals(layerName, "FILENAME") || String.Equals(layerName, "SCALE"))
                                                            {
                                                                continue;
                                                            }

                                                            if (dbt.TextString.Trim().ToUpper().EndsWith("(REF)"))
                                                            {
                                                                // Skip old refdess
                                                                continue;
                                                            }

                                                            // P/0 instead of P/O
                                                            if (String.Equals(dbt.TextString.Trim(), "P/0"))
                                                            {
                                                                _Logger.Log("P/0 found instead of P/O on layer: " + dbt.Layer + " in DWG: " + dwgNameNoExt);
                                                            }

                                                            if (calloutRegex.IsMatch(dbt.TextString) && !dbt.TextString.Contains("-") && layerName.StartsWith("m"))
                                                            {
                                                                callouts.Add(dbt);
                                                                continue;
                                                            }

                                                            // Skip view letters that are on layers that aren't m layers
                                                            if (viewLetterRegex.IsMatch(dbt.TextString.Trim()) && layerName.StartsWith("m"))
                                                            {
                                                                viewLetters.Add(dbt);
                                                                continue;
                                                            }

                                                            // Skip msnos that are on layers that aren't m layers
                                                            if ((msRegex.IsMatch(dbt.TextString.ToUpper()) || letterRegex.IsMatch(dbt.TextString.ToUpper())) && layerName.StartsWith("m"))
                                                            {
                                                                msNos.Add(dbt);
                                                                continue;
                                                            }

                                                            if (refDesRegex.IsMatch(dbt.TextString.ToUpper().Trim())) //&&
                                                            {
                                                                // Do nothing
                                                                // Look for refdess on a per-callout basis because refdes
                                                                // must be within a certain range of each callout
                                                                continue;
                                                            }

                                                            // Check if text refdes
                                                            if (dbt.TextString.Contains("DESIGNATIONS"))
                                                            {
                                                                // Get prefix from line with Designations
                                                                prefixString.Append(dbt.TextString.Substring(dbt.TextString.IndexOf("WITH") + 5).Trim());

                                                                // Get each line below Designations line and add it to list 
                                                                linesBelowDesignations = ToEach.toEachDBText(db,
                                                                                                             acTrans,
                                                                                                             delegate(DBText dbtUnderDes)
                                                                                                             {
                                                                                                                 double yDiff = dbt.Position.Y - dbtUnderDes.Position.Y;
                                                                                                                 double xDiff = Math.Abs(dbt.Position.X - dbtUnderDes.Position.X);

                                                                                                                 return ((dbt.Id != dbtUnderDes.Id) &&
                                                                                                                         (dbt.LayerId == dbtUnderDes.LayerId) &&
                                                                                                                         (dbt.Position.Y > dbtUnderDes.Position.Y) &&
                                                                                                                         (yDiff < .5) &&
                                                                                                                         (xDiff < .02));
                                                                                                             });

                                                                // Append each line to String and format it with spaces and commas
                                                                foreach (DBText lineBelowDes in linesBelowDesignations)
                                                                {
                                                                    if (prefixString[prefixString.Length - 1] != ',')
                                                                    {
                                                                        prefixString.Append(String.Concat(", ", lineBelowDes.TextString.Trim()));
                                                                        continue;
                                                                    }
                                                                    if (prefixString[prefixString.Length - 1] == ',')
                                                                    {
                                                                        prefixString.Append(String.Concat(" ", lineBelowDes.TextString.Trim()));
                                                                        continue;
                                                                    }
                                                                }

                                                                // Delete trailing comma if it exists
                                                                if (prefixString[prefixString.Length - 1] == ',') { prefixString.Length--; }
                                                                // TODO ?
                                                                continue;
                                                            }

                                                            miscTexts.Add(dbt);
                                                        }

                                                        if (ent.GetType() == typeof(BlockReference))
                                                        {
                                                            BlockReference br = acTrans.GetObject(oidbtr, OpenMode.ForRead) as BlockReference;

                                                            if (br.Name.Contains("PRENOTE"))
                                                            {
                                                                _Logger.Log("DWG: " + dwgNameExt + ": layer: " + LayerIdToLayerName[br.LayerId] + ": " + "block reference prefix note found");
                                                            }
                                                            if (br.Name.Contains("HCRDTBL"))
                                                            {
                                                                cards.Add(br);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        xmlW.WriteStartElement("layer");
                                        xmlW.WriteAttributeString("text", ltr.Name);

                                        if (prefixString.Length > 0) { xmlW.WriteAttributeString("refdes_prefix", prefixString.ToString()); }

                                        foreach (BlockReference br in cards)
                                        {
                                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol = br.AttributeCollection;

                                            AttributeReference attRefRefDes = acTrans.GetObject(br.AttributeCollection[0], OpenMode.ForRead) as AttributeReference;
                                            AttributeReference attRefItem = acTrans.GetObject(br.AttributeCollection[1], OpenMode.ForRead) as AttributeReference;

                                            xmlW.WriteStartElement("callout");
                                            xmlW.WriteAttributeString("text", (String.IsNullOrWhiteSpace(attRefItem.TextString)) ? "no_itemno_for_card_slot" : attRefItem.TextString);
                                            xmlW.writeOutAttRefMinMax(attRefItem);

                                            xmlW.WriteStartElement("refdes");
                                            xmlW.WriteAttributeString("text", attRefRefDes.TextString);
                                            xmlW.writeOutAttRefMinMax(attRefRefDes);
                                            xmlW.WriteEndElement();

                                            xmlW.WriteEndElement();
                                        }
                                        foreach (DBText viewLetter in viewLetters)
                                        {
                                            TextEntity tent = new TextEntity(viewLetter);
                                            xmlW.WriteStartElement("view_letter");
                                            xmlW.WriteAttributeString("text", tent.text);
                                            xmlW.writeOutMinMax(tent, _Logger);
                                            xmlW.WriteEndElement();
                                        }

                                        foreach (DBText msNo in msNos)
                                        {
                                            TextEntity tent = new TextEntity(msNo);
                                            xmlW.WriteStartElement("msno");
                                            xmlW.WriteAttributeString("text", tent.text);
                                            xmlW.writeOutMinMax(tent, _Logger);
                                            xmlW.WriteEndElement();
                                        }

                                        foreach (DBText callout in callouts)
                                        {
                                            TextEntity tent = new TextEntity(callout);
                                            xmlW.WriteStartElement("callout");
                                            xmlW.writeOutMinMax(tent, _Logger);
                                            List<DBText> refDess = new List<DBText>();
                                            refDess = ToEach.toEachDBText(db,
                                                                            acTrans,
                                                                            delegate(DBText possibleRefDes)
                                                                            {
                                                                                const System.Double yRange = .15;
                                                                                const System.Double xRange = .15;

                                                                                return (possibleRefDes.Id != callout.Id &&
                                                                                        possibleRefDes.LayerId == callout.LayerId &&
                                                                                        refDesRegex.IsMatch(possibleRefDes.TextString.ToUpper().Trim()) &&
                                                                                    //!possibleRefDes.TextString.Contains(",") &&
                                                                                    //!possibleRefDes.TextString.Contains("-") &&
                                                                                        (possibleRefDes.IsAbove(callout, .1135 /*.5*/) ||
                                                                                            possibleRefDes.IsInRect(callout.AlignmentPoint.X - xRange,
                                                                                                                    callout.AlignmentPoint.Y,
                                                                                                                    callout.AlignmentPoint.X,
                                                                                                                    callout.AlignmentPoint.Y + yRange) ||
                                                                                            possibleRefDes.IsInRect(callout.AlignmentPoint.X,
                                                                                                                    callout.AlignmentPoint.Y,
                                                                                                                    callout.AlignmentPoint.X + xRange,
                                                                                                                    callout.AlignmentPoint.Y + yRange)) &&
                                                                                        !callout.IsRefDesBetween(possibleRefDes, acTrans, db)
                                                                                        );
                                                                            });
                                            if (refDess.Count > 1)
                                            {
                                                if (callout.Justify != AttachmentPoint.MiddleCenter)
                                                {
                                                    _Logger.Log("DWG: " +
                                                                 dwgNameExt +
                                                                 " layer: " +
                                                                 LayerIdToLayerName[tent.dbText.LayerId] +
                                                                 " callout: " +
                                                                 tent.text +
                                                                 " has an alignment point that isn't middle center" +
                                                                 " (" +
                                                                 tent.dbText.Position.X.truncstring(3) +
                                                                 ", " +
                                                                 tent.dbText.Position.Y.truncstring(3) +
                                                                 ") ");
                                                }
                                                else
                                                {
                                                    _Logger.Log("DWG: " +
                                                                 dwgNameExt +
                                                                 " layer: " +
                                                                 LayerIdToLayerName[tent.dbText.LayerId] +
                                                                 " has more than one refdes for callout: " +
                                                                 tent.text +
                                                                 " (" +
                                                                 tent.dbText.Position.X.truncstring(3) +
                                                                 ", " +
                                                                 tent.dbText.Position.Y.truncstring(3) +
                                                                 ") ");//tent.dbText.Layer);
                                                }
                                            }

                                            xmlW.WriteAttributeString("text", tent.text);

                                            foreach (DBText refdes in refDess)
                                            {
                                                TextEntity tentRefDes = new TextEntity(refdes);
                                                xmlW.WriteStartElement("refdes");
                                                xmlW.WriteAttributeString("text", tentRefDes.text);
                                                xmlW.writeOutMinMax(tentRefDes, _Logger);
                                                xmlW.WriteEndElement();
                                            }
                                            xmlW.WriteEndElement();
                                        }
                                        foreach (DBText otherText in miscTexts)
                                        {
                                            if (!linesBelowDesignations.Contains(otherText))
                                            {
                                                TextEntity tent = new TextEntity(otherText);
                                                xmlW.WriteStartElement("other_text");
                                                xmlW.WriteAttributeString("text", tent.text);
                                                xmlW.writeOutMinMax(tent, _Logger);
                                                xmlW.WriteEndElement();
                                            }
                                        }
                                        xmlW.WriteEndElement();
                                    }
                                }
                            }
                        }
                        catch (System.Exception se)
                        {
                            _Logger.Log("DWG: " + dwgNameExt + ": processing error of type: " + se.GetType().ToString() + " : " + se.Message);
                        }
                        finally
                        {
                            xmlW.WriteEndElement();
                            xmlW.WriteRaw(Environment.NewLine);
                            xmlW.WriteEndDocument();
                            xmlW.Flush();
                            xmlW.Close();
                        }
                        acTrans.Commit();
                    }
                }
            }

            AfterProcessing();

            return (String.Concat(DwgCounter.ToString(),
                                  " out of ",
                                  NumDwgs,
                                  " DWGs processed in ",
                                  TimePassed,
                                  ". ",
                                  (_Logger.ErrorCount > 0) ? ("Error Log: " + _Logger.Path) : ("No errors found.")));
        }
    }
}
