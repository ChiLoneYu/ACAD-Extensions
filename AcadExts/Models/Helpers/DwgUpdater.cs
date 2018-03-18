using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.ComponentModel;

namespace AcadExts
{
    // Used by FBD updater to update reference values based on a mapping faile of old and new values
    public class DwgUpdater
    {
        private enum FigType { fig = 1, sheet = 2, zone = 3 };

        // Defaults for vm
        public static Double LeftOfXDefault = .04;
        public static Double BelowYDefault = .33;
        public static Double RightOfXDefault = .8;
        public static Double AboveYDefault = .29;

        private readonly Double LeftOfX;
        private readonly Double BelowY;
        private readonly Double RightOfX;
        private readonly Double AboveY;

        private String oldfname { get; set; }
        private String newfname { get; set; }
        private Logger logger { get; set; }
        private Dictionary<String, Dictionary<String, String>> map { get; set; }

        private readonly Regex appendixRegex = new Regex(@"APPENDIX\s(((\w){1}(\s|$))|((\w){1}(\.(\d{1,3}(\.)?)?)($|\s)))");
        private readonly Regex chapterRegex = new Regex(@"CHAPTER\s(((\d){1,3}(\s|$))|((\d){1,3}(\.(\d{1,3}(\.)?)?)($|\s)))");
        private readonly Regex para0Regex = new Regex(@"(?<=^|\s)PARAGRAPH\s(\d{1,3}|[A-Z])-\d{1,3}(\.(\d{1,3}(\.)?)?)?(?=\s|$)");
        private readonly Regex tableRegex = new Regex(@"(^TEST\sPOINT$)|((?<=TABLE\s)(\d){1,3}-(\d){1,3}(\.((\d){1,3}(\.)?)?)?)(?=$|\s|,)");
        private readonly Regex tmNumRegex = new Regex(@"(\s|^)\d{1,4}-\d{1,4}-\d{1,4}-\d{1,4}-\d{1,4}($|\s|\.$)");
        private readonly Regex sectionRegex = new Regex(@"(^|\s)SECTION\s(\w){1,10}(\.(\d{1,3}(\.)?)?)?(\s|$)");
        private readonly Regex tmRegex = new Regex(@"(^|\s)(SA)?TM(\s|$)");
        private readonly Regex figValRegex = new Regex(@"(?<=\s|^|\()\d{1,3}-\d{1,3}(\.(\d{1,3}(\.)?)?)?(?=\s|$|\,|\)|\/|\()");
        private readonly Regex figsheetRegex = new Regex(@"((\(|\s)SH(S\s|\s))|(\d{1,3}-\d{1,3}/\d{1,3})");
        private readonly Regex figzoneRegex = new Regex(@"((\s\()|(^\())[A-Z]{1}\d{1,3}(\.|,|\)|\s|-)");
        private readonly Regex figureRegex = new Regex(@"(^|\s|\()FIGS?($|\s)");
        private readonly Regex figSheetNoFigRegex = new Regex(@"^\d{1,3}-\d{1,3}/\d{1,3}(,\d{1,3})?$");

        public DwgUpdater(String oldfnameIN, String newfnameIN, Dictionary<String, Dictionary<String, String>> mapIN, Logger loggerIN, Tuple<Double, Double, Double, Double> coordinatesIN)
        {
            LeftOfX = coordinatesIN.Item1;
            BelowY= coordinatesIN.Item2;
            RightOfX = coordinatesIN.Item3;
            AboveY = coordinatesIN.Item4;

            logger = loggerIN;
            oldfname = oldfnameIN;
            newfname = newfnameIN;
            map = mapIN;
        }

        /// <summary>
        /// Get figure type of figure text and object ids of nearby text
        /// </summary>
        /// <param name="dbtFigIn"></param>
        /// <param name="dbIn"></param>
        /// <param name="transIn"></param>
        /// <returns></returns>
        private Tuple<FigType, ObjectIdCollection, ObjectId> getFigInfo(DBText dbtFigIn, Database dbIn, Transaction transIn)
        {
            ObjectIdCollection oidc = new ObjectIdCollection();
            FigType figType = FigType.fig;
            DBText topmostDBT = dbtFigIn;

            using (BlockTableRecord btr2 = transIn.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(dbIn), OpenMode.ForRead) /*acTrans.GetObject(oid2, OpenMode.ForWrite)*/ as BlockTableRecord)
            {
                foreach (ObjectId oid in btr2)
                {
                    Entity ent = transIn.GetObject(oid, OpenMode.ForRead) as Entity;

                    if (ent.GetType() == typeof(DBText))
                    {
                        DBText dbt = ent as DBText;

                        if (dbt.IsPositionInRect(dbtFigIn.Position.X - LeftOfX, dbtFigIn.Position.Y - AboveY, dbtFigIn.Position.X + RightOfX, dbtFigIn.Position.Y + BelowY/*.006*/))
                        {
                            if (dbt.Position.Y > topmostDBT.Position.Y) { topmostDBT = dbt; }

                            if (!oidc.Contains(dbt.ObjectId)) { oidc.Add(dbt.ObjectId); }

                            if (figsheetRegex.IsMatch(dbt.TextString)) { figType = FigType.sheet; }

                            if (figzoneRegex.IsMatch(dbt.TextString)) { figType = FigType.zone; }
                        }
                    }
                }
            }
            return new Tuple<FigType, ObjectIdCollection, ObjectId>(figType, oidc, topmostDBT.Id);
        }

        // Use map from XML file to replace old values with new values
        public void Convert()
        {
            HashSet<ObjectId> changedFigIds = new HashSet<ObjectId>();
            HashSet<ObjectId> placedTopWpIds = new HashSet<ObjectId>();

            String newWP = "";

            // Get the new wp
            // Key doesn't matter because dwg can only have one WP
            if (map["wp"].Values.Count > 0)
            {
                if (!map["wp"].TryGetValue(map["wp"].Keys.FirstOrDefault(), out newWP))
                {
                    logger.Log("New WP could not be found");
                }
            }

            using (Database db = new Database(false, true))
            {
                try
                {
                    db.ReadDwgFile(oldfname, FileOpenMode.OpenForReadAndWriteNoShare, true, String.Empty);
                    db.CloseInput(true);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    throw new System.IO.IOException(String.Concat("Error reading DWG: ", e.Message));
                }

                try
                {
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                        {
                            foreach (ObjectId oidOuter in bt)
                            {
                                using (BlockTableRecord btr = acTrans.GetObject(oidOuter, OpenMode.ForWrite) as BlockTableRecord)
                                {
                                    foreach (ObjectId oidInner in btr)
                                    {
                                        using (Entity ent = acTrans.GetObject(oidInner, OpenMode.ForWrite) as Entity)
                                        {
                                            if (ent.GetType() == typeof(DBText))
                                            {
                                                using (DBText dbt = ent as DBText)
                                                {
                                                    String dbtText = dbt.TextString.Trim().ToUpper();

                                                    String oldValue = String.Empty;
                                                    List<String> oldValues = new List<String>();

                                                    String newValue = String.Empty;

                                                    DBText NearbyValue = new DBText();
                                                    System.Boolean changeNV = false;

                                                    #region APPENDIX
                                                    if (appendixRegex.IsMatch(dbtText))
                                                    {
                                                        oldValue = appendixRegex.Match(dbtText).Value.Trim().Replace("APPENDIX", "").Trim();

                                                        // Look up new value and replace old value
                                                        if (map["appendix"].TryGetValue(oldValue, out newValue))
                                                        {
                                                            try { dbt.TextString = dbtText.Replace(oldValue, newValue); }
                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                        }
                                                        else
                                                        {
                                                            logger.Log("New APPENDIX value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                        }
                                                        continue;
                                                    }
                                                    #endregion

                                                    #region CHAPTER
                                                    if (chapterRegex.IsMatch(dbtText))
                                                    {
                                                        oldValue = chapterRegex.Match(dbtText).Value.Trim().Replace("CHAPTER", "").Trim();

                                                        // Look up new value and replace old value
                                                        if (map["chapter"].TryGetValue(oldValue, out newValue))
                                                        {
                                                            try { dbt.TextString = dbtText.Replace(oldValue, newValue); }
                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                        }
                                                        else
                                                        {
                                                            logger.Log("New CHAPTER value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                        }
                                                        continue;
                                                    }
                                                    #endregion

                                                    #region PARA0
                                                    if (para0Regex.IsMatch(dbtText))
                                                    {
                                                        oldValue = para0Regex.Match(dbtText).Value.Trim().Replace("PARAGRAPH", "").Trim();

                                                        // Look up new value and replace old value
                                                        if (map["para0"].TryGetValue(oldValue, out newValue))
                                                        {
                                                            try { dbt.TextString = dbtText.Replace(oldValue, newValue); }
                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                        }
                                                        else
                                                        {
                                                            logger.Log("New PARA0 value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                        }
                                                        continue;
                                                    }
                                                    #endregion

                                                    #region TABLE
                                                    if (dbtText.Contains("TEST POINT") || dbtText.Contains("TABLE"))
                                                    {
                                                        if (String.Equals("TABLE", dbtText))
                                                        {
                                                            BlockTableRecord btrF = acTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite) as BlockTableRecord;
                                                            // Look nearby for #
                                                            using (BlockTableRecord btr2 = acTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite) /*acTrans.GetObject(oid2, OpenMode.ForWrite)*/ as BlockTableRecord)
                                                            {
                                                                foreach (ObjectId oid2Inner in btr2)
                                                                {
                                                                    Entity ent2 = acTrans.GetObject(oid2Inner, OpenMode.ForWrite) as Entity;
                                                                    if (ent2.GetType() == typeof(DBText))
                                                                    {
                                                                        NearbyValue = ent2 as DBText;
                                                                        if ((NearbyValue.Position.X == dbt.Position.X) &&
                                                                            (NearbyValue.ObjectId != dbt.ObjectId) &&
                                                                            (((NearbyValue.Position.Y - dbt.Position.Y) < 1) || ((dbt.Position.Y - NearbyValue.Position.Y) < 1)))
                                                                        {
                                                                            oldValue = NearbyValue.TextString.Trim();
                                                                            changeNV = true;
                                                                            break;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else if (tableRegex.IsMatch(dbtText))
                                                        {
                                                            oldValue = tableRegex.Match(dbtText).Value.Trim();
                                                        }
                                                        else if (String.Equals(dbtText, "TEST POINT"))
                                                        {
                                                            continue;
                                                        }
                                                        else
                                                        {
                                                            logger.Log("possible TABLE found in incorrect form: " + dbtText + " in DWG: " + oldfname);
                                                            continue;
                                                        }

                                                        if (map["table"].TryGetValue(oldValue, out newValue))
                                                        {
                                                            try
                                                            {
                                                                if (changeNV)
                                                                {
                                                                    NearbyValue.TextString = NearbyValue.TextString.Replace(oldValue, newValue);
                                                                }
                                                                else
                                                                {
                                                                    dbt.TextString = dbtText.Replace(oldValue, newValue);
                                                                }
                                                            }
                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                        }
                                                        else
                                                        {
                                                            logger.Log("New TABLE value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                        }
                                                        continue;
                                                    }
                                                    #endregion

                                                    #region SECTION

                                                    if (dbtText.Contains("SECTION"))
                                                    {
                                                        if (sectionRegex.IsMatch(dbtText))
                                                        {
                                                            oldValue = sectionRegex.Match(dbtText).Value.Trim().Replace("SECTION", "").Trim();
                                                        }
                                                        else
                                                        {
                                                            logger.Log("possible SECTION found in incorrect form: " + dbtText + " in DWG: " + oldfname);
                                                            continue;
                                                        }
                                                        if (map["section"].TryGetValue(oldValue, out newValue))
                                                        {
                                                            try { dbt.TextString = dbt.TextString.Replace(oldValue, newValue); }
                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                        }
                                                        else
                                                        {
                                                            logger.Log("New SECTION value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                        }
                                                        continue;
                                                    }

                                                    #endregion

                                                    #region TM

                                                    if (tmRegex.IsMatch(dbtText))
                                                    {
                                                        if (tmNumRegex.IsMatch(dbtText))
                                                        {
                                                            //String is tm and #, get just tm#. String can have more than one tm#
                                                            //oldValue = tmNumRegex.Match(dbtText).Value .Trim();
                                                            foreach (Match match in tmNumRegex.Matches(dbtText))
                                                            {
                                                                oldValues.Add(match.Value.Replace(".", "").Trim());
                                                            }
                                                        }
                                                        else
                                                        {
                                                            // check if # is below tm
                                                            using (BlockTableRecord btr2 = acTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite) /*acTrans.GetObject(oid2, OpenMode.ForWrite)*/ as BlockTableRecord)
                                                            {
                                                                foreach (ObjectId oid2Inner in btr2)
                                                                {
                                                                    Entity ent2 = acTrans.GetObject(oid2Inner, OpenMode.ForWrite) as Entity;
                                                                    if (ent2.GetType() == typeof(DBText))
                                                                    {
                                                                        NearbyValue = ent2 as DBText;

                                                                        if ((tmNumRegex.IsMatch(NearbyValue.TextString)) &&
                                                                            (NearbyValue.ObjectId != dbt.ObjectId) &&
                                                                            (((dbt.Position.Y - NearbyValue.Position.Y) < .14) && (dbt.Position.Y - NearbyValue.Position.Y) > 0))
                                                                        {
                                                                            oldValue = tmNumRegex.Match(NearbyValue.TextString).Value.Trim();
                                                                            changeNV = true;
                                                                            break;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        if (changeNV)
                                                        {
                                                            // # is below tm
                                                            if (map["tm"].TryGetValue(oldValue, out newValue))
                                                            {
                                                                NearbyValue.TextString = NearbyValue.TextString.Replace(oldValue, newValue);
                                                            }
                                                            else
                                                            {
                                                                logger.Log("New TM value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //String newString = dbt.TextString;
                                                            // # is same line as tm, but there might be more than one tm#
                                                            foreach (String ov in oldValues)
                                                            {
                                                                if (map["tm"].TryGetValue(ov, out newValue))
                                                                {
                                                                    //newString = newString.Replace(ov, newValue);
                                                                    try { dbt.TextString = dbt.TextString.Replace(ov, newValue); }
                                                                    catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                                }
                                                                else
                                                                {
                                                                    logger.Log("New TM value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                                }
                                                            }
                                                            //try { dbt.TextString = newString; }
                                                            //catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                        }
                                                        continue;//?
                                                    }
                                                    #endregion

                                                    #region FIG
                                                    if (figureRegex.IsMatch(dbtText))
                                                    {
                                                        Tuple<FigType, ObjectIdCollection, ObjectId> figInfo = getFigInfo(dbt, db, acTrans);

                                                        ObjectId topMostID = figInfo.Item3;
                                                        ObjectIdCollection oidcSet = figInfo.Item2;
                                                        FigType figType = figInfo.Item1;

                                                        switch (figType)
                                                        {
                                                            case FigType.sheet:

                                                                foreach (ObjectId oid in oidcSet)
                                                                {
                                                                    DBText line = acTrans.GetObject(oid, OpenMode.ForWrite) as DBText;

                                                                    if (oid == topMostID)
                                                                    {
                                                                            // Add Work Package above text
                                                                            DBText wp = new DBText();
                                                                            wp.TextString = newWP;
                                                                            wp.Position = new Point3d(line.Position.X, line.Position.Y + .1, line.Position.Z);
                                                                            btr.AppendEntity(wp);
                                                                            acTrans.AddNewlyCreatedDBObject(wp, true);
                                                                            placedTopWpIds.Add(wp.Id);
                                                                    }

                                                                    // get and replace only old fig value in this line
                                                                    if (figValRegex.IsMatch(line.TextString))
                                                                    {
                                                                        if (changedFigIds.Contains(line.ObjectId)) { continue; }

                                                                        oldValue = figValRegex.Match(line.TextString).Value.Trim();

                                                                        // BUG: Sometimes sheet # can look like figure number
                                                                        if (map["figsheet"].TryGetValue(oldValue, out newValue))
                                                                        {
                                                                            try
                                                                            {
                                                                                line.TextString = line.TextString.Replace(oldValue, newValue); 
                                                                                changedFigIds.Add(line.ObjectId);
                                                                            }
                                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                                        }
                                                                        else
                                                                        {
                                                                            logger.Log("New FIGSHEET value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                                        }
                                                                    }
                                                                }
                                                                break;

                                                            case FigType.fig:

                                                                foreach (ObjectId oid in oidcSet)
                                                                {
                                                                    DBText line = acTrans.GetObject(oid, OpenMode.ForWrite) as DBText;

                                                                    if (oid == topMostID)
                                                                    {
                                                                        // Add Work Package above text
                                                                        DBText wp = new DBText();
                                                                        wp.TextString = newWP;
                                                                        wp.Position = new Point3d(line.Position.X, line.Position.Y + .1, line.Position.Z);
                                                                        btr.AppendEntity(wp);
                                                                        acTrans.AddNewlyCreatedDBObject(wp, true);
                                                                        placedTopWpIds.Add(wp.Id);
                                                                    }

                                                                    // get and replace only old fig value in this line

                                                                    if (figValRegex.IsMatch(line.TextString))
                                                                    {
                                                                        if (changedFigIds.Contains(line.ObjectId)) { continue; }

                                                                        foreach (Match match in figValRegex.Matches(line.TextString))
                                                                        {
                                                                            //newValue = String.Empty;
                                                                            if (map["figure"].TryGetValue(match.Value, out newValue))
                                                                            {
                                                                                try
                                                                                {
                                                                                    line.TextString = line.TextString.Replace(match.Value, newValue);
                                                                                    changedFigIds.Add(line.ObjectId);
                                                                                }
                                                                                catch { logger.Log(match.Value + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                                            }
                                                                            else
                                                                            {
                                                                                logger.Log("New FIG value not found in XML file for old value: " + match.Value + " in DWG: " + oldfname);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                break;

                                                            case FigType.zone:

                                                                foreach (ObjectId oid in oidcSet)
                                                                {
                                                                    DBText line = acTrans.GetObject(oid, OpenMode.ForWrite) as DBText;

                                                                    if (oid == topMostID && !placedTopWpIds.Contains(oid))
                                                                    {
                                                                        // Add Work Package above text
                                                                        DBText wp = new DBText();
                                                                        wp.TextString = newWP;
                                                                        wp.Position = new Point3d(line.Position.X, line.Position.Y + .1, line.Position.Z);
                                                                        btr.AppendEntity(wp);
                                                                        acTrans.AddNewlyCreatedDBObject(wp, true);
                                                                        placedTopWpIds.Add(wp.Id);
                                                                    }

                                                                    // get and replace only old fig value in this line
                                                                    if (figValRegex.IsMatch(line.TextString))
                                                                    {
                                                                        if (changedFigIds.Contains(line.ObjectId)) { continue; }

                                                                        oldValue = figValRegex.Match(line.TextString).Value.Trim();

                                                                        if (map["figzone"].TryGetValue(oldValue, out newValue))
                                                                        {
                                                                            try 
                                                                            {
                                                                                line.TextString = line.TextString.Replace(oldValue, newValue);
                                                                                changedFigIds.Add(line.ObjectId); 
                                                                            }
                                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }
                                                                        }
                                                                        else
                                                                        {
                                                                            logger.Log("New FIGZONE value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname + " at Y= " + line.Position.Y);
                                                                        }
                                                                    }
                                                                }
                                                                break;

                                                            default:
                                                                // Won't get here
                                                                break;
                                                        }
                                                        continue;
                                                    }
                                                    #endregion

                                                    #region FIGSHEET (no FIG)
                                                    if (figSheetNoFigRegex.IsMatch(dbtText))
                                                    {
                                                        oldValue = figValRegex.Match(dbtText).Value.Trim();

                                                        if (map["figsheet"].TryGetValue(oldValue, out newValue))
                                                        {
                                                            // Replace old value with new value
                                                            try { dbt.TextString = dbt.TextString.Replace(oldValue, newValue); }
                                                            catch { logger.Log(oldValue + " could not be replaced with new value: " + newValue + " in DWG: " + oldfname); }

                                                            // Add Work Package above text
                                                            DBText wp = new DBText();
                                                            wp.TextString = newWP;
                                                            wp.Position = new Point3d(dbt.Position.X, dbt.Position.Y +.1, dbt.Position.Z);
                                                            btr.AppendEntity(wp);
                                                            acTrans.AddNewlyCreatedDBObject(wp ,true);
                                                        }
                                                        else
                                                        {
                                                            logger.Log("New FIGSHEET (no fig) value not found in XML file for old value: " + oldValue + " in DWG: " + oldfname);
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (System.Exception se) { throw new System.Exception("Error processing DWG: " + se.Message); }

                try { db.SaveAs(newfname, DwgVersion.Current); }
                catch (System.IO.IOException se) { throw new System.IO.IOException("Error saving new DWG: " + se.Message); }
            }
        }
    }
}
