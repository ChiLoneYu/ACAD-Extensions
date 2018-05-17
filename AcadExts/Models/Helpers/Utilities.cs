using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Xml;
using System.Text.RegularExpressions;
using System.ComponentModel;

namespace AcadExts
{
    // Misc. static helper functions
    public static class Utilities
    {
        public static String nl = Environment.NewLine;

        public static String n(int num)
        {
            String s = String.Empty;
            if (num < 1) { return s; }

            try { s = String.Concat(Enumerable.Repeat<String>(Environment.NewLine, num)); }
            catch { }

            return s;
        }

        // Takes seconds arg, rounds and returns string of minutes if seconds are >60 and seconds if <60
        public static String PrintTimeFromSeconds(this double inSeconds, int places = 2)
        {
            //const int places = 2;

            if ( inSeconds <= 0.0) { return "0 seconds"; }

            try
            {
                return (inSeconds < 60) ? (String.Concat((Math.Round(inSeconds, places, MidpointRounding.AwayFromZero)).ToString(), " seconds"))
                                        : (String.Concat((Math.Round((inSeconds / 60.00d), places)).ToString(), " minutes"));
            }
            catch { return ""; }
        }

        public static double TextWidth(this DBText dbTextIn)
        {
            const double Length8Chars = .4329;
            const double Length9Chars = .4757;

            if (dbTextIn.TextString.Length == 8) { return Length8Chars; }

            if (dbTextIn.TextString.Length == 9) { return Length9Chars; }

            return -1;

            // String txt = dbTextIn.TextString.Trim().ToUpper();
            //if (txt.Substring(txt.LastIndexOf("-") + 1).Length == 1) { return .5571d; }
            //if (txt.Substring(txt.LastIndexOf("-") + 1).Length == 2) { return  }

            //double width = dbTextIn.AlignmentPoint.X - dbTextIn.Position.X;
            //return ( (width <= 0) ? (-1.0d) : (dbTextIn.AlignmentPoint.X - dbTextIn.Position.X) );
        }

        public static IEnumerable<DBObject> getAllDwgObjs(Database db, Transaction acTrans, OpenMode openType, Logger inLogger)
        {
            BlockTable bt = acTrans.GetObject(db.BlockTableId, openType) as BlockTable;
            foreach (ObjectId btrID in bt)
            {
                BlockTableRecord btr = acTrans.GetObject(btrID, openType) as BlockTableRecord;
                foreach (ObjectId id in btr)
                {
                    DBObject dbo = null;

                    try
                    {
                        dbo = acTrans.GetObject(id, openType) as DBObject;
                    }
                    catch
                    {
                        if (dbo != null)
                        {
                            Entity ent = dbo as Entity;
                            inLogger.Log(db.Filename + ": Attempt to access object on locked layer: " + ent.Layer);
                        }
                        else
                        {
                            inLogger.Log(db.Filename + ": Attempt to access null object in: " + db.Filename);
                        }
                        continue;
                    }
                    yield return dbo;
                }
            }
        }

        public static Boolean isFramePageLayerName(this String inStringLayerName)
        {
            if ((inStringLayerName.Trim().Length >= 15) && char.IsLetter((inStringLayerName[0])))
            {
                return true;
            }
            else { return false; }
        }

        public static void delLayer(Transaction acTrans, Database db, String layerName)
        {
            // If the layer is locked, unlock it
            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
            {
                if (lt.Has(layerName))
                {
                    ObjectId layerOID = lt[layerName];

                    using (LayerTableRecord ltr = acTrans.GetObject(layerOID, OpenMode.ForWrite) as LayerTableRecord)
                    {
                        if (ltr.IsLocked) { ltr.IsLocked = false; }
                    }
                }
            }

            // Delete all objects on layer
            ObjectIdCollection coutObjs = new ObjectIdCollection();
            using (BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
            {
                foreach (ObjectId oid in bt)
                {
                    using (BlockTableRecord btr = acTrans.GetObject(oid, OpenMode.ForWrite) as BlockTableRecord)
                    {
                        foreach (ObjectId oidEnt in btr)
                        {
                            Entity ent = acTrans.GetObject(oidEnt, OpenMode.ForRead) as Entity;
                            if (ent.Layer == layerName)
                            {
                                ent.UpgradeOpen();
                                coutObjs.Add(ent.Id);
                                ent.Erase(true);
                            }
                        }
                    }
                }
            }
            db.Purge(coutObjs);

            // Delete layer
            using (LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable)
            {

                if (lt.Has(layerName))
                {
                    // Check to see if it is safe to erase layer
                    ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(lt[layerName]);
                    db.Purge(acObjIdColl);
                    if (acObjIdColl.Count > 0)
                    {
                        LayerTableRecord acLyrTblRec = acTrans.GetObject(acObjIdColl[0], OpenMode.ForWrite) as LayerTableRecord;

                        acLyrTblRec.Erase(true);
                    }
                    else
                    {
                        //Wrong: https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2016/ENU/AutoCAD-NET/files/GUID-DF8A64D3-AE09-4BCE-B9E1-1B642DA4FCFF-htm.html
                        LayerTableRecord acLyrTblRec = acTrans.GetObject(acObjIdColl[0], OpenMode.ForWrite) as LayerTableRecord;

                        acLyrTblRec.Erase(true);
                    }
                    db.Purge(acObjIdColl);
                }
            }
        }

        public static System.Boolean HasNums(this String inString)
        {
            if (String.IsNullOrWhiteSpace(inString)) { return false; }

            String stringNums = new String(inString.ToCharArray().Where(c => char.IsDigit(c)).ToArray());

            return (!String.IsNullOrWhiteSpace(stringNums));
        }

        public static System.Boolean IsPositionInRect(this DBText refDes, double LLX, double LLY, double TRX, double TRY)
        {
            return ((refDes.Position.X >= LLX) &&
                    (refDes.Position.X <= TRX) &&
                    (refDes.Position.Y >= LLY) &&
                    (refDes.Position.Y <= TRY)
                   );
        }

        // Returns true if file is a valid and accessible file
        //public static Boolean isFilePathOK(this String inFilePath)
        //{
        //    FileInfo fi = null;

        //    if (String.IsNullOrWhiteSpace(inFilePath)) { return false; }

        //    try { fi = new FileInfo(inFilePath); }
        //    catch { return false; }

        //    if (!File.Exists(inFilePath) || fi == null) { return false; }

        //    return true;
        //}

        // Returns true if directory is a valid and accessible directory
        public static Boolean isDirectoryPathOK(this String inDirectoryPath)
        {
            DirectoryInfo di = null;

            if (String.IsNullOrWhiteSpace(inDirectoryPath)) { return false; }

            try { di = new DirectoryInfo(inDirectoryPath); }
            catch { return false; }

            if (!Directory.Exists(inDirectoryPath) || di == null) { return false; }
            
            return true;
        }

        // Gives percentage value of current page
        public static Int32 GetPercentage(Int32 current, Int32 total)
        {
            //Math.Ceiling(Convert.ToInt32())
            if (current < 0) { throw new ArgumentException("Invalid argument", "current"); }
            if (total < 1) { throw new ArgumentException("Invalid argument", "total"); }

            return ((Int32) Math.Round((((double) current / (double) total) * 100), MidpointRounding.AwayFromZero));
        }

        // Truncates double after given # of decimal places
        public static String truncstring(this double inDouble, int numPlaces = 3)
        {
            if (numPlaces < 0)
            {
                throw new ArgumentOutOfRangeException("numPlaces", "Cannot truncate argument to less than 0 places");
            }

            return inDouble.ToString(String.Concat("###.", String.Concat(Enumerable.Repeat<String>("0", numPlaces))));

            // https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-numeric-format-strings
            //Char[] placeArray;
            //placeArray = new Char[numPlaces];
        }

        public static System.Boolean IsAbove(this DBText refDes, DBText inCallout, double inYRange)
        {
            return (// Refdes is above callout
                   (refDes.Position.Y > inCallout.Position.Y) &&
                // X coordinates are the same
                   (refDes.AlignmentPoint.X == inCallout.AlignmentPoint.X) &&
                // Callout Y is within yrange units units away from REFDES Y
                   (Math.Abs(refDes.AlignmentPoint.Y) - Math.Abs(inCallout.AlignmentPoint.Y)) < inYRange);
        }

        // Returns true if there is at least 1 refdes between the given callout and given refdes, otherwise false.
        public static System.Boolean IsRefDesBetween(this DBText inCallout, DBText possibleRefDes, Transaction acTrans, Database db)
        {
            //Regex refDesRegex = new Regex(@"^[A-Z]+(\S)*(\d)+(\S)*$");
            Regex refDesRegex = new Regex(@"^[A-Z]+(\d)+((\S*)(\s)*)*$");
            return (ToEach.toEachDBText(db, acTrans, delegate(DBText possibleMiddleRefDes)
                {
                    return ((possibleMiddleRefDes.AlignmentPoint.X == inCallout.AlignmentPoint.X) &&
                            (inCallout.AlignmentPoint.X == possibleRefDes.AlignmentPoint.X) &&
                            (possibleMiddleRefDes.AlignmentPoint.Y < possibleRefDes.AlignmentPoint.Y) &&
                            (possibleMiddleRefDes.AlignmentPoint.Y > inCallout.AlignmentPoint.Y) &&
                            (refDesRegex.IsMatch(possibleMiddleRefDes.TextString.Trim().ToUpper())) &&
                            (possibleMiddleRefDes.LayerId == possibleRefDes.LayerId) &&
                            (possibleRefDes.Id != possibleMiddleRefDes.Id)
                           );
                }).Count >= 1);
        }

        // Returns true if there is at least 1 text object within the rectangle of the given coordinates.
        public static System.Boolean IsInRect(this DBText refDes, double LLX, double LLY, double TRX, double TRY)
        {
            return ((refDes.AlignmentPoint.X >= LLX) &&
                    (refDes.AlignmentPoint.X <= TRX) &&
                    (refDes.AlignmentPoint.Y >= LLY) &&
                    (refDes.AlignmentPoint.Y <= TRY));
        }

        // XML writer extension method to write out bottom left and top right coordinates of a given textEntity as XML attributes
        public static void writeOutMinMax(this XmlWriter xmlWriterIn, TextEntity textObjIn, Logger inLogger)
        {
            // If text entity is null write blanks
            if (textObjIn.id.IsNull)
            {
                xmlWriterIn.WriteAttributeString("x1", "");
                xmlWriterIn.WriteAttributeString("y1", "");
                xmlWriterIn.WriteAttributeString("x2", "");
                xmlWriterIn.WriteAttributeString("y2", "");
                return;
            }
            // Get bottom left coordinates using position property
            double BLX = textObjIn.dbText.Position.X;
            double BLY = textObjIn.dbText.Position.Y;

            // Write bottom left coordinates to XML file 
            xmlWriterIn.WriteAttributeString("x1", BLX.truncstring(3));
            xmlWriterIn.WriteAttributeString("y1", BLY.truncstring(3));

            // Get top right coordinates using geometric extents or alignment point + vector or approximation
            double TRX = textObjIn.TopRightL.X;
            double TRY = textObjIn.TopRightL.Y;

            // Write top right extents to XML file
            xmlWriterIn.WriteAttributeString("x2", TRX.truncstring(3));
            xmlWriterIn.WriteAttributeString("y2", TRY.truncstring(3));

            if (textObjIn.topRightApproximated)
            {
                xmlWriterIn.WriteAttributeString("comment", "x2 and y2 approximated");
            }

            // If extents seem invalid, write comment to XML file
            if ((TRX == BLX) || (TRY == BLY))
            {
                xmlWriterIn.WriteAttributeString("comment2", "Invalid x2/y2 extents because text has no max extents or has a non-centered alignment point");
                inLogger.Log("DWG: " + Path.GetFileName(textObjIn.dbText.Database.Filename) + " has text object with non-centered alignment point and invalid max extents: " + "\"" + textObjIn.text + "\"" + " on layer: " + textObjIn.dbText.Layer);
            }
        }

        public static void writeOutAttRefMinMax(this XmlWriter xmlWriterIn, AttributeReference attRefIn)
        {
            if (attRefIn.Id.IsNull) { return; }

            xmlWriterIn.WriteAttributeString("x1", attRefIn.Position.X.truncstring(3));
            xmlWriterIn.WriteAttributeString("y1", attRefIn.Position.Y.truncstring(3));

            Point2d maxPoint;

            try
            {
                maxPoint = new Point2d(attRefIn.GeometricExtents.MaxPoint.X, attRefIn.GeometricExtents.MaxPoint.Y);
                xmlWriterIn.WriteAttributeString("x2", maxPoint.X.truncstring(3));
                xmlWriterIn.WriteAttributeString("y2", maxPoint.Y.truncstring(3));
            }
            catch
            {
                Vector3d vector = attRefIn.Position.GetVectorTo(attRefIn.AlignmentPoint);

                maxPoint = new Point2d(((attRefIn.AlignmentPoint.X + vector.X > 0) ? attRefIn.AlignmentPoint.X + vector.X : attRefIn.Position.X),
                                                                        ((attRefIn.AlignmentPoint.Y + vector.Y > 0) ? attRefIn.AlignmentPoint.Y + vector.Y : attRefIn.Position.Y));
                xmlWriterIn.WriteAttributeString("x2", maxPoint.X.truncstring(3));
                xmlWriterIn.WriteAttributeString("y2", maxPoint.Y.truncstring(3));
            }
        }

        // Returns the MS# for the given dwg
        public static String getDwgMsNum(Database db, Transaction tr)
        {
            Regex msNo = new Regex(@"^(\s)*MS\d{6}(\S)*(\s)*$");
            Regex letterNo = new Regex(@"^(\s)*[A-Z]\d{6}(\S)*(\s)*$");
            String msNum = String.Empty;

            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            foreach (ObjectId btrId in blockTable)
            {
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                if (btr.IsLayout)
                {
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                        DBText dbt = tr.GetObject(id, OpenMode.ForRead) as DBText;

                        if (dbt == null) { continue; }

                        if (msNo.IsMatch(dbt.TextString.ToUpper()) && ent.Layer == "0")
                        {
                            return msNum = dbt.TextString.Trim();
                        }
                        if (letterNo.IsMatch(dbt.TextString.ToUpper()) && ent.Layer == "0")
                        {
                            return msNum = dbt.TextString.Trim();
                        }
                    }
                }
            }
            throw new System.Exception("MS # not found");
        }

        // Get MS number for specified layer
        public static String msText(Database db, string layerName)
        {
            Regex msNo = new Regex(@"^\d{6}");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    if (btr.IsLayout)
                    {
                        foreach (ObjectId id in btr)
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent.Layer.Equals(layerName, System.StringComparison.CurrentCultureIgnoreCase))
                            {
                                DBText dbt = tr.GetObject(id, OpenMode.ForRead) as DBText;

                                if (dbt == null) { dbt.Dispose(); continue; }
                                else
                                {
                                    if (dbt.TextString.Contains("MS") || msNo.IsMatch(dbt.TextString))
                                    {
                                        //System.Windows.Forms.MessageBox.Show("MS: " + dbt.TextString);
                                        return " (" + dbt.TextString + ") ";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return "";
        }

        // Deletes given layer
        internal static System.Boolean delLayer(Database db, String layerName, LayerTableRecord layer)
        {
            ObjectIdCollection idCollection = new ObjectIdCollection();

            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (ObjectId btrId in blockTable)
                {
                    BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForWrite) as BlockTableRecord;
                    if (btr.IsLayout)
                    {
                        foreach (ObjectId id in btr)
                        {
                            Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                            if (ent.Layer.Equals(layerName, System.StringComparison.CurrentCultureIgnoreCase))
                            {
                                // Delete entity
                                idCollection.Add(ent.Id);
                                ent.Erase();
                            }
                        }
                    }

                    db.Purge(idCollection);
                    // Delete layer
                    layer.Erase();
                }
                return true;
            }
        }

        // Returns true if file is valid, accessible, and has correct file extension
        public static Boolean isFilePathOK(this String inFilePath, String extension)
        {
            return (inFilePath.isFilePathOK() &&
                    String.Equals(System.IO.Path.GetExtension(inFilePath).ToUpper().Trim(),
                                  extension.ToUpper().Trim())
                   );
        }

        // Returns true if file is valid and accessible
        public static Boolean isFilePathOK(this String inFilePath)
        {
            FileInfo fi = null;

            if (String.IsNullOrWhiteSpace(inFilePath)) { return false; }

            try { fi = new FileInfo(inFilePath); }
            catch { return false; }

            if (!File.Exists(inFilePath) || fi == null) { return false; }

            return true;
        }

    }
}
