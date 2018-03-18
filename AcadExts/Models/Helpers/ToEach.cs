using System;
using System.Text;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;

namespace AcadExts
{
    /// <summary>
    /// This class contains methods that take in callback delegates that perform actions on different objects.
    /// These exist so actions can be performed on these types without having to re-write iterations through
    ///  layer tables and block tables to access each entity.
    /// </summary>

    public static class ToEach
    {
        public delegate void delLAYTABLERECORD(LayerTableRecord inLAYTABLERECORD);

        public static void toEachLayerRecord(Database inDB, Transaction inTR, OpenMode inTYPE, delLAYTABLERECORD inMETHOD)
        {
            using (LayerTable lt = inTR.GetObject(inDB.LayerTableId, inTYPE) as LayerTable)
            {
                foreach (ObjectId layerId in lt)
                {
                    // Visit each layer
                    using (LayerTableRecord layer = inTR.GetObject(layerId, inTYPE) as LayerTableRecord)
                    {
                        try { inMETHOD(layer); }
                        catch { throw; }
                    }
                }
            }
        }

        public static List<Entity> toEachEntity(Database inDb, Transaction inTR, Func<Entity, System.Boolean> inMethod)
        {
            List<Entity> affectedEnts = new List<Entity>();

            using (BlockTable bt = inTR.GetObject(inDb.BlockTableId, OpenMode.ForWrite) as BlockTable)
            {
                // foreach (ObjectId oid in bt)
                // {
                using (BlockTableRecord btr = inTR.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                {
                    foreach (ObjectId oid2 in btr)
                    {
                        Entity ent = inTR.GetObject(oid2, OpenMode.ForWrite) as Entity;
                        if (inMethod(ent))
                        {
                            affectedEnts.Add(ent);
                        }
                    }
                }
                //  }
            }
            return affectedEnts;
        }

        public static List<LayerTableRecord> toEachLayer(Database inDb, Transaction inTR, Func<LayerTableRecord, System.Boolean> inMethod)
        {
            List<LayerTableRecord> affectedLayers = new List<LayerTableRecord>();

            using (LayerTable lt = inTR.GetObject(inDb.LayerTableId, OpenMode.ForRead) as LayerTable)
            {
                foreach (ObjectId oid in lt)
                {
                    using (LayerTableRecord layer = inTR.GetObject(oid, OpenMode.ForWrite) as LayerTableRecord)
                    {
                        if (inMethod(layer))
                        {
                            affectedLayers.Add(layer);
                        }
                    }
                }
            }
            return affectedLayers;
        }

        public static List<DBText> toEachDBText(Database inDb, Transaction inTR, Func<DBText, System.Boolean> inMethod)
        {
            List<DBText> affectedDBTs = new List<DBText>();

            using (BlockTable bt = inTR.GetObject(inDb.BlockTableId, OpenMode.ForWrite) as BlockTable)
            {
                using (BlockTableRecord btr = inTR.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                {
                    foreach (ObjectId oid2 in btr)
                    {
                        Entity ent = inTR.GetObject(oid2, OpenMode.ForWrite) as Entity;
                        if (ent.GetType() == typeof(DBText))
                        {
                            DBText dbt = ent as DBText;
                            if (inMethod(dbt))
                            {
                                affectedDBTs.Add(dbt);
                            }
                        }
                    }
                }
            }
            return affectedDBTs;
        }

        //Generic ToEach for DBObjects and derived types using reflection
        // This method only iterates through Entity objects in model space

        // Example Uses:
        //To get all line objs:
        //      IList<Line> textObjLines = ToEach.toEach<Line>(db, acTrans, (s) => true);
        //To get all dwg text objs that dont contain numberic chars:
        //      IList<DBText> textObjs = ToEach.toEach<DBText>(db, acTrans, (s) => { return !s.TextString.HasNums(); });

        public static IList<T> toEach<T>(Database inDb, Transaction inTR, Func<T, System.Boolean> inMethod) where T : DBObject//, new()
        {
            List<T> affectedObjs = new List<T>();

            Type type = typeof(T);

            // Objects whose classes are derived from Entity (which have graphical representation)
            if (type.IsSubclassOf(typeof(Entity)) && typeof(Entity).IsAssignableFrom(type))
            {
                using (BlockTable bt = inTR.GetObject(inDb.BlockTableId, OpenMode.ForRead) as BlockTable)
                {
                    using (BlockTableRecord btr = inTR.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                    {
                        if (typeof(T) == typeof(DBText))
                        {
                            foreach (ObjectId oid2 in btr)
                            {
                                if (oid2.ObjectClass.DxfName == "TEXT")
                                {
                                    DBText dbt = inTR.GetObject(oid2, OpenMode.ForWrite) as DBText;
                                    if (inMethod(dbt as T)) { affectedObjs.Add(dbt as T); }
                                }
                            }
                        }
                        if (typeof(T) == typeof(Line))
                        {
                            foreach (ObjectId oid2 in btr)
                            {
                                if (oid2.ObjectClass.DxfName == "LINE")
                                {
                                    Line line = inTR.GetObject(oid2, OpenMode.ForWrite) as Line;
                                    if (inMethod(line as T)) { affectedObjs.Add(line as T); }
                                }
                            }
                        }
                        //Keep adding here as needed.....
                    }
                }
            }

            // For LayerTableRecords
            if (typeof(T).BaseType == typeof(SymbolTableRecord) && typeof(T) == typeof(LayerTableRecord))
            {
                using (LayerTable lt = inTR.GetObject(inDb.LayerTableId, OpenMode.ForRead) as LayerTable)
                {
                    foreach (ObjectId oid in lt)
                    {
                        using (LayerTableRecord layer = inTR.GetObject(oid, OpenMode.ForRead) as LayerTableRecord)
                        {
                            if (inMethod(layer as T))
                            {
                                affectedObjs.Add(layer as T);
                            }
                        }
                    }
                }
            }
            return affectedObjs;
        }
    }
}