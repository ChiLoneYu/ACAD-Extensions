using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace AcadExts
{
    [rtn("Lists info about all text objects in all blocks in the currently opened DWG")]
    internal sealed class ObjectLister
    {
        //private readonly String path = String.Empty;
        public String _Path { get; set; }
        public ObjectLister()
        {
            //this.path = Application.DocumentManager.MdiActiveDocument.Database.Filename.Trim();
        }

        // List info about all text objects in current (open) dwg and print info to XAML text box
        public List<String> Process()
        {
            List<String> TextList = new List<String>();

            try
            {
                // Putting this in a using statement messes up the dwg instance for the ed.writemesssage() in finally {} of entrycmdclass.cs 
                Document currentDoc = Application.DocumentManager.MdiActiveDocument;
                _Path = currentDoc.Database.Filename;
                //Database oldDb = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;

                using (Database db = currentDoc.Database)
                {
                    using (Transaction acTrans = db.TransactionManager.StartTransaction())
                    {
                        // Dont use polymorphic ToEach way because here we want text objects from ALL blocks, not just model space
                        //IList<DBText> textObjs = ToEach.toEach<DBText>(db, acTrans, (s) => true);

                        //foreach (DBText dbt in textObjs)
                        //{
                        BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = acTrans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;

                            foreach (ObjectId entId in btr)
                            {
                                if (entId.ObjectClass.DxfName == "TEXT")
                                {
                                    using (DBText dbt = acTrans.GetObject(entId, OpenMode.ForRead) as DBText)
                                    {
                                        StringBuilder currentString = new StringBuilder();
                                        // Print the block, layer, text value, and truncated x,y values
                                        currentString.Append(dbt.BlockName)
                                                      .Append("   :   ")
                                                      .Append(dbt.Layer)
                                                      .Append("   :   ")
                                                      .Append(String.Concat("(", dbt.Position.X.truncstring(2),", ", dbt.Position.Y.truncstring(2),")"))
                                                      .Append("   :   ")
                                                      .Append(dbt.TextString);

                                        TextList.Add(currentString.ToString());
                                    }
                                }
                            }
                            if (!btr.IsDisposed) { btr.Dispose(); }
                        }
                        if (!bt.IsDisposed) { bt.Dispose(); }
                        //}
                        acTrans.Commit();
                    }
                }
                return (TextList.Count == 0) ? (new List<String>() { "No text found." }) : TextList.OrderBy<String, String>((s) => s).ToList();
            }
            catch
            {
                throw; //   return String.Concat("Error: ", se.Message);
            }
        }
    }
}
