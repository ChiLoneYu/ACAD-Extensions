using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Configuration;
using System.Reflection;
using System.IO;
using System.Data.SQLite;
using Autodesk.AutoCAD.DatabaseServices;


namespace DA
{
    internal static class DBUtilities
    {
        public static String ToStr(this Double? input)
        {
            if (input.HasValue) { return ((double)input).ToString(); }//input.Value.ToString(); }
            else { return "NULL"; }
        }
    }

    // Provides asynchronous database access
    public class DataAccessor : IDataAccessor
    {
        private SQLiteConnection conn { get; set; }

        public async Task SetupAsync()
        {
            try
            {
                 String currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                // TODO: Use config file
                //AppDomain.CurrentDomain.SetData("DataDirectory", currentPath);
                //ConfigurationManager.RefreshSection("connectionStrings");

                //ConnectionStringSettingsCollection conStrings = ConfigurationManager.ConnectionStrings;
                //String conString = conStrings[0].ToString();
                //var appSettings = ConfigurationManager.AppSettings;


                if (!System.IO.File.Exists(currentPath + "\\db.sqlite"))
                {
                    SQLiteConnection.CreateFile(currentPath + "\\db.sqlite");
                }

                conn = new SQLiteConnection("Data Source=" + currentPath + "\\db.sqlite;Version=3");

                conn.Open();
                string drop = @"DROP TABLE IF EXISTS DWGOBJECTS;";
                string create = drop + @"CREATE TABLE IF NOT EXISTS DWGOBJECTS (type VARCHAR(60) NOT NULL,
                                                                                  objectid VARCHAR(60) PRIMARY KEY NOT NULL,
                                                                                  layer VARCHAR(60) NOT NULL,
                                                                                  x DOUBLE,
                                                                                  y DOUBLE  )";

                await new SQLiteCommand(create, conn).ExecuteNonQueryAsync();
            }
            catch 
            {
                conn.Close();
                throw;
            }
            finally 
            {

            }
        }

        public async Task InsertAsync(Entity entity)
        {
            await InsertAsync(entity.GetType().ToString(), entity.ObjectId.ToString(), entity.Layer, null, null);
            return;
        }

        public async Task InsertAsync(String type, String oid, String layer, Double? x = null, Double? y = null)
        {
            if ((type.Length > 60) || (oid.Length > 60) || (layer.Length > 60))
            {
                throw new ArgumentException("Argument more than 60 characters long");
            }

            if (conn == null) { throw new System.Data.SQLite.SQLiteException("Attempt to Insert with bad DB connection"); }

            string insert = "INSERT INTO DwgObjects (type, objectid, layer, x, y) VALUES ('" + type + "','" + oid + "','" + layer + "'," + x.ToStr() + "," + y.ToStr() +")";
            
            await new SQLiteCommand(insert, conn).ExecuteNonQueryAsync();

            return;
        }

        public void SelectAsync()
        {
            //SQLiteCommand select = new SQLiteCommand("SELECT * from DwgObjects", conn);

            //SQLiteDataReader reader = select.ExecuteReader();

            //List<DwgObj> dwgobjs = new List<DwgObj>();
            //while(reader.Read())
            //{
            //    dwgobjs.Add(new DwgObj() {  });
            //}

            //List<String> strs = new List<string>();
            //while (reader.Read())
            //{
            //    string k = reader["objectid"].ToString();
            //    strs.Add(k);
            //}

    //            public class DwgObj
    //{
    //    String type;
    //    String objectid;
    //    string layer;
    //    Double? x;
    //    Double? y;
    //}

        }

        public void Dispose()
        {
            conn.Close();
        }
    }


}
