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
    public interface IDataAccessor : IDisposable
    {
        Task SetupAsync();

        Task InsertAsync(Entity ent);

        Task InsertAsync(String type, String oid, String layer, Double? x = null, Double? y = null);

        void SelectAsync();

    }
}
