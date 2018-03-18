using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace AcadExts
{
    // Class for custom attribute that marks a method as a custom Raytheon AutoCAD .NET command
    // Used by GetInfoString() to display all custom commands and an optional description at runtime with reflection
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    internal class rtnAttribute : Attribute
    {
        public String description { get; private set; }

        public rtnAttribute(String inDescription)
        {
            this.description = inDescription;
        }
        // Chain Ctor
        public rtnAttribute() : this("") { }
    }

    //https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/concepts/attributes/accessing-attributes-by-using-reflection
}
