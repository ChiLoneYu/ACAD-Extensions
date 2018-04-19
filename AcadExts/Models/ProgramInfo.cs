using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.Versioning;
using Autodesk.AutoCAD;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace AcadExts
{
    internal static class ProgramInfo
    {
        const long toKil = 1024;
        const long toMeg = 1048576;
        const long toGig = 1073741824;

        readonly static Regex TraceRegex = new Regex(@"\sat");

        readonly static String[] AsmNamesNeeded = new String[] { "mscorlib", "Acdbmgd", "System", "accoremgd", "System.Core", 
                                                         "System.Windows.Interactivity", "PresentationFramework", "System.Xml.Linq",
                                                         "System.Xml", "PresentationCore", "System.Windows.Forms",
                                                         "WindowsBase", "System.Xaml", "Microsoft.CSharp", "System.Drawing"};

        // Get methods with custom rtn attributes and return the attribute descriptions
        public static IList<String> GetCmdInfo()
        {
            List<String> InfoList = new List<String>();

            IEnumerable<Type> classes;

            // Get all Classes. Works because we only want classes in this AppDomain
            try
            {
                classes = AppDomain.CurrentDomain.GetAssemblies().SelectMany(t => t.GetTypes())
                                                                 .Where(t => t.IsClass && t.Namespace == "AcadExts");
            }
            catch { InfoList.Add("Error retrieving command info"); return InfoList; }

            InfoList.Add("Commands: ");

            foreach (Type ModelClass in classes)
            {
                try
                {
                    // Use Reflection to get custom attributes for this class
                    System.Attribute[] attributes = System.Attribute.GetCustomAttributes(ModelClass);

                    // Get rtnAttributes and take first one
                    System.Attribute customRtnAttribute = attributes.Where(s => s.GetType() == typeof(rtnAttribute)).FirstOrDefault();

                    rtnAttribute att = customRtnAttribute as rtnAttribute;

                    // Add class name to string, and description if there is one

                    InfoList.Add(String.Concat(ModelClass.Name, (!String.IsNullOrWhiteSpace(att.description)) ? (" : " + att.description) : ("")));
                }
                catch { continue; /* Class isn't a valid custom rtn command, keep going */ }
            }
            InfoList.Add("");
            return InfoList;
        }

        // Get AcadExts assembly info, info about currently loaded assemblies, and AutoCAD version info
        public static IList<String> GetAssemblyInfo()
        {
            List<String> InfoList = new List<String>();

            InfoList.Add("Assembly Info: ");

            #if DEBUG
                InfoList.Add("Acad Exts is running in DEBUG mode");
            #else
                InfoList.Add("Acad Exts is running in RELEASE mode");
            #endif

            try
            {
                // Get Assembly info

                System.IO.FileInfo asmInfo = new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);

                InfoList.Add(String.Concat("Using Acad Exts from: ", asmInfo.LastWriteTime.ToString("M/d/yyyy h:mm tt")));

                InfoList.Add(String.Concat("Using Acad Exts in: ", asmInfo.Directory.FullName, "\\"));

                // Check loaded Assemblies

                IEnumerable<String> AsmNamesReferenced = System.Reflection.Assembly.GetExecutingAssembly().GetReferencedAssemblies().Select(asmName => asmName.Name);

                foreach (String missingAsm in AsmNamesNeeded.Where<String>(delegate(String neededAsm)
                                                                            {
                                                                                return (!AsmNamesReferenced.Contains(neededAsm));
                                                                            }))
                {
                    InfoList.Add(String.Concat("Error loading Assembly: ", missingAsm));
                }

                foreach (String unrecognizedAsm in AsmNamesReferenced.Where<String>(delegate(String referencedAsm)
                                                                                    {
                                                                                        return (!AsmNamesNeeded.Contains(referencedAsm));
                                                                                    }))
                {
                    InfoList.Add(String.Concat("Unknown Assembly loaded: ", unrecognizedAsm));
                }
            }
            catch { }
            //AutoCAD 2016, 2017, 2018 use DWG 2016 format ? 
            //AutoCAD 2013, 2014, 2015 use DWG 2013 format
            //AutoCAD 2010, 2011, 2012 use DWG 2010 format
            //AutoCAD 2007, 2008, 2009 use DWG 2007 format
            //AutoCAD 2004, 2005, 2006 use DWG 2004 format
            //AutoCAD 2000, 2000i, 2002 use DWG 2000 format

            //AC1015 = AutoCAD 2000 
            //AC1018 = AutoCAD 2004 
            //AC1021 = AutoCAD 2007 
            //AC1024 = AutoCAD 2010 
            //AC1027 = AutoCAD 2013 
            //AC1032 = AutoCAD 2018 
            // Get AutoCAD version
            //InfoList.Add(String.Concat("AutoCAD default file save format: ", Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.DefaultFormatForSave.ToString()));
            //Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument;
            //String origFileVersion = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.OriginalFileVersion.ToString();
            InfoList.Add(String.Concat("AutoCAD version: ", Autodesk.AutoCAD.ApplicationServices.Application.Version.ToString()));

            // Get .NET version
            try
            {
                object[] ObjAttributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(TargetFrameworkAttribute), false);

                TargetFrameworkAttribute[] TFAttributes = Array.ConvertAll<object, TargetFrameworkAttribute>(ObjAttributes, delegate(object obj) { return obj as TargetFrameworkAttribute; });

                InfoList.Add(String.Concat(".NET Version: ", TFAttributes.FirstOrDefault().FrameworkDisplayName));
            }
            catch { }

            // App Domain
            //message.Append(String.Concat("App Domain: ", System.AppDomain.CurrentDomain.FriendlyName, Environment.NewLine));
            return InfoList;
        }

        //public static String GetDebugInfo()
        public static IList<String> GetDebugInfo()
        {
            List<String> detailsList = new List<String>();

            try
            {
                detailsList.Add("");
                detailsList.Add(String.Concat("Thread State: ", System.Threading.Thread.CurrentThread.ThreadState));
                detailsList.Add(String.Concat("Thread Name: ", System.Threading.Thread.CurrentThread.Name));
                detailsList.Add(String.Concat("Is Thread Alive? ", System.Threading.Thread.CurrentThread.IsAlive));
                detailsList.Add(String.Concat("Does thread belong to the managed thread pool? ", System.Threading.Thread.CurrentThread.IsThreadPoolThread.ToString()));
                detailsList.Add(String.Concat("Is Process Responding? ", Process.GetCurrentProcess().Responding.ToString()));
                detailsList.Add(String.Concat("Has Process Exited? ", Process.GetCurrentProcess().HasExited));
                detailsList.Add(String.Concat("Nonpaged Memory Size: ", Process.GetCurrentProcess().NonpagedSystemMemorySize64 / toMeg, " MB"));
                detailsList.Add(String.Concat("Working Set: ", Process.GetCurrentProcess().WorkingSet64 / toMeg, " MB"));
                detailsList.Add(String.Concat("Env. UIM? ", System.Environment.UserInteractive));
                detailsList.Add(String.Concat(Utilities.nl, "Stack Trace: "));

                // Format stack trace
                List<String> stackTraceLines = TraceRegex.Split(System.Environment.StackTrace).ToList<String>();
                stackTraceLines.RemoveAt(0);
                stackTraceLines.ForEach(str => detailsList.Add(str.Replace(Utilities.nl, "")));

                ////detailsString.Append(String.Concat(Utilities.nl, "Virtual Memory: ", Process.GetCurrentProcess().VirtualMemorySize64 / toMeg, " MB"));
                ////detailsString.Append(String.Concat(Utilities.nl, "Private Memory: ", Process.GetCurrentProcess().PrivateMemorySize64 / toMeg, " MB"));
                ////detailsString.Append(String.Concat(Utilities.nl, "Paged Memory Size: ", Process.GetCurrentProcess().PagedMemorySize64 / toMeg, " MB"));
                //detailsString.Append(String.Concat(Utilities.nl, "Process Exit Time: ", Process.GetCurrentProcess().ExitTime, " MB"));

            }
            catch (System.Exception se)
            {
                detailsList.Add(String.Concat("Error: ", se.Message));
            }
            return detailsList;
        }
    }
}
