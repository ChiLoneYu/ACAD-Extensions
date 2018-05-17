using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xaml;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System.Windows.Threading;
using System.Runtime.Versioning;
using System.Reflection;
using System.IO;

#region Solution Info
/*Solution Info********************************************************************************/
// Build with EntryPoint set as startup project                                              

// Options -> Environment -> Documents -> enable 'Detect when file is changed outside the environment' and 'Auto-load cahnges, if saved'

// Options -> Debugging -> disable 'just my code'                                                        

// Options -> Debugging -> enable managed compatibility mode to run in debug mode on acad 2017           

// Set 'Copy Local' property of 3 acad references to false                                   

// Contains APIs related to the AutoCAD application login (such as selections set, commands and keywords).
// C:\Program Files\Autodesk\AutoCAD 2017\accoremgd.dll

// Contains the APIs for controlling the AutoCAD application itself – defining custom commands opening and closing documents, plotting, etc. and contains ‘User Interface’ related APIs (such as dialogs)
// C:\Program Files\Autodesk\AutoCAD 2017\acmgd.dll

// Contains the APIs for creating, editing or querying the contents of a DWG file
// C:\Program Files\Autodesk\AutoCAD 2017\acdbmgd.dll

// Reference these DLLs for AutoCad .NET COM Interop
// C:\Program Files\Autodesk\AutoCAD 2017\Autodesk.AutoCAD.Interop.dll
// C:\Program Files\Autodesk\AutoCAD 2017\Autodesk.AutoCAD.Interop.Common.dll

// Don't add this reference (Interop)
// C:\Program Files\Common Files\Autodesk Shared\acax19enu.tlb

// Add references to: AcadExts       
//                    System.Drawing                                                                 
//                    System.Xaml   
//                    System.Xml.Linq                                                                                                        
// AutoCAD:
//                    accoremgd                                                              
//                    acdbmgd                                                                
//                    acmgd         
// WPF Core:                                               
//                    PresentationCore      
//                    PresentationFramework                                                     
//                    WindowsBase 
                                             
// In AcadExts Add references to:                                                            
//                    Microsoft.Office.Interop.Excel 14         
                             
// Todo: 
    // Remove unnecessary property changed events for buttons                                                                                   
    // Use async. file operations instead of sync.

// Acdbmgd assembly calls into EntryPoint assembly which calls into AcadExts assembly
                                                                      
// Can help improve load performance
// These 2 assembly level attributes help AutoCAD load the .NET module faster by telling
// it  which object is the Command class and which one is the Extension Application class, so
// AutoCAD doesn't need to search each object.
/*********************************************************************************************/
#endregion

[assembly: ExtensionApplication(typeof(EntryPoint.Init))]
[assembly: CommandClass(typeof(EntryPoint.EntryCmdClass))]
namespace EntryPoint
{
    //Registers this assembly as an AutoCAD extension
    public class Init : Autodesk.AutoCAD.Runtime.IExtensionApplication
    {
        // Runs on AutoCAD open when this assembly is 'netload'ed into AutoCAD from .lsp file
        // Display information to AutoCAD command window
        public void Initialize()
        {          
            StringBuilder message = new StringBuilder();

            message.Append(String.Concat(Environment.NewLine,
                                         "-----------------------------------------",
                                         Environment.NewLine,
                                         "Run 'AcadExtsMenu' command to open AutoCAD Extensions dialog",
                                         Environment.NewLine,
                                         "-----------------------------------------",
                                         Environment.NewLine));

            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(message.ToString());
        }

        // By the time Terminate() is called,  AutoCAD is already closing and unloading .NET modules/assemblies,
        // so by this time there is no editor to write to
        public void Terminate()
        {
        }
    }

    public static class EntryCmdClass
    {
        // This command opens WPF app from ACAD cmd line, must be public to expose cmd to AutoCAD
        [CommandMethod("AcadExtsMenu", CommandFlags.Session)]
        public static void AcadExts()
        {
            if (System.Threading.Thread.CurrentThread.Name == null) { System.Threading.Thread.CurrentThread.Name = "MainThreadInEntryPointAssembly"; }

            AcadExts.MainWindow mainWin = null;

            if (Application.DocumentManager.MdiActiveDocument == null) { return; }

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            
            ed.WriteMessage(Environment.NewLine + "Opening Acad Exts Menu..." + Environment.NewLine);
            
            try
            {
                // https://www.theswamp.org/index.php?topic=50815.0
                mainWin = new AcadExts.MainWindow();
                //mainWin.Topmost = true;
                mainWin.ShowActivated = true;
                //mainWin.Closed += (sender, e) => { global::System.Windows.Forms.MessageBox.Show("Exiting"); };
                //mainWin.Closed += mainWin_Closed;
                mainWin.Closing += mainWin_Closing;

                Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(mainWin);
            }
            catch(System.Exception se)
            {
                ed.WriteMessage(String.Concat(Environment.NewLine, 
                                              "Exception occured: ",
                                              se.Message,
                                              Environment.NewLine,
                                              Environment.NewLine,
                                              ((se.InnerException != null) ? ("Inner Exception: " + se.InnerException.Message) : (Environment.NewLine)),
                                              Environment.NewLine));
            }
            finally
            {
                // On return of WPF ShowModalWindow call
                try { mainWin.Close(); }
                catch { }

                try { ed.WriteMessage(Environment.NewLine + "Closing Acad Exts Menu..." + Environment.NewLine); }
                catch { }
            }
        }

        static void mainWin_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        }
    }
}
