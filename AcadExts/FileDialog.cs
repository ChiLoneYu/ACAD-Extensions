using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;
using WinForms = System.Windows.Forms;

namespace AcadExts
{
    public class FileDialog : Button
    {
        public static readonly DependencyProperty FilePathProperty = DependencyProperty.Register("FilePath",
                                                                                                 typeof(String),
                                                                                                 typeof(FileDialog));

        public String FilePath
        {
            get { return (String)GetValue(FilePathProperty); }
            set { SetValue(FilePathProperty, value); }
        }

        protected override void OnClick()
        {
            base.OnClick();
            ShowDialog();
        }

        public void ShowDialog()
        {
            WinForms.OpenFileDialog FileBrowser = new WinForms.OpenFileDialog();

            FileBrowser.Title = "Select a File";

            String currentPath = GetValue(FilePathProperty) as String;
            
            if (currentPath.isFilePathOK())
            {
                try
                {
                    FileBrowser.InitialDirectory = System.IO.Directory.GetParent(currentPath).FullName;
                }
                catch { }
            }

            if (FileBrowser.ShowDialog() == WinForms.DialogResult.OK)
            {
                FilePath = FileBrowser.FileName;
            }
        }
    }
}