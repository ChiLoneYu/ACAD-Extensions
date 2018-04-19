using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Windows.Interactivity;
using Microsoft.Expression.Interactivity;
using WinForms = System.Windows.Forms;
using Button = System.Windows.Controls.Button;

namespace AcadExts
{
    #region FYI
    // https://stackoverflow.com/questions/33364827/wpf-mvvm-creating-a-dialog-using-a-behavior
    // A behavior encapsulates pieces of functionality into a reusable component,
    // which we later on can attach to an element in a view. Emphasis is on reusable.
    // One can do the same code in codebehind or perhaps directly in XAML so it is nothing magic about a behavior.
    // Behaviors also have the benefit of keeping the MVVM pattern intact,
    // since we can move code from codebehind to behaviors. 
    // One example is if we want to scroll in selected item in a ListBox and the selected item is chosen from code,
    // e g from a search function.
    // The ViewModel doesn't know that the view use a ListBox to show the list so it can not be used
    // to scroll in the selected item. And we don´t want to put code in the codebehind,
    // but if we use a behavior we solve this problem and creates a reusable component which can be used again.
    #endregion
    //Attached Property / Blend Behavior
    //https://blog.jayway.com/2013/03/20/behaviors-in-wpf-introduction/
    //https://stackoverflow.com/questions/4007882/select-folder-dialog-wpf/17712949#17712949
    //https://docs.microsoft.com/en-us/dotnet/framework/wpf/advanced/attached-properties-overview
    //https://documentation.devexpress.com/WPF/17458/MVVM-Framework/Behaviors/How-to-Create-a-Custom-Behavior
    // TODO:

    public class FolderDialogBehavior : Behavior<Button>
    {
        public static readonly DependencyProperty FolderPathProperty = DependencyProperty.RegisterAttached("FolderPath",
                                                                                           typeof(String),
                                                                                           typeof(FolderDialogBehavior));

        protected override void OnAttached()
        {
            base.OnAttached();
            AssociatedObject.Click += AssociatedObject_Click;
        }

        protected override void OnDetaching()
        {
            AssociatedObject.Click -= AssociatedObject_Click;
            base.OnDetaching();
        }

        void AssociatedObject_Click(object sender, EventArgs e)
        {
            WinForms.FolderBrowserDialog fbd = new WinForms.FolderBrowserDialog();
            
            fbd.Description = "Select a Folder";

            String currentPath = GetValue(FolderPathProperty) as String;

            if (currentPath.isDirectoryPathOK()) 
            {
                fbd.SelectedPath = currentPath; 
            }
                       
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                SetValue(FolderPathProperty, fbd.SelectedPath);
            }
        }

        public static String GetFolderPath(DependencyObject objIn)
        {
            return (String)objIn.GetValue(FolderPathProperty);
        }

        public static void SetFolderPath(DependencyObject objIn, String valueIn)
        {
            objIn.SetValue(FolderPathProperty, valueIn);
        }
    }
}
