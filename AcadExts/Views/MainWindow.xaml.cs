using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Interactivity;
using Microsoft.Expression.Interactivity;

namespace AcadExts
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();

            // Because setting the context in the view isnt working because of VS bug
            //base.DataContext = new Presenter();
        }

        // Added this code to code-behind because there is no other way to access
        // the instance of the viewmodel to cancel the current backgroundworker
        // when the window's exit button is clicked.
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);

            AcadExts.Presenter vmI = (AcadExts.Presenter) this.DataContext;
            
            if (vmI != null) { vmI.StopWorker(); }
        }
    }
}
