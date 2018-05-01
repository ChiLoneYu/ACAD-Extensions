using System;
using System.ComponentModel;
using System.Windows.Input;
using System.Reflection;
using System.Windows.Threading;
using System.Threading;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Configuration;

namespace AcadExts
{
    // View Model class 
    public class Presenter : ObservableObject
    {
        // Save instances so ICommand property doesn't need to return new DelegateCommands every time property get is called

        //DelegateCommand[] DelegateCommands = new DelegateCommand[] {  };

        DelegateCommand _LayerListerCommand = null;
        DelegateCommand _GenerateKeyfileCommand = null;
        DelegateCommand _UpdateFBDCommand = null;
        DelegateCommand _InsXRefCommand = null;
        DelegateCommand _FormatForDeliveryCommand = null;
        DelegateCommand _CheckLayersCommand = null;
        DelegateCommand _LowercaseLayersCommand = null;
        DelegateCommand _ListObjectsCommand = null;
        DelegateCommand _ExtractCommand = null;
        DelegateCommand _DebugDetailsCommand = null;
        DelegateCommand _ConverterTo2000Command = null;
        DelegateCommand _DefaultRectangleCheckedCommand = null;
        DelegateCommand _CopyLayerCommand = null;

        private Boolean DebugInfoDisplayed = false;

        // Use one background worker instance so the code behind knows which bw is current 
        // so it can close it when window exit button is pressed.
        public BackgroundWorker worker = null;

        // Initializes worker with new BackgroundWorker instance
        private BackgroundWorker GetWorker()
        {
            // Add any worker initialization steps here
            return new BackgroundWorker() { WorkerSupportsCancellation = true, WorkerReportsProgress = true };
        }

        // Get Dispatcher for this thread (main thread)
        private Dispatcher presenterDispatcher = Dispatcher.CurrentDispatcher;

        // Takes in an Action and runs it asynchronously on main thread
        public void CallMethodOnMThread(Action method)
        {
            //presenterDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action (  () => _LayerListerCommand.RaiseCanExecuteChanged()   ));
            //presenterDispatcher.BeginInvoke(DispatcherPriority.Normal, ((Action) delegate { _LayerListerCommand.RaiseCanExecuteChanged();}) );
            //Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Normal, method);
            presenterDispatcher.BeginInvoke(DispatcherPriority.Normal, method);
        }

        public Presenter()
        {
            if (System.Threading.Thread.CurrentThread.Name == null) { System.Threading.Thread.CurrentThread.Name = "VM Thread"; }

            // In Info tab, show command descriptions based on custom class attributes and runtime assembly info using reflection
            CmdInfoList = ProgramInfo.GetCmdInfo().Concat(ProgramInfo.GetAssemblyInfo()).ToList<String>();
        }

        // Used by commands to make sure the backgroundworker isn't already running for another command, so
        // two commands never run at the same time
        public Boolean canWork()
        {
            if (worker == null) { return true; }
            else { return !worker.IsBusy; }
        }

        // Cancels current background worker
        public void StopWorker()
        {
            if ((worker != null) && (worker.WorkerSupportsCancellation) && worker.IsBusy)
            {
                try { worker.CancelAsync(); worker.Dispose(); }
                catch { }
            }
            return;
        }

        /********************/
        /* Command bindings */
        /********************/

        public ICommand FormatForDeliveryCommand
        {
            get { return _FormatForDeliveryCommand ??
                         (_FormatForDeliveryCommand = new DelegateCommand(this.FormatForDelivery, (s) => canWork())); }
        }

        public ICommand DebugDetails
        {
            get { return _DebugDetailsCommand ??
                         (_DebugDetailsCommand = new DelegateCommand(this.ToggleDebugInfo)); }
        }

        public ICommand GenerateKeyfileCommand
        {
            get { return _GenerateKeyfileCommand ??
                         (_GenerateKeyfileCommand = new DelegateCommand(this.GenerateKeyfiles, (s) => canWork())); }
        }

        public ICommand UpdateFBDCommand
        {
            get { return _UpdateFBDCommand ??
                         (_UpdateFBDCommand = new DelegateCommand(this.UpdateFBDs, (s) => canWork())); }
        }

        // Sets default figure rectangle coordinates for UpdateFBD command
        public ICommand DefaultRectangleCheckedCommand
        {
            //get { return _DefaultRectangleCheckedCommand ?? (_DefaultRectangleCheckedCommand = new DelegateCommand(DefaultRectangleChecked)); }
            get { return _DefaultRectangleCheckedCommand ?? 
                         (_DefaultRectangleCheckedCommand = new DelegateCommand(() => { LeftOfX = DwgUpdater.LeftOfXDefault;
                                                                                        BelowY = DwgUpdater.BelowYDefault;
                                                                                        RightOfX = DwgUpdater.RightOfXDefault;
                                                                                        AboveY = DwgUpdater.AboveYDefault;
                                                                                       }));
            }
        }

        public ICommand InsXRefCommand
        {
            get { return _InsXRefCommand ??
                         (_InsXRefCommand = new DelegateCommand(this.InsXRef, (s) => canWork())); }
        }

        public ICommand LayerListerCommand
        {
            //get { return _LayerListerCommand ?? (_LayerListerCommand = new DelegateCommand(this.ListLayers, delegate(object param) { if (worker == null) { return true; } return !worker.IsBusy; })); }
            get { return _LayerListerCommand ??
                         (_LayerListerCommand = new DelegateCommand(this.ListLayers, (s) => canWork())); }
        }

        public ICommand CheckLayersCommand
        {
            get { return _CheckLayersCommand ??
                         (_CheckLayersCommand = new DelegateCommand(this.CheckLayers, (s) => canWork())); }
        }

        public ICommand ListObjectsCommand
        {
            get { return _ListObjectsCommand ??
                         (_ListObjectsCommand = new DelegateCommand(this.ListObjects));}
        }

        public ICommand ExtractCommand
        {
            get { return _ExtractCommand ??
                         (_ExtractCommand = new DelegateCommand(this.ExtractFiles, (s) => canWork())); }
        }

        public ICommand LowercaseLayersCommand
        {
            get { return _LowercaseLayersCommand ??
                         (_LowercaseLayersCommand = new DelegateCommand(this.LowercaseLayers, (s) => canWork())); }
        }

        public ICommand ConverterTo2000Command
        {
            get { return _ConverterTo2000Command ??
                         (_ConverterTo2000Command = new DelegateCommand(this.ConverterTo2000, (s) => canWork())); }
        }

        public ICommand CopyLayerCommand
        {
            get { return _CopyLayerCommand ??
                         (_CopyLayerCommand = new DelegateCommand(this.CopyLayer, (s) => canWork())); }
        }

        /****************************/
        /* Binded Props for methods */
        /****************************/

        // Single folder path used for all dwg processing commands

        private String _Path = "Folder Path";
        public String Path
        {
            get { return _Path; }
            set { _Path = value; RaisePropertyChangedEvent("Path"); }
        }

        // 2000 Converter

        private Int32 _ValueTo2000;
        public Int32 ValueTo2000
        {
            get { return _ValueTo2000; }
            set { _ValueTo2000 = value; RaisePropertyChangedEvent("ValueTo2000"); }
        }

        // FBD Updater

        private Int32 _ValueFU;
        public Int32 ValueFU
        {
            get { return _ValueFU; }
            set { _ValueFU = value; RaisePropertyChangedEvent("ValueFU"); }
        }

        private Boolean _FilesNotSpecified = false;
        public Boolean FilesNotSpecified
        {
            get { return _FilesNotSpecified; }
            set { _FilesNotSpecified = value; RaisePropertyChangedEvent("FilesNotSpecified"); }
        }

        private String _PathFUxml;
        public String PathFUxml {
            get { return _PathFUxml; }
            set { _PathFUxml = value; RaisePropertyChangedEvent("PathFUxml"); }
        }

        // Delivery Formatter

        private Int32 _ValueDF;
        public Int32 ValueDF {
            get { return _ValueDF; }
            set { _ValueDF = value; RaisePropertyChangedEvent("ValueDF"); }
        }

        private String _SuffixDF;
        public String SuffixDF {
            get { return _SuffixDF; }
            set { _SuffixDF = value; RaisePropertyChangedEvent("SuffixDF"); }
        }

        // Keyfile Generator

        private Int32 _ValueKF;
        public Int32 ValueKF
        {
            get { return _ValueKF; }
            set { _ValueKF = value; RaisePropertyChangedEvent("ValueKF"); }
        }

        //XRef Inserter

        private Int32 _ValueXI;
        public Int32 ValueXI
        {
            get { return _ValueXI; }
            set { _ValueXI = value; RaisePropertyChangedEvent("ValueXI"); }
        }

        private String _PathXIxlsx;
        public String PathXIxlsx
        {
            get { return _PathXIxlsx; }
            set { _PathXIxlsx = value; RaisePropertyChangedEvent("PathXIxlsx"); }
        }

        // Layer Checker

        private Boolean _MultiLC;
        public Boolean MultiLC
        { 
            get { return _MultiLC; }
            set { _MultiLC = value; RaisePropertyChangedEvent("MultiLC"); }
        }

        private Boolean _ChangesLC;
        public Boolean ChangesLC
        { 
            get { return _ChangesLC; }
            set { _ChangesLC = value; RaisePropertyChangedEvent("ChangesLC"); }
        }

        private Int32 _ValueLC;
        public Int32 ValueLC
        { 
            get { return _ValueLC; }
            set { _ValueLC = value; RaisePropertyChangedEvent("ValueLC"); }
        }

        // Text Obj Lister

        private List<String> _TextList;
        public List<String> TextList
        { 
            get { return _TextList; }
            set { _TextList = value; RaisePropertyChangedEvent("TextList"); }
        }

        // Layer Lister

        private Int32 _ValueLL;
        public Int32 ValueLL
        { 
            get { return _ValueLL; }
            set { _ValueLL = value; RaisePropertyChangedEvent("ValueLL"); }
        }

        //Extractor

        private Int32 _ValueE;
        public Int32 ValueE
        {
            get { return _ValueE; }
            set { _ValueE = value; RaisePropertyChangedEvent("ValueE"); }
        }

        // Lowercase Layers

        private Boolean _MultiLCL;
        public Boolean MultiLCL
        { 
            get { return _MultiLCL; }
            set { _MultiLCL = value; RaisePropertyChangedEvent("MultiLCL"); }
        }

        private Boolean _ChangesLCL;
        public Boolean ChangesLCL 
        {
            get { return _ChangesLCL; }
            set { _ChangesLCL = value; RaisePropertyChangedEvent("ChangesLCL"); }
        }

        private Int32 _ValueLCL;
        public Int32 ValueLCL
        { 
            get { return _ValueLCL; }
            set { _ValueLCL = value; RaisePropertyChangedEvent("ValueLCL"); }
        }

        // List info for all available commands in info tab

        private List<String> _CmdInfoList;
        public List<String> CmdInfoList
        { 
            get { return _CmdInfoList; }
            set { _CmdInfoList = value; RaisePropertyChangedEvent("CmdInfoList"); }
        }
        
        // Figure Rectangle coordinates for FBD Updater

        private Double _LeftOfX = DwgUpdater.LeftOfXDefault;
        public Double LeftOfX
        { 
            get { return _LeftOfX; }
            set { _LeftOfX = value; RaisePropertyChangedEvent("LeftOfX"); }
        }

        private Double _BelowY = DwgUpdater.BelowYDefault;
        public Double BelowY
        { 
            get { return _BelowY; }
            set { _BelowY = value; RaisePropertyChangedEvent("BelowY"); }
        }

        private Double _RightOfX = DwgUpdater.RightOfXDefault;
        public Double RightOfX
        { 
            get { return _RightOfX; }
            set { _RightOfX = value; RaisePropertyChangedEvent("RightOfX"); }
        }

        private Double _AboveY = DwgUpdater.AboveYDefault;
        public Double AboveY
        { 
            get { return _AboveY; }
            set { _AboveY = value; RaisePropertyChangedEvent("AboveY"); }
        }

        // Layer Copier

        private Int32 _ValueLCO;
        public Int32 ValueLCO
        {
            get { return _ValueLCO; }
            set { _ValueLCO = value; RaisePropertyChangedEvent("ValueLCO"); }
        }

        private String _PathCopyMap;
        public String PathCopyMap
        {
            get { return _PathCopyMap; }
            set { _PathCopyMap = value; RaisePropertyChangedEvent("PathCopyMap"); }
        }

        /*****************************/
        /* Backend Methods */
        /*****************************/

        private void CopyLayer()
        {
            ValueLCO= 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
                {
                    CallMethodOnMThread(delegate {_CopyLayerCommand.RaiseCanExecuteChanged(); });
                    DwgProcessor lc = new LayerCopier(Path, sender as BackgroundWorker, PathCopyMap);
                    e.Result = lc.Process();
                };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _CopyLayerCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueLCO = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
        }

        private void ConverterTo2000()
        {
            ValueTo2000 = 0;
            worker = GetWorker();
            
            worker.DoWork += new DoWorkEventHandler(delegate(object sender, DoWorkEventArgs e)
                {
                    CallMethodOnMThread(delegate { _ConverterTo2000Command.RaiseCanExecuteChanged(); });
                    DwgProcessor cf2000 = new ConverterTo2000(Path, sender as BackgroundWorker);
                    e.Result = cf2000.Process();
                });

            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(delegate(object sender, RunWorkerCompletedEventArgs e) 
                {
                    CallMethodOnMThread(delegate { _ConverterTo2000Command.RaiseCanExecuteChanged(); });
                    Path = e.Result as String;
                });

            worker.ProgressChanged += new ProgressChangedEventHandler(delegate(object sender, ProgressChangedEventArgs e)
                {
                    ValueTo2000 = e.ProgressPercentage;
                });

            worker.RunWorkerAsync();
        }
        
        private void ToggleDebugInfo()
        {
            if (!DebugInfoDisplayed)
            {
                CmdInfoList = CmdInfoList.Concat(ProgramInfo.GetDebugInfo()).ToList<String>();
                DebugInfoDisplayed = true;
            }
            else
            {
                CmdInfoList = ProgramInfo.GetCmdInfo().Concat(ProgramInfo.GetAssemblyInfo()).ToList<String>();
                DebugInfoDisplayed = false;
            }
            return;
        }

        private void GenerateKeyfiles()
        {
            ValueKF = 0;
            worker = GetWorker();

            #region Alternate ways
            //worker.DoWork += worker_DoWork;

            //void worker_DoWork(object sender, DoWorkEventArgs e)
            //{
            //    KeyfileGenerator kfg = new KeyfileGenerator(Path, sender as BackgroundWorker);
            //    kfg.Generate();
            //}

            //worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            // or
            //worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);

            //void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            //{
            //    ProgressVisibilityKF = Visibility.Hidden;
            //    Path = e.Result as String;
            //}


            //worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            // or
            //worker.ProgressChanged += worker_ProgressChanged;

            //void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
            //{
            //    ValueKF = e.ProgressPercentage;
            //}
            #endregion

            worker.DoWork += new DoWorkEventHandler(delegate(object sender, DoWorkEventArgs e)
                                                   {
                                                       CallMethodOnMThread(delegate { _GenerateKeyfileCommand.RaiseCanExecuteChanged(); });
                                                       DwgProcessor kfg = new KeyfileGenerator(Path, sender as BackgroundWorker);
                                                       e.Result = kfg.Process();
                                                   });

            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(delegate(object sender, RunWorkerCompletedEventArgs e)
                                                                           {
                                                                               CallMethodOnMThread(delegate { _GenerateKeyfileCommand.RaiseCanExecuteChanged(); });
                                                                               Path = e.Result as String;
                                                                           });

            worker.ProgressChanged += new ProgressChangedEventHandler(delegate(object sender, ProgressChangedEventArgs e)
                                                                     {
                                                                         ValueKF = e.ProgressPercentage;
                                                                     });

            // calls DoWork event, so DoWorkEventHandler runs cause its subscribed to the event
            worker.RunWorkerAsync();
            return;
        }

        private void UpdateFBDs()
        {
            ValueFU = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                CallMethodOnMThread(delegate { _UpdateFBDCommand.RaiseCanExecuteChanged(); });
                DwgProcessor fu = new FBDUpdater(Path,
                                               sender as BackgroundWorker,
                                               PathFUxml,
                                               Tuple.Create<double, double, double, double>(LeftOfX, BelowY, RightOfX, AboveY),
                                               FilesNotSpecified);
                e.Result = fu.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _UpdateFBDCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueFU = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }

        private void InsXRef()
        {
            ValueXI = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                CallMethodOnMThread(delegate { _InsXRefCommand.RaiseCanExecuteChanged(); });
                DwgProcessor xi = new XRefInserter(Path, sender as BackgroundWorker, PathXIxlsx);
                e.Result = xi.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _InsXRefCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueXI = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }

        private void FormatForDelivery()
        {
            ValueDF = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                CallMethodOnMThread(delegate { _FormatForDeliveryCommand.RaiseCanExecuteChanged(); });
                DwgProcessor df = new DeliveryFormatter(Path, sender as BackgroundWorker, SuffixDF);
                e.Result = df.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _FormatForDeliveryCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueDF = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }

        private void CheckLayers()
        {
            ValueLC = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                CallMethodOnMThread(delegate { _CheckLayersCommand.RaiseCanExecuteChanged(); });
                DwgProcessor lc = new LayerChecker(Path, sender as BackgroundWorker, MultiLC, ChangesLC);
                e.Result = lc.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _CheckLayersCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueLC = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }

        private void LowercaseLayers()
        {
            ValueLCL = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                CallMethodOnMThread(delegate { _LowercaseLayersCommand.RaiseCanExecuteChanged(); });
                DwgProcessor llm = new LowercaseLayerMaker(Path, sender as BackgroundWorker);
                e.Result = llm.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _LowercaseLayersCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueLCL = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }

        private void ListObjects()
        {
            try
            {
                ObjectLister ol = new ObjectLister();
                TextList = ol.Process();
            }
            catch (System.Exception se) { TextList = new List<String>() { "Error: " + se.Message }; }

            return;
        }

        private void ListLayers()
        {
            ValueLL = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                //using (IKernel kernel = new StandardKernel())
                //{
                //     //Ninject - Test
                //    kernel.Bind<IProcessor>().To<LayerLister>().Named("ll")
                //                             .WithConstructorArgument("inPath", Path);

                //var listerProcessor = kernel.Get<IProcessor>("ll");
                //    e.Result = listerProcessor.Process(sender as BackgroundWorker);
                //}

                CallMethodOnMThread(delegate { _LayerListerCommand.RaiseCanExecuteChanged(); });
                DwgProcessor ll = new LayerLister(Path, sender as BackgroundWorker);
                e.Result = ll.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _LayerListerCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueLL = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }

        private void ExtractFiles()
        {
            ValueE = 0;
            worker = GetWorker();

            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                CallMethodOnMThread(delegate { _ExtractCommand.RaiseCanExecuteChanged(); });
                DwgProcessor er = new Extractor(Path, sender as BackgroundWorker);
                e.Result = er.Process();
            };

            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                CallMethodOnMThread(delegate { _ExtractCommand.RaiseCanExecuteChanged(); });
                Path = e.Result as String;
            };

            worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                ValueE = e.ProgressPercentage;
            };

            worker.RunWorkerAsync();
            return;
        }
    }
}
