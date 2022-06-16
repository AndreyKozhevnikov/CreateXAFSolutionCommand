using DataForSolutionNameSpace;
using DevExpress.ExpressApp.TemplateWizard;
using DevExpress.VisualStudioInterop.Base;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace CreateXAFSolutionCommand {
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class MyCreateXAFCommand {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("24313641-d073-483d-a37d-ca782ec28211");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="MyCreateXAFCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private MyCreateXAFCommand(AsyncPackage package, OleMenuCommandService commandService) {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static MyCreateXAFCommand Instance {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider {
            get {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package, EnvDTE.DTE _dte) {
            // Switch to the main thread - the call to AddCommand in MyCreateXAFCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new MyCreateXAFCommand(package, commandService);
            Instance.dte = _dte;
        }
        EnvDTE.DTE dte;
        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "MyCreateXAFCommand";


            // EnvDTE.DTE dte = LaunchVsDte(isPreRelease: false);
            NewXafSolutionWizard wz = new NewXafSolutionWizard();

            var solutionDataFileName = @"c:\solutiondata\data.xml";
            DataForSolution dataSolution = DeserializeToObject<DataForSolution>(solutionDataFileName);

            var model = new SolutionModel();
            string mySolutionName = dataSolution.Name;
            model.ApplicationName = mySolutionName;
            model.FullXafVersion = "21.2.0.0";
            model.XafVersion = "21.2";
            if(dataSolution.HasSecurity) {
                model.AuthenticationIsStandard = true;
                model.ClientLevelIntegratedSelected = true;
                model.UseSecurity = true;
            }
            switch(dataSolution.Type) {
                case ProjectTypeEnum.Core:
                    model.BlazorMode = true;
                    model.BlazorPlatformSelected = true;
                    model.WinPlatformSelected = true;
                    model.NetCoreMode = true;
                    break;
                case ProjectTypeEnum.Framework:
                    model.WinPlatformSelected = true;
                    model.WebPlatformSelected = true;
                    model.NetCoreMode = false;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            model.Lang = DevExpress.VisualStudioInterop.Base.Language.CSharp;
            model.OrmIsXpo = true;
            model.CollectModules(true);
            model.WebApiPlatformSelected = false;


            if(dataSolution.Modules.Contains(ModulesEnum.ConditionalAppearance)) {
                var m2 = model.AllModules.Where(x => x is ConditionalAppearanceModuleInfo).First();
                ((ISelectable)m2).Selected = true;
            }
            if(dataSolution.Modules.Contains(ModulesEnum.FileAttachments)) {
                var m2 = model.AllModules.Where(x => x is FileAttachmentsModuleInfo).First();
                ((ISelectable)m2).Selected = true;
            }
            if(dataSolution.Modules.Contains(ModulesEnum.Office)) {
                var m3 = model.AllModules.Where(x => x is OfficeModuleInfo).First();
                ((ISelectable)m3).Selected = true;
            }
            if(dataSolution.Modules.Contains(ModulesEnum.Validation)) {
                var m4 = model.AllModules.Where(x => x is ValidationModuleInfo).First();
                ((ISelectable)m4).Selected = true;
            }
            if(dataSolution.Modules.Contains(ModulesEnum.Report)) {
                var m4 = model.AllModules.Where(x => x is ReportModuleInfo).First();
                ((ISelectable)m4).Selected = true;
            }
            if(dataSolution.Modules.Contains(ModulesEnum.Scheduler)) {
                var m4 = model.AllModules.Where(x => x is SchedulerModuleInfo).First();
                ((ISelectable)m4).Selected = true;
            }
            if(dataSolution.Modules.Contains(ModulesEnum.Dashboards)) {
                var m4 = model.AllModules.Where(x => x is DashboardsModuleInfo).First();
                ((ISelectable)m4).Selected = true;
            }
            model.SolutionName = mySolutionName;
            model.TargetFrameworkVersion = "4.5.2";
            model.VSVersion = "17.0";

            var dxDte = VisualStudioInterop.GetDTE(dte);
            Type vs = typeof(NewXafSolutionWizard);

            FieldInfo r_dte = vs.GetField("dte", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
            r_dte.SetValue(wz, dxDte);

            FieldInfo r_model = vs.GetField("model", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
            r_model.SetValue(wz, model);

            FieldInfo r_baseDirectory = vs.GetField("baseDirectory", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
            var solutionDirectory = Path.Combine(@"c:\!Tickets\", dataSolution.FolderName, dataSolution.Name);
            r_baseDirectory.SetValue(wz, solutionDirectory);

            wz.RunFinished();
            dte.Solution.SaveAs(Path.Combine(solutionDirectory, mySolutionName + ".sln"));

            CopyClasses(solutionDirectory, mySolutionName, dataSolution);
            AddUpdaterToSolution(solutionDirectory, mySolutionName);
            FixConfig(solutionDirectory, mySolutionName, dataSolution);
            CreateGit(solutionDirectory);
            // Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.package,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        public void CopyClasses(string folderName, string solutionName, DataForSolution dataSolution) {

            List<String> addedFiles = new List<String>();
            var modulePath = Path.Combine(folderName, solutionName + ".Module");
            var moduleWinPath = Path.Combine(folderName, solutionName + ".Module.Win");
            var moduleWebPath = Path.Combine(folderName, solutionName + ".Module.Web");
            var moduleBlazorPath = Path.Combine(folderName, solutionName + ".Module.Blazor");
            var sourceSolutionPath = @"c:\Dropbox\work\Templates\MainSolution\FilesToCreateSolution\";
            var csProjName = Path.Combine(folderName, solutionName + ".Module", solutionName + ".Module.csproj");

            var fileNames = new List<string>();
            fileNames.Add(@"BusinessObjects\Contact.cs");
            fileNames.Add(@"BusinessObjects\MyTask.cs");
            fileNames.Add(@"BusinessObjects\CustomClass.cs");
            foreach(string file in fileNames) {
                var filePath = Path.Combine(modulePath, file);
                File.Copy(Path.Combine(sourceSolutionPath, file), filePath);
                addedFiles.Add(file);

            }

            File.Copy(Path.Combine(sourceSolutionPath, @"BusinessObjects\Updater.cs"), Path.Combine(folderName, solutionName + @".Module\DatabaseUpdate\MyUpdater.cs"));
            addedFiles.Add(@"DatabaseUpdate\MyUpdater.cs");
            File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllers.cs"), Path.Combine(folderName, solutionName + @".Module\Controllers\CustomControllers.cs"));
            addedFiles.Add(@"Controllers\CustomControllers.cs");


            if(dataSolution.Modules.Contains(ModulesEnum.Report)) {
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\ClearReportCacheController.cs"), Path.Combine(folderName, solutionName + @".Module\Controllers\ClearReportCacheController.cs"));
                addedFiles.Add(@"Controllers\ClearReportCacheController.cs");
            }
            if(dataSolution.Type == ProjectTypeEnum.Framework) {
                //var p = new Microsoft.Build.Evaluation.Project(csProjName);
                //foreach(var st in addedFiles) {
                //    p.AddItem("Compile", st);
                //}
                //p.Save();
            }
            File.Copy(Path.Combine(sourceSolutionPath, @"delbinobj.bat"), Path.Combine(folderName, @"delbinobj.bat"));
            File.Copy(Path.Combine(sourceSolutionPath, @".gitignore"), Path.Combine(folderName, @".gitignore"));
            File.Copy(Path.Combine(sourceSolutionPath, @"createGit.bat"), Path.Combine(folderName, @"createGit.bat"));
            switch(dataSolution.Type) {
                case ProjectTypeEnum.Framework:
                    File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWin.cs"), Path.Combine(moduleWinPath, @"Controllers\CustomControllerWin.cs"));
                    File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWeb.cs"), Path.Combine(moduleWebPath, @"Controllers\CustomControllerWeb.cs"));
                    break;
                case ProjectTypeEnum.Core:
                    File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWin.cs"), Path.Combine(moduleWinPath, @"Controllers\CustomControllerWin.cs"));
                    File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerBlazor.cs"), Path.Combine(moduleBlazorPath, @"Controllers\CustomControllerBlazor.cs"));
                    break;
            }
        }
        public void AddUpdaterToSolution(string folderName, string solutionName) {
            var updaterPath = Path.Combine(folderName, solutionName + @".Module\Module.cs");
            string text = File.ReadAllText(updaterPath);
            text = text.Replace("using DevExpress.ExpressApp;", "using DevExpress.ExpressApp;\r\nusing dxTestSolution.Module.DatabaseUpdate;");
            text = text.Replace("return new ModuleUpdater[] { updater };", "return new ModuleUpdater[] { updater, new MyUpdater(objectSpace,versionFromDB) };");
            File.WriteAllText(updaterPath, text);
        }
        public void FixConfig(string folderName, string solutionName, DataForSolution dataSolution) {
            List<string> configFiles = new List<string>();
            switch(dataSolution.Type) {
                case ProjectTypeEnum.Core:
                    var configPath = Path.Combine(folderName, solutionName + ".Blazor.Server", "appsettings.json");
                    configFiles.Add(configPath);
                    var configPathWin = Path.Combine(folderName, solutionName + ".Win", "app.config");
                    configFiles.Add(configPathWin);
                    break;
                case ProjectTypeEnum.Framework:
                    var webconfigPath = Path.Combine(folderName, solutionName + ".Web", "web.config");
                    configFiles.Add(webconfigPath);
                    var configPathWin2 = Path.Combine(folderName, solutionName + ".Win", "app.config");
                    configFiles.Add(configPathWin2);
                    break;

            }
            foreach(var file in configFiles) {
                string text = File.ReadAllText(file);
                string intialText = "Initial Catalog=" + solutionName;
                string newText = string.Format("Initial Catalog=d{0}-{1}", DateTime.Today.DayOfYear, solutionName);
                text = text.Replace(intialText, newText);
                File.WriteAllText(file, text);
            }
        }
        public void CreateGit(string folderName) {
            Process.Start(Path.Combine(folderName, @"createGit.bat"));
        }

        public T DeserializeToObject<T>(string filepath) where T : class {
            System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(typeof(T));

            using(StreamReader sr = new StreamReader(filepath)) {
                return (T)ser.Deserialize(sr);
            }
        }
    }
}
