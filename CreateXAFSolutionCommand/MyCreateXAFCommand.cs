using DevExpress.ExpressApp.TemplateWizard;
using DevExpress.VisualStudioInterop.Base;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
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
            var model = new SolutionModel();
            string mySolutionName= "MyNewXAFSolution6";
            model.ApplicationName = mySolutionName;
            model.FullXafVersion = "21.2.0.0";
            model.AuthenticationIsStandard = true;
            model.BlazorMode = true;
            model.BlazorPlatformSelected = true;
            model.ClientLevelIntegratedSelected = true;
            model.HasXafLicense = true;
            model.Lang = DevExpress.VisualStudioInterop.Base.Language.CSharp;
            //model.MainModuleClassName = "Module";
            //model.MainModuleNamespace = mySolutionName+".Module";
            model.NetCoreMode = true;
            model.OrmIsXpo = true;
            model.CollectModules(true);
            model.WebApiPlatformSelected = false;
            var m1 = model.AllModules.Where(x => x is BusinessClassLibraryCustomizationModuleInfo).First();
            ((ISelectable)m1).Selected = true;

            var m2 = model.AllModules.Where(x => x is ConditionalAppearanceModuleInfo).First();
            ((ISelectable)m2).Selected = true;

            var m3 = model.AllModules.Where(x => x is OfficeModuleInfo).First();
            ((ISelectable)m3).Selected = true;

            var m4 = model.AllModules.Where(x => x is ValidationModuleInfo).First();
            ((ISelectable)m4).Selected = true;

            model.SolutionName = mySolutionName;

            model.TargetFrameworkVersion = "4.5.2";
            model.UseSecurity = true;
            model.VSVersion = "17.0";
            model.XafVersion = "21.2";
            wz.SetModel(model);
            var dxDte = VisualStudioInterop.GetDTE(dte); ;
            wz.SetDTE(dxDte);
            wz.SetBaseDirectory(@"c:\!Tickets\!Test\"+ mySolutionName);
            //wz.mo
            wz.RunFinished();
            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
