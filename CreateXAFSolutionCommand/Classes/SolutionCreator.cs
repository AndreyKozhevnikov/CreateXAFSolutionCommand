using DataForSolutionNameSpace;
using DevExpress.ExpressApp.TemplateWizard;
using DevExpress.VisualStudioInterop.Base;
using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CreateXAFSolutionCommand.Classes {
    public class SolutionCreator {
        public void CreateSolution() {
            NewXafSolutionWizard wz = new NewXafSolutionWizard();

            var solutionDataFileName = @"c:\solutiondata\data.xml";
            DataForSolution dataSolution = DeserializeToObject<DataForSolution>(solutionDataFileName);

            var model = new MySolutionModel();
            string mySolutionName = dataSolution.Name;
            model.ApplicationName = mySolutionName;
            model.FullXafVersion = dataSolution.FullXAFVersion;
            model.XafVersion = dataSolution.FullXAFVersion.Substring(0, 4);
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

            SelectModules(dataSolution, model);
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
        }
        EnvDTE.DTE dte;
        public void SetDTE(EnvDTE.DTE _dte) {
            dte = _dte;
        }
        public void SelectModules(DataForSolution dataSolution, ISolutionModel model) {
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

            File.Copy(Path.Combine(sourceSolutionPath, @"BusinessObjects\MyUpdater.cs"), Path.Combine(folderName, solutionName + @".Module\DatabaseUpdate\MyUpdater.cs"));
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
            if(dataSolution.Type == ProjectTypeEnum.Framework) {
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWin.cs"), Path.Combine(moduleWinPath, @"Controllers\CustomControllerWin.cs"));
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWeb.cs"), Path.Combine(moduleWebPath, @"Controllers\CustomControllerWeb.cs"));
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
            System.Diagnostics.Process.Start(Path.Combine(folderName, @"createGit.bat"));
        }

        public T DeserializeToObject<T>(string filepath) where T : class {
            System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(typeof(T));

            using(StreamReader sr = new StreamReader(filepath)) {
                return (T)ser.Deserialize(sr);
            }
        }
    }
}
