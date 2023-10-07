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
using System.Xml.Linq;

namespace CreateXAFSolutionCommand.Classes {
    public class SolutionCreator {
        const string sourceSolutionPath = @"c:\Dropbox\work\Templates\MainSolution\FilesToCreateSolution\";
        public void CreateSolution() {
            NewXafSolutionWizard wz = new NewXafSolutionWizard();

            var solutionDataFileName = @"c:\solutiondata\data.xml";
            DataForSolution dataSolution = DeserializeToObject<DataForSolution>(solutionDataFileName);

            var model = new MySolutionModel();
            string mySolutionName = dataSolution.Name;
            model.ApplicationName = mySolutionName;
            var dxAssembly = Assembly.GetAssembly(typeof(DevExpress.VisualStudioInterop.Base.BuildAction));
            var dxAssemblyVersion = dxAssembly.GetCustomAttribute<AssemblyFileVersionAttribute>().Version;
            model.FullXafVersion = dxAssemblyVersion;
            model.XafVersion = dxAssemblyVersion.Substring(0, 4);
            if (dataSolution.HasSecurity) {
                model.AuthenticationIsStandard = true;
                model.ClientLevelIntegratedSelected = true;
                model.UseSecurity = true;
            }
            switch (dataSolution.Type) {
                case ProjectTypeEnum.Core:
                    model.BlazorMode = true;
                    model.BlazorPlatformSelected = true;
                    model.WinPlatformSelected = true;
                    model.NetCoreMode = true;
                    if (dataSolution.HasWebAPI) {
                        model.WebApiPlatformSelected = true;
                        model.WebApiIntegratedModeSelected = false;
                    }
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
            switch (dataSolution.ORMType) {
                case ORMEnum.XPO:
                    model.OrmIsXpo = true;
                    break;
                case ORMEnum.EF:
                    model.OrmIsEntityFrameworkCore = true;
                    break;
            }

            model.CollectModules(true);
          //  model.WebApiPlatformSelected = false;

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
            try {
                wz.RunFinished();
                dte.Solution.SaveAs(Path.Combine(solutionDirectory, mySolutionName + ".sln"));
            }
            catch (Exception e) {

            }
            CopyServiceClasses(solutionDirectory, mySolutionName, dataSolution);
            CopyBOClasses(solutionDirectory, mySolutionName, dataSolution, dataSolution.ORMType);
            if (dataSolution.ORMType == ORMEnum.EF) {
                AddBOToDBContext(solutionDirectory, mySolutionName);
            }
            if (dataSolution.HasWebAPI) {
                AddBOToWebApi(solutionDirectory, mySolutionName);
            }
            AddUpdaterToSolution(solutionDirectory, mySolutionName);
            FixConfig(solutionDirectory, mySolutionName, dataSolution);
            CreateGit(solutionDirectory);
        }
        EnvDTE.DTE dte;
        public void SetDTE(EnvDTE.DTE _dte) {
            dte = _dte;
        }
        public void SelectModules(DataForSolution dataSolution, ISolutionModel model) {
            Dictionary<ModulesEnum, Type> modulesDictionary = new Dictionary<ModulesEnum, Type>();
            modulesDictionary.Add(ModulesEnum.FileAttachments, typeof(FileAttachmentsModuleInfo));
            modulesDictionary.Add(ModulesEnum.ConditionalAppearance, typeof(ConditionalAppearanceModuleInfo));
            modulesDictionary.Add(ModulesEnum.Office, typeof(OfficeModuleInfo));
            modulesDictionary.Add(ModulesEnum.Report, typeof(ReportModuleInfo));
            modulesDictionary.Add(ModulesEnum.Validation, typeof(ValidationModuleInfo));
            modulesDictionary.Add(ModulesEnum.Scheduler, typeof(SchedulerModuleInfo));
            modulesDictionary.Add(ModulesEnum.Dashboards, typeof(DashboardsModuleInfo));
            if (dataSolution.ORMType == ORMEnum.XPO) {
                modulesDictionary.Add(ModulesEnum.AuditTrail, typeof(AuditTrailModuleInfo));
            } else {
                modulesDictionary.Add(ModulesEnum.AuditTrail, typeof(AuditTrailModuleEFCoreInfo));
            }
            modulesDictionary.Add(ModulesEnum.TreeList, typeof(TreeListEditorsModuleInfo));
            modulesDictionary.Add(ModulesEnum.Notification, typeof(NotificationsModuleInfo));
            modulesDictionary.Add(ModulesEnum.ViewVariant, typeof(ViewVariantsModuleInfo));
            modulesDictionary.Add(ModulesEnum.StateMachine, typeof(StateMachineModuleInfo));


            foreach (var module in dataSolution.Modules) {
                var moduleInfo = modulesDictionary[module];
                var realModule = model.AllModules.FirstOrDefault(x => x.GetType() == moduleInfo);
                if (realModule == null)
                    continue;
                ((ISelectable)realModule).Selected = true;

            }
        }
        public void CopyServiceClasses(string folderName, string solutionName, DataForSolution dataSolution) {
            File.Copy(Path.Combine(sourceSolutionPath, @"delbinobj.bat"), Path.Combine(folderName, @"delbinobj.bat"));
            File.Copy(Path.Combine(sourceSolutionPath, @"delbinobjWOVS.bat"), Path.Combine(folderName, @"delbinobjWOVS.bat"));
            File.Copy(Path.Combine(sourceSolutionPath, @".gitignore"), Path.Combine(folderName, @".gitignore"));
            File.Copy(Path.Combine(sourceSolutionPath, @"createGit.bat"), Path.Combine(folderName, @"createGit.bat"));
        }

        public void CopyBOClasses(string folderName, string solutionName, DataForSolution dataSolution, ORMEnum ormType) {

            List<Tuple<String, string>> addedFiles = new List<Tuple<String, string>>();
            var modulePath = Path.Combine(folderName, solutionName + ".Module");
            var moduleWinPath = Path.Combine(folderName, solutionName + ".Module.Win");
            var moduleWebPath = Path.Combine(folderName, solutionName + ".Module.Web");
            var moduleBlazorCorePath = Path.Combine(folderName, solutionName + ".Blazor.Server");
            var moduleWinCorePath = Path.Combine(folderName, solutionName + ".Win");

            var modulecsProjName = Path.Combine(folderName, solutionName + ".Module", solutionName + ".Module.csproj");
            var wincsProjName = Path.Combine(folderName, solutionName + ".Module.Win", solutionName + ".Module.Win.csproj");
            var webcsProjName = Path.Combine(folderName, solutionName + ".Module.Web", solutionName + ".Module.Web.csproj");


            //var suoPathSource = Path.Combine(sourceSolutionPath, @".vs\dxT1121016v5\v17\.suo");
            //var suoPathTargetFolder = Path.Combine(folderName, @".vs", solutionName, @"v17");
            //Directory.CreateDirectory(suoPathTargetFolder);
            //var suoPathTarget = Path.Combine(suoPathTargetFolder, @".suo");
            //File.Copy(suoPathSource, suoPathTarget);


            var fileNames = new List<string>();
            switch (ormType) {
                case ORMEnum.XPO:
                    fileNames.Add(@"BusinessObjects\Contact.cs");
                    fileNames.Add(@"BusinessObjects\MyTask.cs");
                    fileNames.Add(@"BusinessObjects\CustomClass.cs");
                    break;
                case ORMEnum.EF:
                    fileNames.Add(@"BusinessObjects\EFBusinessClasses.cs");
          
                    break;
            }
            foreach (string file in fileNames) {
                var filePath = Path.Combine(modulePath, file);
                File.Copy(Path.Combine(sourceSolutionPath, file), filePath);
                //  addedFiles.Add(file);
                addedFiles.Add(new Tuple<string, string>(file, modulecsProjName));
            }

            File.Copy(Path.Combine(sourceSolutionPath, @"BusinessObjects\MyUpdater.cs"), Path.Combine(folderName, solutionName + @".Module\DatabaseUpdate\MyUpdater.cs"));
            addedFiles.Add(new Tuple<string, string>(@"DatabaseUpdate\MyUpdater.cs", modulecsProjName));
            File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllers.cs"), Path.Combine(folderName, solutionName + @".Module\Controllers\CustomControllers.cs"));
            addedFiles.Add(new Tuple<string, string>(@"Controllers\CustomControllers.cs", modulecsProjName));


            if (dataSolution.Modules.Contains(ModulesEnum.Report)) {
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\ClearReportCacheController.cs"), Path.Combine(folderName, solutionName + @".Module\Controllers\ClearReportCacheController.cs"));
                addedFiles.Add(new Tuple<string, string>(@"Controllers\ClearReportCacheController.cs", modulecsProjName));
            }


            if (dataSolution.Type == ProjectTypeEnum.Framework) {
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWin.cs"), Path.Combine(moduleWinPath, @"Controllers\CustomControllerWin.cs"));
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWeb.cs"), Path.Combine(moduleWebPath, @"Controllers\CustomControllerWeb.cs"));
                //addedFiles.Add(new Tuple<string, string>(@"Controllers\CustomControllerWin.cs", wincsProjName));
                //addedFiles.Add(new Tuple<string, string>(@"Controllers\CustomControllerWeb.cs", webcsProjName));
            }
            if (dataSolution.Type == ProjectTypeEnum.Core) {
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerWin.cs"), Path.Combine(moduleWinCorePath, @"Controllers\CustomControllerWin.cs"));
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\GridViewControllerWin.cs"), Path.Combine(moduleWinCorePath, @"Controllers\GridViewControllerWin.cs"));
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\CustomControllerBlazor.cs"), Path.Combine(moduleBlazorCorePath, @"Controllers\CustomControllerBlazor.cs"));
                File.Copy(Path.Combine(sourceSolutionPath, @"Controllers\GridViewControllerBlazor.cs"), Path.Combine(moduleBlazorCorePath, @"Controllers\GridViewControllerBlazor.cs"));
            }
            if (dataSolution.Type == ProjectTypeEnum.Framework) {
                AddFilesToCSprojFiles(addedFiles);
            }
        }

        void AddFilesToCSprojFiles(List<Tuple<string, string>> files) {
            var csprojDict = new Dictionary<string, List<string>>();
            foreach (var file in files) {
                if (!csprojDict.ContainsKey(file.Item2)) {
                    csprojDict[file.Item2] = new List<string>();
                }
                csprojDict[file.Item2].Add(file.Item1);
            }

            foreach (var csproj in csprojDict) {
                var csprojName = csproj.Key;
                var xFile = XDocument.Load(csprojName);
                var itemGroups = xFile.Root.Elements().Where(x => x.Name.LocalName == "ItemGroup");
                var itemGroup = itemGroups.Where(x => x.Elements().Where(y => y.Name.LocalName == "Compile").Count() > 0).FirstOrDefault();
                var itemGroupFirstElement = itemGroup.Elements().First();
                foreach (var file in csproj.Value) {
                    var fileElement = new XElement(itemGroupFirstElement);
                    fileElement.Attribute("Include").Value = file;
                    itemGroup.Add(fileElement);
                }
                xFile.Save(csprojName);
            }

        }


        public void AddBOToWebApi(string folderName, string solutionName) {
            var dbContextPah = Path.Combine(folderName, solutionName + string.Format(@".WebApi\Startup.cs", solutionName));
            string text = File.ReadAllText(dbContextPah);
            text = text.Replace("using Microsoft.OpenApi.Models;", "using Microsoft.OpenApi.Models;\r\nusing dxTestSolution.Module.BusinessObjects;");
            text = text.Replace("options.BusinessObject<YourBusinessObject>();",
              @"options.BusinessObject<YourBusinessObject>();
                options.BusinessObject<Contact>();
                options.BusinessObject<MyTask>();
");
            File.WriteAllText(dbContextPah, text);
        }
        public void AddBOToDBContext(string folderName, string solutionName) {
            var dbContextPah = Path.Combine(folderName, solutionName + string.Format(@".Module\BusinessObjects\{0}DBContext.cs",solutionName));
            string text = File.ReadAllText(dbContextPah);
            text = text.Replace("using Microsoft.EntityFrameworkCore;", "using Microsoft.EntityFrameworkCore;\r\nusing dxTestSolution.Module.BusinessObjects;");
            text = text.Replace("public DbSet<ModuleInfo> ModulesInfo { get; set; }",
                @"public DbSet<ModuleInfo> ModulesInfo { get; set; }
                  public DbSet<MyTask> MyTasks { get; set; }
                  public DbSet<Contact> Contacts { get; set; }
");
            File.WriteAllText(dbContextPah, text);
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
            string dataBaseName;
            DataBaseCreatorLib.DataBaseCreator.CreateSQLDataBaseIfNotExists(solutionName, out dataBaseName);
            switch (dataSolution.Type) {
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
            foreach (var file in configFiles) {
                string text = File.ReadAllText(file);
                string intialText = string.Format("Initial Catalog={0}", solutionName);
                string newText = string.Format("Initial Catalog={0}", dataBaseName);
                text = text.Replace(intialText, newText);
                File.WriteAllText(file, text);
            }
        }
        public void CreateGit(string folderName) {
            System.Diagnostics.Process.Start(Path.Combine(folderName, @"createGit.bat"));
        }

        public T DeserializeToObject<T>(string filepath) where T : class {
            System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(typeof(T));

            using (StreamReader sr = new StreamReader(filepath)) {
                return (T)ser.Deserialize(sr);
            }
        }
    }
}
