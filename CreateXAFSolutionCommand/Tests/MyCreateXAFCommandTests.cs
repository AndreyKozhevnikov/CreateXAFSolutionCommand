using CreateXAFSolutionCommand.Classes;
using DataForSolutionNameSpace;
using DevExpress.ExpressApp.TemplateWizard;
using Moq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateXAFSolutionCommand.Tests {
    [TestFixture]
    public class MyCreateXAFCommandTests {
        [Test]
        public void SelectModules() {
            //arrange
            var creator= new SolutionCreator();
            DataForSolution dataSolution = new DataForSolution();
            dataSolution.Modules = new List<ModulesEnum>();
            dataSolution.Modules.Add(ModulesEnum.Appearance);
            dataSolution.Modules.Add(ModulesEnum.Office);

            var moqModel = new Mock<ISolutionModel>();
            ObservableCollection<IModuleInfo> listModules = new ObservableCollection<IModuleInfo>();
            listModules.Add(new FileAttachmentsModuleInfo());
            listModules.Add(new ConditionalAppearanceModuleInfo());
            listModules.Add(new OfficeModuleInfo());
            
            moqModel.Setup(x => x.AllModules).Returns(listModules);
            //act
            creator.SelectModules(dataSolution, moqModel.Object);
            var finalModulesList = moqModel.Object.AllModules;
            //assert
            Assert.AreEqual(false, ((ISelectable)finalModulesList.First(x => x is FileAttachmentsModuleInfo)).Selected);
            Assert.AreEqual(true, ((ISelectable)finalModulesList.First(x => x is ConditionalAppearanceModuleInfo)).Selected);
            Assert.AreEqual(true, ((ISelectable)finalModulesList.First(x => x is OfficeModuleInfo)).Selected);


        }
    }

  
}
