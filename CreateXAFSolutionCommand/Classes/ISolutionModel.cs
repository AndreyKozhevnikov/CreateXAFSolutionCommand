using DevExpress.ExpressApp.TemplateWizard;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace CreateXAFSolutionCommand.Classes {
    public interface ISolutionModel {
        ObservableCollection<IModuleInfo> AllModules { get; }
    }
    public class MySolutionModel : SolutionModel, ISolutionModel {

    }
}
