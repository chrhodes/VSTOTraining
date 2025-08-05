using System;
using System.Windows;

//using ModuleA;

using Prism.Ioc;
using Prism.Modularity;

using VNC;

namespace SupportTools_Visio.Application
{
    //public class PrismApp : Prism.Unity.PrismApplication
    //{
    //    protected override Window CreateShell()
    //    {
    //        long startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY, 0);
    //        //if (SupportTools_Visio.Ribbon.windowHostLocal is null)
    //        //{
    //        //    Ribbon.windowHostLocal = new Presentation.Views.WindowHost();
    //        //}

    //        Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, 0, startTicks);

    //        return Ribbon.windowHostLocal;
    //        ////return null;
    //        ////throw new NotImplementedException();
    //    }

    //    protected override void RegisterTypes(IContainerRegistry containerRegistry)
    //    {
    //        long startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY, 0);
    //        //throw new NotImplementedException();
    //        Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, 0, startTicks);
    //    }

    //    protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
    //    {
    //        long startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY, 0);
    //        Type moduleAType = typeof(ModuleAModule);

    //        moduleCatalog.AddModule(new ModuleInfo()
    //        {
    //            ModuleName = moduleAType.Name,
    //            ModuleType = moduleAType.AssemblyQualifiedName,
    //            InitializationMode = InitializationMode.WhenAvailable
    //            // InitializationMode = InitializationMode.OnDemand
    //        });

    //        base.ConfigureModuleCatalog(moduleCatalog);

    //        Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, 0, startTicks);
    //    }
    //}
}
