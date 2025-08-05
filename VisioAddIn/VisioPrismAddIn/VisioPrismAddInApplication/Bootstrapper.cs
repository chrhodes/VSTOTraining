using System;
using System.Windows;
using System.Windows.Controls;

//using ModuleA;
//using Explore;

using Prism;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using Prism.Unity;

using VisioPrismAddInApplication.Modules;
//using VisioPrismAddInApplication.Presentation.ViewModels;
//using VisioPrismAddInApplication.Presentation.Views;

using VNC;
using VNC.Core.Mvvm.Prism;
using VNC.Core.Services;

//using VNC.Core.Mvvm;
//using VNC.Core.Mvvm.Prism;

namespace VisioPrismAddInApplication
{
    public class Bootstrapper : PrismBootstrapperBase
    {
        // Step 1 - Create the Unity Container

        protected override IContainerExtension CreateContainerExtension()
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("CreateContainerExtension()", true);

            return new UnityContainerExtension();
        }

        // Step 2 - Create the catalog of Modules

        protected override IModuleCatalog CreateModuleCatalog()
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("CreateModuleCatalog()", true);

            return new ConfigurationModuleCatalog();
        }

        // Step 3 - Configure the container

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("RegisterRequiredTypes()", true);

            // Registers all types that are required by Prism to function with the container.

            base.RegisterRequiredTypes(containerRegistry);

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("RegisterRequiredTypes()-Exit", startTicks, true);
        }

        // Step 4 - Register Types to be used

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("RegisterTypes()", true);

            // Used to register types with the container that will be used by your application.

            containerRegistry.Register<IMessageDialogService, MessageDialogService>();

            //containerRegistry.Register<IEditTextViewModel, EditTextViewModel>();
            //containerRegistry.Register<EditText>();

            //containerRegistry.Register<EditParagraphViewModel>();
            //containerRegistry.Register<EditParagraph>();

            //containerRegistry.Register<EditControlRowsViewModel>();
            //containerRegistry.Register<EditControlRows>();

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("RegisterTypes()-Exit", startTicks, true);
        }

        // Step 5 - Configure the catalog of modules
        // Modules are loaded at Startup and must be a project reference

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("ConfigureModuleCatalog()", true);

            moduleCatalog.AddModule(typeof(SupportTools_VisioModule));

            //moduleCatalog.AddModule(typeof(ModuleA.ModuleAModule));

            //moduleCatalog.AddModule(typeof(EditTextModule));

            //moduleCatalog.AddModule(typeof(Explore.ExploreModule));
            //moduleCatalog.AddModule(typeof(Explore.CarModule));

            base.ConfigureModuleCatalog(moduleCatalog);

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("ConfigureModuleCatalog()-Exit", startTicks, true);
        }

        // Step 6 - Configure the RegionAdapters if any custom ones have been created

        protected override void ConfigureRegionAdapterMappings(RegionAdapterMappings regionAdapterMappings)
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("ConfigureRegionAdapterMappings()", true);

            base.ConfigureRegionAdapterMappings(regionAdapterMappings);
            regionAdapterMappings.RegisterMapping(typeof(StackPanel), Container.Resolve<StackPanelRegionAdapter>());

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("ConfigureRegionAdapterMappings()-Exit", startTicks, true);
        }

        // Step 7 - Configure any Region Behaviors
        protected override void ConfigureDefaultRegionBehaviors(IRegionBehaviorFactory regionBehaviors)
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("ConfigureDefaultRegionBehaviors()", true);

            base.ConfigureDefaultRegionBehaviors(regionBehaviors);

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("ConfigureDefaultRegionBehaviors()-Exit", startTicks, true);
        }

        // Step 8 - Create the Shell that will hold the modules in designated regions.

        protected override DependencyObject CreateShell()
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //Log.APPLICATION_INITIALIZE("Exit (null)", Common.LOG_CATEGORY, startTicks);

            Common.WriteToDebugWindow("CreateShell()", true);

            return null;
            //return Container.Resolve<Views.MainWindow>();
            //return Container.TryResolve<Views.MainWindow>();
        }
    }
}
