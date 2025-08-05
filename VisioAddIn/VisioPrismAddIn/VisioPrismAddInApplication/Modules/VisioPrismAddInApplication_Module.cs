using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

//using SupportTools_Visio.Presentation.ViewModels;

using Unity;

using VNC;

namespace VisioPrismAddInApplication.Modules
{
    public class SupportTools_VisioModule : IModule
    {
        // 01
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            //long startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //containerRegistry.Register<IViewAViewModel, ViewAViewModel>();
            //containerRegistry.Register<IViewBViewModel, ViewBViewModel>();
            //containerRegistry.Register<ICommandCockpitViewModel, CommandCockpitViewModel>();
            //////containerRegistry.Register<IViewA, ViewA>();
            //containerRegistry.Register<IViewCViewModel, ViewCViewModel>();

            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02
        public void OnInitialized(IContainerProvider containerProvider)
        {
            //long startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            var regionManager = containerProvider.Resolve<IRegionManager>();

            // Multiple ToolBar Regions

            //IRegion region = regionManager.Regions[RegionNames.ToolBarRegionA];

            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());

            //regionManager.RegisterViewWithRegion(RegionNames.ToolBarRegionA, typeof(ToolBarView));

            // NOTE(crhodes)
            // Can't get this to work.  Hum.  May have to eschew multiple toolbars :)

            //IRegion region = regionManager.Regions[RegionNames.ToolBarRegionA];

            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());

            //regionManager.RegisterViewWithRegion(RegionNames.EditTextRegion, typeof(EditText));

            //regionManager.RegisterViewWithRegion(RegionNames.EditControlPointsRegion, typeof(EditControlPoints));

            //regionManager.RegisterViewWithRegion(RegionNames.EditParagraphRegion, typeof(EditParagraph));

            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
