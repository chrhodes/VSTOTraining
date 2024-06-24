using System;
using System.Linq;
using System.Reflection;

using Microsoft.Office.Interop.Visio;

namespace VisioAddInApplication.Events
{
    public class VisioAppEvents
    {
        private Application _VisioApplication;

        public Application VisioApplication
        {
            get
            {
                return _VisioApplication;
            }
            set
            {
                if (_VisioApplication != null)
                {
                    // Should remove all the event handlers;
                }

                _VisioApplication = value;

                if (_VisioApplication != null)
                {
                    _VisioApplication.AfterModal += new EApplication_AfterModalEventHandler(_VisioApplication_AfterModal);
                    _VisioApplication.AfterRemoveHiddenInformation += new EApplication_AfterRemoveHiddenInformationEventHandler(_VisioApplication_AfterRemoveHiddenInformation);
                    _VisioApplication.AfterResume += new EApplication_AfterResumeEventHandler(_VisioApplication_AfterResume);
                    _VisioApplication.AfterResumeEvents += new EApplication_AfterResumeEventsEventHandler(_VisioApplication_AfterResumeEvents);
                    _VisioApplication.AppActivated += new EApplication_AppActivatedEventHandler(_VisioApplication_AppActivated);
                    _VisioApplication.AppDeactivated += new EApplication_AppDeactivatedEventHandler(_VisioApplication_AppDeactivated);
                    _VisioApplication.AppObjActivated += new EApplication_AppObjActivatedEventHandler(_VisioApplication_AppObjActivated);
                    _VisioApplication.AppObjDeactivated += new EApplication_AppObjDeactivatedEventHandler(_VisioApplication_AppObjDeactivated);
                    _VisioApplication.BeforeDataRecordsetDelete += new EApplication_BeforeDataRecordsetDeleteEventHandler(_VisioApplication_BeforeDataRecordsetDelete);
                    _VisioApplication.BeforeDocumentClose += new EApplication_BeforeDocumentCloseEventHandler(_VisioApplication_BeforeDocumentClose);
                    _VisioApplication.BeforeDocumentSave += new EApplication_BeforeDocumentSaveEventHandler(_VisioApplication_BeforeDocumentSave);
                    _VisioApplication.BeforeDocumentSaveAs += new EApplication_BeforeDocumentSaveAsEventHandler(_VisioApplication_BeforeDocumentSaveAs);
                    _VisioApplication.BeforeMasterDelete += new EApplication_BeforeMasterDeleteEventHandler(_VisioApplication_BeforeMasterDelete);
                    _VisioApplication.BeforeModal += new EApplication_BeforeModalEventHandler(_VisioApplication_BeforeModal);
                    _VisioApplication.BeforePageDelete += new EApplication_BeforePageDeleteEventHandler(_VisioApplication_BeforePageDelete);
                    _VisioApplication.BeforeQuit += new EApplication_BeforeQuitEventHandler(_VisioApplication_BeforeQuit);
                    _VisioApplication.BeforeSelectionDelete += new EApplication_BeforeSelectionDeleteEventHandler(_VisioApplication_BeforeSelectionDelete);
                    _VisioApplication.BeforeShapeDelete += new EApplication_BeforeShapeDeleteEventHandler(_VisioApplication_BeforeShapeDelete);
                    _VisioApplication.BeforeShapeTextEdit += new EApplication_BeforeShapeTextEditEventHandler(_VisioApplication_BeforeShapeTextEdit);
                    _VisioApplication.BeforeStyleDelete += new EApplication_BeforeStyleDeleteEventHandler(_VisioApplication_BeforeStyleDelete);
                    _VisioApplication.BeforeSuspend += new EApplication_BeforeSuspendEventHandler(_VisioApplication_BeforeSuspend);
                    _VisioApplication.BeforeSuspendEvents += new EApplication_BeforeSuspendEventsEventHandler(_VisioApplication_BeforeSuspendEvents);
                    _VisioApplication.BeforeWindowClosed += new EApplication_BeforeWindowClosedEventHandler(_VisioApplication_BeforeWindowClosed);
                    _VisioApplication.BeforeWindowPageTurn += new EApplication_BeforeWindowPageTurnEventHandler(_VisioApplication_BeforeWindowPageTurn);
                    _VisioApplication.BeforeWindowSelDelete += new EApplication_BeforeWindowSelDeleteEventHandler(_VisioApplication_BeforeWindowSelDelete);
                    _VisioApplication.CalloutRelationshipAdded += new EApplication_CalloutRelationshipAddedEventHandler(_VisioApplication_CalloutRelationshipAdded);
                    _VisioApplication.CalloutRelationshipDeleted += new EApplication_CalloutRelationshipDeletedEventHandler(_VisioApplication_CalloutRelationshipDeleted);
                    _VisioApplication.CellChanged += new EApplication_CellChangedEventHandler(_VisioApplication_CellChanged);
                    _VisioApplication.ConnectionsAdded += new EApplication_ConnectionsAddedEventHandler(_VisioApplication_ConnectionsAdded);
                    _VisioApplication.ConnectionsDeleted += new EApplication_ConnectionsDeletedEventHandler(_VisioApplication_ConnectionsDeleted);
                    _VisioApplication.ContainerRelationshipAdded += new EApplication_ContainerRelationshipAddedEventHandler(_VisioApplication_ContainerRelationshipAdded);
                    _VisioApplication.ContainerRelationshipDeleted += new EApplication_ContainerRelationshipDeletedEventHandler(_VisioApplication_ContainerRelationshipDeleted);
                    _VisioApplication.ConvertToGroupCanceled += new EApplication_ConvertToGroupCanceledEventHandler(_VisioApplication_ConvertToGroupCanceled);
                    _VisioApplication.DataRecordsetAdded += new EApplication_DataRecordsetAddedEventHandler(_VisioApplication_DataRecordsetAdded);
                    _VisioApplication.DataRecordsetChanged += new EApplication_DataRecordsetChangedEventHandler(_VisioApplication_DataRecordsetChanged);
                    _VisioApplication.DesignModeEntered += new EApplication_DesignModeEnteredEventHandler(_VisioApplication_DesignModeEntered);
                    _VisioApplication.DocumentChanged += new EApplication_DocumentChangedEventHandler(_VisioApplication_DocumentChanged);
                    _VisioApplication.DocumentCloseCanceled += new EApplication_DocumentCloseCanceledEventHandler(_VisioApplication_DocumentCloseCanceled);
                    _VisioApplication.DocumentCreated += new EApplication_DocumentCreatedEventHandler(_VisioApplication_DocumentCreated);
                    _VisioApplication.DocumentOpened += new EApplication_DocumentOpenedEventHandler(_VisioApplication_DocumentOpened);
                    _VisioApplication.DocumentSaved += new EApplication_DocumentSavedEventHandler(_VisioApplication_DocumentSaved);
                    _VisioApplication.DocumentSavedAs += new EApplication_DocumentSavedAsEventHandler(_VisioApplication_DocumentSavedAs);
                    _VisioApplication.EnterScope += new EApplication_EnterScopeEventHandler(_VisioApplication_EnterScope);
                    _VisioApplication.ExitScope += new EApplication_ExitScopeEventHandler(_VisioApplication_ExitScope);
                    _VisioApplication.FormulaChanged += new EApplication_FormulaChangedEventHandler(_VisioApplication_FormulaChanged);
                    _VisioApplication.GroupCanceled += new EApplication_GroupCanceledEventHandler(_VisioApplication_GroupCanceled);
                    _VisioApplication.KeyDown += new EApplication_KeyDownEventHandler(_VisioApplication_KeyDown);
                    _VisioApplication.KeyPress += new EApplication_KeyPressEventHandler(_VisioApplication_KeyPress);
                    _VisioApplication.KeyUp += new EApplication_KeyUpEventHandler(_VisioApplication_KeyUp);
                    _VisioApplication.MarkerEvent += new EApplication_MarkerEventEventHandler(_VisioApplication_MarkerEvent);
                    _VisioApplication.MasterAdded += new EApplication_MasterAddedEventHandler(_VisioApplication_MasterAdded);
                    _VisioApplication.MasterChanged += new EApplication_MasterChangedEventHandler(_VisioApplication_MasterChanged);
                    _VisioApplication.MasterDeleteCanceled += new EApplication_MasterDeleteCanceledEventHandler(_VisioApplication_MasterDeleteCanceled);
                    _VisioApplication.MouseDown += new EApplication_MouseDownEventHandler(_VisioApplication_MouseDown);
                    _VisioApplication.MouseMove += new EApplication_MouseMoveEventHandler(_VisioApplication_MouseMove);
                    _VisioApplication.MouseUp += new EApplication_MouseUpEventHandler(_VisioApplication_MouseUp);
                    _VisioApplication.MustFlushScopeBeginning += new EApplication_MustFlushScopeBeginningEventHandler(_VisioApplication_MustFlushScopeBeginning);
                    _VisioApplication.MustFlushScopeEnded += new EApplication_MustFlushScopeEndedEventHandler(_VisioApplication_MustFlushScopeEnded);
                    _VisioApplication.NoEventsPending += new EApplication_NoEventsPendingEventHandler(_VisioApplication_NoEventsPending);
                    _VisioApplication.OnKeystrokeMessageForAddon += new EApplication_OnKeystrokeMessageForAddonEventHandler(_VisioApplication_OnKeystrokeMessageForAddon);
                    _VisioApplication.PageAdded += new EApplication_PageAddedEventHandler(_VisioApplication_PageAdded);
                    _VisioApplication.PageChanged += new EApplication_PageChangedEventHandler(_VisioApplication_PageChanged);
                    _VisioApplication.PageDeleteCanceled += new EApplication_PageDeleteCanceledEventHandler(_VisioApplication_PageDeleteCanceled);
                    _VisioApplication.QueryCancelConvertToGroup += new EApplication_QueryCancelConvertToGroupEventHandler(_VisioApplication_QueryCancelConvertToGroup);
                    _VisioApplication.QueryCancelDocumentClose += new EApplication_QueryCancelDocumentCloseEventHandler(_VisioApplication_QueryCancelDocumentClose);
                    _VisioApplication.QueryCancelGroup += new EApplication_QueryCancelGroupEventHandler(_VisioApplication_QueryCancelGroup);
                    _VisioApplication.QueryCancelMasterDelete += new EApplication_QueryCancelMasterDeleteEventHandler(_VisioApplication_QueryCancelMasterDelete);
                    _VisioApplication.QueryCancelPageDelete += new EApplication_QueryCancelPageDeleteEventHandler(_VisioApplication_QueryCancelPageDelete);
                    _VisioApplication.QueryCancelQuit += new EApplication_QueryCancelQuitEventHandler(_VisioApplication_QueryCancelQuit);
                    _VisioApplication.QueryCancelSelectionDelete += new EApplication_QueryCancelSelectionDeleteEventHandler(_VisioApplication_QueryCancelSelectionDelete);
                    _VisioApplication.QueryCancelStyleDelete += new EApplication_QueryCancelStyleDeleteEventHandler(_VisioApplication_QueryCancelStyleDelete);
                    _VisioApplication.QueryCancelSuspend += new EApplication_QueryCancelSuspendEventHandler(_VisioApplication_QueryCancelSuspend);
                    _VisioApplication.QueryCancelSuspendEvents += new EApplication_QueryCancelSuspendEventsEventHandler(_VisioApplication_QueryCancelSuspendEvents);
                    _VisioApplication.QueryCancelUngroup += new EApplication_QueryCancelUngroupEventHandler(_VisioApplication_QueryCancelUngroup);
                    _VisioApplication.QueryCancelWindowClose += new EApplication_QueryCancelWindowCloseEventHandler(_VisioApplication_QueryCancelWindowClose);
                    _VisioApplication.QuitCanceled += new EApplication_QuitCanceledEventHandler(_VisioApplication_QuitCanceled);
                    _VisioApplication.RuleSetValidated += new EApplication_RuleSetValidatedEventHandler(_VisioApplication_RuleSetValidated);
                    _VisioApplication.RunModeEntered += new EApplication_RunModeEnteredEventHandler(_VisioApplication_RunModeEntered);
                    _VisioApplication.SelectionAdded += new EApplication_SelectionAddedEventHandler(_VisioApplication_SelectionAdded);
                    _VisioApplication.SelectionChanged += new EApplication_SelectionChangedEventHandler(_VisioApplication_SelectionChanged);
                    _VisioApplication.SelectionDeleteCanceled += new EApplication_SelectionDeleteCanceledEventHandler(_VisioApplication_SelectionDeleteCanceled);
                    _VisioApplication.ShapeAdded += new EApplication_ShapeAddedEventHandler(_VisioApplication_ShapeAdded);
                    _VisioApplication.ShapeChanged += new EApplication_ShapeChangedEventHandler(_VisioApplication_ShapeChanged);
                    _VisioApplication.ShapeDataGraphicChanged += new EApplication_ShapeDataGraphicChangedEventHandler(_VisioApplication_ShapeDataGraphicChanged);
                    _VisioApplication.ShapeExitedTextEdit += new EApplication_ShapeExitedTextEditEventHandler(_VisioApplication_ShapeExitedTextEdit);
                    _VisioApplication.ShapeLinkAdded += new EApplication_ShapeLinkAddedEventHandler(_VisioApplication_ShapeLinkAdded);
                    _VisioApplication.ShapeLinkDeleted += new EApplication_ShapeLinkDeletedEventHandler(_VisioApplication_ShapeLinkDeleted);
                    _VisioApplication.ShapeParentChanged += new EApplication_ShapeParentChangedEventHandler(_VisioApplication_ShapeParentChanged);
                    _VisioApplication.StyleAdded += new EApplication_StyleAddedEventHandler(_VisioApplication_StyleAdded);
                    _VisioApplication.StyleChanged += new EApplication_StyleChangedEventHandler(_VisioApplication_StyleChanged);
                    _VisioApplication.StyleDeleteCanceled += new EApplication_StyleDeleteCanceledEventHandler(_VisioApplication_StyleDeleteCanceled);
                    _VisioApplication.SuspendCanceled += new EApplication_SuspendCanceledEventHandler(_VisioApplication_SuspendCanceled);
                    _VisioApplication.SuspendEventsCanceled += new EApplication_SuspendEventsCanceledEventHandler(_VisioApplication_SuspendEventsCanceled);
                    _VisioApplication.TextChanged += new EApplication_TextChangedEventHandler(_VisioApplication_TextChanged);
                    _VisioApplication.UngroupCanceled += new EApplication_UngroupCanceledEventHandler(_VisioApplication_UngroupCanceled);
                    _VisioApplication.ViewChanged += new EApplication_ViewChangedEventHandler(_VisioApplication_ViewChanged);
                    _VisioApplication.WindowActivated += new EApplication_WindowActivatedEventHandler(_VisioApplication_WindowActivated);
                    _VisioApplication.WindowChanged += new EApplication_WindowChangedEventHandler(_VisioApplication_WindowChanged);
                    _VisioApplication.WindowCloseCanceled += new EApplication_WindowCloseCanceledEventHandler(_VisioApplication_WindowCloseCanceled);
                    _VisioApplication.WindowOpened += new EApplication_WindowOpenedEventHandler(_VisioApplication_WindowOpened);
                    _VisioApplication.WindowTurnedToPage += new EApplication_WindowTurnedToPageEventHandler(_VisioApplication_WindowTurnedToPage);
                }
            }
        }

        #region Regular Events - Just Log

        short countWindowTurnedToPage;
        void _VisioApplication_WindowTurnedToPage(Window Window)
        {
            DisplayEventInWatchWindow(countWindowTurnedToPage++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeAdded;
        void _VisioApplication_ShapeAdded(Shape Shape)
        {
            DisplayEventInWatchWindow(countShapeAdded++, MethodInfo.GetCurrentMethod().Name);
            //Actions.Visio_Shape.HandleShapeAdded(Shape);
        }

        short countPageChanged;
        void _VisioApplication_PageChanged(Page Page)
        {
            DisplayEventInWatchWindow(countPageChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countMarkerEvent;
        void _VisioApplication_MarkerEvent(Application app, int SequenceNum, string ContextString)
        {
            DisplayEventInWatchWindow(countMarkerEvent++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeMasterDelete;
        void _VisioApplication_BeforeMasterDelete(Master Master)
        {
            DisplayEventInWatchWindow(countBeforeMasterDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countMasterDeleteCanceled;
        void _VisioApplication_MasterDeleteCanceled(Master Master)
        {
            DisplayEventInWatchWindow(countMasterDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countMasterChanged;
        void _VisioApplication_MasterChanged(Master Master)
        {
            DisplayEventInWatchWindow(countMasterChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countMasterAdded;
        void _VisioApplication_MasterAdded(Master Master)
        {
            DisplayEventInWatchWindow(countMasterAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countTextChange;
        void _VisioApplication_TextChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(countTextChange++, MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowOpened;
        void _VisioApplication_WindowOpened(Window Window)
        {
            DisplayEventInWatchWindow(countWindowOpened++, MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowCloseCanceled;
        void _VisioApplication_WindowCloseCanceled(Window Window)
        {
            DisplayEventInWatchWindow(countWindowCloseCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowChanged;
        void _VisioApplication_WindowChanged(Window Window)
        {
            DisplayEventInWatchWindow(countWindowChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowActivated;
        void _VisioApplication_WindowActivated(Window Window)
        {
            DisplayEventInWatchWindow(countWindowActivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countViewChanged;
        void _VisioApplication_ViewChanged(Window Window)
        {
            DisplayEventInWatchWindow(countViewChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countUngroupCanceled;
        void _VisioApplication_UngroupCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(countUngroupCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countSuspendEventsCanceled;
        void _VisioApplication_SuspendEventsCanceled(Application app)
        {
            DisplayEventInWatchWindow(countSuspendEventsCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countSuspendCanceled;
        void _VisioApplication_SuspendCanceled(Application app)
        {
            DisplayEventInWatchWindow(countSuspendCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countStyleDeleteCanceled;
        void _VisioApplication_StyleDeleteCanceled(Style Style)
        {
            DisplayEventInWatchWindow(countStyleDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countStyleChanged;
        void _VisioApplication_StyleChanged(Style Style)
        {
            DisplayEventInWatchWindow(countStyleChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countStyleAdded;
        void _VisioApplication_StyleAdded(Style Style)
        {
            DisplayEventInWatchWindow(countStyleAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeParentChanged;
        void _VisioApplication_ShapeParentChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(countShapeParentChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeLinkDeleted;
        void _VisioApplication_ShapeLinkDeleted(Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayEventInWatchWindow(countShapeLinkDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeLinkAdded;
        void _VisioApplication_ShapeLinkAdded(Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayEventInWatchWindow(countShapeLinkAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeExitedTextEdit;
        void _VisioApplication_ShapeExitedTextEdit(Shape Shape)
        {
            DisplayEventInWatchWindow(countShapeExitedTextEdit++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeDataGraphicChanged;
        void _VisioApplication_ShapeDataGraphicChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(countShapeDataGraphicChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeChanged;
        void _VisioApplication_ShapeChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(countShapeChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countSelectionDeleteCanceled;
        void _VisioApplication_SelectionDeleteCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(countSelectionDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countSelectionChanged;
        void _VisioApplication_SelectionChanged(Window Window)
        {
            //Common.EventAggregator.GetEvent<SelectionChangedEvent>().Publish();
            DisplayEventInWatchWindow(countSelectionChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countSelectionAdded;
        void _VisioApplication_SelectionAdded(Selection Selection)
        {
            DisplayEventInWatchWindow(countSelectionAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countRunModeEntered;
        void _VisioApplication_RunModeEntered(Document Doc)
        {
            DisplayEventInWatchWindow(countRunModeEntered++, MethodInfo.GetCurrentMethod().Name);
        }

        short countRuleSetValidated;
        void _VisioApplication_RuleSetValidated(ValidationRuleSet RuleSet)
        {
            DisplayEventInWatchWindow(countRuleSetValidated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countQuitCanceled;
        void _VisioApplication_QuitCanceled(Application app)
        {
            DisplayEventInWatchWindow(countQuitCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countQueryCancelWindowClose;
        bool _VisioApplication_QueryCancelWindowClose(Window Window)
        {
            DisplayEventInWatchWindow(countQueryCancelWindowClose++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelUngroup;
        bool _VisioApplication_QueryCancelUngroup(Selection Selection)
        {
            DisplayEventInWatchWindow(countQueryCancelUngroup++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelSuspendEvents;
        bool _VisioApplication_QueryCancelSuspendEvents(Application app)
        {
            DisplayEventInWatchWindow(countQueryCancelSuspendEvents++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelSuspend;
        bool _VisioApplication_QueryCancelSuspend(Application app)
        {
            DisplayEventInWatchWindow(countQueryCancelSuspend++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelStyleDelete;
        bool _VisioApplication_QueryCancelStyleDelete(Style Style)
        {
            DisplayEventInWatchWindow(countQueryCancelStyleDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelSelectionDelete;
        bool _VisioApplication_QueryCancelSelectionDelete(Selection Selection)
        {
            DisplayEventInWatchWindow(countQueryCancelSelectionDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelQuit;
        bool _VisioApplication_QueryCancelQuit(Application app)
        {
            DisplayEventInWatchWindow(countQueryCancelQuit++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelPageDelete;
        bool _VisioApplication_QueryCancelPageDelete(Page Page)
        {
            DisplayEventInWatchWindow(countQueryCancelPageDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelMasterDelete;
        bool _VisioApplication_QueryCancelMasterDelete(Master Master)
        {
            DisplayEventInWatchWindow(countQueryCancelMasterDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelGroup;
        bool _VisioApplication_QueryCancelGroup(Selection Selection)
        {
            DisplayEventInWatchWindow(countQueryCancelGroup++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelDocumentClose;
        bool _VisioApplication_QueryCancelDocumentClose(Document Doc)
        {
            DisplayEventInWatchWindow(countQueryCancelDocumentClose++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelConvertToGroup;
        bool _VisioApplication_QueryCancelConvertToGroup(Selection Selection)
        {
            DisplayEventInWatchWindow(countQueryCancelConvertToGroup++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countPageDeleteCanceled;
        void _VisioApplication_PageDeleteCanceled(Page Page)
        {
            DisplayEventInWatchWindow(countPageDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countPageAdded;
        void _VisioApplication_PageAdded(Page Page)
        {
            DisplayEventInWatchWindow(countPageAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countOnKeystrokeMessageForAddon;
        bool _VisioApplication_OnKeystrokeMessageForAddon(MSGWrap MSG)
        {
            DisplayEventInWatchWindow(countOnKeystrokeMessageForAddon++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countGroupCanceled;
        void _VisioApplication_GroupCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(countGroupCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countFormulaChanged;
        void _VisioApplication_FormulaChanged(Cell Cell)
        {
            DisplayEventInWatchWindow(countFormulaChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countExitScope;
        void _VisioApplication_ExitScope(Application app, int nScopeID, string bstrDescription, bool bErrOrCancelled)
        {
            DisplayEventInWatchWindow(countExitScope++, MethodInfo.GetCurrentMethod().Name);
        }

        short countEnterScope;
        void _VisioApplication_EnterScope(Application app, int nScopeID, string bstrDescription)
        {
            DisplayEventInWatchWindow(countEnterScope++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentSavedAs;
        void _VisioApplication_DocumentSavedAs(Document Doc)
        {
            DisplayEventInWatchWindow(countDocumentSavedAs++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentSaved;
        void _VisioApplication_DocumentSaved(Document Doc)
        {
            DisplayEventInWatchWindow(countDocumentSaved++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentOpened;
        void _VisioApplication_DocumentOpened(Document Doc)
        {
            DisplayEventInWatchWindow(countDocumentOpened++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentCreated;
        void _VisioApplication_DocumentCreated(Document Doc)
        {
            DisplayEventInWatchWindow(countDocumentCreated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentCloseCanceled;
        void _VisioApplication_DocumentCloseCanceled(Document Doc)
        {
            DisplayEventInWatchWindow(countDocumentCloseCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentChanged;
        void _VisioApplication_DocumentChanged(Document Doc)
        {
            DisplayEventInWatchWindow(countDocumentChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDesignModeEntered;
        void _VisioApplication_DesignModeEntered(Document Doc)
        {
            DisplayEventInWatchWindow(countDesignModeEntered++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDataRecordsetChanged;
        void _VisioApplication_DataRecordsetChanged(DataRecordsetChangedEvent DataRecordsetChanged)
        {
            DisplayEventInWatchWindow(countDataRecordsetChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countDataRecordsetAdded;
        void _VisioApplication_DataRecordsetAdded(DataRecordset DataRecordset)
        {
            DisplayEventInWatchWindow(countDataRecordsetAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countConvertToGroupCanceled;
        void _VisioApplication_ConvertToGroupCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(countConvertToGroupCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short countContainerRelationshipDeleted;
        void _VisioApplication_ContainerRelationshipDeleted(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(countContainerRelationshipDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short countContainerRelationshipAdded;
        void _VisioApplication_ContainerRelationshipAdded(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(countContainerRelationshipAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countConnectionsDeleted;
        void _VisioApplication_ConnectionsDeleted(Connects Connects)
        {
            DisplayEventInWatchWindow(countConnectionsDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short countConnectionsAdded;
        void _VisioApplication_ConnectionsAdded(Connects Connects)
        {
            DisplayEventInWatchWindow(countConnectionsAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countCellChanged;
        void _VisioApplication_CellChanged(Cell Cell)
        {
            DisplayEventInWatchWindow(countCellChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short countCalloutRelationshipDeleted;
        void _VisioApplication_CalloutRelationshipDeleted(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(countCalloutRelationshipDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short countCalloutRelationshipAdded;
        void _VisioApplication_CalloutRelationshipAdded(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(countCalloutRelationshipAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeWindowSelDelete;
        void _VisioApplication_BeforeWindowSelDelete(Window Window)
        {
            DisplayEventInWatchWindow(countBeforeWindowSelDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeWindowPageTurn;
        void _VisioApplication_BeforeWindowPageTurn(Window Window)
        {
            DisplayEventInWatchWindow(countBeforeWindowPageTurn++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeWindowClosed;
        void _VisioApplication_BeforeWindowClosed(Window Window)
        {
            DisplayEventInWatchWindow(countBeforeWindowClosed++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSuspendEvents;
        void _VisioApplication_BeforeSuspendEvents(Application app)
        {
            DisplayEventInWatchWindow(countBeforeSuspendEvents++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSuspend;
        void _VisioApplication_BeforeSuspend(Application app)
        {
            DisplayEventInWatchWindow(countBeforeSuspend++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeStyleDelete;
        void _VisioApplication_BeforeStyleDelete(Style Style)
        {
            DisplayEventInWatchWindow(countBeforeStyleDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeShapeTextEdit;
        void _VisioApplication_BeforeShapeTextEdit(Shape Shape)
        {
            DisplayEventInWatchWindow(countBeforeShapeTextEdit++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeShapeDelete;
        void _VisioApplication_BeforeShapeDelete(Shape Shape)
        {
            DisplayEventInWatchWindow(countBeforeShapeDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSelectionDelete;
        void _VisioApplication_BeforeSelectionDelete(Selection Selection)
        {
            DisplayEventInWatchWindow(countBeforeSelectionDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeQuit;
        void _VisioApplication_BeforeQuit(Application app)
        {
            DisplayEventInWatchWindow(countBeforeQuit++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforePageDelete;
        void _VisioApplication_BeforePageDelete(Page Page)
        {
            DisplayEventInWatchWindow(countBeforePageDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeModal;
        void _VisioApplication_BeforeModal(Application app)
        {
            DisplayEventInWatchWindow(countBeforeModal++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDocumentSaveAs;
        void _VisioApplication_BeforeDocumentSaveAs(Document Doc)
        {
            DisplayEventInWatchWindow(countBeforeDocumentSaveAs++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDocumentSave;
        void _VisioApplication_BeforeDocumentSave(Document Doc)
        {
            DisplayEventInWatchWindow(countBeforeDocumentSave++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDocumentClose;
        void _VisioApplication_BeforeDocumentClose(Document Doc)
        {
            DisplayEventInWatchWindow(countBeforeDocumentClose++, MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDataRecordsetDelete;
        void _VisioApplication_BeforeDataRecordsetDelete(DataRecordset DataRecordset)
        {
            DisplayEventInWatchWindow(countBeforeDataRecordsetDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAppObjDeactivated;
        void _VisioApplication_AppObjDeactivated(Application app)
        {
            DisplayEventInWatchWindow(countAppObjDeactivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAppObjActivated;
        void _VisioApplication_AppObjActivated(Application app)
        {
            DisplayEventInWatchWindow(countAppObjActivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAppDeactivated;
        void _VisioApplication_AppDeactivated(Application app)
        {
            DisplayEventInWatchWindow(countAppDeactivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAppActivated;
        void _VisioApplication_AppActivated(Application app)
        {
            DisplayEventInWatchWindow(countAppActivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterResumeEvents;
        void _VisioApplication_AfterResumeEvents(Application app)
        {
            DisplayEventInWatchWindow(countAfterResumeEvents++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterResume;
        void _VisioApplication_AfterResume(Application app)
        {
            DisplayEventInWatchWindow(countAfterResume++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterRemoveHiddenInformation;
        void _VisioApplication_AfterRemoveHiddenInformation(Document Doc)
        {
            DisplayEventInWatchWindow(countAfterRemoveHiddenInformation++, MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterModal;
        void _VisioApplication_AfterModal(Application app)
        {
            DisplayEventInWatchWindow(countAfterModal++, MethodInfo.GetCurrentMethod().Name);
        }

        #endregion

        #region Chatty Events - Log if DisplayChattyEvents

        short countKeyUp;
        void _VisioApplication_KeyUp(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countKeyUp++, MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countKeyUp++;
            }
        }

        short countKeyPress;
        void _VisioApplication_KeyPress(int KeyAscii, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countKeyPress++, MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countKeyPress++;
            }
        }

        short countKeyDown;
        void _VisioApplication_KeyDown(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countKeyDown++, MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countKeyDown++;
            }
        }
        short countNoEventsPending;
        void _VisioApplication_NoEventsPending(Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countNoEventsPending++, MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countNoEventsPending++;
            }
        }

        short countMustFlushScopeEnded;
        void _VisioApplication_MustFlushScopeEnded(Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countMustFlushScopeEnded++, MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countMustFlushScopeEnded++;
            }
        }

        short countMustFlushScopeBeginning;
        void _VisioApplication_MustFlushScopeBeginning(Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countMustFlushScopeBeginning++, MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countMustFlushScopeBeginning++;
            }
        }

        short countMouseDown;
        void _VisioApplication_MouseDown(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countMouseDown++, MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countMouseDown++;
            }
        }

        short countMouseUp;
        void _VisioApplication_MouseUp(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countMouseUp++, MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countMouseUp++;
            }
        }

        short countMouseMove;
        void _VisioApplication_MouseMove(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayEventInWatchWindow(countMouseMove++, MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countMouseMove++;
            }
        }

        #endregion Chatty Events

        private void DisplayEventInWatchWindow(short i, string outputLine)
        {
            if (Common.DisplayEvents)
            {
                Common.WriteToWatchWindow($"{outputLine}:{i}");
            }
        }

    }
}
