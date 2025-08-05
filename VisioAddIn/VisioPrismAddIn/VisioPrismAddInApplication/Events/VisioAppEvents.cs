using System;
using System.Reflection;

using Microsoft.Office.Interop.Visio;

namespace VisioPrismAddInApplication.Events
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
                    _VisioApplication.AfterModal += new EApplication_AfterModalEventHandler(VisioApplication_AfterModal);
                    _VisioApplication.AfterRemoveHiddenInformation += new EApplication_AfterRemoveHiddenInformationEventHandler(VisioApplication_AfterRemoveHiddenInformation);
                    _VisioApplication.AfterResume += new EApplication_AfterResumeEventHandler(VisioApplication_AfterResume);
                    _VisioApplication.AfterResumeEvents += new EApplication_AfterResumeEventsEventHandler(VisioApplication_AfterResumeEvents);
                    _VisioApplication.AppActivated += new EApplication_AppActivatedEventHandler(VisioApplication_AppActivated);
                    _VisioApplication.AppDeactivated += new EApplication_AppDeactivatedEventHandler(VisioApplication_AppDeactivated);
                    _VisioApplication.AppObjActivated += new EApplication_AppObjActivatedEventHandler(VisioApplication_AppObjActivated);
                    _VisioApplication.AppObjDeactivated += new EApplication_AppObjDeactivatedEventHandler(VisioApplication_AppObjDeactivated);
                    _VisioApplication.BeforeDataRecordsetDelete += new EApplication_BeforeDataRecordsetDeleteEventHandler(VisioApplication_BeforeDataRecordsetDelete);
                    _VisioApplication.BeforeDocumentClose += new EApplication_BeforeDocumentCloseEventHandler(VisioApplication_BeforeDocumentClose);
                    _VisioApplication.BeforeDocumentSave += new EApplication_BeforeDocumentSaveEventHandler(VisioApplication_BeforeDocumentSave);
                    _VisioApplication.BeforeDocumentSaveAs += new EApplication_BeforeDocumentSaveAsEventHandler(VisioApplication_BeforeDocumentSaveAs);
                    _VisioApplication.BeforeMasterDelete += new EApplication_BeforeMasterDeleteEventHandler(VisioApplication_BeforeMasterDelete);
                    _VisioApplication.BeforeModal += new EApplication_BeforeModalEventHandler(VisioApplication_BeforeModal);
                    _VisioApplication.BeforePageDelete += new EApplication_BeforePageDeleteEventHandler(VisioApplication_BeforePageDelete);
                    _VisioApplication.BeforeQuit += new EApplication_BeforeQuitEventHandler(VisioApplication_BeforeQuit);
                    _VisioApplication.BeforeSelectionDelete += new EApplication_BeforeSelectionDeleteEventHandler(VisioApplication_BeforeSelectionDelete);
                    _VisioApplication.BeforeShapeDelete += new EApplication_BeforeShapeDeleteEventHandler(VisioApplication_BeforeShapeDelete);
                    _VisioApplication.BeforeShapeTextEdit += new EApplication_BeforeShapeTextEditEventHandler(VisioApplication_BeforeShapeTextEdit);
                    _VisioApplication.BeforeStyleDelete += new EApplication_BeforeStyleDeleteEventHandler(VisioApplication_BeforeStyleDelete);
                    _VisioApplication.BeforeSuspend += new EApplication_BeforeSuspendEventHandler(VisioApplication_BeforeSuspend);
                    _VisioApplication.BeforeSuspendEvents += new EApplication_BeforeSuspendEventsEventHandler(VisioApplication_BeforeSuspendEvents);
                    _VisioApplication.BeforeWindowClosed += new EApplication_BeforeWindowClosedEventHandler(VisioApplication_BeforeWindowClosed);
                    _VisioApplication.BeforeWindowPageTurn += new EApplication_BeforeWindowPageTurnEventHandler(VisioApplication_BeforeWindowPageTurn);
                    _VisioApplication.BeforeWindowSelDelete += new EApplication_BeforeWindowSelDeleteEventHandler(VisioApplication_BeforeWindowSelDelete);
                    _VisioApplication.CalloutRelationshipAdded += new EApplication_CalloutRelationshipAddedEventHandler(VisioApplication_CalloutRelationshipAdded);
                    _VisioApplication.CalloutRelationshipDeleted += new EApplication_CalloutRelationshipDeletedEventHandler(VisioApplication_CalloutRelationshipDeleted);
                    _VisioApplication.CellChanged += new EApplication_CellChangedEventHandler(VisioApplication_CellChanged);
                    _VisioApplication.ConnectionsAdded += new EApplication_ConnectionsAddedEventHandler(VisioApplication_ConnectionsAdded);
                    _VisioApplication.ConnectionsDeleted += new EApplication_ConnectionsDeletedEventHandler(VisioApplication_ConnectionsDeleted);
                    _VisioApplication.ContainerRelationshipAdded += new EApplication_ContainerRelationshipAddedEventHandler(VisioApplication_ContainerRelationshipAdded);
                    _VisioApplication.ContainerRelationshipDeleted += new EApplication_ContainerRelationshipDeletedEventHandler(VisioApplication_ContainerRelationshipDeleted);
                    _VisioApplication.ConvertToGroupCanceled += new EApplication_ConvertToGroupCanceledEventHandler(VisioApplication_ConvertToGroupCanceled);
                    _VisioApplication.DataRecordsetAdded += new EApplication_DataRecordsetAddedEventHandler(VisioApplication_DataRecordsetAdded);
                    _VisioApplication.DataRecordsetChanged += new EApplication_DataRecordsetChangedEventHandler(VisioApplication_DataRecordsetChanged);
                    _VisioApplication.DesignModeEntered += new EApplication_DesignModeEnteredEventHandler(VisioApplication_DesignModeEntered);
                    _VisioApplication.DocumentChanged += new EApplication_DocumentChangedEventHandler(VisioApplication_DocumentChanged);
                    _VisioApplication.DocumentCloseCanceled += new EApplication_DocumentCloseCanceledEventHandler(VisioApplication_DocumentCloseCanceled);
                    _VisioApplication.DocumentCreated += new EApplication_DocumentCreatedEventHandler(VisioApplication_DocumentCreated);
                    _VisioApplication.DocumentOpened += new EApplication_DocumentOpenedEventHandler(VisioApplication_DocumentOpened);
                    _VisioApplication.DocumentSaved += new EApplication_DocumentSavedEventHandler(VisioApplication_DocumentSaved);
                    _VisioApplication.DocumentSavedAs += new EApplication_DocumentSavedAsEventHandler(VisioApplication_DocumentSavedAs);
                    _VisioApplication.EnterScope += new EApplication_EnterScopeEventHandler(VisioApplication_EnterScope);
                    _VisioApplication.ExitScope += new EApplication_ExitScopeEventHandler(VisioApplication_ExitScope);
                    _VisioApplication.FormulaChanged += new EApplication_FormulaChangedEventHandler(VisioApplication_FormulaChanged);
                    _VisioApplication.GroupCanceled += new EApplication_GroupCanceledEventHandler(VisioApplication_GroupCanceled);
                    _VisioApplication.KeyDown += new EApplication_KeyDownEventHandler(VisioApplication_KeyDown);
                    _VisioApplication.KeyPress += new EApplication_KeyPressEventHandler(VisioApplication_KeyPress);
                    _VisioApplication.KeyUp += new EApplication_KeyUpEventHandler(VisioApplication_KeyUp);
                    _VisioApplication.MarkerEvent += new EApplication_MarkerEventEventHandler(VisioApplication_MarkerEvent);
                    _VisioApplication.MasterAdded += new EApplication_MasterAddedEventHandler(VisioApplication_MasterAdded);
                    _VisioApplication.MasterChanged += new EApplication_MasterChangedEventHandler(VisioApplication_MasterChanged);
                    _VisioApplication.MasterDeleteCanceled += new EApplication_MasterDeleteCanceledEventHandler(VisioApplication_MasterDeleteCanceled);
                    _VisioApplication.MouseDown += new EApplication_MouseDownEventHandler(VisioApplication_MouseDown);
                    _VisioApplication.MouseMove += new EApplication_MouseMoveEventHandler(VisioApplication_MouseMove);
                    _VisioApplication.MouseUp += new EApplication_MouseUpEventHandler(VisioApplication_MouseUp);
                    _VisioApplication.MustFlushScopeBeginning += new EApplication_MustFlushScopeBeginningEventHandler(VisioApplication_MustFlushScopeBeginning);
                    _VisioApplication.MustFlushScopeEnded += new EApplication_MustFlushScopeEndedEventHandler(VisioApplication_MustFlushScopeEnded);
                    _VisioApplication.NoEventsPending += new EApplication_NoEventsPendingEventHandler(VisioApplication_NoEventsPending);
                    _VisioApplication.OnKeystrokeMessageForAddon += new EApplication_OnKeystrokeMessageForAddonEventHandler(VisioApplication_OnKeystrokeMessageForAddon);
                    _VisioApplication.PageAdded += new EApplication_PageAddedEventHandler(VisioApplication_PageAdded);
                    _VisioApplication.PageChanged += new EApplication_PageChangedEventHandler(VisioApplication_PageChanged);
                    _VisioApplication.PageDeleteCanceled += new EApplication_PageDeleteCanceledEventHandler(VisioApplication_PageDeleteCanceled);
                    _VisioApplication.QueryCancelConvertToGroup += new EApplication_QueryCancelConvertToGroupEventHandler(VisioApplication_QueryCancelConvertToGroup);
                    _VisioApplication.QueryCancelDocumentClose += new EApplication_QueryCancelDocumentCloseEventHandler(VisioApplication_QueryCancelDocumentClose);
                    _VisioApplication.QueryCancelGroup += new EApplication_QueryCancelGroupEventHandler(VisioApplication_QueryCancelGroup);
                    _VisioApplication.QueryCancelMasterDelete += new EApplication_QueryCancelMasterDeleteEventHandler(VisioApplication_QueryCancelMasterDelete);
                    _VisioApplication.QueryCancelPageDelete += new EApplication_QueryCancelPageDeleteEventHandler(VisioApplication_QueryCancelPageDelete);
                    _VisioApplication.QueryCancelQuit += new EApplication_QueryCancelQuitEventHandler(VisioApplication_QueryCancelQuit);
                    _VisioApplication.QueryCancelSelectionDelete += new EApplication_QueryCancelSelectionDeleteEventHandler(VisioApplication_QueryCancelSelectionDelete);
                    _VisioApplication.QueryCancelStyleDelete += new EApplication_QueryCancelStyleDeleteEventHandler(VisioApplication_QueryCancelStyleDelete);
                    _VisioApplication.QueryCancelSuspend += new EApplication_QueryCancelSuspendEventHandler(VisioApplication_QueryCancelSuspend);
                    _VisioApplication.QueryCancelSuspendEvents += new EApplication_QueryCancelSuspendEventsEventHandler(VisioApplication_QueryCancelSuspendEvents);
                    _VisioApplication.QueryCancelUngroup += new EApplication_QueryCancelUngroupEventHandler(VisioApplication_QueryCancelUngroup);
                    _VisioApplication.QueryCancelWindowClose += new EApplication_QueryCancelWindowCloseEventHandler(VisioApplication_QueryCancelWindowClose);
                    _VisioApplication.QuitCanceled += new EApplication_QuitCanceledEventHandler(VisioApplication_QuitCanceled);
                    _VisioApplication.RuleSetValidated += new EApplication_RuleSetValidatedEventHandler(VisioApplication_RuleSetValidated);
                    _VisioApplication.RunModeEntered += new EApplication_RunModeEnteredEventHandler(VisioApplication_RunModeEntered);
                    _VisioApplication.SelectionAdded += new EApplication_SelectionAddedEventHandler(VisioApplication_SelectionAdded);
                    _VisioApplication.SelectionChanged += new EApplication_SelectionChangedEventHandler(VisioApplication_SelectionChanged);
                    _VisioApplication.SelectionDeleteCanceled += new EApplication_SelectionDeleteCanceledEventHandler(VisioApplication_SelectionDeleteCanceled);
                    _VisioApplication.ShapeAdded += new EApplication_ShapeAddedEventHandler(VisioApplication_ShapeAdded);
                    _VisioApplication.ShapeChanged += new EApplication_ShapeChangedEventHandler(VisioApplication_ShapeChanged);
                    _VisioApplication.ShapeDataGraphicChanged += new EApplication_ShapeDataGraphicChangedEventHandler(VisioApplication_ShapeDataGraphicChanged);
                    _VisioApplication.ShapeExitedTextEdit += new EApplication_ShapeExitedTextEditEventHandler(VisioApplication_ShapeExitedTextEdit);
                    _VisioApplication.ShapeLinkAdded += new EApplication_ShapeLinkAddedEventHandler(VisioApplication_ShapeLinkAdded);
                    _VisioApplication.ShapeLinkDeleted += new EApplication_ShapeLinkDeletedEventHandler(VisioApplication_ShapeLinkDeleted);
                    _VisioApplication.ShapeParentChanged += new EApplication_ShapeParentChangedEventHandler(VisioApplication_ShapeParentChanged);
                    _VisioApplication.StyleAdded += new EApplication_StyleAddedEventHandler(VisioApplication_StyleAdded);
                    _VisioApplication.StyleChanged += new EApplication_StyleChangedEventHandler(VisioApplication_StyleChanged);
                    _VisioApplication.StyleDeleteCanceled += new EApplication_StyleDeleteCanceledEventHandler(VisioApplication_StyleDeleteCanceled);
                    _VisioApplication.SuspendCanceled += new EApplication_SuspendCanceledEventHandler(VisioApplication_SuspendCanceled);
                    _VisioApplication.SuspendEventsCanceled += new EApplication_SuspendEventsCanceledEventHandler(VisioApplication_SuspendEventsCanceled);
                    _VisioApplication.TextChanged += new EApplication_TextChangedEventHandler(VisioApplication_TextChanged);
                    _VisioApplication.UngroupCanceled += new EApplication_UngroupCanceledEventHandler(VisioApplication_UngroupCanceled);
                    _VisioApplication.ViewChanged += new EApplication_ViewChangedEventHandler(VisioApplication_ViewChanged);
                    _VisioApplication.WindowActivated += new EApplication_WindowActivatedEventHandler(VisioApplication_WindowActivated);
                    _VisioApplication.WindowChanged += new EApplication_WindowChangedEventHandler(VisioApplication_WindowChanged);
                    _VisioApplication.WindowCloseCanceled += new EApplication_WindowCloseCanceledEventHandler(VisioApplication_WindowCloseCanceled);
                    _VisioApplication.WindowOpened += new EApplication_WindowOpenedEventHandler(VisioApplication_WindowOpened);
                    _VisioApplication.WindowTurnedToPage += new EApplication_WindowTurnedToPageEventHandler(VisioApplication_WindowTurnedToPage);
                }
            }
        }

        #region Regular Events      

        short _count_AfterModal;
        void VisioApplication_AfterModal(Application app)
        {
            DisplayEventInWatchWindow(_count_AfterModal++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AfterRemoveHiddenInformation;
        void VisioApplication_AfterRemoveHiddenInformation(Document Doc)
        {
            DisplayEventInWatchWindow(_count_AfterRemoveHiddenInformation++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AfterResume;
        void VisioApplication_AfterResume(Application app)
        {
            DisplayEventInWatchWindow(_count_AfterResume++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AfterResumeEvents;
        void VisioApplication_AfterResumeEvents(Application app)
        {
            DisplayEventInWatchWindow(_count_AfterResumeEvents++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AppActivated;
        void VisioApplication_AppActivated(Application app)
        {
            DisplayEventInWatchWindow(_count_AppActivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AppDeactivated;
        void VisioApplication_AppDeactivated(Application app)
        {
            DisplayEventInWatchWindow(_count_AppDeactivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AppObjActivated;
        void VisioApplication_AppObjActivated(Application app)
        {
            DisplayEventInWatchWindow(_count_AppObjActivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_AppObjDeactivated;
        void VisioApplication_AppObjDeactivated(Application app)
        {
            DisplayEventInWatchWindow(_count_AppObjDeactivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeDataRecordsetDelete;
        void VisioApplication_BeforeDataRecordsetDelete(DataRecordset DataRecordset)
        {
            DisplayEventInWatchWindow(_count_BeforeDataRecordsetDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeDocumentClose;
        void VisioApplication_BeforeDocumentClose(Document Doc)
        {
            DisplayEventInWatchWindow(_count_BeforeDocumentClose++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeDocumentSave;
        void VisioApplication_BeforeDocumentSave(Document Doc)
        {
            DisplayEventInWatchWindow(_count_BeforeDocumentSave++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeDocumentSaveAs;
        void VisioApplication_BeforeDocumentSaveAs(Document Doc)
        {
            DisplayEventInWatchWindow(_count_BeforeDocumentSaveAs++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeMasterDelete;
        void VisioApplication_BeforeMasterDelete(Master Master)
        {
            DisplayEventInWatchWindow(_count_BeforeMasterDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeModal;
        void VisioApplication_BeforeModal(Application app)
        {
            DisplayEventInWatchWindow(_count_BeforeModal++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforePageDelete;
        void VisioApplication_BeforePageDelete(Page Page)
        {
            DisplayEventInWatchWindow(_count_BeforePageDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeQuit;
        void VisioApplication_BeforeQuit(Application app)
        {
            DisplayEventInWatchWindow(_count_BeforeQuit++, MethodInfo.GetCurrentMethod().Name);
        }

        internal short _count_BeforeSelectionDelete;
        void VisioApplication_BeforeSelectionDelete(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_BeforeSelectionDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeShapeDelete;
        void VisioApplication_BeforeShapeDelete(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_BeforeShapeDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeShapeTextEdit;
        void VisioApplication_BeforeShapeTextEdit(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_BeforeShapeTextEdit++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeStyleDelete;
        void VisioApplication_BeforeStyleDelete(Style Style)
        {
            DisplayEventInWatchWindow(_count_BeforeStyleDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeSuspend;
        void VisioApplication_BeforeSuspend(Application app)
        {
            DisplayEventInWatchWindow(_count_BeforeSuspend++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeSuspendEvents;
        void VisioApplication_BeforeSuspendEvents(Application app)
        {
            DisplayEventInWatchWindow(_count_BeforeSuspendEvents++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeWindowClosed;
        void VisioApplication_BeforeWindowClosed(Window Window)
        {
            DisplayEventInWatchWindow(_count_BeforeWindowClosed++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeWindowPageTurn;
        void VisioApplication_BeforeWindowPageTurn(Window Window)
        {
            DisplayEventInWatchWindow(_count_BeforeWindowPageTurn++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_BeforeWindowSelDelete;
        void VisioApplication_BeforeWindowSelDelete(Window Window)
        {
            DisplayEventInWatchWindow(_count_BeforeWindowSelDelete++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_CalloutRelationshipAdded;
        void VisioApplication_CalloutRelationshipAdded(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(_count_CalloutRelationshipAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_CalloutRelationshipDeleted;
        void VisioApplication_CalloutRelationshipDeleted(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(_count_CalloutRelationshipDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ConnectionsAdded;
        void VisioApplication_ConnectionsAdded(Connects Connects)
        {
            DisplayEventInWatchWindow(_count_ConnectionsAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ConnectionsDeleted;
        void VisioApplication_ConnectionsDeleted(Connects Connects)
        {
            DisplayEventInWatchWindow(_count_ConnectionsDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ContainerRelationshipAdded;
        void VisioApplication_ContainerRelationshipAdded(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(_count_ContainerRelationshipAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ContainerRelationshipDeleted;
        void VisioApplication_ContainerRelationshipDeleted(RelatedShapePairEvent ShapePair)
        {
            DisplayEventInWatchWindow(_count_ContainerRelationshipDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ConvertToGroupCanceled;
        void VisioApplication_ConvertToGroupCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_ConvertToGroupCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DataRecordsetAdded;
        void VisioApplication_DataRecordsetAdded(DataRecordset DataRecordset)
        {
            DisplayEventInWatchWindow(_count_DataRecordsetAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DataRecordsetChanged;
        void VisioApplication_DataRecordsetChanged(DataRecordsetChangedEvent DataRecordsetChanged)
        {
            DisplayEventInWatchWindow(_count_DataRecordsetChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DesignModeEntered;
        void VisioApplication_DesignModeEntered(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DesignModeEntered++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DocumentChanged;
        void VisioApplication_DocumentChanged(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DocumentChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DocumentCloseCanceled;
        void VisioApplication_DocumentCloseCanceled(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DocumentCloseCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DocumentCreated;
        void VisioApplication_DocumentCreated(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DocumentCreated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DocumentOpened;
        void VisioApplication_DocumentOpened(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DocumentOpened++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DocumentSaved;
        void VisioApplication_DocumentSaved(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DocumentSaved++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_DocumentSavedAs;
        void VisioApplication_DocumentSavedAs(Document Doc)
        {
            DisplayEventInWatchWindow(_count_DocumentSavedAs++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_EnterScope;
        void VisioApplication_EnterScope(Application app, int nScopeID, string bstrDescription)
        {
            DisplayEventInWatchWindow(_count_EnterScope++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ExitScope;
        void VisioApplication_ExitScope(Application app, int nScopeID, string bstrDescription, bool bErrOrCancelled)
        {
            DisplayEventInWatchWindow(_count_ExitScope++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_GroupCanceled;
        void VisioApplication_GroupCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_GroupCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_MarkerEvent;
        void VisioApplication_MarkerEvent(Application app, int SequenceNum, string ContextString)
        {

            DisplayEventInWatchWindow(_count_MarkerEvent++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_MasterAdded;
        void VisioApplication_MasterAdded(Master Master)
        {
            DisplayEventInWatchWindow(_count_MasterAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_MasterChanged;
        void VisioApplication_MasterChanged(Master Master)
        {
            DisplayEventInWatchWindow(_count_MasterChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_MasterDeleteCanceled;
        void VisioApplication_MasterDeleteCanceled(Master Master)
        {
            DisplayEventInWatchWindow(_count_MasterDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_OnKeystrokeMessageForAddon;
        bool VisioApplication_OnKeystrokeMessageForAddon(MSGWrap MSG)
        {
            DisplayEventInWatchWindow(_count_OnKeystrokeMessageForAddon++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_PageAdded;
        void VisioApplication_PageAdded(Page Page)
        {
            DisplayEventInWatchWindow(_count_PageAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_PageChanged;
        void VisioApplication_PageChanged(Page Page)
        {
            DisplayEventInWatchWindow(_count_PageChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_PageDeleteCanceled;
        void VisioApplication_PageDeleteCanceled(Page Page)
        {
            DisplayEventInWatchWindow(_count_PageDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_QueryCancelConvertToGroup;
        bool VisioApplication_QueryCancelConvertToGroup(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_QueryCancelConvertToGroup++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelDocumentClose;
        bool VisioApplication_QueryCancelDocumentClose(Document Doc)
        {
            DisplayEventInWatchWindow(_count_QueryCancelDocumentClose++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelGroup;
        bool VisioApplication_QueryCancelGroup(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_QueryCancelGroup++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelMasterDelete;
        bool VisioApplication_QueryCancelMasterDelete(Master Master)
        {
            DisplayEventInWatchWindow(_count_QueryCancelMasterDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelPageDelete;
        bool VisioApplication_QueryCancelPageDelete(Page Page)
        {
            DisplayEventInWatchWindow(_count_QueryCancelPageDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelQuit;
        bool VisioApplication_QueryCancelQuit(Application app)
        {
            DisplayEventInWatchWindow(_count_QueryCancelQuit++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelSelectionDelete;
        bool VisioApplication_QueryCancelSelectionDelete(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_QueryCancelSelectionDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelStyleDelete;
        bool VisioApplication_QueryCancelStyleDelete(Style Style)
        {
            DisplayEventInWatchWindow(_count_QueryCancelStyleDelete++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelSuspend;
        bool VisioApplication_QueryCancelSuspend(Application app)
        {
            DisplayEventInWatchWindow(_count_QueryCancelSuspend++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelSuspendEvents;
        bool VisioApplication_QueryCancelSuspendEvents(Application app)
        {
            DisplayEventInWatchWindow(_count_QueryCancelSuspendEvents++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelUngroup;
        bool VisioApplication_QueryCancelUngroup(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_QueryCancelUngroup++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QueryCancelWindowClose;
        bool VisioApplication_QueryCancelWindowClose(Window Window)
        {
            DisplayEventInWatchWindow(_count_QueryCancelWindowClose++, MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short _count_QuitCanceled;
        void VisioApplication_QuitCanceled(Application app)
        {
            DisplayEventInWatchWindow(_count_QuitCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_RuleSetValidated;
        void VisioApplication_RuleSetValidated(ValidationRuleSet RuleSet)
        {
            DisplayEventInWatchWindow(_count_RuleSetValidated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_RunModeEntered;
        void VisioApplication_RunModeEntered(Document Doc)
        {
            DisplayEventInWatchWindow(_count_RunModeEntered++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_SelectionAdded;
        void VisioApplication_SelectionAdded(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_SelectionAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_SelectionChanged;
        void VisioApplication_SelectionChanged(Window Window)
        {

            DisplayEventInWatchWindow(_count_SelectionChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_SelectionDeleteCanceled;
        void VisioApplication_SelectionDeleteCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_SelectionDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ShapeAdded;
        void VisioApplication_ShapeAdded(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_ShapeAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ShapeChanged;
        void VisioApplication_ShapeChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_ShapeChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ShapeDataGraphicChanged;
        void VisioApplication_ShapeDataGraphicChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_ShapeDataGraphicChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ShapeExitedTextEdit;
        void VisioApplication_ShapeExitedTextEdit(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_ShapeExitedTextEdit++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ShapeLinkAdded;
        void VisioApplication_ShapeLinkAdded(Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayEventInWatchWindow(_count_ShapeLinkAdded++, MethodInfo.GetCurrentMethod().Name);
        }
        short _count_ShapeLinkDeleted;
        void VisioApplication_ShapeLinkDeleted(Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayEventInWatchWindow(_count_ShapeLinkDeleted++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ShapeParentChanged;
        void VisioApplication_ShapeParentChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_ShapeParentChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_StyleAdded;
        void VisioApplication_StyleAdded(Style Style)
        {
            DisplayEventInWatchWindow(_count_StyleAdded++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_StyleChanged;
        void VisioApplication_StyleChanged(Style Style)
        {
            DisplayEventInWatchWindow(_count_StyleChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_StyleDeleteCanceled;
        void VisioApplication_StyleDeleteCanceled(Style Style)
        {
            DisplayEventInWatchWindow(_count_StyleDeleteCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_SuspendCanceled;
        void VisioApplication_SuspendCanceled(Application app)
        {
            DisplayEventInWatchWindow(_count_SuspendCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_SuspendEventsCanceled;
        void VisioApplication_SuspendEventsCanceled(Application app)
        {
            DisplayEventInWatchWindow(_count_SuspendEventsCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_TextChange;
        void VisioApplication_TextChanged(Shape Shape)
        {
            DisplayEventInWatchWindow(_count_TextChange++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_UngroupCanceled;
        void VisioApplication_UngroupCanceled(Selection Selection)
        {
            DisplayEventInWatchWindow(_count_UngroupCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_ViewChanged;
        void VisioApplication_ViewChanged(Window Window)
        {
            DisplayEventInWatchWindow(_count_ViewChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_WindowActivated;
        void VisioApplication_WindowActivated(Window Window)
        {
            DisplayEventInWatchWindow(_count_WindowActivated++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_WindowChanged;
        void VisioApplication_WindowChanged(Window Window)
        {
            DisplayEventInWatchWindow(_count_WindowChanged++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_WindowCloseCanceled;
        void VisioApplication_WindowCloseCanceled(Window Window)
        {
            DisplayEventInWatchWindow(_count_WindowCloseCanceled++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_WindowOpened;
        void VisioApplication_WindowOpened(Window Window)
        {
            DisplayEventInWatchWindow(_count_WindowOpened++, MethodInfo.GetCurrentMethod().Name);
        }

        short _count_WindowTurnedToPage;
        void VisioApplication_WindowTurnedToPage(Window Window)
        {
            DisplayEventInWatchWindow(_count_WindowTurnedToPage++, MethodInfo.GetCurrentMethod().Name);
        }

        #endregion

        #region Chatty Events - Log if DisplayChattyEvents

        short _count_CellChanged;
        void VisioApplication_CellChanged(Cell Cell)
        {
            DisplayEventInWatchWindow(_count_CellChanged++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_FormulaChanged;
        void VisioApplication_FormulaChanged(Cell Cell)
        {
            DisplayEventInWatchWindow(_count_FormulaChanged++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_KeyDown;
        void VisioApplication_KeyDown(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            DisplayEventInWatchWindow(_count_KeyDown++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_KeyPress;
        void VisioApplication_KeyPress(int KeyAscii, ref bool CancelDefault)
        {
            DisplayEventInWatchWindow(_count_KeyPress++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_KeyUp;
        void VisioApplication_KeyUp(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            DisplayEventInWatchWindow(_count_KeyUp++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_MouseDown;
        void VisioApplication_MouseDown(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            DisplayEventInWatchWindow(_count_MouseDown++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_MouseMove;
        void VisioApplication_MouseMove(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            DisplayEventInWatchWindow(_count_MouseMove++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_MouseUp;
        void VisioApplication_MouseUp(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            DisplayEventInWatchWindow(_count_MouseUp++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_MustFlushScopeBeginning;
        void VisioApplication_MustFlushScopeBeginning(Application app)
        {
            DisplayEventInWatchWindow(_count_MustFlushScopeBeginning++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_MustFlushScopeEnded;
        void VisioApplication_MustFlushScopeEnded(Application app)
        {
            DisplayEventInWatchWindow(_count_MustFlushScopeEnded++, MethodInfo.GetCurrentMethod().Name, true);
        }

        short _count_NoEventsPending;
        void VisioApplication_NoEventsPending(Application app)
        {
            DisplayEventInWatchWindow(_count_NoEventsPending++, MethodInfo.GetCurrentMethod().Name, true); ;
        }

        #endregion Chatty Events

        internal void DisplayEventInWatchWindow(short i, string outputLine, Boolean isChattyEvent = false)
        {
            if (Common.DisplayEvents)
            {
                if (isChattyEvent && !Common.DisplayChattyEvents) return;

                Common.WriteToWatchWindow($"{outputLine}:{i}");
            }
        }
    }
}
