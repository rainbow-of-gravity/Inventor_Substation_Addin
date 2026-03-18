using Inventor;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.Generic;
using System.Linq;
using Drawing = System.Drawing;


namespace SelectionInfo2
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("25119a9c-3557-4ed2-9bec-e184a99835f3")]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private ApplicationEvents applicationEvents;
        private string ClientId = "{25119A9C-3557-4ED2-9BEC-E184A99835F3}";

        // Inventor application object.
        private Inventor.Application inventor;

        private DockableWindow selectionInfoWnd;
        private DataGridView selectionPropertyGrid;
        private UserInputEvents userInputEvents;
        ButtonDefinition goUpCmd;
        ButtonDefinition goDownCmd;
        private bool _suppressSelectionEvents;
        private const string CategoryName = "000 SectionInfo Navigation";

        public StandardAddInServer()
        {
        }

        /// <summary>
        /// Gets or sets the selected object for display properties.
        /// </summary>
        /// <value>
        /// The selected object.
        /// </value>
        private object _selectedObject;
        private object SelectedObject
        {
            get => _selectedObject;
            set
            {
                _selectedObject = value;
                selectionPropertyGrid.Rows.Clear();

                if (value == null) return;

                // Get all iProperties dynamically
                if (value is DocumentInfo docInfo)
                {
                    foreach (var kvp in docInfo.AlliProperties)
                    {
                        //string propertyName = kvp.Key;
                        string propertyName = kvp.Key.Split('/').Last();

                        selectionPropertyGrid.Rows.Add(propertyName, kvp.Value?.ToString());
                    }
                }
                else
                {
                    // Fallback: use reflection for other entity types
                    foreach (var prop in value.GetType().GetProperties())
                    {
                        try
                        {
                            selectionPropertyGrid.Rows.Add(prop.Name, prop.GetValue(value)?.ToString());
                        }
                        catch { }
                    }
                }
            }
        }

        /// <summary>
        /// Shows the ActiveDocument in the propertyGrid.
        /// </summary>
        private void ShowActiveDocument()
        {
            var entity = new DocumentInfo(inventor.ActiveDocument);
            SelectedObject = entity;
        }

        /// <summary>
        /// Clear the propertyGrid.
        /// </summary>
        private void ClearPalette()
        {
            SelectedObject = null;
        }

        /// <summary>
        /// Shows the sent object in the propertyGrid, provided it is a valid [document type] entity.
        /// </summary>
        /// <param name="entity">Selected object.</param>
        private void ShowEntity(object entity)
        {
            entity = SelectionInfoSelector.GetSelectionInfo(entity);
            if (entity != null)
            {
                SelectedObject = entity;
            }

            {
                //Leave unchanged
            }
        }

        /// <summary>
        /// Show the selected objects (if any) in the propertyGrid.
        /// </summary>
        private void ShowSelected()
        {
            
            var selectSet = inventor.ActiveDocument.SelectSet;

            switch (selectSet.Count)
            {

                case 0:
                    {
                        ShowActiveDocument();
                        break;
                    }

                case 1:
                    {
                        var entity = selectSet[1];
                        ShowEntity(entity);
                        break;
                    }

                default:
                    {
                        ClearPalette();
                        break;
                    }
            }

        }

        private List<ComponentOccurrence> navigationStack = new List<ComponentOccurrence>();
        private void GoUpHierarchy()
        {

            _suppressSelectionEvents = true;
            try
            {
                var selectSet = inventor.ActiveDocument.SelectSet;
                if (selectSet.Count != 1)
                    return;
                
                ComponentOccurrence occ = (ComponentOccurrence)selectSet[1];
                if (!navigationStack.Any(o => o.Name == occ.Name))
                {
                    navigationStack.Add(occ);
                }
                if (occ.ParentOccurrence != null)
                {
                    selectSet.Clear();
                    selectSet.Select(occ.ParentOccurrence);
                }
            }
            finally
            {
                _suppressSelectionEvents = false;
            }
            ShowSelected();
        }

        
        private void GoDownHierarchy()
        {
            _suppressSelectionEvents = true;


            try
            {
                var selectSet = inventor.ActiveDocument.SelectSet;
                if (selectSet.Count != 1) return;
                if (navigationStack.Count <= 0) { return; }

                ComponentOccurrence lastOcc = navigationStack.Last();
                navigationStack.RemoveAt(navigationStack.Count - 1);
                selectSet.Clear();
                selectSet.Select(lastOcc);
            }
            finally
            {
                _suppressSelectionEvents = false;
            }
            ShowSelected();
        }


        /// <summary>
        /// Handle when a document is activated [opened].
        /// </summary>
        private void ApplicationEvents_OnActivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter,
            NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            ShowSelected();
        }


        /// <summary>
        ///Handle when the current document is deactivated [closed].
        /// </summary>

        private void ApplicationEvents_OnDeactivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter,
            NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            if (BeforeOrAfter != EventTimingEnum.kBefore)
                return;

            ClearPalette();
        }




        /// <summary>
        /// Handle user-selections, choose best course based on how many objects are selected.
        /// </summary>
        /// <param name="JustSelectedEntities"></param>
        /// <param name="MoreSelectedEntities"></param>
        /// <param name="SelectionDevice"></param>
        /// <param name="ModelPosition"></param>
        /// <param name="ViewPosition"></param>
        /// <param name="View"></param>
        private void UserInputEvents_OnSelect(ObjectsEnumerator JustSelectedEntities,
            ref ObjectCollection MoreSelectedEntities, SelectionDeviceEnum SelectionDevice, Point ModelPosition,
            Point2d ViewPosition, Inventor.View View)
        {
            if (_suppressSelectionEvents) return;
            navigationStack.Clear();
            ShowSelected();
        }

        /// <summary>
        /// Handle when the active document changes.
        /// </summary>
        private void ApplicationEvents_OnDocumentChange(_Document DocumentObject,
            EventTimingEnum BeforeOrAfter, CommandTypesEnum ReasonsForChange,
            NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
            if (BeforeOrAfter != EventTimingEnum.kBefore)
                return;

            if (inventor.ActiveDocumentType == DocumentTypeEnum.kNoDocument)
            {
                ClearPalette();
            }
            else
            {
                ShowSelected();
            }

        }

        /// <summary>
        /// Handle when user deselects an item.
        /// </summary>
        private void UserInputEvents_OnUnSelect(ObjectsEnumerator UnSelectedEntities, SelectionDeviceEnum SelectionDevice, Point ModelPosition, Point2d ViewPosition, Inventor.View View)
        {
            if (_suppressSelectionEvents) return;
            ShowSelected();
        }

        #region ApplicationAddInServer Members

        /// <summary>
        ///This method is called by Inventor when it loads the AddIn.
        /// The AddInSiteObject provides access to the Inventor Application object.
        /// The FirstTime flag indicates if the AddIn is loaded for the first time.
        /// </summary>
        /// <param name="addInSiteObject"></param>
        /// <param name="firstTime"></param>
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            Debug.WriteLine("Activate!");

            // Initialize AddIn members.
            inventor = addInSiteObject.Application;

            //Create dockable window
            selectionInfoWnd = inventor.UserInterfaceManager.DockableWindows.Add(ClientId,
                "SelectionInfo.StandardAddInServer.selectionInfoWnd", "Selection2");
            selectionInfoWnd.ShowVisibilityCheckBox = true;

            // Create DataGridView control
            selectionPropertyGrid = new DataGridView();
            selectionPropertyGrid.Columns.Add("Property", "Property");
            selectionPropertyGrid.Columns.Add("Value", "Value");
            selectionPropertyGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            selectionPropertyGrid.ReadOnly = true;
            selectionPropertyGrid.AllowUserToAddRows = false;
            selectionPropertyGrid.RowHeadersVisible = false;
            selectionPropertyGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            selectionPropertyGrid.BackgroundColor = Drawing.Color.White;


            //Add propertyGrid to dockable window
            selectionInfoWnd.AddChild(selectionPropertyGrid.Handle);

            //Setup event handlers
            userInputEvents = inventor.CommandManager.UserInputEvents;
            userInputEvents.OnSelect += UserInputEvents_OnSelect;
            userInputEvents.OnUnSelect += UserInputEvents_OnUnSelect;

            applicationEvents = inventor.ApplicationEvents;
            applicationEvents.OnActivateDocument += ApplicationEvents_OnActivateDocument;
            applicationEvents.OnDeactivateDocument += ApplicationEvents_OnDeactivateDocument;
            applicationEvents.OnDocumentChange += ApplicationEvents_OnDocumentChange;

            CommandCategory cmdCat = GetOrCreateCategory(CategoryName);
            goUpDefinition(firstTime, cmdCat);
            goDownDefinition(firstTime, cmdCat);



        }

        private void goUpDefinition(bool firstTime, CommandCategory cmdCat)
        {
            try
            {
                goUpCmd = inventor.CommandManager.ControlDefinitions["GoUpHierarchy"] as ButtonDefinition;
                Debug.WriteLine("GoUpHierarchy already exists, reusing.");
            }
            catch
            {
                try
                {
                    goUpCmd = inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                        "↑ Go Up Hierarchy",
                        "GoUpHierarchy",
                        CommandTypesEnum.kEditMaskCmdType,
                        ClientId,
                        "Go Up Hierarchy",
                        "Navigate up to parent component",
                        GetIconPath("up_16.png"),
                        GetIconPath("up_32.png"),
                        ButtonDisplayEnum.kDisplayTextInLearningMode);
                    Debug.WriteLine("GoUpHierarchy created successfully.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to create GoUpHierarchy: {ex.Message}");
                }
            }

            if (goUpCmd == null)
                throw new InvalidOperationException("AddButtonDefinition did not return a ButtonDefinition");

            cmdCat.Add(goUpCmd);
            AddToRibbon(firstTime, goUpCmd);
            goUpCmd.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(goUpCmd_OnExecute);
        }
        void goUpCmd_OnExecute(NameValueMap Context)
        {
            GoUpHierarchy();
        }
        private void goDownDefinition(bool firstTime, CommandCategory cmdCat)
        {
            try
            {
                goDownCmd = inventor.CommandManager.ControlDefinitions["GoDownHierarchy"] as ButtonDefinition;
                Debug.WriteLine("GoDownHierarchy already exists, reusing.");
            }
            catch
            {
                try
                {
                    goDownCmd = inventor.CommandManager.ControlDefinitions.AddButtonDefinition(
                        "↓ Go Down Hierarchy",
                        "GoDownHierarchy",
                        CommandTypesEnum.kEditMaskCmdType,
                        ClientId,
                        "Go Down Hierarchy",
                        "Navigate down to child component",
                        GetIconPath("down_16.png"),
                        GetIconPath("down_32.png"),
                        ButtonDisplayEnum.kDisplayTextInLearningMode);
                    Debug.WriteLine("GoDownHierarchy created successfully.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to create GoDownHierarchy: {ex.Message}");
                }
            }

            if (goDownCmd == null)
                throw new InvalidOperationException("AddButtonDefinition did not return a ButtonDefinition");

            cmdCat.Add(goDownCmd);
            AddToRibbon(firstTime, goDownCmd);
            goDownCmd.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(goDownCmd_OnExecute);
        }
        void goDownCmd_OnExecute(NameValueMap Context)
        {
            //System.Windows.Forms.MessageBox.Show("GoUpHierarchyHandler", "SectionInfo2");
            //MessageBox.Show("GoUpHierarchyHandler");
            GoDownHierarchy();
        }
        private CommandCategory GetOrCreateCategory(string name)
        {
            try
            {
                return inventor.CommandManager.CommandCategories[name];
            }
            catch
            {
                return inventor.CommandManager.CommandCategories.Add(name, ClientId);
            }
        }
        private stdole.IPictureDisp GetIconPath(string fileName)
        {
            string resourceName = $"SelectionInfo2.Icons.{fileName}";
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    Debug.WriteLine($"Resource not found: {resourceName}");
                    // List all available resources to help diagnose
                    foreach (var name in assembly.GetManifestResourceNames())
                        Debug.WriteLine($"Available resource: {name}");
                    return null;
                }

                var bitmap = new Drawing.Bitmap(stream);
                return PictureConverter.ImageToIPictureDisp(bitmap);
            }
        }
        internal class PictureConverter : System.Windows.Forms.AxHost
        {
            private PictureConverter() : base("59EE46BA-677D-4D20-BF10-8D8067CB8B33") { }

            public static stdole.IPictureDisp ImageToIPictureDisp(Drawing.Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        private void AddToRibbon(bool firstTime, ButtonDefinition buttonCommand)
        {
            try
            {
                if (inventor.UserInterfaceManager.InterfaceStyle == InterfaceStyleEnum.kRibbonInterface)
                {
                    Ribbon ribbon = inventor.UserInterfaceManager.Ribbons["Assembly"];
                    RibbonTab tab = ribbon.RibbonTabs["id_TabAssemble"];

                    RibbonPanel panel;
                    try
                    {
                        panel = tab.RibbonPanels["SelectionInfoPanel"];
                    }
                    catch
                    {
                        panel = tab.RibbonPanels.Add("Section Info", "SelectionInfoPanel", ClientId, "", false);
                    }

                    // Check if button already exists in panel before adding
                    try
                    {
                        panel.CommandControls.AddButton(buttonCommand, true, true, "", false);
                    }
                    catch
                    {
                        Debug.WriteLine($"Button {buttonCommand.DisplayName} already exists in panel");
                    }
                }
                else
                {
                    CommandBar oCommandBar = inventor.UserInterfaceManager.CommandBars["PMxPartFeatureCmdBar"];
                    oCommandBar.Controls.AddButton(buttonCommand);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"AddToRibbon failed: {ex.Message}");
            }
        }

        /// <summary>
        /// This method is called by Inventor when the AddIn in unloaded.
        /// The AddIn will be unloaded either manually by the user or
        /// when the Inventor session is terminated.
        /// </summary>
        public void Deactivate()
        {
            //Remove event handlers
            userInputEvents.OnSelect -= UserInputEvents_OnSelect;
            userInputEvents.OnUnSelect -= UserInputEvents_OnUnSelect;
            applicationEvents.OnActivateDocument -= ApplicationEvents_OnActivateDocument;
            applicationEvents.OnDeactivateDocument -= ApplicationEvents_OnDeactivateDocument;
            applicationEvents.OnDocumentChange -= ApplicationEvents_OnDocumentChange;

            //Cleanup selectionPropertyGrid
            SelectedObject = null;
            selectionPropertyGrid = null;

            // Release objects.
            applicationEvents = null;
            userInputEvents = null;

            Marshal.ReleaseComObject(inventor);
            inventor = null;

            Marshal.ReleaseComObject(goUpCmd);
            goUpCmd = null;
            Marshal.ReleaseComObject(goDownCmd);
            goDownCmd = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        #endregion
    }
}