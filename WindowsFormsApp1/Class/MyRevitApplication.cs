using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace WindowsFormsApp1.Class
{
    public class MyRevitApplication : IExternalApplication
    {
        private readonly ICollection<ElementId> _previousSelection = new List<ElementId>();
        //private int _idleCounter = 0;
        //private const int IdleFrequency = 10;
        private ElementId _lastClickedElementId = null;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "Для Илюши";

                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException ex)
                {
                    TaskDialog.Show("Title", ex.Message);
                }

                string panelName = "Название";
                RibbonPanel panel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == panelName) ?? application.CreateRibbonPanel(tabName, panelName);

                PushButtonData buttonData = new PushButtonData(
                    "UpdateSpaceParameters",
                    "Инициализировать или рассчитать",
                    Assembly.GetExecutingAssembly().Location,
                    "WindowsFormsApp1.Class.UpdateSpaceParameters");

                PushButton pushButton = panel.AddItem(buttonData) as PushButton;
                pushButton.ToolTip = "Секретный текст :D";

                application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                application.Idling += OnIdling;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            application.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            application.Idling -= OnIdling;
            DoubleClickTracker.Stop();
            return Result.Succeeded;
        }

        private void OnDocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e)
        {
            DoubleClickTracker.Start();
        }

        private void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            //_idleCounter++;
            //if (_idleCounter < IdleFrequency)
            //{
            //    return;
            //}
            //_idleCounter = 0;

            UIApplication uiapp = sender as UIApplication;
            UIDocument uidoc = uiapp?.ActiveUIDocument;

            if (uidoc != null)
            {
                Selection selection = uidoc.Selection;
                ICollection<ElementId> currentSelection = selection.GetElementIds();
                GlobalSettings.CurrentSelection = currentSelection;

                if (currentSelection.Count == 1)
                {
                    ElementId currentElementId = currentSelection.First();

                    if (_lastClickedElementId != null && _lastClickedElementId == currentElementId && GlobalSettings.IsDoubleClick)
                    {
                        Element element = uidoc.Document.GetElement(currentElementId);
                        if (element.Category != null)
                        {
                            if (element.Category.Id.Value == (int)BuiltInCategory.OST_MEPSpaces)
                            {
                                try
                                {
                                    CreateSpaceForm(element, uiapp);
                                    GlobalSettings.IsDoubleClick = false;
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Title", ex.Message);
                                }
                            }

                            if (element.Category.Id.Value == (int)BuiltInCategory.OST_HVAC_Zones && element is Zone zone)
                            {
                                try
                                {
                                    CreateZoneForm(zone, uiapp);
                                    GlobalSettings.IsDoubleClick = false;
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Title", ex.Message);
                                }
                            }
                        }
                    }
                    else
                    {
                        GlobalSettings.IsDoubleClick = false;
                    }

                    _lastClickedElementId = currentElementId;
                }
                else if (currentSelection.Count == 0)
                {
                    _previousSelection.Clear();
                }
            }
        }

        public static void CreateSpaceForm(Element space, UIApplication uiapp)
        {
            SpaceForm window = new SpaceForm(space);
            window.CenterToRevitWindow(uiapp);
            Application.Run(window);
        }

        public static void CreateZoneForm(Zone space, UIApplication uiapp)
        {
            ZoneForm window = new ZoneForm(space, uiapp);
            window.CenterToRevitWindow(uiapp);
            Application.Run(window);
        }
    }
}