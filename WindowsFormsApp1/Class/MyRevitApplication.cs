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
        private ElementId _lastClickedElementId = null;
        private static DoubleClickTracker _doubleClickTracker;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "I.L.U.S.H.A.";

                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException ex)
                {
                    TaskDialog.Show("Title", ex.Message);
                }

                string panelName = "Intelligent Library for Unified Smart Handling Automation";
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

                // Инициализируем трекер двойного клика
                _doubleClickTracker = new DoubleClickTracker();
                
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
            try
            {
                // Отписываемся от событий
                application.ControlledApplication.DocumentOpened -= OnDocumentOpened;
                application.Idling -= OnIdling;
                
                // Останавливаем трекер двойного клика
                DoubleClickTracker.Stop();
                
                // Освобождаем ресурсы
                _doubleClickTracker?.Dispose();
                _doubleClickTracker = null;
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Логируем ошибку, но не показываем диалог при закрытии
                System.Diagnostics.Debug.WriteLine($"Ошибка при закрытии приложения: {ex.Message}");
                return Result.Failed;
            }
        }

        private void OnDocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e)
        {
            try
            {
                DoubleClickTracker.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка при открытии документа: {ex.Message}");
            }
        }

        private void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            try
            {
                // Обрабатываем отложенные сохранения
                SpaceForm.ProcessPendingSaves();

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
                
                // Сбрасываем флаг активности Revit при каждом событии Idling
                DoubleClickTracker.ResetRevitActiveFlag();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка в OnIdling: {ex.Message}");
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