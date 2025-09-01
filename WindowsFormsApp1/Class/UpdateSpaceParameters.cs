using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Globalization;

namespace WindowsFormsApp1.Class
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UpdateSpaceParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                GlobalSettings.CommandData = commandData;

                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ICollection<Element> spaces = collector.OfCategory(BuiltInCategory.OST_MEPSpaces).ToElements();

                using (Transaction t = new Transaction(doc, "Update Space Parameters"))
                {
                    t.Start();

                    foreach (Element space in spaces)
                    {
                        ProcessSpaceParameters(space, commandData, doc);
                    }

                    t.Commit();
                }

                uidoc.RefreshActiveView();

                //DoubleClickTracker.Start();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Title", ex.Message);
                return Result.Failed;
            }
        }

        private void ProcessSpaceParameters(Element space, ExternalCommandData commandData, Document doc)
        {
            var paramsToCheck = new[] { "Вытяжка", "Приток", "Кратность", "Человеки", "П_КолЧел", "В_КолЧел", "Норма", "В_Крат", "П_Крат" };

            foreach (var paramName in paramsToCheck)
            {
                Parameter param = space.LookupParameter(paramName);
                if (param == null) AddSharedParameter(doc, space, paramName, commandData);
            }
        }

        public static void Calculate(Element space)
        {
            try
            {
                ExternalCommandData commandData = GlobalSettings.CommandData;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                using (Transaction t = new Transaction(doc, "Update Space Parameters"))
                {
                    t.Start();

                    double exhaustResult = GetParameterValue(space, "Вытяжка");
                    double supplyResult = GetParameterValue(space, "Приток");
                    
                    double volume = GetVolume(space);
                    double v_krat = GetParameterValue(space, "В_Крат");
                    double p_krat = GetParameterValue(space, "П_Крат");
                    double v_kolChel = GetParam(space, "В_КолЧел");
                    double p_kolChel = GetParam(space, "П_КолЧел");
                    int norm = GetParam(space, "Норма");

                    if (volume != 0)
                    {
                        if (v_krat != 0 && exhaustResult == 0)
                        {
                            exhaustResult = RoundToNearestMultipleOfFive(volume * v_krat);
                            CalculateResult(space, "Вытяжка", exhaustResult);
                        }

                        if (p_krat != 0 && supplyResult == 0)
                        {
                            supplyResult = RoundToNearestMultipleOfFive(volume * p_krat);
                            CalculateResult(space, "Приток", supplyResult);
                        }
                    }
                    
                    if (norm != 0)
                    {
                        if (v_kolChel != 0 && exhaustResult == 0)
                        {
                            exhaustResult = RoundToNearestMultipleOfFive(norm * v_kolChel);
                            CalculateResult(space, "Вытяжка", exhaustResult);
                        }

                        if (p_kolChel != 0 && supplyResult == 0)
                        {
                            supplyResult = RoundToNearestMultipleOfFive(norm * p_kolChel);
                            CalculateResult(space, "Приток", supplyResult);
                        }
                    }

                    t.Commit();
                }
                
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Title", ex.Message);
            }
        }

        private static void CalculateResult(Element space, string param, double data)
        {
            SetParameterValue(space, param, data);
        }

        static int RoundToNearestMultipleOfFive(double value)
        {
            return (int)(Math.Round(value / 5.0) * 5);
        }

        private static double GetVolume(Element space)
        {
            Parameter volumeParam = space.LookupParameter("Объем");
            return volumeParam != null && volumeParam.HasValue ? UnitUtils.ConvertFromInternalUnits(volumeParam.AsDouble(), UnitTypeId.CubicMeters) : 0;
        }

        private static int GetParam(Element space, string key)
        {
            Parameter normParam = space.LookupParameter(key);
            if (normParam != null && normParam.HasValue)
            {
                int normParamInt = (int)double.Parse(normParam.AsValueString().Replace(" м³", ""), CultureInfo.InvariantCulture);
                return normParamInt;
            }
            return 40; // Значение по умолчанию - 40
        }

        private static double GetParameterValue(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue)
            {
                return param.StorageType == StorageType.Double ? UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.CubicMeters) : param.AsDouble();
            }
            return 0;
        }

        private static void SetParameterValue(Element space, string paramName, double value)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null)
            {
                if (param.StorageType == StorageType.Double)
                {
                    param.Set(UnitUtils.ConvertToInternalUnits(value, UnitTypeId.CubicMeters));
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    param.Set((int)value);
                }
            }
        }

        //private static bool GetBooleanParameterValue(Element space, string paramName)
        //{
        //    Parameter param = space.LookupParameter(paramName);
        //    if (param != null && param.HasValue && param.StorageType == StorageType.Integer)
        //    {
        //        return param.AsInteger() == 1;
        //    }
        //    return false;
        //}

        private static void AddSharedParameter(Document doc, Element element, string paramName, ExternalCommandData commandData)
        {
            Autodesk.Revit.ApplicationServices.Application app = commandData?.Application?.Application;
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app), "Не удалось получить доступ к приложению Revit.");
            }
            string sharedParameterFile = Configuration.GetLogFilePath("FOP_FILE_PATH");
            app.SharedParametersFilename = sharedParameterFile;

            DefinitionFile defFile = app.OpenSharedParameterFile() ?? throw new Exception("Файл общих параметров не найден. Путь: " + sharedParameterFile);
            Definition definition = null;
            foreach (DefinitionGroup group in defFile.Groups)
            {
                definition = group.Definitions?.get_Item(paramName);
                if (definition != null)
                    break;
            }

            if (definition == null)
            {
                throw new Exception("Определение параметра '" + paramName + "' не найдено в файле общих параметров.");
            }

            CategorySet categories = new CategorySet();
            categories.Insert(element.Category);

            InstanceBinding binding = app.Create.NewInstanceBinding(categories);
            BindingMap bindingMap = doc.ParameterBindings;

            bindingMap.Insert(definition, binding, GroupTypeId.Data);
        }
    }
}
