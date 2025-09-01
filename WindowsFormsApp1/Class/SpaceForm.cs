using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;

namespace WindowsFormsApp1.Class
{
    public class SpaceForm : System.Windows.Forms.Form
    {
        private readonly System.Windows.Forms.TextBox txtVentilation;
        private readonly System.Windows.Forms.TextBox txtIntake;
        //private readonly CheckBox chkMultiplicity;
        //private readonly CheckBox chkPeople;
        private readonly System.Windows.Forms.TextBox txtP_KolPeople;
        private readonly System.Windows.Forms.TextBox txtV_KolPeople;
        private readonly System.Windows.Forms.TextBox txtStandard;
        private readonly System.Windows.Forms.TextBox txtV_Multiplicity;
        private readonly System.Windows.Forms.TextBox txtP_Multiplicity;
        private readonly System.Windows.Forms.TextBox txtNumber;
        private readonly System.Windows.Forms.TextBox txtName;
        private readonly System.Windows.Forms.TextBox txtArea;
        private readonly System.Windows.Forms.TextBox txtOffset;
        private readonly System.Windows.Forms.TextBox txtVolume;
        private readonly System.Windows.Forms.TextBox txtCategory;
        private readonly System.Windows.Forms.TextBox txtTemperature;
        private readonly System.Windows.Forms.TextBox txtCleanliness;
        private readonly System.Windows.Forms.TextBox txtHeatLoss;
        private readonly Element _space;

        public SpaceForm(Element space)
        {
            _space = space;
            Text = "Space Parameters";
            Size = new System.Drawing.Size(500, 800);
            StartPosition = FormStartPosition.Manual;

            Deactivate += (s, e) => Close();

            FormClosing += (s, e) =>
            {
                SaveSpaceParameters();
                UpdateSpaceParameters.Calculate(space);
            };

            KeyPreview = true;
            KeyDown += new KeyEventHandler(Form1_KeyDown);

            var otherParametersGroup = CreateGroupBox("Остальные параметры", 10, 10, 450, 300);
            txtNumber = CreateTextBox("Номер", 40, 30, otherParametersGroup);
            txtName = CreateTextBox("Имя", 40, 60, otherParametersGroup);
            txtArea = CreateTextBox("Площадь", 40, 90, otherParametersGroup);
            txtOffset = CreateTextBox("Смещение сверху", 40, 120, otherParametersGroup);
            txtVolume = CreateTextBox("Объем", 40, 150, otherParametersGroup);
            txtCategory = CreateTextBox("ADSK_Категория помещения", 40, 180, otherParametersGroup);
            txtTemperature = CreateTextBox("ADSK_Температура в помещении", 40, 210, otherParametersGroup);
            txtCleanliness = CreateTextBox("Категория_помещения_по_чистоте", 40, 240, otherParametersGroup);
            txtHeatLoss = CreateTextBox("Теплопотери", 40, 270, otherParametersGroup);

            var multiplicityGroup = CreateGroupBox("Кратность", 10, 310, 450, 120);
            //chkMultiplicity = CreateCheckBox("Кратность", 40, 30, multiplicityGroup);
            txtP_Multiplicity = CreateTextBox("Приток кратность", 40, 30, multiplicityGroup);
            txtV_Multiplicity = CreateTextBox("Вытяжка кратность", 40, 60, multiplicityGroup);

            var peopleGroup = CreateGroupBox("Количество человек", 10, 440, 450, 180);
            //chkPeople = CreateCheckBox("Кол-во человек", 40, 30, peopleGroup);
            txtStandard = CreateTextBox("Норма", 40, 30, peopleGroup);
            txtP_KolPeople = CreateTextBox("Приток кол-во человек", 40, 60, peopleGroup);
            txtV_KolPeople = CreateTextBox("Вытяжка кол-во человек", 40, 90, peopleGroup);

            var ventilationGroup = CreateGroupBox("Вентиляция", 10, 630, 450, 100);
            txtVentilation = CreateTextBox("Приток", 40, 30, ventilationGroup);
            txtIntake = CreateTextBox("Вытяжка", 40, 60, ventilationGroup);

            LoadSpaceParameters(space);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void LoadSpaceParameters(Element space)
        {
            txtVentilation.Text = GetParameterValueInCubicMeters(space, "Приток").ToString();
            txtIntake.Text = GetParameterValueInCubicMeters(space, "Вытяжка").ToString();
            //chkMultiplicity.Checked = GetBooleanParameterValue(space, "Кратность");
            txtP_Multiplicity.Text = GetParameterValueInCubicMeters(space, "П_Крат").ToString();
            txtV_Multiplicity.Text = GetParameterValueInCubicMeters(space, "В_Крат").ToString();
            //chkPeople.Checked = GetBooleanParameterValue(space, "Человеки");
            txtStandard.Text = GetParameterValueInCubicMeters(space, "Норма").ToString();
            txtP_KolPeople.Text = GetParameterValueInCubicMeters(space, "П_КолЧел").ToString();
            txtV_KolPeople.Text = GetParameterValueInCubicMeters(space, "В_КолЧел").ToString();
            txtNumber.Text = GetStringParameterValue(space, "Номер");
            txtName.Text = GetStringParameterValue(space, "Имя");
            txtArea.Text = GetParameterValueInSquareMeters(space, "Площадь").ToString("F2");
            txtOffset.Text = GetParameterValueInMillimeters(space, "Смещение сверху").ToString("F2");
            txtVolume.Text = GetParameterValueInCubicMeters(space, "Объем").ToString("F2");
            txtCategory.Text = GetStringParameterValue(space, "ADSK_Категория помещения");
            txtTemperature.Text = GetParameterValueInCelsius(space, "ADSK_Температура в помещении").ToString("F2");
            txtCleanliness.Text = GetStringParameterValue(space, "Категория_помещения_по_чистоте");
            txtHeatLoss.Text = GetStringParameterValue(space, "Теплопотери");
        }

        private void SaveSpaceParameters()
        {
            using (Transaction transaction = new Transaction(_space.Document, "Update Space Parameters"))
            {
                transaction.Start();
                SetParameterValue(_space, "Приток", txtVentilation.Text);
                SetParameterValue(_space, "Вытяжка", txtIntake.Text);
                //SetParameterValue(_space, "Кратность", chkMultiplicity.Checked ? "1" : "0");
                SetParameterValue(_space, "П_Крат", txtP_Multiplicity.Text);
                SetParameterValue(_space, "В_Крат", txtV_Multiplicity.Text);
                //SetParameterValue(_space, "Человеки", chkPeople.Checked ? "1" : "0");
                SetParameterValue(_space, "Норма", txtStandard.Text);
                SetParameterValue(_space, "П_КолЧел", txtP_KolPeople.Text);
                SetParameterValue(_space, "В_КолЧел", txtV_KolPeople.Text);
                SetParameterValue(_space, "Номер", txtNumber.Text);
                SetParameterValue(_space, "Имя", txtName.Text);
                SetParameterValue(_space, "Смещение сверху", txtOffset.Text, UnitTypeId.Millimeters);
                SetParameterValue(_space, "ADSK_Категория помещения", txtCategory.Text);
                SetParameterValue(_space, "ADSK_Температура в помещении", txtTemperature.Text, UnitTypeId.Celsius);
                SetParameterValue(_space, "Категория_помещения_по_чистоте", txtCleanliness.Text);
                SetParameterValue(_space, "Теплопотери", txtHeatLoss.Text);
                transaction.Commit();
            }
        }

        private System.Windows.Forms.TextBox CreateTextBox(string label, int x, int y, System.Windows.Forms.Control parent)
        {
            var lbl = new Label
            {
                Text = label,
                Location = new System.Drawing.Point(x, y),
                AutoSize = true
            };
            parent.Controls.Add(lbl);

            var txt = new System.Windows.Forms.TextBox
            {
                Location = new System.Drawing.Point(x + 200, y),
                Width = 150
            };
            parent.Controls.Add(txt);

            return txt;
        }

        //private CheckBox CreateCheckBox(string label, int x, int y, System.Windows.Forms.Control parent)
        //{
        //    var lbl = new Label
        //    {
        //        Text = label,
        //        Location = new System.Drawing.Point(x, y),
        //        AutoSize = true
        //    };
        //    parent.Controls.Add(lbl);

        //    var chk = new CheckBox
        //    {
        //        Location = new System.Drawing.Point(x + 200, y),
        //        AutoSize = true
        //    };
        //    parent.Controls.Add(chk);

        //    return chk;
        //}

        private GroupBox CreateGroupBox(string text, int x, int y, int width, int height)
        {
            var groupBox = new GroupBox
            {
                Text = text,
                Location = new System.Drawing.Point(x, y),
                Size = new System.Drawing.Size(width, height)
            };
            Controls.Add(groupBox);
            return groupBox;
        }

        private static double GetParameterValueInCelsius(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
            {
                return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Celsius);
            }
            return 0;
        }

        private static double GetParameterValueInSquareMeters(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
            {
                return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.SquareMeters);
            }
            return 0;
        }

        private static double GetParameterValueInMillimeters(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
            {
                return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
            }
            return 0;
        }

        private static double GetParameterValueInCubicMeters(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
            {
                return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.CubicMeters);
            }
            return 0;
        }

        private static string GetStringParameterValue(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.String)
            {
                return param.AsString();
            }
            return string.Empty;
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

        private static void SetParameterValue(Element space, string paramName, string value, ForgeTypeId unitTypeId = null)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.IsReadOnly == false)
            {
                if (param.StorageType == StorageType.Double)
                {
                    if (double.TryParse(value, out double doubleValue))
                    {
                        param.Set(UnitUtils.ConvertToInternalUnits(doubleValue, unitTypeId ?? UnitTypeId.CubicMeters));
                    }
                    else
                    {
                        param.Set(0);
                    }
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    if (int.TryParse(value, out int intValue))
                    {
                        param.Set(intValue);
                    }
                    else
                    {
                        param.Set(0);
                    }
                }
                else if (param.StorageType == StorageType.String)
                {
                    param.Set(string.IsNullOrEmpty(value) ? "0" : value);
                }
            }
        }

        public void CenterToRevitWindow(UIApplication uiapp)
        {
            // Получаем дескриптор главного окна Revit
            IntPtr revitHandle = uiapp.MainWindowHandle;

            // Получаем размеры окна Revit
            GetWindowRect(revitHandle, out RECT revitRect);

            // Определяем центральные координаты окна Revit
            int revitCenterX = (revitRect.Left + revitRect.Right) / 2;
            int revitCenterY = (revitRect.Top + revitRect.Bottom) / 2;

            // Устанавливаем позицию формы по центру Revit
            Left = revitCenterX - (Width / 2);
            Top = revitCenterY - (Height / 2);
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
