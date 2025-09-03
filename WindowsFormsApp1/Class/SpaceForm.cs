using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;

namespace WindowsFormsApp1.Class
{
    public class SpaceForm : System.Windows.Forms.Form
    {
        private readonly System.Windows.Forms.TextBox txtVentilation;
        private readonly System.Windows.Forms.TextBox txtIntake;
        private readonly System.Windows.Forms.TextBox txtSystemNameVentilation;
        private readonly System.Windows.Forms.TextBox txtSystemNameIntake;
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
        
        // Константы для компактного дизайна
        private const int MARGIN = 5;
        private const int GROUP_MARGIN = 5;
        private const int LABEL_WIDTH = 150;
        private const int TEXTBOX_WIDTH = 120;
        private const int TEXTBOX_HEIGHT = 20;
        private const int LABEL_HEIGHT = 18;
        private const int GROUP_PADDING = 5;
        private const int ROW_SPACING = 3;
        private const int COLUMN_WIDTH = 300;
        private const int GROUP_HEADER_HEIGHT = 25;
        
        // Флаг для отслеживания необходимости сохранения
        private bool _needsSave = false;
        
        // Статическая очередь для отложенных сохранений
        private static readonly Queue<Action> _pendingSaves = new Queue<Action>();

        public SpaceForm(Element space)
        {
            _space = space;
            Text = "Space Parameters";
            StartPosition = FormStartPosition.Manual;
            AutoSize = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            Deactivate += (s, e) => Close();

            FormClosing += (s, e) =>
            {
                _needsSave = true;
                
                try
                {
                    // Проверяем, можно ли модифицировать документ
                    //if (_space.Document.IsReadOnly)
                    //{
                    //    TaskDialog.Show("Предупреждение", "Документ доступен только для чтения. Изменения не будут сохранены.");
                    //    return; // Выходим без попытки сохранения
                    //}

                    // Проверяем, что пространство все еще существует и валидно
                    if (_space == null || _space.IsValidObject == false)
                    {
                        TaskDialog.Show("Ошибка", "Пространство больше не доступно для редактирования.");
                        return;
                    }

                    // Пытаемся сохранить
                    SaveAndCalculateParameters();
                    _needsSave = false;
                    
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                {
                    TaskDialog.Show("Информация", "Документ заблокирован или недоступен для редактирования. Изменения будут сохранены позже.\n\nОшибка: " + ex.Message);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", $"Произошла ошибка при сохранении: {ex.Message}\nИзменения будут сохранены позже.");
                }
                
                // Если сохранение не удалось, добавляем в очередь отложенных сохранений
                if (_needsSave)
                {
                    _pendingSaves.Enqueue(() => SaveAndCalculateParameters());
                }
            };

            KeyPreview = true;
            KeyDown += new KeyEventHandler(Form1_KeyDown);

            // Создаем GroupBox'ы в две колонки с аккордеоном
            var leftColumn = CreateColumn(MARGIN, MARGIN);
            var rightColumn = CreateColumn(MARGIN + COLUMN_WIDTH, MARGIN);

            // Левая колонка
            var otherParametersGroup = CreateCollapsibleGroupBox("Остальные параметры", leftColumn, 0);
            txtNumber = CreateLabeledTextBox("Номер", otherParametersGroup, 0);
            txtName = CreateLabeledTextBox("Имя", otherParametersGroup, 1);
            txtArea = CreateLabeledTextBox("Площадь", otherParametersGroup, 2);
            txtOffset = CreateLabeledTextBox("Смещение сверху", otherParametersGroup, 3);
            txtVolume = CreateLabeledTextBox("Объем", otherParametersGroup, 4);
            txtCategory = CreateLabeledTextBox("ADSK_Категория помещения", otherParametersGroup, 5);
            txtTemperature = CreateLabeledTextBox("ADSK_Температура в помещении", otherParametersGroup, 6);
            txtCleanliness = CreateLabeledTextBox("Категория_помещения_по_чистоте", otherParametersGroup, 7);
            txtHeatLoss = CreateLabeledTextBox("Теплопотери", otherParametersGroup, 8);

            var multiplicityGroup = CreateCollapsibleGroupBox("Кратность", leftColumn, 1);
            txtP_Multiplicity = CreateLabeledTextBox("Приток кратность", multiplicityGroup, 0);
            txtV_Multiplicity = CreateLabeledTextBox("Вытяжка кратность", multiplicityGroup, 1);

            var peopleGroup = CreateCollapsibleGroupBox("Количество человек", leftColumn, 2);
            txtStandard = CreateLabeledTextBox("Норма", peopleGroup, 0);
            txtP_KolPeople = CreateLabeledTextBox("Приток кол-во человек", peopleGroup, 1);
            txtV_KolPeople = CreateLabeledTextBox("Вытяжка кол-во человек", peopleGroup, 2);

            // Правая колонка
            var ventilationGroup = CreateCollapsibleGroupBox("Вентиляция", rightColumn, 0);
            txtVentilation = CreateLabeledTextBox("Приток", ventilationGroup, 0);
            txtIntake = CreateLabeledTextBox("Вытяжка", ventilationGroup, 1);

            var systemNameGroup = CreateCollapsibleGroupBox("Имя системы", rightColumn, 1);
            txtSystemNameVentilation = CreateLabeledTextBox("Приток", systemNameGroup, 0);
            txtSystemNameIntake = CreateLabeledTextBox("Вытяжка", systemNameGroup, 1);

            // Пересчитываем позиции всех групп после создания
            RecalculateAllGroupPositions(leftColumn);
            RecalculateAllGroupPositions(rightColumn);

            // Устанавливаем фиксированный размер формы
            SetFormSize();

            LoadSpaceParameters(space);
        }

        private System.Windows.Forms.Panel CreateColumn(int x, int y)
        {
            var panel = new System.Windows.Forms.Panel
            {
                Location = new System.Drawing.Point(x, y),
                AutoSize = false,
                Size = new System.Drawing.Size(COLUMN_WIDTH - MARGIN * 2, 600),
                BorderStyle = BorderStyle.None
            };
            Controls.Add(panel);
            return panel;
        }

        private void SetFormSize()
        {
            Width = COLUMN_WIDTH * 2 + MARGIN * 3;
            Height = 700;
        }

        private System.Windows.Forms.GroupBox CreateCollapsibleGroupBox(string title, System.Windows.Forms.Panel parent, int groupIndex)
        {
            // Временно создаем группу с минимальной высотой
            var groupBox = new System.Windows.Forms.GroupBox
            {
                Text = title,
                Location = new System.Drawing.Point(0, 0), // Временная позиция
                Size = new System.Drawing.Size(COLUMN_WIDTH - MARGIN * 2, GROUP_HEADER_HEIGHT),
                Padding = new Padding(GROUP_PADDING),
                Tag = groupIndex
            };

            // Добавляем кнопку сворачивания/разворачивания
            var toggleButton = new System.Windows.Forms.Button
            {
                Text = "▼",
                Size = new System.Drawing.Size(20, 20),
                Location = new System.Drawing.Point(groupBox.Width - 25, 2),
                FlatStyle = FlatStyle.Flat,
                Tag = groupBox
            };

            toggleButton.Click += (s, e) => ToggleGroup(groupBox, toggleButton);
            groupBox.Controls.Add(toggleButton);

            parent.Controls.Add(groupBox);
            return groupBox;
        }

        private void RecalculateAllGroupPositions(System.Windows.Forms.Panel column)
        {
            int currentY = 0;
            
            // Сортируем группы по индексу
            var groups = new List<System.Windows.Forms.GroupBox>();
            foreach (System.Windows.Forms.Control control in column.Controls)
            {
                if (control is System.Windows.Forms.GroupBox groupBox)
                {
                    groups.Add(groupBox);
                }
            }
            
            // Более безопасная сортировка с проверкой типов
            groups.Sort((a, b) => 
            {
                int indexA = a.Tag is int intA ? intA : 0;
                int indexB = b.Tag is int intB ? intB : 0;
                return indexA.CompareTo(indexB);
            });

            // Устанавливаем позиции групп
            foreach (var groupBox in groups)
            {
                groupBox.Location = new System.Drawing.Point(0, currentY);
                currentY += groupBox.Height + GROUP_MARGIN;
            }
        }

        private void ToggleGroup(System.Windows.Forms.GroupBox groupBox, System.Windows.Forms.Button toggleButton)
        {
            bool isCollapsed = groupBox.Height <= GROUP_HEADER_HEIGHT;
            
            if (isCollapsed)
            {
                // Разворачиваем группу
                int contentHeight = CalculateGroupContentHeight(groupBox);
                groupBox.Height = contentHeight + GROUP_PADDING * 2;
                toggleButton.Text = "▼";
                
                // Показываем все элементы управления
                foreach (System.Windows.Forms.Control control in groupBox.Controls)
                {
                    if (control != toggleButton)
                        control.Visible = true;
                }
            }
            else
            {
                // Сворачиваем группу
                groupBox.Height = GROUP_HEADER_HEIGHT;
                toggleButton.Text = "▶";
                
                // Скрываем все элементы управления кроме кнопки
                foreach (System.Windows.Forms.Control control in groupBox.Controls)
                {
                    if (control != toggleButton)
                        control.Visible = false;
                }
            }

            // Пересчитываем позиции всех групп в колонке
            var parent = groupBox.Parent as System.Windows.Forms.Panel;
            if (parent != null)
            {
                RecalculateAllGroupPositions(parent);
            }
        }

        private int CalculateGroupContentHeight(System.Windows.Forms.GroupBox groupBox)
        {
            int maxBottom = 0;
            foreach (System.Windows.Forms.Control control in groupBox.Controls)
            {
                if (control != groupBox.Controls[groupBox.Controls.Count - 1]) // Исключаем кнопку
                {
                    if (control.Bottom > maxBottom) maxBottom = control.Bottom;
                }
            }
            return maxBottom + GROUP_PADDING;
        }

        private System.Windows.Forms.TextBox CreateLabeledTextBox(string labelText, System.Windows.Forms.GroupBox parent, int rowIndex)
        {
            // Создаем Label
            var label = new System.Windows.Forms.Label
            {
                Text = labelText,
                Location = new System.Drawing.Point(GROUP_PADDING, GROUP_HEADER_HEIGHT + rowIndex * (LABEL_HEIGHT + ROW_SPACING)),
                Size = new System.Drawing.Size(LABEL_WIDTH, LABEL_HEIGHT),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Создаем TextBox с фиксированной шириной и умной прокруткой
            var textBox = new System.Windows.Forms.TextBox
            {
                Location = new System.Drawing.Point(GROUP_PADDING + LABEL_WIDTH + 5, GROUP_HEADER_HEIGHT + rowIndex * (LABEL_HEIGHT + ROW_SPACING)),
                Size = new System.Drawing.Size(TEXTBOX_WIDTH, TEXTBOX_HEIGHT),
                AutoSize = false,
                Multiline = true,
                WordWrap = true,
                ScrollBars = ScrollBars.Vertical, // Только вертикальная прокрутка
                MaxLength = 1000
            };

            // Добавляем обработчик для изменения только высоты
            textBox.TextChanged += (s, e) => HeightOnlyResizeTextBox(textBox);
            
            // Добавляем элементы в GroupBox
            parent.Controls.Add(label);
            parent.Controls.Add(textBox);

            // Автоматически подстраиваем размер GroupBox
            int contentHeight = GROUP_HEADER_HEIGHT + (rowIndex + 1) * (LABEL_HEIGHT + ROW_SPACING) + GROUP_PADDING;
            parent.Size = new System.Drawing.Size(parent.Width, contentHeight + GROUP_PADDING);

            return textBox;
        }

        private void HeightOnlyResizeTextBox(System.Windows.Forms.TextBox textBox)
        {
            if (string.IsNullOrEmpty(textBox.Text))
            {
                textBox.Height = TEXTBOX_HEIGHT;
                return;
            }

            using (var graphics = textBox.CreateGraphics())
            {
                var textSize = graphics.MeasureString(textBox.Text, textBox.Font);
                
                // Изменяем только высоту, ширина остается фиксированной
                int newHeight = Math.Max(TEXTBOX_HEIGHT, Math.Min((int)textSize.Height + 10, 80));
                
                if (textBox.Height != newHeight)
                {
                    textBox.Height = newHeight;
                    
                    // Пересчитываем позиции элементов ниже
                    RepositionControlsBelowTextBox(textBox);
                    
                    // Пересчитываем размер GroupBox
                    ResizeGroupBox(textBox.Parent as System.Windows.Forms.GroupBox);
                }
            }
        }

        private void RepositionControlsAfterTextBox(System.Windows.Forms.TextBox changedTextBox)
        {
            var parent = changedTextBox.Parent as System.Windows.Forms.GroupBox;
            if (parent == null) return;

            // Находим все элементы справа от измененного TextBox
            foreach (System.Windows.Forms.Control control in parent.Controls)
            {
                if (control is System.Windows.Forms.TextBox textBox && 
                    textBox != changedTextBox && 
                    textBox.Location.X > changedTextBox.Location.X)
                {
                    // Сдвигаем элемент вправо на разницу в ширине
                    int widthDifference = changedTextBox.Width - TEXTBOX_WIDTH;
                    textBox.Location = new System.Drawing.Point(
                        textBox.Location.X + widthDifference, 
                        textBox.Location.Y
                    );
                }
            }
        }

        private void RepositionControlsBelowTextBox(System.Windows.Forms.TextBox changedTextBox)
        {
            var parent = changedTextBox.Parent as System.Windows.Forms.GroupBox;
            if (parent == null) return;

            int heightDifference = changedTextBox.Height - TEXTBOX_HEIGHT;
            if (heightDifference == 0) return;

            // Находим все элементы ниже измененного TextBox
            foreach (System.Windows.Forms.Control control in parent.Controls)
            {
                if (control != changedTextBox && 
                    control.Location.Y > changedTextBox.Location.Y)
                {
                    // Сдвигаем элемент вниз на разницу в высоте
                    control.Location = new System.Drawing.Point(
                        control.Location.X, 
                        control.Location.Y + heightDifference
                    );
                }
            }
        }

        private void ResizeGroupBox(System.Windows.Forms.GroupBox groupBox)
        {
            if (groupBox == null) return;

            // Находим самый нижний элемент в GroupBox
            int maxBottom = 0;
            foreach (System.Windows.Forms.Control control in groupBox.Controls)
            {
                if (control.Bottom > maxBottom) maxBottom = control.Bottom;
            }

            // Устанавливаем новую высоту GroupBox
            int newHeight = maxBottom + GROUP_PADDING;
            groupBox.Height = newHeight;

            // Пересчитываем позиции всех групп в колонке
            var column = groupBox.Parent as System.Windows.Forms.Panel;
            if (column != null)
            {
                RecalculateAllGroupPositions(column);
            }
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
            txtSystemNameVentilation.Text = GetStringParameterValue(space, "П_ИмяСистемы");
            txtSystemNameIntake.Text = GetStringParameterValue(space, "В_ИмяСистемы");
            txtP_Multiplicity.Text = GetParameterValueInCubicMeters(space, "П_Крат").ToString();
            txtV_Multiplicity.Text = GetParameterValueInCubicMeters(space, "В_Крат").ToString();
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

        private void SaveAndCalculateParameters()
        {
            // Дополнительные проверки перед сохранением
            if (_space == null || _space.IsValidObject == false)
            {
                throw new InvalidOperationException("Пространство недоступно для редактирования");
            }

            //if (_space.Document.IsReadOnly)
            //{
            //    throw new InvalidOperationException("Документ доступен только для чтения");
            //}

            using (Transaction transaction = new Transaction(_space.Document, "Update Space Parameters"))
            {
                try
                {
                    transaction.Start();

                    // Сохраняем параметры
                    SetParameterValue(_space, "Приток", txtVentilation.Text);
                    SetParameterValue(_space, "Вытяжка", txtIntake.Text);
                    SetParameterValue(_space, "П_ИмяСистемы", txtSystemNameVentilation.Text);
                    SetParameterValue(_space, "В_ИмяСистемы", txtSystemNameIntake.Text);
                    SetParameterValue(_space, "П_Крат", txtP_Multiplicity.Text);
                    SetParameterValue(_space, "В_Крат", txtV_Multiplicity.Text);
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

                    // Пересчитываем параметры
                    CalculateSpaceParameters(_space);

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    throw new InvalidOperationException($"Ошибка при сохранении параметров: {ex.Message}", ex);
                }
            }
        }

        private void CalculateSpaceParameters(Element space)
        {
            try
            {
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
                        SetParameterValue(space, "Вытяжка", exhaustResult.ToString());
                    }

                    if (p_krat != 0 && supplyResult == 0)
                    {
                        supplyResult = RoundToNearestMultipleOfFive(volume * p_krat);
                        SetParameterValue(space, "Приток", supplyResult.ToString());
                    }
                }
                
                if (norm != 0)
                {
                    if (v_kolChel != 0 && exhaustResult == 0)
                    {
                        exhaustResult = RoundToNearestMultipleOfFive(norm * v_kolChel);
                        SetParameterValue(_space, "Вытяжка", exhaustResult.ToString());
                    }

                    if (p_kolChel != 0 && supplyResult == 0)
                    {
                        supplyResult = RoundToNearestMultipleOfFive(norm * p_kolChel);
                        SetParameterValue(_space, "Приток", supplyResult.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Ошибка при расчете параметров: {ex.Message}", ex);
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
            IntPtr revitHandle = uiapp.MainWindowHandle;
            GetWindowRect(revitHandle, out RECT revitRect);

            int revitCenterX = (revitRect.Left + revitRect.Right) / 2;
            int revitCenterY = (revitRect.Top + revitRect.Bottom) / 2;

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

        // Вспомогательные методы для расчета
        private static double GetParameterValue(Element space, string paramName)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.HasValue)
            {
                return param.StorageType == StorageType.Double ? UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.CubicMeters) : param.AsDouble();
            }
            return 0;
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
                int normParamInt = (int)double.Parse(normParam.AsValueString().Replace(" м³", ""), System.Globalization.CultureInfo.InvariantCulture);
                return normParamInt;
            }
            return 40;
        }

        private static int RoundToNearestMultipleOfFive(double value)
        {
            return (int)(Math.Round(value / 5.0) * 5);
        }

        // Статический метод для обработки отложенных сохранений
        public static void ProcessPendingSaves()
        {
            int processedCount = 0;
            int maxAttempts = 10;
            
            while (_pendingSaves.Count > 0 && processedCount < maxAttempts)
            {
                try
                {
                    var saveAction = _pendingSaves.Dequeue();
                    saveAction();
                    processedCount++;
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка при отложенном сохранении: {ex.Message}");
                    processedCount++;
                }
            }
        }
    }
}
