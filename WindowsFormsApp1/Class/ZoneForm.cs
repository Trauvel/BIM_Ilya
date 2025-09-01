using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Mechanical;

using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1.Class
{
    public class ZoneForm : System.Windows.Forms.Form
    {
        private readonly System.Windows.Forms.TextBox txtZoneName; // Имя зоны
        private readonly System.Windows.Forms.TextBox txtLevel; // Уровень
        private readonly System.Windows.Forms.TextBox txtTotalSupplyAir; // Сумма притоков пространств зоны
        private readonly System.Windows.Forms.TextBox txtTotalExhaustAir; // Сумма вытяжек пространств зоны
        private readonly System.Windows.Forms.TextBox txtAirDifference; // Разность притока с вытяжкой
        private readonly ListView listViewSpaces; // Список пространств
        private readonly Zone _zone;
        private readonly UIApplication _uiapp;

        public ZoneForm(Zone zone, UIApplication uiapp)
        {
            _zone = zone;
            _uiapp = uiapp;

            Text = "Zone Parameters";
            Size = new System.Drawing.Size(600, 600);
            StartPosition = FormStartPosition.Manual;

            //Deactivate += (s, e) => Close();

            KeyPreview = true;
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form_KeyDown);

            var parametersGroup = CreateGroupBox("Zone Parameters", 10, 10, 550, 300);
            txtZoneName = CreateTextBox("Имя зоны", 40, 30, parametersGroup);
            txtLevel = CreateTextBox("Уровень", 40, 70, parametersGroup);
            txtTotalSupplyAir = CreateTextBox("Сумма притока", 40, 110, parametersGroup);
            txtTotalExhaustAir = CreateTextBox("Сумма вытяжки", 40, 150, parametersGroup);
            txtAirDifference = CreateTextBox("Разность притока и вытяжки", 40, 190, parametersGroup);

            var spacesGroup = CreateGroupBox("Пространства", 10, 320, 550, 200);
            listViewSpaces = CreateListView(20, 30, 500, 150, spacesGroup);

            LoadZoneParameters(zone);

            // Enable resizing of the controls when the form is resized
            parametersGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtZoneName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtLevel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtTotalSupplyAir.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtTotalExhaustAir.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtAirDifference.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            spacesGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            listViewSpaces.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

            // Добавляем обработчик события для выделения пространства на плане при нажатии на элемент списка
            listViewSpaces.ItemActivate += ListViewSpaces_ItemActivate;
            listViewSpaces.MouseClick += ListViewSpaces_MouseClick;
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void LoadZoneParameters(Zone zone)
        {
            txtZoneName.Text = zone.Name;
            txtLevel.Text = GetLevelName(zone);
            double totalSupply = GetTotalSupplyAir(zone);
            double totalExhaust = GetTotalExhaustAir(zone);
            txtTotalSupplyAir.Text = totalSupply.ToString("F2");
            txtTotalExhaustAir.Text = totalExhaust.ToString("F2");
            txtAirDifference.Text = (totalSupply - totalExhaust).ToString("F2");

            LoadSpacesListView(zone);
        }

        private void LoadSpacesListView(Zone zone)
        {
            listViewSpaces.Items.Clear();
            var spaces = GetSpacesInZone(zone);
            foreach (var space in spaces)
            {
                string spaceName = space.Name;
                string supplyAir = GetParameterValueInCubicMeters(space, "Приток").ToString("F2");
                string exhaustAir = GetParameterValueInCubicMeters(space, "Вытяжка").ToString("F2");

                ListViewItem item = new System.Windows.Forms.ListViewItem(new[] { spaceName, supplyAir, exhaustAir })
                {
                    Tag = space // Сохраняем пространство в свойстве Tag, чтобы затем использовать его для выделения
                };
                listViewSpaces.Items.Add(item);
            }

            listViewSpaces.LabelEdit = true; // Включаем возможность редактирования
        }

        private void ListViewSpaces_ItemActivate(object sender, EventArgs e)
        {
            if (listViewSpaces.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = listViewSpaces.SelectedItems[0];

                if (selectedItem.Tag is Element space)
                {
                    UIDocument uidoc = _uiapp.ActiveUIDocument;
                    uidoc.Selection.SetElementIds(new List<ElementId> { space.Id });
                    uidoc.ShowElements(space);
                }
            }
        }

        private void ListViewSpaces_MouseClick(object sender, MouseEventArgs e)
        {
            if (listViewSpaces.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = listViewSpaces.GetItemAt(e.X, e.Y);
                if (selectedItem != null && selectedItem.Tag is Element space)
                {
                    int subItemIndex = GetSubItemIndex(selectedItem, e.X);
                    if (subItemIndex == 1 || subItemIndex == 2) // Проверяем, что клик по колонке "Приток" или "Вытяжка"
                    {
                        string columnName = listViewSpaces.Columns[subItemIndex].Text;

                        using (var inputDialog = new System.Windows.Forms.Form())
                        {
                            inputDialog.Text = "Редактировать " + columnName;
                            var textBox = new System.Windows.Forms.TextBox { Width = 200, Text = selectedItem.SubItems[subItemIndex].Text };
                            var okButton = new System.Windows.Forms.Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Bottom };

                            inputDialog.Controls.Add(textBox);
                            inputDialog.Controls.Add(okButton);
                            inputDialog.AcceptButton = okButton;
                            inputDialog.StartPosition = FormStartPosition.CenterParent;

                            if (inputDialog.ShowDialog() == DialogResult.OK)
                            {
                                if (double.TryParse(textBox.Text, out double newValue))
                                {
                                    using (Transaction trans = new Transaction(_uiapp.ActiveUIDocument.Document, "Update Space Parameters"))
                                    {
                                        trans.Start();
                                        SetParameterValue(space, columnName, newValue);
                                        trans.Commit();
                                    }
                                    selectedItem.SubItems[subItemIndex].Text = newValue.ToString("F2");
                                }
                            }
                        }
                    }
                }
            }
        }

        private int GetSubItemIndex(ListViewItem item, int x)
        {
            int currentX = 0;
            for (int i = 0; i < listViewSpaces.Columns.Count; i++)
            {
                int columnWidth = listViewSpaces.Columns[i].Width;
                if (x >= currentX && x < currentX + columnWidth)
                {
                    return i;
                }
                currentX += columnWidth;
            }
            return -1;
        }

        private double GetTotalSupplyAir(Zone zone)
        {
            double totalSupply = 0;
            var spaces = GetSpacesInZone(zone);
            foreach (var space in spaces)
            {
                totalSupply += GetParameterValueInCubicMeters(space, "Приток");
            }
            return totalSupply;
        }

        private double GetTotalExhaustAir(Zone zone)
        {
            double totalExhaust = 0;
            var spaces = GetSpacesInZone(zone);
            foreach (var space in spaces)
            {
                totalExhaust += GetParameterValueInCubicMeters(space, "Вытяжка");
            }
            return totalExhaust;
        }

        private IEnumerable<Element> GetSpacesInZone(Zone zone)
        {
            var spaces = new List<Element>();

            // Используем параметр Spaces для получения SpaceSet
            var spaceSet = zone.Spaces;
            if (spaceSet != null)
            {
                foreach (var space in spaceSet)
                {
                    if (space is Element element)
                    {
                        spaces.Add(element);
                    }
                }
            }
            return spaces;
        }

        private string GetLevelName(Element element)
        {
            ElementId levelId = element.LevelId;
            if (levelId != ElementId.InvalidElementId)
            {
                if (element.Document.GetElement(levelId) is Level level)
                {
                    return level.Name;
                }
            }
            return string.Empty;
        }

        private System.Windows.Forms.TextBox CreateTextBox(string label, int x, int y, System.Windows.Forms.Control parent)
        {
            var lbl = new System.Windows.Forms.Label
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

        private GroupBox CreateGroupBox(string text, int x, int y, int width, int height)
        {
            var groupBox = new System.Windows.Forms.GroupBox
            {
                Text = text,
                Location = new System.Drawing.Point(x, y),
                Size = new System.Drawing.Size(width, height)
            };
            Controls.Add(groupBox);
            return groupBox;
        }

        private ListView CreateListView(int x, int y, int width, int height, System.Windows.Forms.Control parent)
        {
            var listView = new System.Windows.Forms.ListView
            {
                Location = new System.Drawing.Point(x, y),
                Size = new System.Drawing.Size(width, height),
                View = System.Windows.Forms.View.Details,
                FullRowSelect = true,
                GridLines = true
            };

            listView.Columns.Add("Имя пространства", 150);
            listView.Columns.Add("Приток", 100);
            listView.Columns.Add("Вытяжка", 100);

            parent.Controls.Add(listView);

            return listView;
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

        private static void SetParameterValue(Element space, string paramName, double value)
        {
            Parameter param = space.LookupParameter(paramName);
            if (param != null && param.IsReadOnly == false && param.StorageType == StorageType.Double)
            {
                param.Set(UnitUtils.ConvertToInternalUnits(value, UnitTypeId.CubicMeters));
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
