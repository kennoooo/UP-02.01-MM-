using ClosedXML.Excel;
using Microsoft.Win32;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private const int MaxSize = 20;

        private DataTable _inputTable;
        private DataTable _resultTable;
        private int _rows = 3;
        private int _cols = 5;
        private string _resultMethod = "";
        private string _resultCost = "";

        public MainWindow()
        {
            InitializeComponent();
            BuildInputGrid(_rows, _cols);
        }

        private void BuildInputGrid(int rows, int cols)
        {
            _rows = rows;
            _cols = cols;

            _inputTable = new DataTable();
            _inputTable.Columns.Add("Поставщик \\ Потребитель", typeof(string));
            for (int j = 0; j < _cols; j++)
                _inputTable.Columns.Add($"B{j + 1}", typeof(int));
            _inputTable.Columns.Add("Запасы (ai)", typeof(int));

            for (int i = 0; i < _rows; i++)
            {
                DataRow row = _inputTable.NewRow();
                row[0] = $"A{i + 1}";
                for (int j = 0; j <= _cols; j++)
                    row[j + 1] = 0;
                _inputTable.Rows.Add(row);
            }

            DataRow demandRow = _inputTable.NewRow();
            demandRow[0] = "Потребности (bj)";
            for (int j = 0; j <= _cols; j++)
                demandRow[j + 1] = 0;
            _inputTable.Rows.Add(demandRow);

            RefreshInputGridColumns();
            TxtInputTitle.Text = $"Ввод исходных данных (поставщиков: {_rows}, потребителей: {_cols})";
            ResetResult();
        }

        private void RefreshInputGridColumns()
        {
            InputGrid.Columns.Clear();
            for (int c = 0; c < _inputTable.Columns.Count; c++)
            {
                DataColumn col = _inputTable.Columns[c];
                InputGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = col.ColumnName,
                    Binding = new Binding($"[{col.ColumnName}]"),
                    IsReadOnly = (c == 0)
                });
            }
            InputGrid.ItemsSource = _inputTable.DefaultView;
        }

        private void ResetResult()
        {
            ResultGrid.ItemsSource = null;
            ResultGrid.Columns.Clear();
            _resultTable = null;
            _resultMethod = "";
            _resultCost = "";
            TxtMethodInfo.Text = "Выберите метод решения";
            TxtTotalCost.Text = "Целевая функция F(x) = 0";
            SetButtonStyle(BtnNorthWest, false);
            SetButtonStyle(BtnMinElement, false);
        }

        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(TxtRows.Text.Trim(), out int rows) ||
                !int.TryParse(TxtCols.Text.Trim(), out int cols) ||
                rows < 1 || cols < 1 || rows > MaxSize || cols > MaxSize)
            {
                MessageBox.Show($"Введите целые числа от 1 до {MaxSize} для количества строк и столбцов.",
                    "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show("Таблица будет создана заново, введённые данные пропадут. Продолжить?",
                "Новая таблица", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            InputGrid.CommitEdit(DataGridEditingUnit.Row, true);
            BuildInputGrid(rows, cols);
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Стереть все введённые данные и результат?",
                "Очистка", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            InputGrid.CommitEdit(DataGridEditingUnit.Row, true);
            BuildInputGrid(_rows, _cols);
        }

        private bool TryReadData(out int[] supply, out int[] demand, out int[,] costs)
        {
            supply = new int[_rows];
            demand = new int[_cols];
            costs = new int[_rows, _cols];

            try
            {
                for (int i = 0; i < _rows; i++)
                {
                    for (int j = 0; j < _cols; j++)
                    {
                        costs[i, j] = Convert.ToInt32(_inputTable.Rows[i][j + 1]);
                        if (costs[i, j] < 0) throw new Exception();
                    }
                    supply[i] = Convert.ToInt32(_inputTable.Rows[i][_cols + 1]);
                    if (supply[i] < 0) throw new Exception();
                }

                for (int j = 0; j < _cols; j++)
                {
                    demand[j] = Convert.ToInt32(_inputTable.Rows[_rows][j + 1]);
                    if (demand[j] < 0) throw new Exception();
                }

                int totalSupply = supply.Sum();
                int totalDemand = demand.Sum();

                if (totalSupply != totalDemand)
                {
                    int diff = Math.Abs(totalSupply - totalDemand);
                    string addTarget = totalSupply > totalDemand
                        ? $"фиктивного потребителя с потребностью {diff}"
                        : $"фиктивного поставщика с запасом {diff}";

                    MessageBoxResult choice = MessageBox.Show(
                        $"Задача не сбалансирована!\n" +
                        $"Сумма запасов: {totalSupply}, сумма потребностей: {totalDemand}.\n\n" +
                        $"Нажмите «Да», чтобы сбалансировать (добавить {addTarget}),\n" +
                        $"или «Нет», чтобы закрыть окно.",
                        "Ошибка баланса",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Error);

                    if (choice != MessageBoxResult.Yes)
                        return false;

                    return BalanceTask(ref supply, ref demand, ref costs, totalSupply, totalDemand);
                }

                return true;
            }
            catch
            {
                MessageBox.Show("Проверьте корректность введенных чисел. Все значения должны быть целыми неотрицательными числами.",
                    "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private bool BalanceTask(ref int[] supply, ref int[] demand, ref int[,] costs, int totalSupply, int totalDemand)
        {
            if (totalSupply > totalDemand)
            {
                if (_cols + 1 > MaxSize)
                {
                    MessageBox.Show($"Невозможно добавить фиктивного потребителя: превышен лимит столбцов ({MaxSize}).",
                        "Ошибка балансировки", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                int diff = totalSupply - totalDemand;
                int newCols = _cols + 1;

                var newDemand = new int[newCols];
                Array.Copy(demand, newDemand, _cols);
                newDemand[_cols] = diff;

                var newCosts = new int[_rows, newCols];
                for (int i = 0; i < _rows; i++)
                {
                    for (int j = 0; j < _cols; j++)
                        newCosts[i, j] = costs[i, j];
                    newCosts[i, _cols] = 0;
                }

                // Обновляем таблицу ввода в интерфейсе
                string newColName = $"B{newCols} (фикт.)";
                _inputTable.Columns.Add(newColName, typeof(int));
                _inputTable.Columns[newColName].SetOrdinal(_cols + 1);

                for (int i = 0; i < _rows; i++)
                    _inputTable.Rows[i][_cols + 1] = 0;
                _inputTable.Rows[_rows][_cols + 1] = diff;

                _cols = newCols;
                demand = newDemand;
                costs = newCosts;
                TxtCols.Text = _cols.ToString();
            }
            else
            {
                if (_rows + 1 > MaxSize)
                {
                    MessageBox.Show($"Невозможно добавить фиктивного поставщика: превышен лимит строк ({MaxSize}).",
                        "Ошибка балансировки", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                int diff = totalDemand - totalSupply;
                int newRows = _rows + 1;

                var newSupply = new int[newRows];
                Array.Copy(supply, newSupply, _rows);
                newSupply[_rows] = diff;

                var newCosts = new int[newRows, _cols];
                for (int i = 0; i < _rows; i++)
                    for (int j = 0; j < _cols; j++)
                        newCosts[i, j] = costs[i, j];

                for (int j = 0; j < _cols; j++)
                    newCosts[_rows, j] = 0;

                // Обновляем таблицу ввода в интерфейсе
                DataRow fictRow = _inputTable.NewRow();
                fictRow[0] = $"A{newRows} (фикт.)";
                for (int j = 0; j < _cols; j++)
                    fictRow[j + 1] = 0;
                fictRow[_cols + 1] = diff;
                _inputTable.Rows.InsertAt(fictRow, _rows);

                _rows = newRows;
                supply = newSupply;
                costs = newCosts;
                TxtRows.Text = _rows.ToString();
            }

            RefreshInputGridColumns();
            TxtInputTitle.Text = $"Ввод исходных данных (поставщиков: {_rows}, потребителей: {_cols})";
            return true;
        }

        private void BtnNorthWest_Click(object sender, RoutedEventArgs e)
        {
            InputGrid.CommitEdit(DataGridEditingUnit.Row, true);
if (!TryReadData(out int[] supply, out int[] demand, out int[,] costs)) return;
            HighlightActiveButton(BtnNorthWest, BtnMinElement);
            SolveNorthWest(supply, demand, costs);
        }
        private void BtnMinElement_Click(object sender, RoutedEventArgs e)
        {
            InputGrid.CommitEdit(DataGridEditingUnit.Row, true);
            if (!TryReadData(out int[] supply, out int[] demand, out int[,] costs)) return;
            HighlightActiveButton(BtnMinElement, BtnNorthWest);
            SolveMinElement(supply, demand, costs);
        }
        private static void SetButtonStyle(Button button, bool active)
        {
            if (active)
            {
                button.Background = new SolidColorBrush(Color.FromRgb(225, 240, 255));
                button.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 120, 215));
                button.FontWeight = FontWeights.SemiBold;
            }
            else
            {
                button.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
                button.BorderBrush = new SolidColorBrush(Color.FromRgb(171, 171, 171));
                button.FontWeight = FontWeights.Normal;
            }
        }
        private void HighlightActiveButton(Button active, Button inactive)
        {
            SetButtonStyle(active, true);
            SetButtonStyle(inactive, false);
        }
        private void SolveNorthWest(int[] supply, int[] demand, int[,] costs)
        {
            int[,] plan = new int[_rows, _cols];
            int[] a = (int[])supply.Clone();
            int[] b = (int[])demand.Clone();
            int i = 0, j = 0;
            while (i < _rows && j < _cols)
            {
                int quantity = Math.Min(a[i], b[j]);
                plan[i, j] = quantity;
                a[i] -= quantity;
                b[j] -= quantity;
                if (a[i] == 0) i++;
                else j++;
            }
            DisplayResult(plan, costs, supply, demand, "Метод северо-западного угла");
        }
        private void SolveMinElement(int[] supply, int[] demand, int[,] costs)
        {
            int[,] plan = new int[_rows, _cols];
            int[] a = (int[])supply.Clone();
            int[] b = (int[])demand.Clone();
            bool[] rowClosed = new bool[_rows];
            bool[] colClosed = new bool[_cols];
            while (true)
            {
                int minCost = int.MaxValue;
                int minI = -1, minJ = -1;
                for (int i = 0; i < _rows; i++)
                {
                    if (rowClosed[i]) continue;
                    for (int j = 0; j < _cols; j++)
                    {
                        if (colClosed[j]) continue;
                        if (costs[i, j] < minCost)
                        {
                            minCost = costs[i, j];
                            minI = i;
                            minJ = j;
                        }
                    }
                }
                if (minI == -1 || minJ == -1) break;
                int quantity = Math.Min(a[minI], b[minJ]);
                plan[minI, minJ] = quantity;
                a[minI] -= quantity;
                b[minJ] -= quantity;
                if (a[minI] == 0) rowClosed[minI] = true;
                else if (b[minJ] == 0) colClosed[minJ] = true;
            }
            DisplayResult(plan, costs, supply, demand, "Метод минимальных элементов");
        }
        private void DisplayResult(int[,] plan, int[,] costs, int[] supply, int[] demand, string methodName)
        {
            int totalCost = 0;
            DataTable dt = new DataTable();
            dt.Columns.Add("Поставщик \\ Потребитель", typeof(string));
            for (int j = 0; j < _cols; j++)
            {
                string colHeader = _inputTable.Columns[j + 1].ColumnName;
                dt.Columns.Add($"{colHeader} ({demand[j]})", typeof(string));
            }
            dt.Columns.Add("Запасы (ai)", typeof(int));
            for (int i = 0; i < _rows; i++)
            {
                DataRow row = dt.NewRow();
                row[0] = _inputTable.Rows[i][0];
                for (int j = 0; j < _cols; j++)
                {
                    int x = plan[i, j];
                    int c = costs[i, j];
                    totalCost += x * c;
                    row[j + 1] = x > 0 ? $"{x} [тариф: {c}]" : $"- [тариф: {c}]";
                }
                row[_cols + 1] = supply[i];
                dt.Rows.Add(row);
            }
            ResultGrid.ItemsSource = null;
            ResultGrid.Columns.Clear();
            foreach (DataColumn col in dt.Columns)
            {
                ResultGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = col.ColumnName,
                    Binding = new Binding($"[{col.ColumnName}]")
                });
            }
            ResultGrid.ItemsSource = dt.DefaultView;
            _resultTable = dt;
            _resultMethod = $"Выбран метод: {methodName}";
            _resultCost = $"Целевая функция F(x) = {totalCost} у.е.";
            TxtMethodInfo.Text = _resultMethod;
            TxtTotalCost.Text = _resultCost;
        }
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            InputGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var dialog = new SaveFileDialog
            {
                Filter = "Excel (.xlsx)|.xlsx",
                FileName = "transport_task.xlsx"
            };
            if (dialog.ShowDialog() != true) return;
            try
            {
                using (var wb = new XLWorkbook())
                {
                    var wsInput = wb.Worksheets.Add("Исходные данные");
                    WriteTable(wsInput, _inputTable);
                    if (_resultTable != null)
                    {
                        var wsResult = wb.Worksheets.Add("Результат");
                        WriteTable(wsResult, _resultTable);
                        int r = _resultTable.Rows.Count + 3;
                        wsResult.Cell(r, 1).Value = _resultMethod;
                        wsResult.Cell(r + 1, 1).Value = _resultCost;
                        wsResult.Cell(r, 1).Style.Font.Bold = true;
                        wsResult.Cell(r + 1, 1).Style.Font.Bold = true;
                    }
                    wb.SaveAs(dialog.FileName);
                }
                MessageBox.Show("Данные сохранены в файл.", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось сохранить файл (возможно, он открыт в Excel):\n" + ex.Message,
                "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private static void WriteTable(IXLWorksheet ws, DataTable dt)
        {
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var header = ws.Cell(1, c + 1);
                header.Value = dt.Columns[c].ColumnName;
                header.Style.Font.Bold = true;
            }
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    object v = dt.Rows[r][c];
                    var cell = ws.Cell(r + 2, c + 1);
                    if (v is int number) cell.Value = number;
                    else cell.Value = Convert.ToString(v);
                }
            }
            ws.Columns().AdjustToContents();
        }
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "Excel (.xlsx)|.xlsx" };
            if (dialog.ShowDialog() != true) return;
            try
            {
                int[,] values;
                int rows, cols;
                using (var wb = new XLWorkbook(dialog.FileName))
                {
                    var ws = wb.Worksheet(1);
                    var lastRowUsed = ws.LastRowUsed();
                    var lastColUsed = ws.LastColumnUsed();
                    if (lastRowUsed == null || lastColUsed == null)
                        throw new FormatException("Лист пуст.");
                    rows = lastRowUsed.RowNumber() - 2;
                    cols = lastColUsed.ColumnNumber() - 2;
                    if (rows < 1 || cols < 1 || rows > MaxSize || cols > MaxSize)
                        throw new FormatException(
                        $"Неверный размер таблицы в файле. Допустимо от 1 до {MaxSize} поставщиков и потребителей.");
                    values = new int[rows + 1, cols + 1];
                    for (int i = 0; i <= rows; i++)
                    {
                        for (int j = 0; j <= cols; j++)
                        {
                            var cell = ws.Cell(i + 2, j + 2);
                            string text = cell.GetString().Trim();
                            if (text.Length == 0)
                            {
                                values[i, j] = 0;
                            }
                            else if (!int.TryParse(text, out int parsed) || parsed < 0)
                            {
                                throw new FormatException(
                                $"Ячейка {cell.Address}: нужно целое неотрицательное число, а там «{text}».");
                            }
                            else
                            {
                                values[i, j] = parsed;
                            }
                        }
                    }
                }
                BuildInputGrid(rows, cols);
                TxtRows.Text = rows.ToString();
                TxtCols.Text = cols.ToString();
                for (int i = 0; i <= rows; i++)
                {
                    for (int j = 0; j <= cols; j++)
                    {
                        if (i == rows && j == cols) continue;
                        _inputTable.Rows[i][j + 1] = values[i, j];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось загрузить данные:\n" + ex.Message,
                "Ошибка импорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}