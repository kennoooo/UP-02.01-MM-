using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private DataTable _inputTable;
        private const int RowsCount = 3;
        private const int ColsCount = 5;

        public MainWindow()
        {
            InitializeComponent();
            InitializeInputGrid();
        }

        private void InitializeInputGrid()
        {
            _inputTable = new DataTable();
            _inputTable.Columns.Add("Поставщик \\ Потребитель", typeof(string));

            for (int j = 0; j < ColsCount; j++)
            {
                _inputTable.Columns.Add($"B{j + 1}", typeof(int));
            }
            _inputTable.Columns.Add("Запасы (ai)", typeof(int));

            int[,] defaultCosts =
            {
                { 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0 }
            };
            int[] defaultSupply = { 200, 350, 300 };
            int[] defaultDemand = { 270, 130, 190, 150, 110 };

            for (int i = 0; i < RowsCount; i++)
            {
                DataRow row = _inputTable.NewRow();
                row[0] = $"A{i + 1}";
                for (int j = 0; j < ColsCount; j++)
                {
                    row[j + 1] = defaultCosts[i, j];
                }
                row[ColsCount + 1] = defaultSupply[i];
                _inputTable.Rows.Add(row);
            }

            DataRow demandRow = _inputTable.NewRow();
            demandRow[0] = "Потребности (bj)";
            for (int j = 0; j < ColsCount; j++)
            {
                demandRow[j + 1] = defaultDemand[j];
            }
            demandRow[ColsCount + 1] = 0;
            _inputTable.Rows.Add(demandRow);

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

        private bool TryReadData(out int[] supply, out int[] demand, out int[,] costs)
        {
            supply = new int[RowsCount];
            demand = new int[ColsCount];
            costs = new int[RowsCount, ColsCount];

            try
            {
                for (int i = 0; i < RowsCount; i++)
                {
                    for (int j = 0; j < ColsCount; j++)
                    {
                        costs[i, j] = Convert.ToInt32(_inputTable.Rows[i][j + 1]);
                        if (costs[i, j] < 0) throw new Exception();
                    }
                    supply[i] = Convert.ToInt32(_inputTable.Rows[i][ColsCount + 1]);
                    if (supply[i] < 0) throw new Exception();
                }

                for (int j = 0; j < ColsCount; j++)
                {
                    demand[j] = Convert.ToInt32(_inputTable.Rows[RowsCount][j + 1]);
                    if (demand[j] < 0) throw new Exception();
                }

                if (supply.Sum() != demand.Sum())
                {
                    MessageBox.Show($"Задача не сбалансирована! Сумма запасов ({supply.Sum()}) не равна сумме потребностей ({demand.Sum()}).",
                                    "Ошибка баланса", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
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

        private void HighlightActiveButton(Button active, Button inactive)
        {
            active.Background = new SolidColorBrush(Color.FromRgb(225, 240, 255));
            active.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 120, 215));
            active.FontWeight = FontWeights.SemiBold;

            inactive.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
            inactive.BorderBrush = new SolidColorBrush(Color.FromRgb(171, 171, 171));
            inactive.FontWeight = FontWeights.Normal;
        }

        private void SolveNorthWest(int[] supply, int[] demand, int[,] costs)
        {
            int[,] plan = new int[RowsCount, ColsCount];
            int[] a = (int[])supply.Clone();
            int[] b = (int[])demand.Clone();

            int i = 0, j = 0;
            while (i < RowsCount && j < ColsCount)
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
            int[,] plan = new int[RowsCount, ColsCount];
            int[] a = (int[])supply.Clone();
            int[] b = (int[])demand.Clone();
            bool[] rowClosed = new bool[RowsCount];
            bool[] colClosed = new bool[ColsCount];

            while (true)
            {
                int minCost = int.MaxValue;
                int minI = -1, minJ = -1;

                for (int i = 0; i < RowsCount; i++)
                {
                    if (rowClosed[i]) continue;
                    for (int j = 0; j < ColsCount; j++)
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

            for (int j = 0; j < ColsCount; j++)
            {
                dt.Columns.Add($"B{j + 1} ({demand[j]})", typeof(string));
            }
            dt.Columns.Add("Запасы (ai)", typeof(int));

            for (int i = 0; i < RowsCount; i++)
            {
                DataRow row = dt.NewRow();
                row[0] = $"A{i + 1}";

                for (int j = 0; j < ColsCount; j++)
                {
                    int x = plan[i, j];
                    int c = costs[i, j];
                    totalCost += x * c;

                    row[j + 1] = x > 0 ? $"{x}  [тариф: {c}]" : $"-  [тариф: {c}]";
                }
                row[ColsCount + 1] = supply[i];
                dt.Rows.Add(row);
            }

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
            TxtMethodInfo.Text = $"Выбран метод: {methodName}";
            TxtTotalCost.Text = $"Целевая функция F(x) = {totalCost} у.е.";
        }
    }
}
