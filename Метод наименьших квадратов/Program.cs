using System;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using OxyPlot.Axes;
using MathNet.Symbolics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Globalization;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace Метод_наименьших_квадратов
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            MainForm mainForm = new MainForm();
            Application.Run(mainForm);
        }
    }

    public partial class MainForm : Form
    {
        private Label labelN;
        private TextBox textBoxN;
        private Button calculateButton;
        private Button clearButton;
        private Button generButton;
        private Button excelButton;
        private TextBox localPointsTextBox;
        private DataGridView dataGridView;
        private PlotView plotView;
        private Label line;
        private TextBox textBox5;

        public MainForm()
        {
            InitializeComponent();
            this.Size = new System.Drawing.Size(1000, 800);
        }

        private void InitializeComponent()
        {
            labelN = new Label();
            labelN.AutoSize = true;
            labelN.Location = new System.Drawing.Point(6, 13);
            labelN.Name = "label2";
            labelN.Size = new System.Drawing.Size(61, 13);
            labelN.TabIndex = 2;
            labelN.Text = "Введите степень:";

            textBoxN = new TextBox();
            textBoxN.Location = new System.Drawing.Point(110, 13);
            textBoxN.Name = "textBox2";
            textBoxN.Size = new System.Drawing.Size(44, 20);
            textBoxN.TabIndex = 3;
            
            line = new Label();
            line.AutoSize = true;
            line.Location = new System.Drawing.Point(6, 50);
            line.Name = "label3";
            line.Size = new System.Drawing.Size(73, 13);
            line.TabIndex = 11;
            line.Text = "Размерность матрицы:";

            textBox5 = new TextBox();
            textBox5.Location = new System.Drawing.Point(140, 45);
            textBox5.Name = "textBox5";
            textBox5.Size = new System.Drawing.Size(36, 20);
            textBox5.TabIndex = 13;

            calculateButton = new Button();
            calculateButton.Location = new System.Drawing.Point(3, 112);
            calculateButton.Name = "button1";
            calculateButton.Size = new System.Drawing.Size(84, 23);
            calculateButton.TabIndex = 4;
            calculateButton.Text = "Рассчитать ";
            calculateButton.UseVisualStyleBackColor = true;
            calculateButton.Click += new System.EventHandler(this.CalculateButton_Click);

            clearButton = new Button();
            clearButton.Location = new System.Drawing.Point(96, 112);
            clearButton.Name = "button2";
            clearButton.Size = new System.Drawing.Size(80, 23);
            clearButton.TabIndex = 5;
            clearButton.Text = "Очистить";
            clearButton.UseVisualStyleBackColor = true;
            clearButton.Click += new System.EventHandler(this.ClearButton_Click);

            generButton = new Button();
            generButton.Location = new System.Drawing.Point(192, 112);
            generButton.Name = "button3";
            generButton.Size = new System.Drawing.Size(158, 23);
            generButton.TabIndex = 6;
            generButton.Text = "Генерировать данные";
            generButton.UseVisualStyleBackColor = true;
            generButton.Click += GenerateDataButton_Click;

           excelButton = new Button();
            excelButton.Location = new System.Drawing.Point(69, 150);
            excelButton.Name = "button4";
            excelButton.Size = new System.Drawing.Size(120, 23);
            excelButton.TabIndex = 7;
            excelButton.Text = "Загрузить из EXCEL";
            excelButton.UseVisualStyleBackColor = true;
            excelButton.Click += ExcelButton_Click;

            localPointsTextBox = new TextBox();
            localPointsTextBox.Location = new System.Drawing.Point(382, 12);
            localPointsTextBox.Multiline = true;
            localPointsTextBox.Name = "textBox3";
            localPointsTextBox.Size = new System.Drawing.Size(215, 237);
            localPointsTextBox.TabIndex = 8;
            localPointsTextBox.UseWaitCursor = true;

            dataGridView = new DataGridView();
            dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new System.Drawing.Point(612, 12);
            dataGridView.Name = "dataGridView";
            dataGridView.Size = new System.Drawing.Size(355, 245);
            dataGridView.TabIndex = 9;

            plotView = new PlotView();
            plotView.Location = new System.Drawing.Point(12, 255);
            plotView.Name = "plotView";
            plotView.PanCursor = System.Windows.Forms.Cursors.Hand;
            plotView.Size = new System.Drawing.Size(900, 509);
            plotView.TabIndex = 10;
            plotView.ZoomHorizontalCursor = System.Windows.Forms.Cursors.SizeWE;
            plotView.ZoomRectangleCursor = System.Windows.Forms.Cursors.SizeNWSE;
            plotView.ZoomVerticalCursor = System.Windows.Forms.Cursors.SizeNS;
            
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(892, 667);
            Controls.Add(plotView);
            Controls.Add(dataGridView);
            Controls.Add(localPointsTextBox);
            Controls.Add(excelButton);
            Controls.Add(generButton);
            Controls.Add(clearButton);
            Controls.Add(calculateButton);
            Controls.Add(textBoxN);
            Controls.Add(labelN);
            Controls.Add(line);
            Controls.Add(textBox5);
            ResumeLayout(false);
            PerformLayout();
        }

        private void CalculateButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем наличие данных в DataGridView
                if (dataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("Добавьте данные в DataGridView.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int degree = int.Parse(textBoxN.Text);

                int rowCount = dataGridView.Rows.Count;
                double[,] data = new double[rowCount, degree + 1]; // Изменяем размерность матрицы

                // Заполняем матрицу данными из DataGridView
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j <= degree; j++)
                    {
                        data[i, j] = Convert.ToDouble(dataGridView.Rows[i].Cells[$"A{j + 1}"].Value); // Используем A1, A2, ..., B
                    }
                }

                // Вызываем метод наименьших квадратов для нахождения коэффициентов
                PolynomialCoefficients coefficients = LeastSquaresMethod(degree, data);

                // Выводим коэффициенты функции в текстовое поле
                StringBuilder coefficientsString = new StringBuilder("Коэффициенты функции: \r\n");
                for (int i = 0; i < coefficients.Coefficients.Length; i++)
                {
                    coefficientsString.Append($"a{i + 1} = {coefficients.Coefficients[i]:F2}\r\n");
                }

                localPointsTextBox.Text = coefficientsString.ToString();

                // Строим график
                PlotModel plotModel = new PlotModel();

                // Добавляем точки из DataGridView на график
                ScatterSeries scatterSeries = new ScatterSeries { MarkerType = MarkerType.Circle };
                for (int i = 0; i < rowCount; i++)
                {
                    scatterSeries.Points.Add(new ScatterPoint(data[i, 0], data[i, degree]));
                }
                plotModel.Series.Add(scatterSeries);

                // Находим минимальное и максимальное значения X для построения функции
                double minX = data[0, 0];
                double maxX = data[0, 0];
                for (int i = 1; i < rowCount; i++)
                {
                    minX = Math.Min(minX, data[i, 0]);
                    maxX = Math.Max(maxX, data[i, 0]);
                }

                // Добавляем аппроксимированную функцию на график
                FunctionSeries functionSeries = new FunctionSeries(x => CalculateApproximation(coefficients.Coefficients, x, degree), minX, maxX, 0.1);
                plotModel.Series.Add(functionSeries);

                // Добавляем сетку на график
                plotModel.Axes.Add(new OxyPlot.Axes.LinearAxis { MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Position = AxisPosition.Bottom });
                plotModel.Axes.Add(new OxyPlot.Axes.LinearAxis { MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Position = AxisPosition.Left });

                plotView.Model = plotModel;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при расчете коэффициентов: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private double CalculateApproximation(double[] coefficients, double x, int degree)
        {
            double result = 0;
            for (int i = 0; i <= degree; i++)
            {
                result += coefficients[i] * Math.Pow(x, i);
            }
            return result;
        }

        private void GenerateDataButton_Click(object sender, EventArgs e)
        {
            try
            {
                int rowCount;
                if (!int.TryParse(textBox5.Text, out rowCount) || rowCount <= 0)
                {
                    MessageBox.Show("Введите корректное число строк.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Очищаем существующие данные в DataGridView и добавляем столбцы, если их нет
                dataGridView.Columns.Clear();
                dataGridView.Rows.Clear();

                // Определяем размерность матрицы и вектора
                int matrixSize = rowCount;
                int vectorSize = rowCount;

                // Добавляем столбцы с соответствующими именами в DataGridView, если их еще нет
                for (int i = 0; i < matrixSize; i++)
                {
                    string columnName = $"A{i + 1}";
                    if (dataGridView.Columns[columnName] == null)
                    {
                        dataGridView.Columns.Add(columnName, columnName);
                    }
                }

                // Добавляем столбец "B", если его еще нет
                if (dataGridView.Columns["B"] == null)
                {
                    dataGridView.Columns.Add("B", "B");
                }

                // Генерируем новые данные и выводим их в DataGridView
                Random random = new Random();
                for (int i = 0; i < rowCount; i++)
                {
                    // Создаем массив данных для строки
                    object[] rowData = new object[matrixSize + 1];

                    // Генерация случайных значений для столбцов A1, A2, ..., B
                    for (int j = 0; j < matrixSize; j++)
                    {
                        string columnName = $"A{j + 1}";
                        double value = Math.Round(random.NextDouble() * 10, 2);
                        rowData[j] = value;
                    }

                    // Генерация случайного значения для столбца B
                    double y = Math.Round(random.NextDouble() * 10, 2);
                    rowData[matrixSize] = y;

                    // Добавляем строку с данными в DataGridView
                    dataGridView.Rows.Add(rowData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при генерации данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private double CalculateApproximation(Func<double, double> function, double x, int degree)
        {
            double result = 0;
            for (int i = 0; i <= degree; i++)
            {
                result += function(x) * Math.Pow(x, i);
            }
            return result;
        }

        private void ExcelButton_Click(object sender, EventArgs e)
        {
            LoadFromExcelAndApply();
        }

        private Tuple<string[,], int> ExportExcel()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Файлы Excel (Spisok.xlsx)|*.xlsx";

            if (!(ofd.ShowDialog() == DialogResult.OK))
                return Tuple.Create<string[,], int>(null, 0);

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            int lastRow = (int)lastCell.Row;
            int lastColumn = (int)lastCell.Column;

            // Изменяем размер массива list, чтобы вместить данные
            string[,] list = new string[lastRow, lastColumn];

            // Заполняем массив list данными из Excel
            for (int i = 0; i < lastRow; i++)
            {
                for (int j = 0; j < lastColumn; j++)
                {
                    // Преобразуем текст ячейки в строковом представлении в тип double
                    if (double.TryParse(ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString(),
                                        System.Globalization.NumberStyles.Any,
                                        CultureInfo.InvariantCulture,
                                        out double cellValue))
                    {
                        // Преобразуем значение ячейки в строку с учетом возможного минуса и дробной части
                        list[i, j] = cellValue.ToString("G2");
                    }
                    else
                    {
                        // Если не удалось преобразовать, сохраняем текст ячейки как есть
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();
                    }
                }
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            return Tuple.Create(list, lastRow);
        }

        private void LoadFromExcelAndApply()
        {
            Tuple<string[,], int> excelData = ExportExcel();

            if (excelData.Item2 > 0)
            {
                string[,] list = excelData.Item1;
                int rowCount = excelData.Item2;

                // Очищаем существующие данные в DataGridView и добавляем столбцы, если их нет
                ClearDataGridView();

                // Определяем размерность матрицы и вектора
                int matrixSize = list.GetLength(1) - 1; // Уменьшаем на 1, чтобы исключить последний столбец
                int vectorSize = rowCount;

                // Добавляем столбцы с соответствующими именами в DataGridView
                for (int i = 0; i < matrixSize; i++)
                {
                    dataGridView.Columns.Add($"A{i + 1}", $"A{i + 1}");
                }
                dataGridView.Columns.Add("B", "B");

                // Заполняем DataGridView данными из Excel
                for (int i = 0; i < rowCount; i++)
                {
                    object[] rowData = new object[matrixSize + 1];

                    for (int j = 0; j < matrixSize; j++)
                    {
                        if (double.TryParse(list[i, j], NumberStyles.Any, CultureInfo.InvariantCulture, out double cellValue))
                        {
                            rowData[j] = cellValue;
                        }
                        else
                        {
                            rowData[j] = list[i, j];
                        }
                    }

                    if (double.TryParse(list[i, matrixSize], NumberStyles.Any, CultureInfo.InvariantCulture, out double lastCellValue))
                    {
                        rowData[matrixSize] = lastCellValue;
                    }
                    else
                    {
                        rowData[matrixSize] = list[i, matrixSize];
                    }

                    dataGridView.Rows.Add(rowData);
                }
            }
        }


        private void ClearDataGridView()
        {
            dataGridView.Columns.Clear();
            dataGridView.Rows.Clear();
        }

        public struct PolynomialCoefficients
        {
            public double[] Coefficients;
        }
        private PolynomialCoefficients LeastSquaresMethod(int degree, double[,] data)
        {
            int rowCount = data.GetLength(0);
            int matrixSize = degree + 1;

            // Создаем матрицу A и вектор B для системы уравнений
            double[,] A = new double[matrixSize, matrixSize];
            double[] B = new double[matrixSize];

            // Заполняем матрицу A и вектор B
            for (int i = 0; i < matrixSize; i++)
            {
                for (int j = 0; j < matrixSize; j++)
                {
                    A[i, j] = 0;
                    for (int k = 0; k < rowCount; k++)
                    {
                        A[i, j] += Math.Pow(data[k, j], i);
                    }
                }

                B[i] = 0;
                for (int k = 0; k < rowCount; k++)
                {
                    B[i] += data[k, matrixSize - 1] * Math.Pow(data[k, i], i);
                }
            }

            // Решаем систему уравнений (Ax = B) для нахождения коэффициентов
            Matrix<double> matrixA = DenseMatrix.OfArray(A);
            Vector<double> vectorB = Vector<double>.Build.Dense(B);
            Vector<double> result = matrixA.Solve(vectorB);

            return new PolynomialCoefficients { Coefficients = result.ToArray() };
        }

        private double[] SolveSystemOfEquations(double[,] A, double[] B)
        {
            return new double[A.GetLength(0)];
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            textBoxN.Clear();
            textBox5.Clear();
            localPointsTextBox.Text = string.Empty;

            dataGridView.Columns.Clear();
            dataGridView.Rows.Clear();
        }

    }
}
