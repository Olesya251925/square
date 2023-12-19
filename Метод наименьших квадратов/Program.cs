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
        private Label labelFunction;
        private TextBox textBoxFunction;
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
            labelFunction = new Label();
            labelFunction.AutoSize = true;
            labelFunction.Location = new System.Drawing.Point(9, 26);
            labelFunction.Name = "labelFunction";
            labelFunction.Size = new System.Drawing.Size(100, 13);
            labelFunction.TabIndex = 0;
            labelFunction.Text = "Введите функцию:";

            textBoxFunction = new TextBox();
            textBoxFunction.Location = new System.Drawing.Point(115, 23);
            textBoxFunction.Name = "textBox1";
            textBoxFunction.Size = new System.Drawing.Size(134, 20);
            textBoxFunction.TabIndex = 1;

            labelN = new Label();
            labelN.AutoSize = true;
            labelN.Location = new System.Drawing.Point(12, 62);
            labelN.Name = "label2";
            labelN.Size = new System.Drawing.Size(61, 13);
            labelN.TabIndex = 2;
            labelN.Text = "Введите n:";

            textBoxN = new TextBox();
            textBoxN.Location = new System.Drawing.Point(96, 62);
            textBoxN.Name = "textBox2";
            textBoxN.Size = new System.Drawing.Size(44, 20);
            textBoxN.TabIndex = 3;
            
            line = new Label();
            line.AutoSize = true;
            line.Location = new System.Drawing.Point(162, 65);
            line.Name = "label3";
            line.Size = new System.Drawing.Size(73, 13);
            line.TabIndex = 11;
            line.Text = "Кол-во строк";

            textBox5 = new TextBox();
            textBox5.Location = new System.Drawing.Point(242, 61);
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
            dataGridView.Size = new System.Drawing.Size(268, 237);
            dataGridView.TabIndex = 9;

            plotView = new PlotView();
            plotView.Location = new System.Drawing.Point(12, 255);
            plotView.Name = "plotView";
            plotView.PanCursor = System.Windows.Forms.Cursors.Hand;
            plotView.Size = new System.Drawing.Size(868, 409);
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
            Controls.Add(textBoxFunction);
            Controls.Add(labelFunction);
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

                double[,] data = new double[dataGridView.Rows.Count, 2];
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    data[i, 0] = Convert.ToDouble(dataGridView.Rows[i].Cells["XColumn"].Value);
                    data[i, 1] = Convert.ToDouble(dataGridView.Rows[i].Cells["YColumn"].Value);
                }

                // Вызываем метод наименьших квадратов для нахождения коэффициентов
                PolynomialCoefficients coefficients = LeastSquaresMethod(degree, data);

                // Выводим коэффициенты функции в текстовое поле
                StringBuilder coefficientsString = new StringBuilder("Коэффициенты функции: \r\n");
                for (int i = 0; i < coefficients.Coefficients.Length; i++)
                {
                    coefficientsString.Append($"a{i} = {coefficients.Coefficients[i]:F2}\r\n  ");
                }

                localPointsTextBox.Text = coefficientsString.ToString();

                // Строим график
                PlotModel plotModel = new PlotModel();

                // Добавляем точки из DataGridView на график
                ScatterSeries scatterSeries = new ScatterSeries { MarkerType = MarkerType.Circle };
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    scatterSeries.Points.Add(new ScatterPoint(data[i, 0], data[i, 1]));
                }
                plotModel.Series.Add(scatterSeries);

                // Находим минимальное и максимальное значения X для построения функции
                double minX = data[0, 0];
                double maxX = data[0, 0];
                for (int i = 1; i < data.GetLength(0); i++)
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

        private Func<double, double> ParseFunction(string functionText)
        {
            // Реализуйте парсинг функции и возвращение делегата Func<double, double>
            try
            {
                // Пример парсинга, замените его на реальный код
                SymbolicExpression expression = SymbolicExpression.Parse(functionText);
                return x => Convert.ToDouble(expression.Compile("x").Invoke(x).ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка при парсинге функции: " + ex.Message);
            }
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
                ClearDataGridView();

                if (dataGridView.Columns.Count == 0)
                {
                    dataGridView.Columns.Add("XColumn", "X");
                    dataGridView.Columns.Add("YColumn", "Y");
                }

                // Генерируем новые данные и выводим их в DataGridView
                Random random = new Random();
                for (int i = 0; i < rowCount; i++)
                {
                    double x = Math.Round(random.NextDouble() * 10, 2); // Генерация случайных значений X с двумя знаками после запятой
                    double y = Math.Round(random.NextDouble() * 10, 2); // Генерация случайных значений Y с двумя знаками после запятой

                    dataGridView.Rows.Add(x, y); // Добавляем строку с x и y
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при генерации данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearDataGridView()
        {
            // Очищаем значения в столбцах XColumn и YColumn в DataGridView
            dataGridView.Rows.Clear();
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

                // Очищаем существующие данные в DataGridView
                ClearDataGridView();

                // Заполняем DataGridView данными из Excel
                for (int i = 0; i < rowCount; i++)
                {
                    dataGridView.Rows.Add(list[i, 0], list[i, 1]); // Здесь предполагается, что у вас есть столбцы "XColumn" и "YColumn"
                }
            }
        }

        public struct PolynomialCoefficients
        {
            public double[] Coefficients;
        }
        private PolynomialCoefficients LeastSquaresMethod(int degree, double[,] data)
        {
            int rowCount = data.GetLength(0);

            // Создаем матрицу A и вектор B для системы уравнений
            double[,] A = new double[degree + 1, degree + 1];
            double[] B = new double[degree + 1];

            // Заполняем матрицу A и вектор B
            for (int i = 0; i <= degree; i++)
            {
                for (int j = 0; j <= degree; j++)
                {
                    A[i, j] = 0;
                    for (int k = 0; k < rowCount; k++)
                    {
                        A[i, j] += Math.Pow(data[k, 0], i + j);
                    }
                }

                B[i] = 0;
                for (int k = 0; k < rowCount; k++)
                {
                    B[i] += data[k, 1] * Math.Pow(data[k, 0], i);
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
            textBoxFunction.Clear();
            localPointsTextBox.Text = string.Empty;

            // Очищаем значения в столбцах "XColumn" и "YColumn" в DataGridView
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.Cells["XColumn"].Value = null;
                row.Cells["YColumn"].Value = null;
            }
        }

    }
}
