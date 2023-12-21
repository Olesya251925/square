using System;
using System.Windows.Forms;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using OxyPlot.Axes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Globalization;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using System.Linq;

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
        private DataGridView dataGridView1;
        private PlotView plotView;
        private Label line;
        private TextBox textBox5;
        private Button applyButton;
        public MainForm()
        {
            InitializeComponent();
            this.Size = new System.Drawing.Size(1100, 800);
        }

        private void InitializeComponent()
        {
            labelN = new Label();
            labelN.AutoSize = true;
            labelN.Location = new System.Drawing.Point(50, 62);
            labelN.Location = new System.Drawing.Point(6, 13);
            labelN.Name = "label2";
            labelN.Size = new System.Drawing.Size(61, 13);
            labelN.TabIndex = 2;
            labelN.Text = "Введите n:";
            labelN.Text = "Введите степень:";

            textBoxN = new TextBox();
            textBoxN.Location = new System.Drawing.Point(96, 62);
            textBoxN.Location = new System.Drawing.Point(110, 13);
            textBoxN.Name = "textBox2";
            textBoxN.Size = new System.Drawing.Size(44, 20);
            textBoxN.TabIndex = 3;

            line = new Label();
            line.AutoSize = true;
            line.Location = new System.Drawing.Point(162, 65);
            line.Location = new System.Drawing.Point(6, 50);
            line.Name = "label3";
            line.Size = new System.Drawing.Size(73, 13);
            line.TabIndex = 11;
            line.Text = "Кол-во строк:";

            textBox5 = new TextBox();
            textBox5.Location = new System.Drawing.Point(100, 45);
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
            excelButton.Location = new System.Drawing.Point(5, 150);
            excelButton.Name = "button4";
            excelButton.Size = new System.Drawing.Size(120, 23);
            excelButton.TabIndex = 7;
            excelButton.Text = "Загрузить из EXCEL";
            excelButton.UseVisualStyleBackColor = true;
            excelButton.Click += ExcelButton_Click;
            
            applyButton = new Button();
            applyButton.Location = new System.Drawing.Point(150, 150);
            applyButton.Size = new System.Drawing.Size(120, 23);
            applyButton.TabIndex = 7;
            applyButton.Text = "Применить";
            applyButton.UseVisualStyleBackColor = true;
            applyButton.Click += ApplyButton_Click;

            localPointsTextBox = new TextBox();
            localPointsTextBox.Location = new System.Drawing.Point(180, 12);
            localPointsTextBox.Multiline = true;
            localPointsTextBox.Name = "textBox3";
            localPointsTextBox.Size = new System.Drawing.Size(250, 90);
            localPointsTextBox.TabIndex = 8;
            localPointsTextBox.UseWaitCursor = true;

            dataGridView = new DataGridView();
            dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new System.Drawing.Point(450, 12);
            dataGridView.Name = "dataGridView";
            dataGridView.Size = new System.Drawing.Size(245, 245);
            dataGridView.TabIndex = 9;

            dataGridView1 = new DataGridView();
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new System.Drawing.Point(700, 12);
            dataGridView1.Size = new System.Drawing.Size(340, 245);
            dataGridView1.TabIndex = 9;

            plotView = new PlotView();
            plotView.Location = new System.Drawing.Point(12, 255);
            plotView.Name = "plotView";
            plotView.PanCursor = System.Windows.Forms.Cursors.Hand;
            plotView.Size = new System.Drawing.Size(1050, 509);
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
            Controls.Add(dataGridView1);
            Controls.Add(line);
            Controls.Add(textBox5);
            Controls.Add(applyButton);
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

                int degree;
                if (!int.TryParse(textBoxN.Text, out degree) || degree < 1)
                {
                    MessageBox.Show("Введите корректное значение для степени полинома.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                double[,] data = new double[dataGridView.Rows.Count, 2];

                // Создаем матрицу A
                double[,] A = new double[degree + 1, degree + 1];

                // Заполняем матрицу данными из DataGridView
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    data[i, 0] = Convert.ToDouble(dataGridView.Rows[i].Cells["xi"].Value);
                    data[i, 1] = Convert.ToDouble(dataGridView.Rows[i].Cells["yi"].Value);
                }

                // Вызываем метод наименьших квадратов для нахождения коэффициентов
                PolynomialCoefficients coefficients = LeastSquaresMethod(degree, data);

                // Выводим коэффициенты функции в текстовое поле
                
                StringBuilder coefficientsString = new StringBuilder("Коэффициенты функции: \r\n");
                for (int i = 0; i < coefficients.Coefficients.Length; i++)
                {
                    coefficientsString.Append($"a{i} = {coefficients.Coefficients[coefficients.Coefficients.Length - 1 - i]:F2}\r\n");
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

        private void ApplyButton_Click(object sender, EventArgs e)
        {
            try
            {
                int rowCount;
                if (!int.TryParse(textBox5.Text, out rowCount) || rowCount < 0)
                {
                    MessageBox.Show("Введите корректное число строк.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Очищаем существующие данные в DataGridView и добавляем столбцы, если их нет
                ClearDataGridView();
                dataGridView.Columns.Clear();
                dataGridView.Rows.Clear();

                // Добавляем столбцы "xi" и "yi" в DataGridView
                dataGridView.Columns.Add("xi", "xi");
                dataGridView.Columns.Add("yi", "yi");

                // Задаем формат отображения ячеек с двумя знаками после запятой
                dataGridView.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.DefaultCellStyle.Format = "0.00");

                // Добавляем строки в DataGridView
                for (int i = 0; i < rowCount; i++)
                {
                    dataGridView.Rows.Add();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при применении данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                ClearDataGridView();
                dataGridView.Columns.Clear();
                dataGridView.Rows.Clear();

                // Определяем размерность матрицы и вектора
                int matrixSize = rowCount;

                // Добавляем столбцы "xi" и "yi" в DataGridView
                dataGridView.Columns.Add("xi", "xi");
                dataGridView.Columns.Add("yi", "yi");

                // Генерируем новые данные и выводим их в DataGridView
                Random random = new Random();
                for (int i = 0; i < rowCount; i++)
                {
                    // Создаем массив данных для строки
                    object[] rowData = new object[matrixSize + 1];

                    // Генерация случайных значений для столбцов "xi" и "yi"
                    double xi = Math.Round((random.NextDouble() - 0.5) * 20, 2); // Генерация от -10 до 10
                    double yi = Math.Round((random.NextDouble() - 0.5) * 20, 2); // Генерация от -10 до 10

                    rowData[0] = xi;
                    rowData[1] = yi;

                    // Добавляем строку с данными в DataGridView
                    dataGridView.Rows.Add(rowData);
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
                    new CultureInfo("ru-RU"),  // Используйте "ru-RU" вместо CultureInfo.InvariantCulture
                    out double cellValue))
                    {
                        // Преобразуем значение ячейки в строку с учетом возможного минуса и дробной части
                        list[i, j] = cellValue.ToString("G2", new CultureInfo("ru-RU"));
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
            try
            {
                ClearDataGridView();

                // Используем ExportExcel для получения данных из Excel
                var excelData = ExportExcel();

                // Проверяем, что данные успешно получены
                if (excelData.Item1 != null && excelData.Item2 > 0)
                {
                    // Задаем имена столбцов в DataGridView
                    dataGridView.Columns.Add("xi", "xi");
                    dataGridView.Columns.Add("yi", "yi");

                    // Заполняем DataGridView данными из Excel
                    for (int i = 0; i < excelData.Item2; i++)
                    {
                        object[] rowData = new object[2];
                        rowData[0] = excelData.Item1[i, 0];
                        rowData[1] = excelData.Item1[i, 1];

                        dataGridView.Rows.Add(rowData);
                    }
                }
                else
                {
                    MessageBox.Show("Не удалось загрузить данные из Excel.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных из Excel: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private PolynomialCoefficients LeastSquaresMethod(int degree, double[,] data)
        {
            // Определяем количество строк в массиве data
            int rowCount = data.GetLength(0) - 1;

            // Определяем размер матрицы A и вектора B, увеличивая на 1 степень полинома
            int matrixSize = degree + 1;

            // Создаем квадратную матрицу A для хранения коэффициентов системы линейных уравнений
            double[,] A = new double[matrixSize, matrixSize];

            // Создаем вектор B для хранения правой части системы линейных уравнений
            double[] B = new double[matrixSize];

            // Цикл для обхода строк матрицы A и вектора B
            for (int i = 0; i < matrixSize; i++)
            {
                // Вложенный цикл для заполнения значений в строке матрицы A
                for (int j = 0; j < matrixSize; j++)
                {
                    // Задаем начальное значение в матрице A
                    A[i, j] = 0;

                    // Вложенный цикл для вычисления суммы степеней data[k, 0] в соответствии с формулой
                    for (int k = 0; k < rowCount; k++)
                    {
                        A[i, j] += Math.Pow(data[k, 0], matrixSize - 1 - i + matrixSize - 1 - j);
                    }
                }

                // Задаем начальное значение в векторе B
                B[i] = 0;

                // Вложенный цикл для вычисления суммы значений вектора B
                for (int k = 0; k < rowCount; k++)
                {
                    B[i] += data[k, 1] * Math.Pow(data[k, 0], matrixSize - 1 - i);
                }
            }


            // Выводим матрицу A и вектор B в DataGridView для визуализации
            FillDataGridViewFromMatrixAndVector("A", A, "B", B);

            // Решаем систему линейных уравнений методом наименьших квадратов
            double[] coefficients = SolveLeastSquares(A, B);

            // Округляем коэффициенты и изменяем порядок, чтобы они соответствовали порядку степеней полинома
            int decimalPlaces = 7;
            coefficients = coefficients.Reverse().Select(c => Math.Round(c, decimalPlaces)).ToArray();

            // Возвращаем структуру PolynomialCoefficients с найденными коэффициентами
            return new PolynomialCoefficients { Coefficients = coefficients };
        }

        public struct PolynomialCoefficients
        {
            public double[] Coefficients;
        }

        // Метод для решения системы линейных уравнений методом наименьших квадратов
        private double[] SolveLeastSquares(double[,] matrixA, double[] vectorB)
        {
            // Преобразуем массивы в объекты библиотеки MathNet.Numerics
            var A = Matrix<double>.Build.DenseOfArray(matrixA);
            var B = Vector<double>.Build.Dense(vectorB);

            // Используем сингулярное разложение (SVD) для решения системы линейных уравнений
            double[] coefficients = A.Svd().Solve(B).ToArray();

            return coefficients;
        }
        private void FillDataGridViewFromMatrixAndVector(string matrixTitle, double[,] matrix, string vectorTitle, double[] vector)
        {
            // Очищаем dataGridView1 перед заполнением новыми данными
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // Создаем новую колонку для каждого столбца матрицы
            for (int j = 0; j < matrix.GetLength(1); j++)
            {
                dataGridView1.Columns.Add($"{matrixTitle}{j}", $"{matrixTitle}{j}");
            }

            // Добавляем колонку для вектора
            dataGridView1.Columns.Add(vectorTitle, vectorTitle);

            // Задаем формат отображения ячеек с двумя знаками после запятой
            dataGridView1.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.DefaultCellStyle.Format = "0.00");

            // Заполняем dataGridView1 данными из матрицы и вектора
            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView1);

                for (int j = 0; j < matrix.GetLength(1); j++)
                {
                    row.Cells[j].Value = matrix[i, j];
                }

                // Заполняем dataGridView1 данными из вектора
                row.Cells[matrix.GetLength(1)].Value = vector[i];

                dataGridView1.Rows.Add(row);
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            textBoxN.Clear();
            textBox5.Clear();
            localPointsTextBox.Text = string.Empty;

            dataGridView.Columns.Clear();
            dataGridView.Rows.Clear();

            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
        }

    }
}
