using OfficeOpenXml;
using System.ComponentModel;
using System.Text;

namespace ExcelToTxtConverterWinFormsApp
{
    public partial class MainForm : Form
    {
        private BackgroundWorker worker;
        public MainForm()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
        }
        private void InitializeBackgroundWorker()
        {
            worker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }
        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = dialog.FileName;
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!File.Exists(txtExcelPath.Text))
            {
                MessageBox.Show("Файл Excel не выбран или не существует!", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            progressBar.Value = 0;
            btnStart.Enabled = false;
            worker.RunWorkerAsync();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sourceExcelPath = txtExcelPath.Text;

                using (var package = new ExcelPackage(new FileInfo(sourceExcelPath), "38214"))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    var rawDate = worksheet.Cells["G2"].Text;
                    var mainDir = worksheet.Cells["I6"].Text;
                    var backupDir = worksheet.Cells["I7"].Text;

                    if (!DateTime.TryParseExact(rawDate, "dd.MM.yy", null,
                        System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                    {
                        throw new FormatException("Неверный формат даты в ячейке G2");
                    }

                    var fileName = $"KTMS0L00_{fileDate:yyyyMMdd}.txt";
                    var sb = new StringBuilder();

                    var totalRows = 40 - 3 + 1; // Расчет общего количества строк
                    var currentRow = 0;

                    for (int row = 3; row <= 40; row++)
                    {
                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                        var values = new List<string>();
                        foreach (int col in new[] { 3, 4, 5, 6, 7 })
                        {
                            values.Add(worksheet.Cells[row, col].Text?.Trim() ?? "");
                        }
                        sb.AppendLine(string.Join(";", values));

                        currentRow++;
                        int progress = (int)((double)currentRow / totalRows * 100);
                        worker.ReportProgress(progress);
                    }

                    SaveToFile(mainDir, fileName, sb.ToString());
                    SaveToFile(backupDir, fileName, sb.ToString());
                }
            }
            catch (Exception ex)
            {
                e.Result = ex;
            }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lblStatus.Text = $"Обработка... {e.ProgressPercentage}%";
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStart.Enabled = true;

            if (e.Error != null)
            {
                MessageBox.Show($"Ошибка: {e.Error.Message}", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                lblStatus.Text = "Операция отменена";
            }
            else if (e.Result is Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                lblStatus.Text = "Готово! Файлы успешно созданы";
                progressBar.Value = 100;
                MessageBox.Show("Операция завершена успешно!", "Успех",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SaveToFile(string directory, string fileName, string content)
        {
            Directory.CreateDirectory(directory);
            var fullPath = Path.Combine(directory, fileName);
            File.WriteAllText(fullPath, content, Encoding.UTF8);
        }
    }
}
