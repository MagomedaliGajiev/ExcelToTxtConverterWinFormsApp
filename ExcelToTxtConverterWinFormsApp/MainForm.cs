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
            InitializePathControls();
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

        private void InitializePathControls()
        {
            // ������������� �������� �� ���������
            txtMainPath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "KTMS");
            txtBackupPath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "KTMS_Backup");
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

        private void btnBrowseMainPath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "�������� �������� ����� ��� ����������";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtMainPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void btnBrowseBackupPath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "�������� ��������� ����� ��� ����������";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtBackupPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            progressBar.Value = 0;
            btnStart.Enabled = false;
            worker.RunWorkerAsync();
        }

        private bool ValidateInputs()
        {
            if (!File.Exists(txtExcelPath.Text))
            {
                MessageBox.Show("���� Excel �� ������ ��� �� ����������!", "������",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtMainPath.Text))
            {
                MessageBox.Show("�������� ����� ��� ���������� �� �������!", "������",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                string sourceExcelPath = txtExcelPath.Text;
                string mainDir = txtMainPath.Text;
                string backupDir = txtBackupPath.Text;

                using (var package = new ExcelPackage(new FileInfo(sourceExcelPath), "38214"))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    string rawDate = worksheet.Cells["G2"].Text;

                    if (!DateTime.TryParseExact(rawDate, "dd.MM.yy", null,
                        System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                    {
                        throw new FormatException("�������� ������ ���� � ������ G2");
                    }

                    string fileName = $"KTMS0L00_{fileDate:yyyyMMdd}.txt";
                    var sb = new StringBuilder();

                    int totalRows = 40 - 3 + 1;
                    int currentRow = 0;

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

                    if (!string.IsNullOrWhiteSpace(backupDir))
                    {
                        SaveToFile(backupDir, fileName, sb.ToString());
                    }
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
            lblStatus.Text = $"���������... {e.ProgressPercentage}%";
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStart.Enabled = true;

            if (e.Error != null)
            {
                MessageBox.Show($"������: {e.Error.Message}", "������",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                lblStatus.Text = "�������� ��������";
            }
            else if (e.Result is Exception ex)
            {
                MessageBox.Show($"������: {ex.Message}", "������",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                lblStatus.Text = "������! ����� ������� �������";
                progressBar.Value = 100;
                MessageBox.Show("�������� ��������� �������!", "�����",
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
