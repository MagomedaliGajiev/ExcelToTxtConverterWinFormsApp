
namespace ExcelToTxtConverterWinFormsApp
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private TextBox txtExcelPath;
        private Button btnBrowseExcel;
        private Button btnStart;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label label1;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();

            // Настройка элементов
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.Location = new System.Drawing.Point(12, 25);
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.Size = new System.Drawing.Size(300, 20);
            this.txtExcelPath.TabIndex = 0;
            // 
            // btnBrowseExcel
            // 
            this.btnBrowseExcel.Location = new System.Drawing.Point(318, 23);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseExcel.TabIndex = 1;
            this.btnBrowseExcel.Text = "Обзор...";
            this.btnBrowseExcel.UseVisualStyleBackColor = true;
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(12, 60);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(381, 30);
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "Начать конвертацию";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 100);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(381, 23);
            this.progressBar.TabIndex = 3;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 130);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(44, 13);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "Готово";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(133, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Выберите файл Excel:";
            // 
            // MainForm
            // 
            this.ClientSize = new System.Drawing.Size(405, 160);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnBrowseExcel);
            this.Controls.Add(this.txtExcelPath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертер Excel в TXT";


        }

        #endregion
    }
}
