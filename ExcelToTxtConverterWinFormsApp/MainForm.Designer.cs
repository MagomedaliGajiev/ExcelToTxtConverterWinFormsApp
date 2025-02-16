

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
        private TextBox txtMainPath;
        private Button btnBrowseMainPath;
        private TextBox txtBackupPath;
        private Button btnBrowseBackupPath;
        private Label label2;
        private Label label3;
        private CheckBox chkBackup;

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
            txtExcelPath = new TextBox();
            btnBrowseExcel = new Button();
            btnStart = new Button();
            progressBar = new ProgressBar();
            lblStatus = new Label();
            label1 = new Label();
            txtMainPath = new TextBox();
            btnBrowseMainPath = new Button();
            txtBackupPath = new TextBox();
            btnBrowseBackupPath = new Button();
            label2 = new Label();
            label3 = new Label();
            chkBackup = new CheckBox();
            SuspendLayout();
            // 
            // txtExcelPath
            // 
            txtExcelPath.Location = new Point(12, 25);
            txtExcelPath.Name = "txtExcelPath";
            txtExcelPath.Size = new Size(300, 23);
            txtExcelPath.TabIndex = 0;
            // 
            // btnBrowseExcel
            // 
            btnBrowseExcel.Location = new Point(318, 23);
            btnBrowseExcel.Name = "btnBrowseExcel";
            btnBrowseExcel.Size = new Size(75, 23);
            btnBrowseExcel.TabIndex = 1;
            btnBrowseExcel.Text = "Обзор...";
            btnBrowseExcel.UseVisualStyleBackColor = true;
            btnBrowseExcel.Click += btnBrowseExcel_Click;
            // 
            // btnStart
            // 
            btnStart.Location = new Point(12, 252);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(381, 30);
            btnStart.TabIndex = 2;
            btnStart.Text = "Начать конвертацию";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(12, 288);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(381, 23);
            progressBar.TabIndex = 3;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(12, 327);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(45, 15);
            lblStatus.TabIndex = 4;
            lblStatus.Text = "Готово";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(125, 15);
            label1.TabIndex = 5;
            label1.Text = "Выберите файл Excel:";
            // 
            // txtMainPath
            // 
            txtMainPath.Location = new Point(12, 90);
            txtMainPath.Name = "txtMainPath";
            txtMainPath.Size = new Size(300, 23);
            txtMainPath.TabIndex = 6;
            // 
            // btnBrowseMainPath
            // 
            btnBrowseMainPath.Location = new Point(318, 88);
            btnBrowseMainPath.Name = "btnBrowseMainPath";
            btnBrowseMainPath.Size = new Size(75, 23);
            btnBrowseMainPath.TabIndex = 7;
            btnBrowseMainPath.Text = "Обзор...";
            btnBrowseMainPath.UseVisualStyleBackColor = true;
            btnBrowseMainPath.Click += btnBrowseMainPath_Click;
            // 
            // txtBackupPath
            // 
            txtBackupPath.Location = new Point(12, 170);
            txtBackupPath.Name = "txtBackupPath";
            txtBackupPath.Size = new Size(300, 23);
            txtBackupPath.TabIndex = 8;
            // 
            // btnBrowseBackupPath
            // 
            btnBrowseBackupPath.Location = new Point(318, 170);
            btnBrowseBackupPath.Name = "btnBrowseBackupPath";
            btnBrowseBackupPath.Size = new Size(75, 23);
            btnBrowseBackupPath.TabIndex = 9;
            btnBrowseBackupPath.Text = "Обзор...";
            btnBrowseBackupPath.UseVisualStyleBackColor = true;
            btnBrowseBackupPath.Click += btnBrowseBackupPath_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 74);
            label2.Name = "label2";
            label2.Size = new Size(166, 15);
            label2.TabIndex = 10;
            label2.Text = "Основная папка сохранения:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(12, 152);
            label3.Name = "label3";
            label3.Size = new Size(168, 15);
            label3.TabIndex = 11;
            label3.Text = "Резервная папка сохранения:";
            // 
            // chkBackup
            // 
            chkBackup.AutoSize = true;
            chkBackup.Checked = true;
            chkBackup.CheckState = CheckState.Checked;
            chkBackup.Location = new Point(12, 130);
            chkBackup.Name = "chkBackup";
            chkBackup.Size = new Size(184, 19);
            chkBackup.TabIndex = 12;
            chkBackup.Text = "Создавать резервную копию";
            chkBackup.UseVisualStyleBackColor = true;
            this.chkBackup.CheckedChanged += new System.EventHandler(this.chkBackup_CheckedChanged);
            // 
            // MainForm
            // 
            ClientSize = new Size(405, 364);
            Controls.Add(label1);
            Controls.Add(lblStatus);
            Controls.Add(progressBar);
            Controls.Add(btnStart);
            Controls.Add(btnBrowseExcel);
            Controls.Add(txtExcelPath);
            Controls.Add(chkBackup);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(btnBrowseBackupPath);
            Controls.Add(txtBackupPath);
            Controls.Add(btnBrowseMainPath);
            Controls.Add(txtMainPath);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Конвертер Excel в TXT";
            ResumeLayout(false);
            PerformLayout();
        }

        private void chkBackup_CheckedChanged(object sender, EventArgs e)
        {
            bool isBackupEnabled = chkBackup.Checked;

            // Включить/выключить связанные элементы
            txtBackupPath.Enabled = isBackupEnabled;
            btnBrowseBackupPath.Enabled = isBackupEnabled;

            if (!isBackupEnabled)
            {
                txtBackupPath.Text = string.Empty;
            }
        }

        #endregion
    }
}
