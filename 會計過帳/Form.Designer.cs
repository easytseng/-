using System;

namespace 會計過帳
{
    partial class Form
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.openXlsFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.label_cashBook = new System.Windows.Forms.Label();
            this.label_fileSetting = new System.Windows.Forms.Label();
            this.button_CashBook = new System.Windows.Forms.Button();
            this.linkLabel_cashBook = new System.Windows.Forms.LinkLabel();
            this.linkLabel_balanceSheet = new System.Windows.Forms.LinkLabel();
            this.button_balanceSheet = new System.Windows.Forms.Button();
            this.label_balanceSheet = new System.Windows.Forms.Label();
            this.comboBox_balanceSheet = new System.Windows.Forms.ComboBox();
            this.label_balanceSheetSelect = new System.Windows.Forms.Label();
            this.comboBox_incomeStatement = new System.Windows.Forms.ComboBox();
            this.label_incomeStatementSelect = new System.Windows.Forms.Label();
            this.linkLabel_incomeStatement = new System.Windows.Forms.LinkLabel();
            this.button_incomeStatement = new System.Windows.Forms.Button();
            this.label_incomeStatement = new System.Windows.Forms.Label();
            this.comboBox_payTransfer = new System.Windows.Forms.ComboBox();
            this.label_payTransferSelect = new System.Windows.Forms.Label();
            this.linkLabel_payTransfer = new System.Windows.Forms.LinkLabel();
            this.button_payTransfer = new System.Windows.Forms.Button();
            this.label_payTransfer = new System.Windows.Forms.Label();
            this.comboBox_accountReceivable = new System.Windows.Forms.ComboBox();
            this.label_accountReceivableSelect = new System.Windows.Forms.Label();
            this.linkLabel_accountReceivable = new System.Windows.Forms.LinkLabel();
            this.button_accountReceivable = new System.Windows.Forms.Button();
            this.label_accountReceivable = new System.Windows.Forms.Label();
            this.comboBox_year = new System.Windows.Forms.ComboBox();
            this.label_year = new System.Windows.Forms.Label();
            this.comboBox_month = new System.Windows.Forms.ComboBox();
            this.label_month = new System.Windows.Forms.Label();
            this.button_start = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openXlsFileDialog
            // 
            this.openXlsFileDialog.Filter = "Text files (*.xls)|*.xls";
            this.openXlsFileDialog.Title = "Open xls file";
            // 
            // label_cashBook
            // 
            this.label_cashBook.AutoSize = true;
            this.label_cashBook.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_cashBook.Location = new System.Drawing.Point(34, 59);
            this.label_cashBook.Name = "label_cashBook";
            this.label_cashBook.Size = new System.Drawing.Size(66, 19);
            this.label_cashBook.TabIndex = 0;
            this.label_cashBook.Text = "現金簿";
            // 
            // label_fileSetting
            // 
            this.label_fileSetting.AutoSize = true;
            this.label_fileSetting.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_fileSetting.Location = new System.Drawing.Point(34, 19);
            this.label_fileSetting.Name = "label_fileSetting";
            this.label_fileSetting.Size = new System.Drawing.Size(106, 24);
            this.label_fileSetting.TabIndex = 1;
            this.label_fileSetting.Text = "檔案設定";
            // 
            // button_CashBook
            // 
            this.button_CashBook.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_CashBook.Location = new System.Drawing.Point(152, 54);
            this.button_CashBook.Name = "button_CashBook";
            this.button_CashBook.Size = new System.Drawing.Size(134, 29);
            this.button_CashBook.TabIndex = 2;
            this.button_CashBook.Text = "請選擇檔案";
            this.button_CashBook.UseVisualStyleBackColor = true;
            this.button_CashBook.Click += new System.EventHandler(this.buttonCashBook_Click);
            // 
            // linkLabel_cashBook
            // 
            this.linkLabel_cashBook.AutoSize = true;
            this.linkLabel_cashBook.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.linkLabel_cashBook.Location = new System.Drawing.Point(315, 61);
            this.linkLabel_cashBook.Name = "linkLabel_cashBook";
            this.linkLabel_cashBook.Size = new System.Drawing.Size(0, 16);
            this.linkLabel_cashBook.TabIndex = 3;
            this.linkLabel_cashBook.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_cashBook_LinkClicked);
            // 
            // linkLabel_balanceSheet
            // 
            this.linkLabel_balanceSheet.AutoSize = true;
            this.linkLabel_balanceSheet.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.linkLabel_balanceSheet.Location = new System.Drawing.Point(315, 111);
            this.linkLabel_balanceSheet.Name = "linkLabel_balanceSheet";
            this.linkLabel_balanceSheet.Size = new System.Drawing.Size(0, 16);
            this.linkLabel_balanceSheet.TabIndex = 8;
            this.linkLabel_balanceSheet.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_balanceSheet_LinkClicked);
            // 
            // button_balanceSheet
            // 
            this.button_balanceSheet.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_balanceSheet.Location = new System.Drawing.Point(152, 104);
            this.button_balanceSheet.Name = "button_balanceSheet";
            this.button_balanceSheet.Size = new System.Drawing.Size(134, 29);
            this.button_balanceSheet.TabIndex = 7;
            this.button_balanceSheet.Text = "請選擇檔案";
            this.button_balanceSheet.UseVisualStyleBackColor = true;
            this.button_balanceSheet.Click += new System.EventHandler(this.button_balanceSheet_Click);
            // 
            // label_balanceSheet
            // 
            this.label_balanceSheet.AutoSize = true;
            this.label_balanceSheet.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_balanceSheet.Location = new System.Drawing.Point(34, 109);
            this.label_balanceSheet.Name = "label_balanceSheet";
            this.label_balanceSheet.Size = new System.Drawing.Size(104, 19);
            this.label_balanceSheet.TabIndex = 6;
            this.label_balanceSheet.Text = "資產負債表";
            // 
            // comboBox_balanceSheet
            // 
            this.comboBox_balanceSheet.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox_balanceSheet.FormattingEnabled = true;
            this.comboBox_balanceSheet.Location = new System.Drawing.Point(629, 106);
            this.comboBox_balanceSheet.Name = "comboBox_balanceSheet";
            this.comboBox_balanceSheet.Size = new System.Drawing.Size(159, 27);
            this.comboBox_balanceSheet.TabIndex = 10;
            // 
            // label_balanceSheetSelect
            // 
            this.label_balanceSheetSelect.AutoSize = true;
            this.label_balanceSheetSelect.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_balanceSheetSelect.Location = new System.Drawing.Point(488, 109);
            this.label_balanceSheetSelect.Name = "label_balanceSheetSelect";
            this.label_balanceSheetSelect.Size = new System.Drawing.Size(123, 19);
            this.label_balanceSheetSelect.TabIndex = 9;
            this.label_balanceSheetSelect.Text = "請選擇工作頁";
            // 
            // comboBox_incomeStatement
            // 
            this.comboBox_incomeStatement.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox_incomeStatement.FormattingEnabled = true;
            this.comboBox_incomeStatement.Location = new System.Drawing.Point(629, 160);
            this.comboBox_incomeStatement.Name = "comboBox_incomeStatement";
            this.comboBox_incomeStatement.Size = new System.Drawing.Size(159, 27);
            this.comboBox_incomeStatement.TabIndex = 15;
            // 
            // label_incomeStatementSelect
            // 
            this.label_incomeStatementSelect.AutoSize = true;
            this.label_incomeStatementSelect.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_incomeStatementSelect.Location = new System.Drawing.Point(488, 163);
            this.label_incomeStatementSelect.Name = "label_incomeStatementSelect";
            this.label_incomeStatementSelect.Size = new System.Drawing.Size(123, 19);
            this.label_incomeStatementSelect.TabIndex = 14;
            this.label_incomeStatementSelect.Text = "請選擇工作頁";
            // 
            // linkLabel_incomeStatement
            // 
            this.linkLabel_incomeStatement.AutoSize = true;
            this.linkLabel_incomeStatement.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.linkLabel_incomeStatement.Location = new System.Drawing.Point(315, 165);
            this.linkLabel_incomeStatement.Name = "linkLabel_incomeStatement";
            this.linkLabel_incomeStatement.Size = new System.Drawing.Size(0, 16);
            this.linkLabel_incomeStatement.TabIndex = 13;
            this.linkLabel_incomeStatement.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_incomeStatement_LinkClicked);
            // 
            // button_incomeStatement
            // 
            this.button_incomeStatement.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_incomeStatement.Location = new System.Drawing.Point(152, 158);
            this.button_incomeStatement.Name = "button_incomeStatement";
            this.button_incomeStatement.Size = new System.Drawing.Size(134, 29);
            this.button_incomeStatement.TabIndex = 12;
            this.button_incomeStatement.Text = "請選擇檔案";
            this.button_incomeStatement.UseVisualStyleBackColor = true;
            this.button_incomeStatement.Click += new System.EventHandler(this.button_incomeStatement_Click);
            // 
            // label_incomeStatement
            // 
            this.label_incomeStatement.AutoSize = true;
            this.label_incomeStatement.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_incomeStatement.Location = new System.Drawing.Point(34, 163);
            this.label_incomeStatement.Name = "label_incomeStatement";
            this.label_incomeStatement.Size = new System.Drawing.Size(66, 19);
            this.label_incomeStatement.TabIndex = 11;
            this.label_incomeStatement.Text = "損益表";
            // 
            // comboBox_payTransfer
            // 
            this.comboBox_payTransfer.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox_payTransfer.FormattingEnabled = true;
            this.comboBox_payTransfer.Location = new System.Drawing.Point(629, 211);
            this.comboBox_payTransfer.Name = "comboBox_payTransfer";
            this.comboBox_payTransfer.Size = new System.Drawing.Size(159, 27);
            this.comboBox_payTransfer.TabIndex = 20;
            // 
            // label_payTransferSelect
            // 
            this.label_payTransferSelect.AutoSize = true;
            this.label_payTransferSelect.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_payTransferSelect.Location = new System.Drawing.Point(488, 214);
            this.label_payTransferSelect.Name = "label_payTransferSelect";
            this.label_payTransferSelect.Size = new System.Drawing.Size(123, 19);
            this.label_payTransferSelect.TabIndex = 19;
            this.label_payTransferSelect.Text = "請選擇工作頁";
            // 
            // linkLabel_payTransfer
            // 
            this.linkLabel_payTransfer.AutoSize = true;
            this.linkLabel_payTransfer.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.linkLabel_payTransfer.Location = new System.Drawing.Point(315, 216);
            this.linkLabel_payTransfer.Name = "linkLabel_payTransfer";
            this.linkLabel_payTransfer.Size = new System.Drawing.Size(0, 16);
            this.linkLabel_payTransfer.TabIndex = 18;
            this.linkLabel_payTransfer.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_payTransfer_LinkClicked);
            // 
            // button_payTransfer
            // 
            this.button_payTransfer.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_payTransfer.Location = new System.Drawing.Point(152, 209);
            this.button_payTransfer.Name = "button_payTransfer";
            this.button_payTransfer.Size = new System.Drawing.Size(134, 29);
            this.button_payTransfer.TabIndex = 17;
            this.button_payTransfer.Text = "請選擇檔案";
            this.button_payTransfer.UseVisualStyleBackColor = true;
            this.button_payTransfer.Click += new System.EventHandler(this.button_payTransfer_Click);
            // 
            // label_payTransfer
            // 
            this.label_payTransfer.AutoSize = true;
            this.label_payTransfer.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_payTransfer.Location = new System.Drawing.Point(34, 214);
            this.label_payTransfer.Name = "label_payTransfer";
            this.label_payTransfer.Size = new System.Drawing.Size(85, 19);
            this.label_payTransfer.TabIndex = 16;
            this.label_payTransfer.Text = "代付轉帳";
            // 
            // comboBox_accountReceivable
            // 
            this.comboBox_accountReceivable.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox_accountReceivable.FormattingEnabled = true;
            this.comboBox_accountReceivable.Location = new System.Drawing.Point(629, 264);
            this.comboBox_accountReceivable.Name = "comboBox_accountReceivable";
            this.comboBox_accountReceivable.Size = new System.Drawing.Size(159, 27);
            this.comboBox_accountReceivable.TabIndex = 25;
            // 
            // label_accountReceivableSelect
            // 
            this.label_accountReceivableSelect.AutoSize = true;
            this.label_accountReceivableSelect.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_accountReceivableSelect.Location = new System.Drawing.Point(488, 267);
            this.label_accountReceivableSelect.Name = "label_accountReceivableSelect";
            this.label_accountReceivableSelect.Size = new System.Drawing.Size(123, 19);
            this.label_accountReceivableSelect.TabIndex = 24;
            this.label_accountReceivableSelect.Text = "請選擇工作頁";
            // 
            // linkLabel_accountReceivable
            // 
            this.linkLabel_accountReceivable.AutoSize = true;
            this.linkLabel_accountReceivable.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.linkLabel_accountReceivable.Location = new System.Drawing.Point(315, 269);
            this.linkLabel_accountReceivable.Name = "linkLabel_accountReceivable";
            this.linkLabel_accountReceivable.Size = new System.Drawing.Size(0, 16);
            this.linkLabel_accountReceivable.TabIndex = 23;
            this.linkLabel_accountReceivable.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_accountReceivable_LinkClicked);
            // 
            // button_accountReceivable
            // 
            this.button_accountReceivable.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_accountReceivable.Location = new System.Drawing.Point(152, 262);
            this.button_accountReceivable.Name = "button_accountReceivable";
            this.button_accountReceivable.Size = new System.Drawing.Size(134, 29);
            this.button_accountReceivable.TabIndex = 22;
            this.button_accountReceivable.Text = "請選擇檔案";
            this.button_accountReceivable.UseVisualStyleBackColor = true;
            this.button_accountReceivable.Click += new System.EventHandler(this.button_accountReceivable_Click);
            // 
            // label_accountReceivable
            // 
            this.label_accountReceivable.AutoSize = true;
            this.label_accountReceivable.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_accountReceivable.Location = new System.Drawing.Point(34, 267);
            this.label_accountReceivable.Name = "label_accountReceivable";
            this.label_accountReceivable.Size = new System.Drawing.Size(85, 19);
            this.label_accountReceivable.TabIndex = 21;
            this.label_accountReceivable.Text = "應收轉帳";
            // 
            // comboBox_year
            // 
            this.comboBox_year.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox_year.FormattingEnabled = true;
            this.comboBox_year.Location = new System.Drawing.Point(152, 326);
            this.comboBox_year.Name = "comboBox_year";
            this.comboBox_year.Size = new System.Drawing.Size(134, 27);
            this.comboBox_year.TabIndex = 27;
            // 
            // label_year
            // 
            this.label_year.AutoSize = true;
            this.label_year.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_year.Location = new System.Drawing.Point(34, 329);
            this.label_year.Name = "label_year";
            this.label_year.Size = new System.Drawing.Size(85, 19);
            this.label_year.TabIndex = 26;
            this.label_year.Text = "財報年份";
            // 
            // comboBox_month
            // 
            this.comboBox_month.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox_month.FormattingEnabled = true;
            this.comboBox_month.Location = new System.Drawing.Point(152, 383);
            this.comboBox_month.Name = "comboBox_month";
            this.comboBox_month.Size = new System.Drawing.Size(134, 27);
            this.comboBox_month.TabIndex = 29;
            // 
            // label_month
            // 
            this.label_month.AutoSize = true;
            this.label_month.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label_month.Location = new System.Drawing.Point(34, 386);
            this.label_month.Name = "label_month";
            this.label_month.Size = new System.Drawing.Size(85, 19);
            this.label_month.TabIndex = 28;
            this.label_month.Text = "財報月份";
            // 
            // button_start
            // 
            this.button_start.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button_start.Location = new System.Drawing.Point(629, 375);
            this.button_start.Name = "button_start";
            this.button_start.Size = new System.Drawing.Size(146, 39);
            this.button_start.TabIndex = 30;
            this.button_start.Text = "產生報表";
            this.button_start.UseVisualStyleBackColor = true;
            this.button_start.Click += new System.EventHandler(this.button_start_Click);
            // 
            // Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button_start);
            this.Controls.Add(this.comboBox_month);
            this.Controls.Add(this.label_month);
            this.Controls.Add(this.comboBox_year);
            this.Controls.Add(this.label_year);
            this.Controls.Add(this.comboBox_accountReceivable);
            this.Controls.Add(this.label_accountReceivableSelect);
            this.Controls.Add(this.linkLabel_accountReceivable);
            this.Controls.Add(this.button_accountReceivable);
            this.Controls.Add(this.label_accountReceivable);
            this.Controls.Add(this.comboBox_payTransfer);
            this.Controls.Add(this.label_payTransferSelect);
            this.Controls.Add(this.linkLabel_payTransfer);
            this.Controls.Add(this.button_payTransfer);
            this.Controls.Add(this.label_payTransfer);
            this.Controls.Add(this.comboBox_incomeStatement);
            this.Controls.Add(this.label_incomeStatementSelect);
            this.Controls.Add(this.linkLabel_incomeStatement);
            this.Controls.Add(this.button_incomeStatement);
            this.Controls.Add(this.label_incomeStatement);
            this.Controls.Add(this.comboBox_balanceSheet);
            this.Controls.Add(this.label_balanceSheetSelect);
            this.Controls.Add(this.linkLabel_balanceSheet);
            this.Controls.Add(this.button_balanceSheet);
            this.Controls.Add(this.label_balanceSheet);
            this.Controls.Add(this.linkLabel_cashBook);
            this.Controls.Add(this.button_CashBook);
            this.Controls.Add(this.label_fileSetting);
            this.Controls.Add(this.label_cashBook);
            this.Name = "Form";
            this.Text = "Form";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openXlsFileDialog;
        private System.Windows.Forms.Label label_fileSetting;

        private System.Windows.Forms.Label label_cashBook;
        private System.Windows.Forms.Label label_balanceSheet;
        private System.Windows.Forms.Label label_incomeStatement;
        private System.Windows.Forms.Label label_payTransfer;
        private System.Windows.Forms.Label label_accountReceivable;

        private System.Windows.Forms.Button button_CashBook;
        private System.Windows.Forms.Button button_balanceSheet;
        private System.Windows.Forms.Button button_incomeStatement;
        private System.Windows.Forms.Button button_payTransfer;
        private System.Windows.Forms.Button button_accountReceivable;

        private System.Windows.Forms.LinkLabel linkLabel_cashBook;
        private System.Windows.Forms.LinkLabel linkLabel_balanceSheet;
        private System.Windows.Forms.LinkLabel linkLabel_incomeStatement;
        private System.Windows.Forms.LinkLabel linkLabel_payTransfer;
        private System.Windows.Forms.LinkLabel linkLabel_accountReceivable;

        private System.Windows.Forms.Label label_balanceSheetSelect;
        private System.Windows.Forms.Label label_incomeStatementSelect;
        private System.Windows.Forms.Label label_payTransferSelect;
        private System.Windows.Forms.Label label_accountReceivableSelect;

        private System.Windows.Forms.ComboBox comboBox_balanceSheet;
        private System.Windows.Forms.ComboBox comboBox_incomeStatement;
        private System.Windows.Forms.ComboBox comboBox_payTransfer;
        private System.Windows.Forms.ComboBox comboBox_accountReceivable;

        private System.Windows.Forms.ComboBox comboBox_year;
        private System.Windows.Forms.Label label_year;
        private System.Windows.Forms.ComboBox comboBox_month;
        private System.Windows.Forms.Label label_month;
        private System.Windows.Forms.Button button_start;
    }
}

