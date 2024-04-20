using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace 會計過帳
{
    public partial class Form : System.Windows.Forms.Form
    {
        private string TAG_CASH_BOOK = "CashBook";
        private string TAG_BALANCE_SHEET = "BalanceSheet";
        private string TAG_INCOME_STATEMENT = "IncomeStatement";
        private string TAG_PAY_TRANSFER = "PayTransfer";
        private string TAG_ACCOUNT_RECEIVABLE = "AccountReceivable";

        private Microsoft.Office.Interop.Excel.Application xlApp;

        private string directoryName = string.Empty;
        private string spreadSheet1Path = string.Empty;
        private string spreadSheet2Path = string.Empty;
        private string balanceSheetResultPath = string.Empty;
        private string incomeStatementResultPath = string.Empty;

        private string cashBookPath = string.Empty;
        private string balanceSheetPath = string.Empty;
        private string payTransferPath = string.Empty;
        private string accountReceivablePath = string.Empty;
        private string incomeStatementPath = string.Empty;

        private string scanAccountDrRange = ConfigurationManager.AppSettings.Get("ScanAccountDrRange");
        private string scanAccountCrRange = ConfigurationManager.AppSettings.Get("ScanAccountCrRange");
        private string[] drVoucherAccountList = ConfigurationManager.AppSettings.Get("DrVoucherAccountList").Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
        private string[] crVoucherAccountList = ConfigurationManager.AppSettings.Get("CrVoucherAccountList").Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
        private Boolean isDebug = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("isDebug"));

        private Regex dateRegex = new Regex(@"^\d{4}\/\d{1,2}\/\d{2}$", RegexOptions.IgnorePatternWhitespace);
        private Regex digitalRegex = new Regex(@"^\d+$");

        private string defaulfPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public Form()
        {
            InitializeComponent();

            initComboBox();

            xlApp = new Microsoft.Office.Interop.Excel.Application();


        }

        public void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            xlApp.Quit();

            foreach (Process clsProcess in Process.GetProcesses())
            {
                try
                {
                    if (clsProcess.ProcessName.Equals("EXCEL"))
                    {
                        clsProcess.Kill();
                        break;
                    }
                }
                catch (Exception ex)
                {

                }
            }
        }

        #region File Select Button
        private void buttonCashBook_Click(object sender, EventArgs e)
        {
            openXlsFileDialog.Tag = TAG_CASH_BOOK;

            openFileSelectDialog();
        }

        private void button_balanceSheet_Click(object sender, EventArgs e)
        {
            openXlsFileDialog.Tag = TAG_BALANCE_SHEET;

            openFileSelectDialog();
        }

        private void button_incomeStatement_Click(object sender, EventArgs e)
        {
            openXlsFileDialog.Tag = TAG_INCOME_STATEMENT;

            openFileSelectDialog();
        }

        private void button_payTransfer_Click(object sender, EventArgs e)
        {
            openXlsFileDialog.Tag = TAG_PAY_TRANSFER;

            openFileSelectDialog();
        }

        private void button_accountReceivable_Click(object sender, EventArgs e)
        {
            openXlsFileDialog.Tag = TAG_ACCOUNT_RECEIVABLE;

            openFileSelectDialog();
        }

        private void openFileSelectDialog()
        {
            openXlsFileDialog.InitialDirectory = defaulfPath;

            if (openXlsFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var sr = new StreamReader(openXlsFileDialog.FileName);

                    string fileName = openXlsFileDialog.FileName.Split('\\').Where(x => !string.IsNullOrWhiteSpace(x)).LastOrDefault();

                    defaulfPath = openXlsFileDialog.FileName.Replace(fileName, string.Empty);

                    switch (openXlsFileDialog.Tag)
                    {
                        case "CashBook":

                            if (!fileName.Contains("現金簿"))
                            {
                                MessageBox.Show(fileName + " 檔案名稱錯誤");
                                return;
                            }
                            linkLabel_cashBook.Text = fileName;
                            linkLabel_cashBook.Visible = true;
                            linkLabel_cashBook.Tag = openXlsFileDialog.FileName;

                            cashBookPath = openXlsFileDialog.FileName;

                            break;

                        case "BalanceSheet":

                            if (!fileName.Contains("資產負債表"))
                            {
                                MessageBox.Show(fileName + " 檔案名稱錯誤");
                                return;
                            }

                            linkLabel_balanceSheet.Text = fileName;
                            linkLabel_balanceSheet.Visible = true;
                            linkLabel_balanceSheet.Tag = openXlsFileDialog.FileName;

                            balanceSheetPath = openXlsFileDialog.FileName;

                            Workbook balanceSheetWB = xlApp.Workbooks.Open(openXlsFileDialog.FileName);

                            comboBox_balanceSheet.Items.Clear();
                            foreach (Worksheet worksheet in balanceSheetWB.Sheets)
                            {
                                comboBox_balanceSheet.Items.Add(worksheet.Name);
                            }

                            balanceSheetWB.Close(false);

                            break;

                        case "IncomeStatement":

                            if (!fileName.Contains("損益表"))
                            {
                                MessageBox.Show(fileName + " 檔案名稱錯誤");
                                return;
                            }

                            linkLabel_incomeStatement.Text = fileName;
                            linkLabel_incomeStatement.Visible = true;
                            linkLabel_incomeStatement.Tag = openXlsFileDialog.FileName;

                            incomeStatementPath = openXlsFileDialog.FileName;

                            Workbook incomeStatementWB = xlApp.Workbooks.Open(openXlsFileDialog.FileName);

                            comboBox_incomeStatement.Items.Clear();
                            foreach (Worksheet worksheet in incomeStatementWB.Sheets)
                            {
                                comboBox_incomeStatement.Items.Add(worksheet.Name);
                            }

                            incomeStatementWB.Close(false);

                            break;

                        case "PayTransfer":

                            if (!fileName.Contains("代付轉帳"))
                            {
                                MessageBox.Show(fileName + " 檔案名稱錯誤");
                                return;
                            }

                            linkLabel_payTransfer.Text = fileName;
                            linkLabel_payTransfer.Visible = true;
                            linkLabel_payTransfer.Tag = openXlsFileDialog.FileName;

                            payTransferPath = openXlsFileDialog.FileName;

                            Workbook payTransferWB = xlApp.Workbooks.Open(openXlsFileDialog.FileName);

                            comboBox_payTransfer.Items.Clear();
                            foreach (Worksheet worksheet in payTransferWB.Sheets)
                            {
                                comboBox_payTransfer.Items.Add(worksheet.Name);
                            }

                            payTransferWB.Close(false);

                            break;

                        case "AccountReceivable":

                            if (!fileName.Contains("應收轉帳"))
                            {
                                MessageBox.Show(fileName + " 檔案名稱錯誤");
                                return;
                            }

                            linkLabel_accountReceivable.Text = fileName;
                            linkLabel_accountReceivable.Visible = true;
                            linkLabel_accountReceivable.Tag = openXlsFileDialog.FileName;

                            accountReceivablePath = openXlsFileDialog.FileName;

                            Workbook accountReceivableWB = xlApp.Workbooks.Open(openXlsFileDialog.FileName);

                            comboBox_accountReceivable.Items.Clear();
                            foreach (Worksheet worksheet in accountReceivableWB.Sheets)
                            {
                                comboBox_accountReceivable.Items.Add(worksheet.Name);
                            }

                            accountReceivableWB.Close(false);

                            break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Exception error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }
        #endregion

        #region linkLabel
        private void linkLabel_cashBook_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFile(linkLabel_cashBook);
        }

        private void linkLabel_balanceSheet_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFile(linkLabel_balanceSheet);
        }


        private void linkLabel_incomeStatement_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFile(linkLabel_incomeStatement);
        }

        private void linkLabel_payTransfer_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFile(linkLabel_payTransfer);
        }

        private void linkLabel_accountReceivable_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            openFile(linkLabel_accountReceivable);
        }

        private void openFile(Control control)
        {
            Process process = new Process();
            process.StartInfo.FileName = (string)control.Tag;
            process.Start();
        }
        #endregion

        private void initComboBox()
        {
            comboBox_year.Items.Clear();
            int nowYear = DateTime.Now.Year;
            for (int i = -3; i <= 3; i++)
            {
                comboBox_year.Items.Add("民國 " + string.Format("{0,3:###}", (nowYear - 1911 + i)) + "年");
            }

            comboBox_month.Items.Clear();
            for (int i = 1; i <= 12; i++)
            {
                comboBox_month.Items.Add(string.Format("{0,2:##}", i) + "月");
            }
        }

        private void initFolder()
        {
            directoryName = "產出結果_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            if (!Directory.Exists(directoryName))
            {
                Directory.CreateDirectory(directoryName);
            }
        }

        private void initOutputFile()
        {
            spreadSheet1Path = "./" + directoryName + "/" + comboBox_year.SelectedItem.ToString().Replace("民國", string.Empty) + comboBox_month.SelectedItem.ToString() + "試算表1.xls";
            spreadSheet2Path = "./" + directoryName + "/" + comboBox_year.SelectedItem.ToString().Replace("民國", string.Empty) + comboBox_month.SelectedItem.ToString() + "試算表2.xls";
            balanceSheetResultPath = "./" + directoryName + "/" + comboBox_year.SelectedItem.ToString().Replace("民國", string.Empty) + comboBox_month.SelectedItem.ToString() + "資產負債表.xls";
            incomeStatementResultPath = "./" + directoryName + "/" + comboBox_year.SelectedItem.ToString().Replace("民國", string.Empty) + comboBox_month.SelectedItem.ToString() + "損益表.xls";

            CopyResourceToFile("會計過帳.1試算表_範本.xls", spreadSheet1Path);
            CopyResourceToFile("會計過帳.2試算表_範本.xls", spreadSheet2Path);
            CopyResourceToFile("會計過帳.資產負債表_範本.xls", balanceSheetResultPath);
            CopyResourceToFile("會計過帳.損益表_範本.xls", incomeStatementResultPath);
        }

        public void CopyResourceToFile(string resourceName, string fileName)
        {
            using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }
            }
        }

        private void button_start_Click(object sender, EventArgs e)
        {
            this.Enabled = false;

            Boolean isPAss = false;
            if (!isDebug)
            {
                try
                {
                    checkLinkLabel(linkLabel_cashBook, "現金簿");

                    checkLinkLabel(linkLabel_balanceSheet, "資產負債表");
                    checkField(comboBox_balanceSheet.SelectedItem, "資產負債表 工作表");

                    checkLinkLabel(linkLabel_incomeStatement, "損益表");
                    checkField(comboBox_incomeStatement.SelectedItem, "損益表 工作表");

                    checkLinkLabel(linkLabel_payTransfer, "代付轉帳");
                    checkField(comboBox_payTransfer.SelectedItem, "代付轉帳 工作表");

                    checkLinkLabel(linkLabel_accountReceivable, "應收轉帳");
                    checkField(comboBox_accountReceivable.SelectedItem, "應收轉帳 工作表");

                    checkField(comboBox_year.SelectedItem, "財報年份");

                    checkField(comboBox_month.SelectedItem, "財報月份");

                    isPAss = true;
                }
                catch (Exception ex)
                {
                    this.Enabled = true;
                    MessageBox.Show(ex.Message);
                    return;
                }
            }

            if (isDebug)
            {
                isPAss = true;
                DirectoryInfo d = new DirectoryInfo("C:\\參照資料");

                foreach (var file in d.GetFiles())
                {
                    if (file.Name.Contains("現金簿"))
                    {
                        cashBookPath = file.FullName;
                    }

                    if (file.Name.Contains("資產負債表"))
                    {
                        balanceSheetPath = file.FullName;
                    }

                    if (file.Name.Contains("損益表"))
                    {
                        incomeStatementPath = file.FullName;
                    }

                    if (file.Name.Contains("代付轉帳"))
                    {
                        payTransferPath = file.FullName;
                    }

                    if (file.Name.Contains("應收轉帳"))
                    {
                        accountReceivablePath = file.FullName;
                    }
                }
            }

            if (isPAss)
            {
                initFolder();
                initOutputFile();

                SortedDictionary<string, double> assetMap = new SortedDictionary<string, double>();
                SortedDictionary<string, double> liabMap = new SortedDictionary<string, double>();

                Workbook balanceSheetWB = xlApp.Workbooks.Open(balanceSheetPath);
                Worksheet balanceSheetWS = null;
                if (isDebug)
                {
                    balanceSheetWS = balanceSheetWB.Sheets[ConfigurationManager.AppSettings.Get("debugBalanceSheetWSName")];
                }
                else
                {
                    balanceSheetWS = balanceSheetWB.Sheets[comboBox_balanceSheet.SelectedItem.ToString()];
                }

                ParseBalanceWS(assetMap, liabMap, balanceSheetWS);

                balanceSheetWB.Close(false);

                Workbook incomeStatementWB = xlApp.Workbooks.Open(incomeStatementPath);
                Worksheet incomeStatementWS = null;
                if (isDebug)
                {
                    incomeStatementWS = incomeStatementWB.Sheets[ConfigurationManager.AppSettings.Get("debugIncomeStatementWSName")];
                }
                else
                {
                    incomeStatementWS = incomeStatementWB.Sheets[comboBox_incomeStatement.SelectedItem.ToString()];
                }

                ParseIncomeStatementWS(assetMap, liabMap, incomeStatementWS);

                incomeStatementWB.Close(false);

                Workbook cashBookWB = xlApp.Workbooks.Open(cashBookPath);

                string taipeiWSName = string.Empty;
                string keelungWSName = string.Empty;

                foreach (Worksheet worksheet in cashBookWB.Sheets)
                {
                    if (worksheet.Name.Contains("台北"))
                    {
                        taipeiWSName = worksheet.Name;
                    }

                    if (worksheet.Name.Contains("基隆"))
                    {
                        keelungWSName = worksheet.Name;
                    }
                }

                Worksheet taipeiWS = cashBookWB.Sheets[taipeiWSName];
                Worksheet keelungWS = cashBookWB.Sheets[keelungWSName];

                CalculateCash(assetMap, taipeiWS, keelungWS);

                ParseCashBookWS(assetMap, liabMap, taipeiWS);
                ParseCashBookWS(assetMap, liabMap, keelungWS);

                cashBookWB.Close(false);

                #region 試算表1
                spreadSheet1Path = Path.GetFullPath(spreadSheet1Path);

                Workbook spreadSheet1WB = xlApp.Workbooks.Open(spreadSheet1Path);
                Worksheet spreadSheet1 = spreadSheet1WB.Sheets[1];

                List<string> drCrAccountList = GenerateDrCrAccountList(spreadSheet1);

                UpdateSpreadSheet(assetMap, liabMap, spreadSheet1, drCrAccountList);

                spreadSheet1WB.Save();
                spreadSheet1WB.Close(false);
                #endregion

                #region 試算表2
                Workbook payTransferWB = xlApp.Workbooks.Open(payTransferPath);
                Worksheet payTransferSheet = null;
                if (isDebug)
                {
                    payTransferSheet = payTransferWB.Sheets[1];
                }
                else
                {
                    payTransferSheet = payTransferWB.Sheets[comboBox_payTransfer.SelectedItem.ToString()];
                }

                ParsePayTransferWS(assetMap, liabMap, payTransferSheet);

                payTransferWB.Close(false);

                Workbook accountReceivableWB = xlApp.Workbooks.Open(accountReceivablePath);
                Worksheet accountReceivableSheet = null;
                if (isDebug)
                {
                    accountReceivableSheet = accountReceivableWB.Sheets[1];
                }
                else
                {
                    accountReceivableSheet = accountReceivableWB.Sheets[comboBox_accountReceivable.SelectedItem.ToString()];
                }

                ParseAccountReceivableWS(assetMap, liabMap, accountReceivableSheet);

                accountReceivableWB.Close(false);

                spreadSheet2Path = Path.GetFullPath(spreadSheet2Path);

                Workbook spreadSheet2WB = xlApp.Workbooks.Open(spreadSheet2Path);
                Worksheet spreadSheet2 = spreadSheet2WB.Sheets[1];

                UpdateSpreadSheet(assetMap, liabMap, spreadSheet2, drCrAccountList);

                spreadSheet2WB.Save();
                spreadSheet2WB.Close(false);
                #endregion

                #region 損益表
                incomeStatementResultPath = Path.GetFullPath(incomeStatementResultPath);

                Workbook incomeStatementResultWB = xlApp.Workbooks.Open(incomeStatementResultPath);
                Worksheet incomeStatementResultWS = incomeStatementResultWB.Sheets[1];

                UpdateIncomeStatementResult(assetMap, liabMap, incomeStatementResultWS);

                Range range = incomeStatementResultWS.UsedRange.Find("累計損益", Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                range = range.Cells.Offset[0, 3];
                double amount = range.Value2;

                liabMap["本期損益"] = amount;

                incomeStatementResultWB.Save();
                incomeStatementResultWB.Close(false);
                #endregion

                #region 資產負債表
                balanceSheetResultPath = Path.GetFullPath(balanceSheetResultPath);

                Workbook balanceSheetResultWB = xlApp.Workbooks.Open(balanceSheetResultPath);
                Worksheet balanceSheetResultWS = balanceSheetResultWB.Sheets[1];

                UpdateBalanceSheetResult(assetMap, liabMap, balanceSheetResultWS);

                balanceSheetResultWB.Save();
                balanceSheetResultWB.Close(false);
                #endregion


                this.Enabled = true;
                MessageBox.Show("已處理完成");
            }
        }

        private void checkLinkLabel(LinkLabel obj, string errorStr)
        {
            if (string.IsNullOrEmpty(obj.Text))
            {
                throw new Exception("請確認欄位:" + errorStr);
            }
        }

        private void checkField(Object obj, string errorStr)
        {
            if (obj == null)
            {
                throw new Exception("請確認欄位:" + errorStr);
            }
        }

        public List<string> GenerateDrCrAccountList(Worksheet ws)
        {
            List<string> accountList = new List<string>();

            #region DR 借方
            string[] drRange = scanAccountDrRange.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
            int drColumn = NumberFromExcelColumn(Regex.Match(drRange[0], @"[a-zA-Z]+").Value);
            int drStart = Convert.ToInt32(Regex.Match(drRange[0], @"\d+").Value);
            int drEnd = Convert.ToInt32(Regex.Match(drRange[1], @"\d+").Value);

            Range xlRange = ws.UsedRange;
            for (int i = drStart; i <= drEnd; i++)
            {

                Range currentRange = (Range)xlRange.Cells[i, drColumn];
                if (currentRange.Value2 == null)
                {
                    continue;
                }
                string account = currentRange.Value2.ToString().Trim();

                if (!"代付款".Equals(account) && !accountList.Contains(account))
                {
                    accountList.Add(account);
                }
            }
            #endregion


            #region CR 貸方
            string[] crRange = scanAccountCrRange.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
            int crColumn = NumberFromExcelColumn(Regex.Match(crRange[0], @"[a-zA-Z]+").Value);
            int crStart = Convert.ToInt32(Regex.Match(drRange[0], @"\d+").Value);
            int crEnd = Convert.ToInt32(Regex.Match(drRange[1], @"\d+").Value);

            for (int i = drStart; i <= drEnd; i++)
            {

                Range currentRange = (Range)xlRange.Cells[i, crColumn];
                if (currentRange.Value2 == null)
                {
                    continue;
                }
                string account = currentRange.Value2.ToString().Trim();

                if (!"代付款".Equals(account) && !accountList.Contains(account))
                {
                    accountList.Add(account);
                }
            }
            #endregion

            return accountList;
        }

        /// <summary>
        /// A -> 1<br/>
        /// B -> 2<br/>
        /// C -> 3<br/>
        /// ...
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public int NumberFromExcelColumn(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        public void ParseBalanceWS(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {

            Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;

            int scanStart = 0;
            int scanEnd = 0;

            #region 找出會計科目範圍
            for (int i = 1; i <= rowCount; i++)
            {
                Range currentRange = (Range)xlRange.Cells[i, 3];
                if (currentRange.Value2 == null)
                {
                    continue;
                }


                string data = currentRange.Value2.ToString().Replace(" ", "");

                if ("負債淨值".Equals(data))
                {
                    scanStart = i + 1;
                }
                if ("總計".Equals(data))
                {
                    scanEnd = i;
                    break;
                }
            }
            #endregion

            if (scanEnd == 0)
            {
                throw new Exception("找不到總計");
            }

            #region 取出會科及金額
            for (int i = scanStart; i < scanEnd; i++)
            {
                try
                {

                    //   資  產
                    Range currentRange = (Range)xlRange.Cells[i, 1];
                    if (currentRange.Value2 == null)
                    {
                        continue;
                    }

                    string account = currentRange.Value2.ToString().Replace(" ", "");

                    currentRange = (Range)xlRange.Cells[i, 2];
                    if (currentRange.Value2 == null)
                    {
                        continue;
                    }

                    double amount = Convert.ToDouble(currentRange.Value2.ToString().Replace(" ", ""));

                    assetMap[account] = amount;


                    // 負  債  淨  值
                    currentRange = (Range)xlRange.Cells[i, 3];
                    if (currentRange.Value2 == null)
                    {
                        continue;
                    }

                    account = currentRange.Value2.ToString().Replace(" ", "");

                    if (!string.IsNullOrEmpty(account))
                    {

                        currentRange = (Range)xlRange.Cells[i, 4];
                        if (currentRange.Value2 == null)
                        {
                            continue;
                        }

                        amount = Convert.ToDouble(currentRange.Value2.ToString().Replace(" ", ""));

                        liabMap[account] = amount;
                    }
                }
                catch (Exception ex)
                {

                }
            }
            #endregion
        }

        public void ParseIncomeStatementWS(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {

            Range xlRange = ws.UsedRange;
            Range range = ws.UsedRange.Find(" 小     計", Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            Range target = range.Cells.Offset[1, 0];
            List<string> incomeAccountList = new List<string> { "營業收入", "其他收入", "利息收入" };

            while (target != null && target.Value2 != null)
            {
                try
                {
                    double income = Convert.ToDouble(target.Value2);
                    target = target.Cells.Offset[0, -1];
                    string account = target.Value2.ToString();

                    if (incomeAccountList.Contains(account))
                    {
                        liabMap[account] = income;
                    }
                    else
                    {
                        assetMap[account] = income;
                    }


                    target = target.Cells.Offset[1, 1];
                }
                catch (Exception ex)
                {
                    target = null;
                }

            }

        }

        public void CalculateCash(SortedDictionary<string, double> assetMap, Worksheet taipeiWS, Worksheet keelungWS)
        {
            double cash = assetMap["現金"];

            cash = cash + GetCash(taipeiWS);
            cash = cash + GetCash(keelungWS);

            assetMap["現金"] = cash;
        }

        public double GetCash(Worksheet ws)
        {
            double cash = 0;

            Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                Range currentRange = (Range)xlRange.Cells[i, 3];
                if (currentRange.Value2 == null)
                {
                    continue;
                }

                string dateString = currentRange.Value2.ToString();

                Match dateMatch = dateRegex.Match(dateString);
                Match digitalMatch = digitalRegex.Match(dateString);

                if (!dateMatch.Success && !digitalMatch.Success)
                {
                    continue;
                }

                double income = 0;
                double expenditure = 0;

                try
                {
                    if ("收入".Equals(xlRange.Cells[i, 5].Value2.ToString()))
                    {
                        continue;
                    }
                }
                catch (Exception ex)
                {

                }
                try
                {
                    income = Convert.ToDouble(xlRange.Cells[i, 5].Value2);
                }
                catch (Exception ex)
                {

                }

                try
                {
                    expenditure = Convert.ToDouble(xlRange.Cells[i, 6].Value2);
                }
                catch (Exception ex)
                {

                }

                cash = cash + income - expenditure;
            }

            return cash;
        }

        public void ParseCashBookWS(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {

            Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {

                Range currentRange = (Range)xlRange.Cells[i, 3];
                if (currentRange.Value2 == null)
                {
                    continue;
                }


                string dateString = currentRange.Value2.ToString();

                Match dateMatch = dateRegex.Match(dateString);
                Match digitalMatch = digitalRegex.Match(dateString);

                if (!dateMatch.Success && !digitalMatch.Success)
                {
                    continue;
                }

                currentRange = (Range)xlRange.Cells[i, 2];
                if (currentRange.Value2 == null)
                {
                    continue;
                }

                string account = xlRange.Cells[i, 2].Value2.ToString();

                if (string.IsNullOrEmpty(account))
                {
                    continue;
                }

                account = account.Replace(" ", "");

                Boolean isNeedUpdateKey = false;
                string accountKey = string.Empty;
                string accountType = getAccountMap(assetMap, liabMap, account, out accountKey, out isNeedUpdateKey);

                double income = 0;
                double expenditure = 0;

                try
                {
                    if ("收入".Equals(xlRange.Cells[i, 5].Value2.ToString()))
                    {
                        continue;
                    }
                }
                catch (Exception ex)
                {

                }
                try
                {
                    income = Convert.ToDouble(xlRange.Cells[i, 5].Value2);
                }
                catch (Exception ex)
                {

                }

                try
                {
                    expenditure = Convert.ToDouble(xlRange.Cells[i, 6].Value2);
                }
                catch (Exception ex)
                {

                }

                double balance;

                if ("asset".Equals(accountType))
                {
                    balance = expenditure - income;
                }
                else
                {
                    balance = income - expenditure;
                }


                if (string.IsNullOrEmpty(accountKey))
                {
                    if ("asset".Equals(accountType))
                    {
                        assetMap[account] = balance; //不存在會科 直接寫入
                    }
                    else
                    {
                        liabMap[account] = balance; //不存在會科 直接寫入
                    }
                }
                else
                {

                    if ("asset".Equals(accountType))
                    {
                        balance = assetMap[accountKey] + balance; //存在會科  取出現在餘額
                    }
                    else
                    {
                        balance = liabMap[accountKey] + balance; //存在會科  取出現在餘額
                    }

                    if (isNeedUpdateKey)
                    {
                        if ("asset".Equals(accountType))
                        {
                            assetMap.Remove(accountKey); //刪除舊的key  
                            assetMap[account] = balance;//存入較完整名稱的新Key
                        }
                        else
                        {
                            liabMap.Remove(accountKey); //刪除舊的key  
                            liabMap[account] = balance;//存入較完整名稱的新Key
                        }
                    }
                    else
                    {
                        if ("asset".Equals(accountType))
                        {
                            assetMap[accountKey] = balance;//將餘額存放在舊key中
                        }
                        else
                        {
                            liabMap[accountKey] = balance;//將餘額存放在舊key中
                        }
                    }
                }
            }
        }

        public string getAccountMap(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, string accountName, out string accountKey, out Boolean isNeedUpdateKey)
        {
            isNeedUpdateKey = false;

            accountKey = checkMap(liabMap, accountName, out isNeedUpdateKey);

            if (!string.IsNullOrEmpty(accountKey))
            {
                return "liab";
            }
            else
            {
                accountKey = checkMap(assetMap, accountName, out isNeedUpdateKey);
                return "asset";
            }
        }

        public string checkMap(SortedDictionary<string, double> accountMap, string accountName, out Boolean isNeedUpdateKey)
        {
            string accountKey = string.Empty;
            isNeedUpdateKey = false;

            string[] keys = accountMap.Keys.ToArray();

            foreach (string key in keys)
            {
                if (accountName.Equals(key))
                {
                    accountKey = key;

                    break;
                }

                //if (key.Contains(accountName) || accountName.Contains(key))
                //{
                //    accountKey = key;

                //    if (accountName.Contains(key))//現在map存的key比較短  要跟換為長的
                //    {
                //        isNeedUpdateKey = true;
                //    }

                //    break;
                //}
            }

            return accountKey;
        }

        public void UpdateSpreadSheet(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws, List<string> drCrAccountList)
        {
            string pendingPayment = "F2"; //代付款 
            string pendingPaymentAccount = "B24";
            string sumStart = "G3"; //代付款 上一行

            pendingPayment = UpdateSpreadSheetByAccountMap(ws, drCrAccountList, assetMap, pendingPayment);
            pendingPayment = UpdateSpreadSheetByAccountMap(ws, drCrAccountList, liabMap, pendingPayment);

            Range target = ws.UsedRange.Range[pendingPayment].Offset[2, 0];


            Range source = ws.Range["A25"];
            source.Copy(target);

            target.Value2 = "TOTAL";


            target = target.Offset[0, 1];

            ws.UsedRange.Range[pendingPaymentAccount].Value2 = "=" + target.Address;


            Range sumEndRange = target.Offset[-2, 0];

            target.Value2 = "=SUM(" + sumStart + ":" + sumEndRange.Address.Replace("$", "") + ")";
        }

        public string UpdateSpreadSheetByAccountMap(Worksheet ws, List<string> drCrAccountList, SortedDictionary<string, double> accountMap, string pendingPayment)
        {
            List<string> skipList = new List<string>() { "本期損益" };
            Range target = null;

            foreach (KeyValuePair<string, double> item in accountMap)
            {
                if (skipList.Contains(item.Key))
                {
                    continue;
                }
                if (drCrAccountList.Contains(item.Key))
                {
                    Range range = ws.UsedRange.Find(item.Key, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                    target = range.Cells.Offset[0, 1];
                    target.Value2 = item.Value;
                }
                else
                {
                    target = ws.UsedRange.Range[pendingPayment].Offset[1, 0];
                    target.Value2 = item.Key;

                    pendingPayment = target.Address;

                    target = target.Offset[0, 1];
                    target.Value2 = item.Value;
                }
            }

            return pendingPayment;
        }

        public void WriteResourceToFile(string resourceName, string fileName)
        {
            using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }
            }
        }

        public void ParsePayTransferWS(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {

            Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                try
                {
                    //借方
                    ParseCell(assetMap, liabMap, xlRange, i, true);

                }
                catch (Exception ex)
                {

                }
                try
                {
                    //貸方
                    ParseCell(assetMap, liabMap, xlRange, i, false);

                }
                catch (Exception ex)
                {

                }
            }
        }

        public void ParseAccountReceivableWS(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {

            Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                try
                {
                    //借方
                    ParseCell(assetMap, liabMap, xlRange, i, true);

                }
                catch (Exception ex)
                {

                }
                try
                {
                    //貸方
                    ParseCell(assetMap, liabMap, xlRange, i, false);

                }
                catch (Exception ex)
                {

                }
            }
        }

        public void ParseCell(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Range xlRange, int rowIndex, Boolean isDr)
        {
            int accountColumnIndex = 0;
            int amountColumnIndex = 0;
            int abstractColumnIndex = 0;

            if (isDr)
            {
                accountColumnIndex = 1;
                amountColumnIndex = 5;
                abstractColumnIndex = 3;
            }
            else
            {
                accountColumnIndex = 6;
                amountColumnIndex = 11;
                abstractColumnIndex = 8;
            }

            Range currentRange = (Range)xlRange.Cells[rowIndex, accountColumnIndex];
            if (currentRange.Value2 == null)
            {
                return;
            }

            string account = currentRange.Value2.ToString().Replace(" ", "");
            if (!drVoucherAccountList.Contains(account) && !crVoucherAccountList.Contains(account))
            {
                return;
            }

            if ("代付款".Equals(account))
            {
                currentRange = (Range)xlRange.Cells[rowIndex, abstractColumnIndex];
                if (currentRange.Value2 == null)
                {
                    return;
                }

                account = currentRange.Value2.ToString().Replace(" ", "");

                Boolean isNeedUpdateKey = false;
                string accountKey = checkMap(assetMap, account, out isNeedUpdateKey);


                Double amount = Convert.ToDouble(xlRange.Cells[rowIndex, amountColumnIndex].Value2);

                if (!isDr)
                {
                    amount = -amount;
                }

                if (string.IsNullOrEmpty(accountKey))
                {

                    assetMap[account] = amount; //不存在會科 直接寫入

                }
                else
                {

                    amount = assetMap[accountKey] + amount; //存在會科  取出現在餘額


                    if (isNeedUpdateKey)
                    {

                        assetMap.Remove(accountKey); //刪除舊的key  
                        assetMap[account] = amount;//存入較完整名稱的新Key

                    }
                    else
                    {
                        assetMap[accountKey] = amount;//將餘額存放在舊key中
                    }
                }
            }
            else
            {
                Double amount = Convert.ToDouble(xlRange.Cells[rowIndex, amountColumnIndex].Value2);

                if (isDr && crVoucherAccountList.Contains(account))
                {
                    amount = -amount;
                }
                else if (!isDr && drVoucherAccountList.Contains(account))
                {
                    amount = -amount;
                }

                if (assetMap.ContainsKey(account))
                {
                    assetMap[account] = assetMap[account] + amount;
                }
                else
                {
                    liabMap[account] = liabMap[account] + amount;
                }
            }
        }

        public void UpdateBalanceSheetResult(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {
            UpdateWSbyMap(assetMap, ws);
            UpdateWSbyMap(liabMap, ws);

            UpdatePendingPayment(assetMap, ws);

            int daysInMonth = DateTime.DaysInMonth(Convert.ToInt32(Regex.Match(comboBox_year.SelectedItem.ToString(), @"\d+").Value) + 1911, Convert.ToInt32(Regex.Match(comboBox_month.SelectedItem.ToString(), @"\d+").Value));

            Range range = ws.UsedRange.Find("總計", Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            range = range.Cells.Offset[1, 0];
            range.Value2 = comboBox_year.SelectedItem.ToString() + comboBox_month.SelectedItem.ToString() + "01日  至  " + comboBox_year.SelectedItem.ToString() + comboBox_month.SelectedItem.ToString() + daysInMonth + "日";

            ws.Name = comboBox_year.SelectedItem.ToString().Replace("民國", string.Empty) + comboBox_month.SelectedItem.ToString();
        }

        public void UpdateIncomeStatementResult(SortedDictionary<string, double> assetMap, SortedDictionary<string, double> liabMap, Worksheet ws)
        {
            UpdateWSbyMap(assetMap, ws);
            //暫存 累計損益
            double amount = liabMap["累計損益"];
            liabMap.Remove("累計損益");
            UpdateWSbyMap(liabMap, ws);
            liabMap["累計損益"] = amount;

            int daysInMonth = DateTime.DaysInMonth(Convert.ToInt32(Regex.Match(comboBox_year.SelectedItem.ToString(), @"\d+").Value) + 1911, Convert.ToInt32(Regex.Match(comboBox_month.SelectedItem.ToString(), @"\d+").Value));
            ws.UsedRange.Range["B2"].Value2 = comboBox_year.SelectedItem.ToString() + comboBox_month.SelectedItem.ToString() + "01日  至  " + comboBox_year.SelectedItem.ToString() + comboBox_month.SelectedItem.ToString() + daysInMonth + "日";

            ws.Name = comboBox_year.SelectedItem.ToString().Replace("民國", string.Empty) + comboBox_month.SelectedItem.ToString();
        }

        public void UpdateWSbyMap(SortedDictionary<string, double> map, Worksheet ws)
        {
            List<string> removeKeyList = new List<string>();

            foreach (KeyValuePair<string, double> item in map)
            {
                Range range = ws.UsedRange.Find(item.Key, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

                if (range != null)
                {
                    Range target = range.Cells.Offset[0, 1];
                    target.Value2 = item.Value;

                    removeKeyList.Add(item.Key);
                }
            }

            foreach (string key in removeKeyList)
            {
                map.Remove(key);
            }
        }

        public void UpdatePendingPayment(SortedDictionary<string, double> map, Worksheet ws)
        {
            int rowIndex = 11;
            string cellTarget = "A11";
            foreach (KeyValuePair<string, double> item in map)
            {
                // Get the rows to copy (rows 4157 to 4178 in your example).
                Range copyRange = ws.Rows[rowIndex + ":" + rowIndex];

                // Insert enough new rows to fit the rows we're copying.
                copyRange.Insert(XlInsertShiftDirection.xlShiftDown);

                //公司名稱
                Range target = ws.UsedRange.Range[cellTarget];
                target.Value2 = item.Key;

                //金額
                target = target.Offset[0, 1];
                target.Value2 = item.Value;

                target = target.Offset[1, -1];

                cellTarget = target.Address;

                rowIndex++;
            }

        }


    }
}
