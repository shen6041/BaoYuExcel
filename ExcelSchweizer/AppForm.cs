//-------------------------------------------------------------------------------------
// All Rights Reserved , Copyright (C) 2014 , DZD , Ltd .
//-------------------------------------------------------------------------------------

using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System;

namespace MatchStall
{
    /// <summary>
    /// 主窗体
    ///
    /// 修改纪录
    ///
    ///		2014-01-12 版本：1.0 yanghenglian 创建主键，注意命名空间的排序。
    ///		
    /// 版本：1.0
    ///
    /// <author>
    ///		<name>yanghenglian</name>
    ///		<date>2014-01-12</date>
    /// </author>
    /// </summary>
    public partial class AppForm : Form
    {
        private DataTable chooseDt = new DataTable();
        private readonly string _tempDir = Environment.CurrentDirectory + "/temp";
        public AppForm()
        {
            InitializeComponent();
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
                Directory.CreateDirectory(_tempDir);
            }
            catch (Exception)
            {
                MessageBox.Show(@"临时目录创建失败！");
            }
            this.Closed += AppForm_Closed;
        }

        void AppForm_Closed(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
            }
            catch (Exception)
            {
                MessageBox.Show(@"临时目录删除失败！");
            }
        }

        private Dictionary<string, double> moneyDic = new Dictionary<string, double>();

        /// <summary>
        /// 选择Excel文件导入DataGridView
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChooseExcelFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = @"Excel文件|*.xls;*.xlsx", Multiselect = true })//;*.xlsx
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    foreach (var fileName in ofd.FileNames)
                    {
                        try
                        {
                            var fi = new FileInfo(fileName);
                            var tempName = _tempDir + "/" + fi.Name +
                                DateTime.Now.Ticks + fi.Extension;//临时文件
                            fi.CopyTo(tempName);//拷贝至临时文件夹
                            var data = ExcelHelper.ExcelToDataTable(tempName);

                            var index = data.Columns.IndexOf("交易金额");
                            var money = 0.0;

                            foreach (DataRow row in data.Rows)
                            {

                                var temp = 0.0;
                                double.TryParse(row.ItemArray[index].ToString(), out temp);
                                money += temp;
                            }
                            moneyDic.Add(fi.Name, money);

                            chooseDt.Merge(data);//选中文件DataTable
                            this.lblExcelFileName.Text = ofd.FileName;
                        }
                        catch (Exception ex)
                        {
                            XiaoHan.LogWriter.WriteToDefaultDirectory(ex);
                            MessageBox.Show(fileName + @"导入失败！");
                        }
                    }
                    //统计总数

                    dgvInfo.DataSource = chooseDt;

                }
            }
        }

        private void btnSummary_Click(object sender, EventArgs e)
        {
            var column = new string[dgvInfo.Columns.Count];
            for (int index = 0; index < dgvInfo.Columns.Count; index++)
            {
                DataGridViewTextBoxColumn column1 = dgvInfo.Columns[index] as DataGridViewTextBoxColumn;
                column[index] = column1.Name;
            }
            dgvInfo.SummaryColumns = column;
            dgvInfo.SummaryRowVisible = true;
            //var column = dgvInfo.SelectedCells[0].OwningColumn;
            //dgvInfo.SummaryColumns = new string[] { column.Name };
            //dgvInfo.SummaryRowVisible = true;
            //try
            //{
            //    Clipboard.SetText(dgvInfo.Sum);
            //    MessageBox.Show(@"统计数据已拷贝到剪切板中！");
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show(@"剪切板操作失败，数据在列的下方！");
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chooseDt = new DataTable();
            dgvInfo.DataSource = chooseDt;
            dgvInfo.SummaryRowVisible = false;
            dgvInfo.SummaryColumns = new string[0];
            moneyDic.Clear();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string text = "";
            var totalMoney = 0.0;
            foreach (var item in moneyDic)
            {
                text += "文件名：" + item.Key;
                var fileStr = item.Key.ToString();
                var blanks = 20 - fileStr.Length;
                if (blanks > 0)
                    for (int i = 0; i < blanks; i++)
                    {
                        text += " ";
                    }
                text += "金额：" + item.Value + "\r\n";
                totalMoney += item.Value;
            }

            text += "总  计：";
            for (int i = 0; i < 20; i++)
            {
                text += " ";
            }
            text += "金额：" + totalMoney + "\r\n";

            var textFile = _tempDir + "/" + DateTime.Now.Ticks + ".txt";
            File.WriteAllText(textFile, text);
            Process.Start(textFile);
        }

        /// <summary>
        /// 导出excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
                {
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    DefaultExt = ".xls",
                    CheckPathExists = true,
                    CreatePrompt = true,
                    Filter = @"xls Document(*.xls)|*.xls"
                };
            if (sfd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                FileInfo info = new FileInfo(sfd.FileName);
                try
                {
                    ExcelHelper.DataGridViewToExcel(dgvInfo, info.Name, true, info.Directory.FullName);
                }
                catch (Exception)
                {
                    MessageBox.Show(@"导出失败！");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var codes = AppConfigUtils.GetAppConfig("DealCode");
            var defaultCode = string.Empty;
            foreach (var code in codes.Split('|'))
            {
                defaultCode += "交易码='" + code + "' or ";
            }
            defaultCode = defaultCode.Remove(defaultCode.Length - 3, 3);
            var result = InputBox("输入过滤公式", "", defaultCode);
            if (!string.IsNullOrEmpty(result))
            {
                DataRow[] drArr = chooseDt.Select(result);
                DataTable dtNew = chooseDt.Clone();
                for (int i = 0; i < drArr.Length; i++)
                {
                    dtNew.ImportRow(drArr[i]);
                }
                dgvInfo.DataSource = dtNew;
            }
        }

        private string InputBox(string Caption, string Hint, string Default)
        {
            Form InputForm = new Form();
            InputForm.MinimizeBox = false;
            InputForm.MaximizeBox = false;
            InputForm.StartPosition = FormStartPosition.CenterScreen;
            InputForm.Width = 420;
            InputForm.Height = 150;

            InputForm.Text = Caption;
            Label lbl = new Label();
            lbl.Text = Hint;
            lbl.Left = 10;
            lbl.Top = 20;
            lbl.Parent = InputForm;
            lbl.AutoSize = true;
            TextBox tb = new TextBox();
            tb.Left = 30;
            tb.Top = 45;
            tb.Width = 360;
            tb.Parent = InputForm;
            tb.Text = Default;
            tb.SelectAll();
            Button btnok = new Button();
            btnok.Left = 120;
            btnok.Top = 80;
            btnok.Parent = InputForm;
            btnok.Text = "确定";
            InputForm.AcceptButton = btnok;//回车响应

            btnok.DialogResult = DialogResult.OK;
            Button btncancal = new Button();
            btncancal.Left = 220;
            btncancal.Top = 80;
            btncancal.Parent = InputForm;
            btncancal.Text = "取消";
            btncancal.DialogResult = DialogResult.Cancel;
            try
            {
                if (InputForm.ShowDialog() == DialogResult.OK)
                {
                    return tb.Text;
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                InputForm.Dispose();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataTable table = dgvInfo.DataSource as DataTable;
            var codeIndex = table.Columns.IndexOf("帐号");
            var moneyIndex = table.Columns.IndexOf("交易金额");
            //在配置文件中写明01账号对应的银行账号号码
            var acounts = AppConfigUtils.GetAppConfig("AcountCode");
            var acount = acounts.Split('|');
            var text = "";
            foreach (var s in acount)
            {
                var numble = s.Split(',')[0];//01账号
                var code = s.Split(',')[1];//真实对应的银行账号
                var money = 0.00;
                foreach (DataRow row in table.Rows)
                {
                    if (row[codeIndex].ToString() == code)
                    {
                        var temp = 0.0;
                        double.TryParse(row.ItemArray[moneyIndex].ToString(), out temp);
                        money += temp;
                    }
                }
                text += "账号：" + numble;
                text += "        金额：" + money + "\r\n";
            }
            var textFile = _tempDir + "/" + DateTime.Now.Ticks + ".txt";
            File.WriteAllText(textFile, text);
            Process.Start(textFile);
        }
    }
}

