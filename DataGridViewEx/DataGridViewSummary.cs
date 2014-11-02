using System;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace Ruijc.UI.DataGridViewEx
{
    //文件        ： DataGridViewSummary.cs
    //描述        ： 扩展DataGridView且具有数据统计及键盘快捷操作功能
    //创建人      ： Ruijc
    //创建时间    ： 2010年05月12日
    //最后修改人  ：（无）
    //最后修改时间：（无）
    //版权所有 (C)：（无）
    public partial class DataGridViewSummary : DataGridView
    {
        /// <summary>
        /// 合计文本窗口
        /// </summary>
        private Panel _SummaryContainer = new Panel();
        /// <summary>
        /// 合计行头标签
        /// </summary>
        private SummaryTextBox _SummaryRowHeaderLabel = new SummaryTextBox();

        /// <summary>
        /// 缓存统计文本框控件
        /// </summary>
        private Hashtable _SummaryTextHashTable = new Hashtable();
        private SummaryTextType _FirstColSumTextType = SummaryTextType.None;

        /// <summary>
        /// 合计文本类型
        /// </summary>
        private enum SummaryTextType
        {
            Number,
            Text,
            None
        }

        #region 可设计属性

        /// <summary>
        /// 统计文本标题
        /// </summary>
        private string _SummaryHeaderText;
        [Browsable(true), Category("Summary")]
        public string SummaryHeaderText
        {
            get { return _SummaryHeaderText; }
            set
            {
                _SummaryHeaderText = value;
            }
        }

        /// <summary>
        /// 是否显示统计行
        /// </summary>
        private bool _SummaryRowVisible = true;
        [Browsable(true), Category("Summary")]
        public bool SummaryRowVisible
        {
            get { return _SummaryRowVisible; }
            set
            {
                _SummaryRowVisible = value;
                if (this._SummaryContainer != null)
                {
                    _SummaryContainer.Visible = value;
                }
            }
        }

        /// <summary>
        /// 统计行背景色
        /// </summary>
        private Color _SummaryRowBackColor = Color.White;
        [Browsable(true), Category("Summary")]
        public Color SummaryRowBackColor
        {
            get
            {
                if (_SummaryRowBackColor == Color.White)
                    return this.GridColor;
                return _SummaryRowBackColor;
            }
            set
            {
                _SummaryRowBackColor = value;
            }
        }

        /// <summary>
        /// 统计行字体色
        /// </summary>
        private Color _SummaryRowForeColor = Color.Blue;
        [Browsable(true), Category("Summary")]
        public Color SummaryRowForeColor
        {
            get
            {
                return _SummaryRowForeColor;
            }
            set
            {
                _SummaryRowForeColor = value;
            }
        }
        /// <summary>
        ///统计文本粗体字
        /// </summary>
        private bool _SummaryHeaderBold;
        [Browsable(true), Category("Summary")]
        public bool SummaryHeaderBold
        {
            get { return _SummaryHeaderBold; }
            set { _SummaryHeaderBold = value; }
        }

        /// <summary>
        /// 统计列,要求设置为要统计Column的Name或DataPropertyName
        /// </summary>
        private string[] _SummaryColumns;
        [Browsable(true), Category("Summary")]
        public string[] SummaryColumns
        {
            get { return _SummaryColumns; }
            set
            {
                _SummaryColumns = value;
                this._SummaryTextHashTable.Clear();
                this._SummaryContainer.Controls.Clear();
                this.RefreshSummaryTextBoxCache();
                this.AdjustSummaryTextBoxWidth();
                this.ShowSummaryTextInfo();
            }
        }

        /// <summary>
        /// 表格数据列
        /// </summary>
        private Dictionary<string, DataGridViewColumn> GridViewColumns
        {
            get
            {
                Dictionary<string, DataGridViewColumn> result = new Dictionary<string, DataGridViewColumn>();
                DataGridViewColumn column = this.Columns.GetFirstColumn(DataGridViewElementStates.None);
                if (column == null || column.Name.Trim() == "")
                    return result;
                result.Add(column.Name, column);
                while ((column = this.Columns.GetNextColumn(column, DataGridViewElementStates.None, DataGridViewElementStates.None)) != null)
                {
                    if (column.Visible == false) continue;
                    if (!result.ContainsKey(column.Name) && column.Name.Trim() != "")
                        result.Add(column.Name, column);
                }
                return result;
            }
        }

        private void RefreshSummaryTextBoxCache()
        {
            if (this._SummaryColumns == null || this._SummaryColumns.Length == 0)
            {
                this._SummaryContainer.Visible = false;
                return;
            }
            this._SummaryContainer.Visible = this._SummaryRowVisible;
            //统计列
            for (int i = 0; i < this._SummaryColumns.Length; i++)
            {
                if (this.GridViewColumns.ContainsKey(_SummaryColumns[i]))
                {
                    SummaryTextBox sumTextBox = new SummaryTextBox();
                    DataGridViewColumn currCol = this.GridViewColumns[_SummaryColumns[i]];
                    sumTextBox.Name = currCol.DataPropertyName.Trim() == "" ? currCol.Name : currCol.DataPropertyName;
                    sumTextBox.IsSummary = true;
                    sumTextBox.IsHeaderLabel = false;
                    if (!this._SummaryTextHashTable.ContainsKey(currCol))
                        this._SummaryTextHashTable.Add(currCol, sumTextBox);
                    if (currCol.DisplayIndex == 0)
                        _FirstColSumTextType = SummaryTextType.Number;
                }
            }

            //非统计列
            foreach (DataGridViewColumn currCol in this.GridViewColumns.Values)
            {
                if (!_SummaryTextHashTable.ContainsKey(currCol))
                {
                    SummaryTextBox sumTextBox = new SummaryTextBox();
                    sumTextBox.Name = currCol.DataPropertyName.Trim() == "" ? currCol.Name : currCol.DataPropertyName;
                    sumTextBox.IsSummary = false;
                    sumTextBox.IsHeaderLabel = false;
                    this._SummaryTextHashTable.Add(currCol, sumTextBox);
                }
            }

            foreach (DataGridViewColumn currCol in this.GridViewColumns.Values)
            {
                if (_FirstColSumTextType == SummaryTextType.None)
                {
                    //取得当前列的下一列
                    DataGridViewColumn NextCol = this.Columns.GetNextColumn(currCol, DataGridViewElementStates.None, DataGridViewElementStates.None);
                    if (NextCol == null) break;
                    if (NextCol.Name.Trim() == "") continue;
                    SummaryTextBox nextSumTextBox = this._SummaryTextHashTable[NextCol] as SummaryTextBox;
                    SummaryTextBox currSumTextBox = this._SummaryTextHashTable[currCol] as SummaryTextBox;
                    currSumTextBox.Name = currCol.DataPropertyName.Trim() == "" ? currCol.Name : currCol.DataPropertyName;

                    //如果下一列对应是SummaryTextBox实例是合计，且第一列不是头列时设置第一列为头列类型
                    if (nextSumTextBox.IsSummary && _FirstColSumTextType != SummaryTextType.Text)
                    {
                        currSumTextBox.IsHeaderLabel = true;
                        currSumTextBox.IsSummary = false;
                        currSumTextBox.Visible = true;
                        _FirstColSumTextType = SummaryTextType.Text;
                    }
                    else
                    {
                        currSumTextBox.IsHeaderLabel = false;
                        currSumTextBox.IsSummary = false;
                        currSumTextBox.Visible = false;
                    }
                }
            }
        }

        #endregion

        #region 构造函数
        public DataGridViewSummary()
        {
            this._SummaryContainer.AutoSize = false;
            this._SummaryContainer.BorderStyle = this.BorderStyle;
            this._SummaryContainer.Height = this.RowTemplate.Height - 1;
            this._SummaryContainer.Padding = new Padding(0);
            this._SummaryContainer.Margin = new Padding(0);

            //加入统计区域,并能过OnPaint(PaintEventArgs e)调整统计区域位置
            this.Controls.Add(this._SummaryContainer);
            //垂直滚动条改变
            this.VerticalScrollBar.VisibleChanged += new EventHandler(VerticalScrollBar_VisibleChanged);
            //表格尺寸改变
            this.Resize += new EventHandler(DataGridViewSummary_Resize);
            //单元格值改变
            this.CellValueChanged += new DataGridViewCellEventHandler(DataGridViewSummary_CellValueChanged);
            //单元格验证完成
            this.CellValidated += new DataGridViewCellEventHandler(DataGridViewSummary_CellValidated);
            //列显示顺序改变
            this.ColumnDisplayIndexChanged += new DataGridViewColumnEventHandler(DataGridViewSummary_ColumnDisplayIndexChanged);
            //列宽改变
            this.ColumnWidthChanged += new DataGridViewColumnEventHandler(DataGridViewSummary_ColumnWidthChanged);
            //添加列
            this.ColumnAdded += new DataGridViewColumnEventHandler(DataGridViewSummary_ColumnAdded);
            //移除列
            this.ColumnRemoved += new DataGridViewColumnEventHandler(DataGridViewSummary_ColumnRemoved);
            //数据源改变
            this.DataSourceChanged += new EventHandler(DataGridViewSummary_DataSourceChanged);
            //数据输入错误
            this.DataError += new DataGridViewDataErrorEventHandler(DataGridViewSummary_DataError);
            //数据绑定完成
            this.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(DataGridViewSummary_DataBindingComplete);
        }

        #endregion

        #region 事件处理
        void DataGridViewSummary_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            this.ResetSummaryContainer();
        }

        void DataGridViewSummary_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.RowIndex == -1) return;
            this.Rows[e.RowIndex].ErrorText = "输入数据不正确!";
            e.Cancel = true;
        }

        void DataGridViewSummary_DataSourceChanged(object sender, EventArgs e)
        {
            this.ResetSummaryContainer();
        }

        void DataGridViewSummary_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1) return;
            this.Rows[e.RowIndex].ErrorText = "";
        }

        void DataGridViewSummary_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            this.ShowSummaryTextInfo();
        }

        void DataGridViewSummary_ColumnRemoved(object sender, DataGridViewColumnEventArgs e)
        {
            this._SummaryTextHashTable.Clear();
            this._SummaryContainer.Controls.Clear();
            this.ResetSummaryContainer();
        }

        void DataGridViewSummary_ColumnDisplayIndexChanged(object sender, DataGridViewColumnEventArgs e)
        {
            this._SummaryTextHashTable.Clear();
            this._SummaryContainer.Controls.Clear();
            this.ResetSummaryContainer();
        }

        void DataGridViewSummary_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            this._SummaryTextHashTable.Clear();
            this._SummaryContainer.Controls.Clear();
            this.ResetSummaryContainer();
        }

        void DataGridViewSummary_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            this.ResetSummaryContainer();
        }

        private void DataGridViewSummary_Resize(object sender, EventArgs e)
        {
            this.ResetSummaryContainer();
        }

        void VerticalScrollBar_VisibleChanged(object sender, EventArgs e)
        {
            if (this.VerticalScrollBar.Visible)
            {
                this.VerticalScrollBar.SmallChange = this.RowTemplate.Height * 2;
            }
        }
        #endregion

        #region 应用处理
        /// <summary>
        /// 设置统计区域大小、位置，并构建统计
        /// </summary>
        private void ResetSummaryContainer()
        {
            this.RefreshSummaryTextBoxCache();

            int realHeight = this._SummaryContainer.Height;

            if (this.HorizontalScrollBar.Visible)
            {
                this._SummaryContainer.Top = this.Height - realHeight - this.HorizontalScrollBar.Height - 1;
            }
            else
            {
                this._SummaryContainer.Top = this.Height - realHeight - 1;
            }
            this._SummaryRowHeaderLabel.Name = "RowsHeader_Label";
            this._SummaryRowHeaderLabel.IsHeaderLabel = true;
            this._SummaryRowHeaderLabel.Text = "√";
            this._SummaryRowHeaderLabel.Top = 0;
            this._SummaryRowHeaderLabel.Left = 0;
            this._SummaryRowHeaderLabel.BorderColor = this.GridColor;
            this._SummaryRowHeaderLabel.Height = this.RowTemplate.Height;
            this._SummaryRowHeaderLabel.Width = this.RowHeadersWidth;
            if (!this._SummaryContainer.Controls.ContainsKey("RowsHeader_Label"))
                this._SummaryContainer.Controls.Add(this._SummaryRowHeaderLabel);

            this._SummaryContainer.Left = 0;
            this._SummaryContainer.Width = CalcSummaryContainerWidth();
            this._SummaryContainer.Invalidate();
            //调整统计项SummaryTextBox尺寸
            this.AdjustSummaryTextBoxWidth();
            //显示最终计算结果
            this.ShowSummaryTextInfo();
        }

        /// <summary>
        /// 调整统计项SummaryTextBox尺寸
        /// </summary>
        private void AdjustSummaryTextBoxWidth()
        {
            int rowHeaderWidth = this.RowHeadersVisible ? this.RowHeadersWidth - 1 : 0;
            int curPos = rowHeaderWidth;
            int labelWidth = 0;
            foreach (DataGridViewColumn col in this.GridViewColumns.Values)
            {
                SummaryTextBox sumTextBox = (SummaryTextBox)this._SummaryTextHashTable[col];
                if (sumTextBox == null) continue;
                //计算统计头文本宽度
                if (!sumTextBox.Visible)
                {
                    labelWidth += col.Width;
                    continue;
                }

                sumTextBox.BorderColor = this.GridColor;
                sumTextBox.BackColor = this._SummaryRowBackColor;

                if (!col.Visible)
                {
                    sumTextBox.Visible = false;
                    continue;
                }
                int startX = curPos;
                if (this.HorizontalScrollBar.Visible)
                    startX = curPos - this.HorizontalScrollingOffset;

                int currWidth = col.Width;
                if (sumTextBox.IsHeaderLabel)
                    currWidth = labelWidth + col.Width;

                if (startX < rowHeaderWidth)
                {
                    currWidth -= rowHeaderWidth - startX;
                    startX = rowHeaderWidth;
                }

                if (startX + currWidth > this.Width)
                    currWidth = this.Width - startX;

                if (this.RightToLeft == RightToLeft.Yes)
                    startX = this.Width - startX - currWidth;

                if (sumTextBox.Left != startX || sumTextBox.Width != currWidth)
                {
                    sumTextBox.SetBounds(startX, 0, currWidth + 1, this.RowTemplate.Height);
                    sumTextBox.BorderColor = this.GridColor;
                    sumTextBox.Visible = true;
                    if (this._SummaryContainer.Controls.ContainsKey(sumTextBox.Name))
                    {
                        SummaryTextBox originalTextBox = this._SummaryContainer.Controls[sumTextBox.Name] as SummaryTextBox;
                        originalTextBox.SetBounds(startX, 0, currWidth + 1, this.RowTemplate.Height);
                        originalTextBox.Invalidate();
                        originalTextBox.BringToFront();
                    }
                    else
                    {
                        this._SummaryContainer.Controls.Add(sumTextBox);
                    }
                    sumTextBox.BringToFront();
                }
                if (sumTextBox.IsHeaderLabel)
                {
                    curPos += labelWidth + col.Width;
                }
                else
                {
                    curPos += col.Width;
                }
                sumTextBox.Invalidate();
            }
            this._SummaryContainer.Refresh();
        }

        /// <summary>
        /// 计算统计控件的宽度
        /// </summary>
        /// <returns></returns>
        private int CalcSummaryContainerWidth()
        {
            int columnWidth = this.RowHeadersVisible ? this.RowHeadersWidth : 1;
            foreach (DataGridViewColumn col in this.GridViewColumns.Values)
            {
                if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill)
                {
                    columnWidth += col.MinimumWidth;
                }
                else
                    columnWidth += col.Width;
            }
            if (columnWidth == 0)
                columnWidth = this.Width;
            return columnWidth;
        }

        /// <summary>
        /// 是否是整形
        /// </summary>
        /// <param name="o">对象</param>
        /// <returns>true 或 false</returns>
        protected bool IsInteger(object o)
        {
            bool IsInt = false;
            if (o is Int64)
                IsInt = true;
            if (o is Int32)
                IsInt = true;
            if (o is Int16)
                IsInt = true;
            return IsInt;
        }

        /// <summary>
        /// 是否是小数点型
        /// </summary>
        /// <param name="o">对象</param>
        /// <returns>true 或 false</returns>
        protected bool IsDecimal(object o)
        {
            bool IsDec = false;
            Decimal dec;
            IsDec = Decimal.TryParse(o.ToString(), out dec);
            if (IsDec) return true;

            Single sin;
            IsDec = Single.TryParse(o.ToString(), out sin);
            if (IsDec) return true;

            Double dou;
            IsDec = Double.TryParse(o.ToString(), out dou);
            if (IsDec) return true;

            return IsDec;
        }

        public string Sum { get; set; }

        /// <summary>
        /// 列合计计算
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        private string CalcSum(DataGridViewColumn col)
        {
            SummaryTextBox sumTextBox = this._SummaryTextHashTable[col] as SummaryTextBox;
            object valObj = sumTextBox.Tag;
            foreach (DataGridViewRow dgvRow in this.Rows)
            {
                if (this._SummaryTextHashTable.ContainsKey(col))
                {
                    if (sumTextBox != null && sumTextBox.IsSummary)
                    {
                        if (valObj == null)
                            valObj = 0;
                        if (dgvRow.Cells[col.Index].Value != null
                            && !(dgvRow.Cells[col.Index].Value is DBNull))
                        {
                            if (IsInteger(dgvRow.Cells[col.Index].Value))
                            {
                                valObj = Convert.ToInt64(valObj) + Convert.ToInt64(dgvRow.Cells[col.Index].Value == null || dgvRow.Cells[col.Index].Value.ToString() == "" ? 0 : dgvRow.Cells[col.Index].Value);
                            }
                            else if (IsDecimal(dgvRow.Cells[col.Index].Value))
                            {
                                valObj = Convert.ToDecimal(valObj) + Convert.ToDecimal(dgvRow.Cells[col.Index].Value == null || dgvRow.Cells[col.Index].Value.ToString() == "" ? 0 : dgvRow.Cells[col.Index].Value);
                            }
                        }
                    }
                }
            }
            Sum = string.Format("{0}", valObj);
            return Sum;
        }

        /// <summary>
        /// 显示统计信息
        /// </summary>
        private void ShowSummaryTextInfo()
        {
            int icnt = this._SummaryContainer.Controls.Count;
            _FirstColSumTextType = SummaryTextType.None;
            foreach (Control ctr in this._SummaryContainer.Controls)
            {
                SummaryTextBox sumTextBox = ctr as SummaryTextBox;
                if (sumTextBox == null)
                    continue;
                if (sumTextBox.Name == "RowsHeader_Label")
                    continue;

                DataGridViewColumn currCell = this.GridViewColumns[sumTextBox.Name];
                if (currCell == null) continue;
                sumTextBox.ForeColor = this._SummaryRowForeColor;
                sumTextBox.BackColor = this._SummaryRowBackColor;
                if (sumTextBox.IsHeaderLabel)
                {//如果SummaryTextBox的IsHeaderLabel为true则显示统计标题文本
                    sumTextBox.Text = this._SummaryHeaderText;
                    sumTextBox.TextAlign = HorizontalAlignment.Center;
                    sumTextBox.Font = new Font(this.DefaultCellStyle.Font, this._SummaryHeaderBold ? FontStyle.Bold : FontStyle.Regular);
                    _FirstColSumTextType = SummaryTextType.Text;
                    sumTextBox.Invalidate();
                    continue;
                }
                if (sumTextBox.IsSummary)
                {//如果是SummaryTextBox的IsSummary为true则计算对应列的合计
                    sumTextBox.Text = this.CalcSum(currCell);
                    sumTextBox.FormatString = currCell.DefaultCellStyle.Format;
                    sumTextBox.TextAlign = AligmentHelper.TranslateGridColumnAligment(currCell.DefaultCellStyle.Alignment);
                    sumTextBox.Invalidate();
                    continue;
                }
                if (!sumTextBox.IsHeaderLabel && !sumTextBox.IsSummary)
                {
                    sumTextBox.Text = "";
                    sumTextBox.Invalidate();
                }
            }
            this._SummaryRowHeaderLabel.Text = "√";
            //如果第一列是文本,则将统计文本设置到_SummaryRowHeaderLabel
            if (_FirstColSumTextType != SummaryTextType.Text)
            {
                this._SummaryRowHeaderLabel.Text = this._SummaryHeaderText;
            }
            this._SummaryRowHeaderLabel.TextAlign = HorizontalAlignment.Center;
            this._SummaryRowHeaderLabel.Font = new Font(this.DefaultCellStyle.Font, this._SummaryHeaderBold ? FontStyle.Bold : FontStyle.Regular);
            this._SummaryRowHeaderLabel.Invalidate();
        }
        #endregion

        #region 继承重写
        /// <summary>
        /// 键盘操作处理
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="keyData"></param>
        /// <returns></returns>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            int WM_KEYDOWN = 256;
            int WM_SYSKEYDOWN = 260;
            if (msg.Msg == WM_KEYDOWN || msg.Msg == WM_SYSKEYDOWN)
            {
                switch (keyData)
                {
                    case Keys.Shift | Keys.Enter://按下Shift+Enter组合键时退格
                        SendKeys.Send("+{TAB}");
                        return true;
                    case Keys.Enter://按下Enter键时提供DataGridView快捷输入
                        if (this.CurrentCell.IsInEditMode)
                        {
                            SendKeys.Send("{TAB}");
                        }
                        else
                        {
                            if (this.CurrentCell is DataGridViewComboBoxCell)
                            {//下拉控件操作
                                SendKeys.Send("{F2}");
                                SendKeys.Send("{F4}");
                            }
                            else if (this.CurrentCell is DataGridViewCheckBoxCell)
                            {//复择框操作
                                SendKeys.Send(" ");
                                SendKeys.Send("{TAB}");
                            }

                            else if (this.CurrentCell is DataGridViewTextBoxCell)
                            {//文本框编辑
                                SendKeys.Send("{F2}");
                            }
                        }
                        return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            this.ResetSummaryContainer();
            base.OnPaint(e);
        }
        #endregion
    }
}
