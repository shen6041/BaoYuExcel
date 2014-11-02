using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace Ruijc.UI.DataGridViewEx
{
    public partial class SummaryTextBox : Control
    {
        /// <summary>
        /// 文本格式化
        /// </summary>
        private StringFormat _Format;
        public SummaryTextBox()
        {            
            InitializeComponent();

            _Format = new StringFormat( StringFormatFlags.NoWrap  | StringFormatFlags.FitBlackBox | StringFormatFlags.MeasureTrailingSpaces);
            _Format.LineAlignment = StringAlignment.Center;            

            this.Height = 10;
            this.Width = 10;

            this.Padding = new Padding(2);
            this.MouseClick += SummaryTextBox_MouseClick;
        }

        void SummaryTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                Clipboard.SetText(Text);
                MessageBox.Show(@"统计数据已拷贝到剪切板中！");
            }
            catch (Exception)
            {
                MessageBox.Show(@"剪切板操作失败，数据在列的下方！");
            }
        }

        public SummaryTextBox(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
            this.MouseClick += SummaryTextBox_MouseClick;
            this.TextChanged += new EventHandler(SummaryTextBox_TextChanged);
        }

        private void SummaryTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_FormatString) && !string.IsNullOrEmpty(Text))
            {
                Text = string.Format(_FormatString, Text);
            }
        }

        private Color _BorderColor = Color.Black;

        private bool _IsSummary;
        public bool IsSummary
        {
            get { return _IsSummary; }
            set { _IsSummary = value; }
        }

        private bool _IsHeaderLabel=false;
        public bool IsHeaderLabel
        {
            get { return _IsHeaderLabel; }
            set { _IsHeaderLabel = value; }
        }

        private string _FormatString;
        public string FormatString
        {
            get { return _FormatString; }
            set { _FormatString = value; }
        }


        private HorizontalAlignment _TextAlign = HorizontalAlignment.Left;
        [DefaultValue(HorizontalAlignment.Left)]
        public HorizontalAlignment TextAlign
        {
            get { return _TextAlign; }
            set 
            {
                _TextAlign = value;
                SetFormatFlags();
            }
        }

        private Color _ForeColor = Color.Blue;
        [DefaultValue(HorizontalAlignment.Left)]
        public new Color ForeColor
        {
            get { return _ForeColor; }
            set
            {
                _ForeColor = value;
            }
        }

        private StringTrimming _Trimming = StringTrimming.None;
        [DefaultValue(StringTrimming.None)]
        public StringTrimming Trimming
        {
            get { return _Trimming; }
            set
            {
                _Trimming = value;
                SetFormatFlags();
            }
        }

        private void SetFormatFlags()
        {
            _Format.Alignment = AligmentHelper.TranslateAligment(TextAlign);
            _Format.Trimming = _Trimming;                
        }

        public Color BorderColor
        {
            get { return _BorderColor; }
            set { _BorderColor = value; }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle textBounds;

            if (!string.IsNullOrEmpty(_FormatString) && !string.IsNullOrEmpty(Text))
            {
                Text = String.Format("{0:" + _FormatString + "}", Convert.ToDecimal(Text));
            }

            textBounds = new Rectangle(this.ClientRectangle.X + 2, this.ClientRectangle.Y + 2, this.ClientRectangle.Width - 2 , this.ClientRectangle.Height - 2 );
            SolidBrush brush = new SolidBrush(this._ForeColor);
            using(Pen pen = new Pen(_BorderColor))
            {
                pen.Width = 2;

                e.Graphics.FillRectangle(new SolidBrush(this.BackColor), this.ClientRectangle);
                e.Graphics.DrawRectangle(pen, this.ClientRectangle.X, this.ClientRectangle.Y, this.ClientRectangle.Width , this.ClientRectangle.Height - 1);
                e.Graphics.DrawString(Text, Font, brush, textBounds, _Format);
            }
            brush.Dispose();
            e.Graphics.Dispose();
        }
    }
}


