using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Compute
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            /*
            // 获取TableLayoutPanel的第一行
            var firstRow = tableLayoutPanel1.GetRow(tableLayoutPanel1.Controls[0]);

            // 遍历第一行的所有元素
            foreach (Control control in tableLayoutPanel1.Controls)
            {
                if (tableLayoutPanel1.GetRow(control) == firstRow)
                {
                    if (control is Label)
                    {
                        // 设置Label的值为100
                        ((Label)control).Text = "100";
                    }
                    else if (control is NumericUpDown)
                    {
                        // 设置NumericUpDown的值为100
                        ((NumericUpDown)control).Value = 100;
                    }
                }
            }*/
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            /*
            // 获取主屏幕的分辨率
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            // 计算窗体的大小（这里设置为屏幕宽高的一半）
            int formWidth = screenWidth / 2;
            int formHeight = screenHeight / 2;

            // 设置窗体的大小
            Width = formWidth;
            Height = formHeight;
            */
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label40_Click(object sender, EventArgs e)
        {

        }

        private void label44_Click(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void label82_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 复制第一行的控件
            List<Control> firstRowControls = new List<Control>();
            for (int col = 0; col < tableLayoutPanel1.ColumnCount; col++)
            {
                var control = tableLayoutPanel1.GetControlFromPosition(col, 1);
                if (control != null)
                {
                    Control newControl = CloneControl(control);
                    firstRowControls.Add(newControl);
                }
            }

            // 在最后一行后面添加一行
            tableLayoutPanel1.RowCount++;
            int newRow = tableLayoutPanel1.RowCount - 1;

            for (int col = 0; col < tableLayoutPanel1.ColumnCount; col++)
            {
                var control = firstRowControls[col];
                Control newControl = CloneControl(control);
                tableLayoutPanel1.Controls.Add(newControl, col, newRow);
            }
        }

        private Control CloneControl(Control control)
        {
            if (control is Label)
            {
                Label label = new Label();
                label.Text = ((Label)control).Text;
                label.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
                label.AutoSize = true;
                label.Font = new System.Drawing.Font("Microsoft YaHei UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
                label.Size = new System.Drawing.Size(174, 28);
                label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                return label;
            }
            else if (control is NumericUpDown)
            {
                NumericUpDown numericUpDown = new NumericUpDown();
                numericUpDown.Value = 0;
                numericUpDown.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom;
                numericUpDown.Maximum = new decimal(new int[] { 100000000, 0, 0, 0 });
                //numericUpDown.Name = "numericUpDown4";
                numericUpDown.Size = new System.Drawing.Size(142, 23);
                //numericUpDown.TabIndex = 4;
                return numericUpDown;
            }
            else
            {
                // 如果有其他类型的控件，可以根据需要添加相应的处理逻辑
                return null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int row = 1; row < tableLayoutPanel1.RowCount-3; row++)
            {
                for (int col = 1; col <= 3; col++)
                {
                    decimal inAmountScore = 0;
                    decimal inGoalScore = 0;
                    decimal inSpecAmountScore = 0;
                    var control = tableLayoutPanel1.GetControlFromPosition(col, row);
                    var control1 = tableLayoutPanel1.GetControlFromPosition(col+3, row);
                    if (control != null)
                    {

                        if (control is Label)
                        {


                            // 设置Label的值为103
                           // ((Label)control).Text = "103";
                        }
                        else if (control is NumericUpDown)
                        {
                            // 设置NumericUpDown的值为103
                            if (col == 1)
                            {
                                ((Label)control1).Text = ((NumericUpDown)control).Value.ToString();//写公式上去
                            }
                           
                        }
                    }
                }
            }
        }
    }
}
