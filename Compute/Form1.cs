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
            return;
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

            bool rowIsOk = true;
            for (int row = 1; row < tableLayoutPanel1.RowCount - 3; row++)
            {
                if(rowIsOk==false)
                {
                    break;
                }
                decimal inAmountScore = 0;
                decimal inGoalScore = 0;
                decimal inSpecAmountScore = 0;
                for (int col = 1; col <= 3; col++)
                {

                    var inputControl = tableLayoutPanel1.GetControlFromPosition(col, row);
                    var outControl = tableLayoutPanel1.GetControlFromPosition(col + 3, row);
                    if (inputControl is Label)
                    {

                        // 设置Label的值为103
                        // ((Label)control).Text = "103";
                    }
                    else if (inputControl is NumericUpDown)
                    {
                        // 设置NumericUpDown的值为103
                        if (col == 1)
                        {
                            inAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.03));//写公式上去

                        }
                        else if (col == 2)
                        {
                            inGoalScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.000125));

                        }
                        else if (col == 3)
                        {
                            inSpecAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.1));
                        }
                    }
                    if(inSpecAmountScore==Convert.ToDecimal(0)&& inGoalScore== Convert.ToDecimal(0) && inAmountScore == Convert.ToDecimal(0)&&row<=7)
                    {
                        MessageBox.Show("第" + (row ).ToString() + "行还没有输入数据\n最少输入7行数据才能计算");
                        rowIsOk = false;
                        break;
                    }
                }
            }


            if(rowIsOk==false)//必须最少输入7行才可以计算
            {
                return;
            }
            for (int row = 1; row < tableLayoutPanel1.RowCount-3; row++)
            {
                decimal inAmountScore = 0;
                decimal inGoalScore = 0;
                decimal inSpecAmountScore = 0;
                for (int col = 1; col <= 3; col++)
                {
                   
                    var inputControl = tableLayoutPanel1.GetControlFromPosition(col, row);
                    var outControl = tableLayoutPanel1.GetControlFromPosition(col+3, row);
                    if (inputControl != null)
                    {
                        if (inputControl is Label)
                        {

                            // 设置Label的值为103
                           // ((Label)control).Text = "103";
                        }
                        else if (inputControl is NumericUpDown)
                        {
                            // 设置NumericUpDown的值为103
                            if (col == 1)
                            {
                                inAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.03));//写公式上去
                                if(inAmountScore>60)
                                {
                                    inAmountScore = 60;
                                }
                                ((Label)outControl).Text = inAmountScore.ToString();
                            }else if(col == 2)
                            {
                                inGoalScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.000125));
                                if (inGoalScore > 35)
                                {
                                    inGoalScore = 35;
                                }
                                ((Label)outControl).Text =inGoalScore.ToString();//写公式上去
                             }
                            else if (col == 3)
                            {
                                inSpecAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.1));
                                if (inSpecAmountScore > 5)
                                {
                                    inSpecAmountScore = 5;
                                }
                                ((Label)outControl).Text = inSpecAmountScore.ToString();//写公式上去
                            }
                        }
                    }
                }
                if(inAmountScore==0&&inGoalScore==0&&inSpecAmountScore==0)
                {
                    tableLayoutPanel1.GetControlFromPosition(4, row).Text="0";
                    tableLayoutPanel1.GetControlFromPosition(5, row).Text = "0";
                    tableLayoutPanel1.GetControlFromPosition(6, row).Text = "0";
                    break;
                }
                var scoreControl = tableLayoutPanel1.GetControlFromPosition(7, row);
                ((Label)scoreControl).Text = ((inAmountScore * Convert.ToDecimal(0.6) + inGoalScore * Convert.ToDecimal(0.35) + inSpecAmountScore * Convert.ToDecimal(0.05))/Convert.ToDecimal(row)).ToString();
            }
        }
    }
}
