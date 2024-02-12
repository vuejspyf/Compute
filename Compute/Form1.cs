using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
            DataTable tblDatas = new DataTable("Datas");
            DataColumn dc = null;

            dc = tblDatas.Columns.Add("项目", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("得分", Type.GetType("System.String"));
            /*
            for (int col = 1; col < tableLayoutPanel1.ColumnCount; col++)
            {
                dc = tblDatas.Columns.Add( ((Label)tableLayoutPanel1.GetControlFromPosition(col, 0 )).Text, Type.GetType("System.Decimal")); 
              
            }*/
            decimal rowCount = 0;
            decimal[] inAmountScoreArr = new decimal[15];
            decimal[] inGoalScoreArr = new decimal[15];
            decimal[] inSpecAmountScoreArr = new decimal[15];
            for (int row = 1; row < tableLayoutPanel1.RowCount - 1; row++)
            {
                decimal inAmountScore = 0;
                decimal inGoalScore = 0;
                decimal inSpecAmountScore = 0;
              //  DataRow newRow;
              //  newRow = tblDatas.NewRow();
               
                for (int col = 0; col < tableLayoutPanel1.ColumnCount; col++)
                {             
                    var inputControl = tableLayoutPanel1.GetControlFromPosition(col, row);
                    var outControl = tableLayoutPanel1.GetControlFromPosition(col + 3, row);
                    if (inputControl != null)
                    {
                        if (inputControl is Label)
                        {

                            // 设置Label的值为103
                            //newRow[col] = ((Label)inputControl).Text;
                        }
                        else if (inputControl is NumericUpDown)
                        {

                            //newRow[col] = ((NumericUpDown)inputControl).Value;
                            // 设置NumericUpDown的值为103
                            if (col == 1)
                            {
                                inAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.03));//写公式上去
                                if (inAmountScore > 60)
                                {
                                    inAmountScore = 60;
                                }
                            }
                            else if (col == 2)
                            {
                                inGoalScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.000125));
                                if (inGoalScore > 35)
                                {
                                    inGoalScore = 35;
                                }
                            }
                            else if (col == 3)
                            {
                                inSpecAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.1));
                                if (inSpecAmountScore > 5)
                                {
                                    inSpecAmountScore = 5;
                                }
                            }
                        }
                    }
                }

                if (inAmountScore == 0 && inGoalScore == 0 && inSpecAmountScore == 0)
                {
                    break;
                }
                rowCount++;
                inAmountScoreArr[row - 1] = inAmountScore;
                inGoalScoreArr[row - 1] = inGoalScore ;
                inSpecAmountScoreArr[row - 1] = inSpecAmountScore;
                //tblDatas.Rows.Add(newRow);
            }
            if(rowCount==0)
            {
                MessageBox.Show("没有输入数据");
                return;
            }    
            DataRow newRow;
            newRow = tblDatas.NewRow();
            newRow["项目"] = "平均购进条数";
            newRow["得分"] = (inAmountScoreArr.Sum() / rowCount).ToString("0.0000");
            tblDatas.Rows.Add(newRow);

            newRow = tblDatas.NewRow();
            newRow["项目"] = "平均购进金额";
            newRow["得分"] = (inGoalScoreArr.Sum() / rowCount).ToString("0.0000");
            tblDatas.Rows.Add(newRow);

            newRow = tblDatas.NewRow();
            newRow["项目"] = "平均购进品规";
            newRow["得分"] = (inSpecAmountScoreArr.Sum() / rowCount).ToString("0.0000");
            tblDatas.Rows.Add(newRow);
            newRow = tblDatas.NewRow();
            newRow["项目"] = "总分";
            newRow["得分"] = (inSpecAmountScoreArr.Sum() / rowCount+ inGoalScoreArr.Sum() / rowCount+ inAmountScoreArr.Sum() / rowCount).ToString("0.0000");
            tblDatas.Rows.Add(newRow);
            /*
            dc = tblDatas.Columns.Add("ID", Type.GetType("System.Int32"));
            dc.AutoIncrement = true;//自动增加
            dc.AutoIncrementSeed = 1;//起始为1
            dc.AutoIncrementStep = 1;//步长为1
            dc.AllowDBNull = false;//

            dc = tblDatas.Columns.Add("Product", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("Version", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("Description", Type.GetType("System.String"));


            DataRow newRow;
            newRow = tblDatas.NewRow();
            newRow[1] = "大话西游";
            newRow[2] = "2.0";
            newRow[3] = "我很喜欢";
            
            tblDatas.Rows.Add(newRow);

            newRow = tblDatas.NewRow();
            newRow["Product"] = "梦幻西游";
            newRow["Version"] = "3.0";
            newRow["Description"] = "比大话更幼稚";
            tblDatas.Rows.Add(newRow);
            */

            DataTable dtTable = tblDatas;
            string sheetName = "导出数据";//设置工作簿的名称
            IWorkbook wb = new HSSFWorkbook();//创建一个工作簿对象
            ISheet sheet = string.IsNullOrEmpty(sheetName) ? wb.CreateSheet("sheet1") : wb.CreateSheet(sheetName);//设置工作簿，如果名称为空，则默认名称为sheet1
            int rowIndex = 0;//默认行索引为0，也就是第一行
          
            if (dtTable.Columns.Count > 0)//如果工作表里有数据，也就是有列
            {
                IRow header = sheet.CreateRow(rowIndex);//创建第一行
               
                for (int i = 0; i < dtTable.Columns.Count; i++)//遍历所有列
                {
                    ICell cell = header.CreateCell(i);//创建单元格
                    cell.SetCellValue(dtTable.Columns[i].ColumnName);//填充列名
                    ICellStyle style = cell.CellStyle;
                   //style.WrapText = true;  
                    style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    style.VerticalAlignment = VerticalAlignment.Center;
                    IFont font = wb.CreateFont();
                    font.FontName = "微软雅黑";
                    font.IsBold = true;
                    font.FontHeightInPoints = 42;
                    style.SetFont(font);
                    cell.CellStyle = style;
                    
                }
            }
            //添加数据
            if (dtTable.Rows.Count > 0)//如果有数据，也就是有行
            {
                for (int i = 0; i < dtTable.Rows.Count; i++)//遍历行
                {
                    rowIndex++;//行索引＋1
                    IRow row = sheet.CreateRow(rowIndex);//创建行对象
                    for (int j = 0; j < dtTable.Columns.Count; j++)//遍历行里面的每一列
                    {
                        ICell cell = row.CreateCell(j);//创建单元格
                        cell.SetCellValue(dtTable.Rows[i][j].ToString());//将DataTable里的值添加到工作簿中
                    }
                }
            }
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);//自适应单元格大小
            }
            string  fileName = @"F:\AAAB.xlsx";
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                fileName = folder.SelectedPath + "\\数据导出" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                //RunAsAdmin(fileName);
                using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite))//创建文件流
                {
                    wb.Write(fs);//把工作簿写入流
                }
                MessageBox.Show("导出成功");//提示导出成功
            }
          
        }

        public static void RunAsAdmin(string fileName)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo(fileName);
            startInfo.UseShellExecute = true;
            startInfo.Verb = "runas";
            Process.Start(startInfo);
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
            for (int row = 1; row < tableLayoutPanel1.RowCount - 1; row++)
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
                        MessageBox.Show("第" + (row ).ToString() + "次还没有输入数据\n最少输入7次数据才能计算");
                        rowIsOk = false;
                        break;
                    }
                }
            }


            if(rowIsOk==false)//必须最少输入7行才可以计算
            {
                return;
            }
            decimal[] totalScoreArr = new decimal[15];
            for (int row = 1; row < tableLayoutPanel1.RowCount-1; row++)
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
                                ((Label)outControl).Text = inAmountScore.ToString("0.0000");
                            }else if(col == 2)
                            {
                                inGoalScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.000125));
                                if (inGoalScore > 35)
                                {
                                    inGoalScore = 35;
                                }
                                ((Label)outControl).Text =inGoalScore.ToString("0.0000");//写公式上去
                             }
                            else if (col == 3)
                            {
                                inSpecAmountScore = ((((NumericUpDown)inputControl).Value) * Convert.ToDecimal(0.1));
                                if (inSpecAmountScore > 5)
                                {
                                    inSpecAmountScore = 5;
                                }
                                ((Label)outControl).Text = inSpecAmountScore.ToString("0.0000");//写公式上去
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
                decimal totalScore = (inAmountScore + inGoalScore + inSpecAmountScore);
                decimal summary = 0;
                totalScoreArr[row-1] =(totalScore);
                for (global::System.Int32 i = 0; i < row; i++)
                {
                    summary += totalScoreArr[i];
                }
                ((Label)scoreControl).Text = (summary/row).ToString("0.0000");
                
            }
        }
    }
}
