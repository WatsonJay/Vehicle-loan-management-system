using CCWin;
using org.in2bits.MyXls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace 车辆抵押管理
{
    public partial class Form1 : CCSkinMain
    {
        #region 变量
        bool i=false;
        string path;
        int recordCount =0; // 要导出的记录总数  
        int maxRecordCount = 100; // 每个sheet表的最大记录数  
        int sheetCount = 1; // Sheet表的数目 
        string sheetName="车辆抵押信息表";
        private Class1.car_Param carParam;
        int Column = 0;//列数
        #endregion

        #region 初始化启动
        public Form1()
        {
            Helper.eventSend += new SendHandler(ReceiveParam);
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            path = Application.StartupPath;
            dateshow();
            toolbox.Hide();
            date2.text = DateTime.Now.ToString("yyyy/M/d");
            date1.text = DateTime.Parse(date2.text).AddMonths(-1).AddDays(1).ToString("yyyy/M/d");
            maxLine.Text = maxRecordCount.ToString();
            type.SelectedIndex = 0;
            dateType.SelectedIndex = 0;
            userradio.Checked=true;
        }
        #endregion

        #region 显示刷新
        public void dateshow()
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            DataView dv = new DataView();
            try
            {
                da = Dbconnect.datareader();
                
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                return;
            }
            da.Fill(ds);
            dv = ds.Tables[0].DefaultView;
            dv.AllowNew = false;
            this.DATAGRID.DataSource = dv;//表从起始行显示在dataGridView里 
            this.DATAGRID.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            DATAGRID.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DATAGRID.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                this.DATAGRID.Rows[i].Cells[0].Value = "False";
            }
            this.DATAGRID.ReadOnly = false;
            this.DATAGRID.Columns["ID"].Visible = false;
            for (int a = 0; a < ds.Tables[0].Columns.Count;a++ )
            {
                this.DATAGRID.Columns[a+2].ReadOnly = true;
            }
            backgroundchange();
        }

        public void dateshow(string title,string value)
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            DataView dv = new DataView();
            try
            {
                da = Dbconnect.datareader(title,value);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return;
            }
            da.Fill(ds);
            dv = ds.Tables[0].DefaultView;
            dv.AllowNew = false;
            this.DATAGRID.DataSource = dv;//表从起始行显示在dataGridView里 
            this.DATAGRID.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            DATAGRID.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DATAGRID.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                this.DATAGRID.Rows[i].Cells[0].Value = "False";
            }
            this.DATAGRID.ReadOnly = false;
            this.DATAGRID.Columns["ID"].Visible = false;
            for (int a = 0; a < ds.Tables[0].Columns.Count; a++)
            {
                this.DATAGRID.Columns[a + 2].ReadOnly = true;
            }
            backgroundchange();
        }
        #endregion

        #region 超时变色

        private void backgroundchange()
        {
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                string datecontrol = this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 1].Value.ToString();
                string morgate = this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 3].Value.ToString();
                string datenow=DateTime.Now.ToString("yyyy/M/d");
                DateTime dt1 = Convert.ToDateTime(datecontrol);
                DateTime dt2 = Convert.ToDateTime(datenow);
                TimeSpan ts = dt1 - dt2;
                int sub = ts.Days;
                if (sub<=0&&morgate=="False")
                {
                    for (int k = 0; k < DATAGRID.ColumnCount; k++)
                    {
                        DATAGRID.Rows[i].Cells[k].Style.BackColor = System.Drawing.Color.Yellow;
                    }
                }
            }
        }

        #endregion

        #region 批量删除

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ArrayList name = new ArrayList();
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                string b; 
                DataGridViewCell checkbox = (DataGridViewCell)this.DATAGRID.Rows[i].Cells[0];
                if ((string)checkbox.Value == "True")
                {
                    b = DATAGRID.Rows[i].Cells[2].Value.ToString();
                    name.Add(b);
                }
             }
            if (MessageBox.Show("确定删除" + name.Count + "条数据吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk) == DialogResult.OK)
            {
                Dbconnect dbconnect=new Dbconnect();
                foreach (string a in name)
                {
                    try
                    {
                        dbconnect.deletedata(a,path);
                    }
                    catch(Exception ex)
                    {
                       MessageBox.Show(ex.Message);
                    }
                }
                this.DATAGRID.ReadOnly = false;
                dateshow();
            }
        }

        #endregion

        #region 其他

        private void toolStripButton3_Click(object sender, EventArgs e) //刷新
        {
            dateshow();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            dateshow();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)//新增
        {
            add add = new add();
            add.Show();
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            if (!i)
            {
                toolbox.Show();
            }
            else
            {
                toolbox.Hide();
            }
            i = !i;
        }

        private void DATAGRID_Sorted(object sender, EventArgs e)
        {
            backgroundchange();
        }

        private void skinButton3_Click(object sender, EventArgs e) //单表最大行数
        {
            int.TryParse(maxLine.Text, out maxRecordCount);
        }

        private void skinButton4_Click(object sender, EventArgs e) //时间段选择
        {
            DateTime dt1 = new DateTime(), dt2 = new DateTime(), dt3 = new DateTime();
            TimeSpan ts1,ts2;
            string datenow = DateTime.Now.ToString("yyyy/M/d");
            if (userradio.Checked == true)
            {
                dt3 = Convert.ToDateTime(date2.Text);
                dt2 = Convert.ToDateTime(date1.Text);
            }
            else if (threeMonth.Checked)
            {
                dt3 = Convert.ToDateTime(datenow);
                dt2 = (dt3.AddMonths(-3).AddDays(1));
            }
            else if (sixMonth.Checked)
            {
                dt3 = Convert.ToDateTime(datenow);
                dt2 = (dt3.AddMonths(-6).AddDays(1));
            }
            else if (nineMonth.Checked)
            {
                dt3 = Convert.ToDateTime(datenow);
                dt2 = (dt3.AddMonths(-9).AddDays(1));
            }
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                string datecontrol = "";
                switch(dateType.SelectedIndex)
                {
                    case 0:  datecontrol= this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 4].Value.ToString(); break;
                    case 1:  datecontrol = this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 2].Value.ToString(); break;
                    case 2:  datecontrol = this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 1].Value.ToString(); break;
                }
                dt1 = Convert.ToDateTime(datecontrol);
                ts1 = dt3 - dt1; ts2 = dt2 - dt1;
                int sub1 = ts1.Days; int sub2 = ts2.Days;
                if (sub1>0 && sub2<0)
                {
                    this.DATAGRID.Rows[i].Cells[0].Value = "True";
                }
                else
                {
                    this.DATAGRID.Rows[i].Cells[0].Value = "False";
                }
            }
            this.DATAGRID.ReadOnly = false;
        }

        private void chooseAll_Click(object sender, EventArgs e) //全选
        {
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                this.DATAGRID.Rows[i].Cells[0].Value = "True";
            }
            this.DATAGRID.ReadOnly = false;
        }

        private void chooseNone_Click(object sender, EventArgs e) //全不选
        {
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                this.DATAGRID.Rows[i].Cells[0].Value = "False";
            }
            this.DATAGRID.ReadOnly = false;
        }

        private void chooseBack_Click(object sender, EventArgs e) //反选
        {
            this.DATAGRID.ReadOnly = true;
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                if ((string)this.DATAGRID.Rows[i].Cells[0].Value == "True")
                    this.DATAGRID.Rows[i].Cells[0].Value = "False";
                else if ((string)this.DATAGRID.Rows[i].Cells[0].Value == "False")
                    this.DATAGRID.Rows[i].Cells[0].Value = "True";

            }
            this.DATAGRID.ReadOnly = false;
        }

        private void skinButton5_Click(object sender, EventArgs e) // 查询
        {
            string title, value;
            title = type.Text;
            value = typevalue.Text;
            dateshow(title, value);
        }
        #endregion

        #region 保存到Excel

            #region 按键事件
        private void skinButton2_Click(object sender, EventArgs e)
        {
            List<int> list = new List<int>();
            ArrayList title = new ArrayList();
            string oldstring,newstring;
            int i = 0,k = 0;
            DataTable dt = new DataTable();
            this.DATAGRID.ReadOnly = true;
            for (i = 3; i < DATAGRID.ColumnCount; i++)  //列标题
            {
                Column += 1;
                dt.Columns.Add(DATAGRID.Columns[i].HeaderText);
                title.Add(DATAGRID.Columns[i].HeaderText);
            }
            for (i = 0; i < DATAGRID.RowCount; i++)   //列内容导入
            {
                
                DataGridViewCell checkbox = (DataGridViewCell)this.DATAGRID.Rows[i].Cells[0];
                if ((string)checkbox.Value == "True")
                {
                    DataRow dr = dt.NewRow();
                    recordCount += 1;
                    for (k = 3; k < DATAGRID.ColumnCount; k++)
                    {
                        if (DATAGRID.ColumnCount - k < 3 || DATAGRID.ColumnCount - k == 4)
                        {
                            oldstring = DATAGRID.Rows[i].Cells[k].Value.ToString();
                            newstring = oldstring.Replace("0:00:00", "");
                            dr[k - 3] = newstring;
                        }
                        else
                        {
                            if (DATAGRID.ColumnCount - k == 3)
                            {
                                if (DATAGRID.Rows[i].Cells[k].Value.ToString() == "True")
                                {
                                    oldstring = DATAGRID.Rows[i].Cells[k].Value.ToString();
                                    newstring = oldstring.Replace("True", "已抵押");
                                    dr[k - 3] = newstring;
                                }
                                if (DATAGRID.Rows[i].Cells[k].Value.ToString() == "False")
                                {
                                    oldstring = DATAGRID.Rows[i].Cells[k].Value.ToString();
                                    newstring = oldstring.Replace("False", "未抵押");
                                    dr[k - 3] = newstring; ;
                                }
                            }
                            else
                            {
                                if (DATAGRID.ColumnCount - k == 5)
                                {
                                    if (DATAGRID.Rows[i].Cells[k].Value.ToString() == "True")
                                    {
                                        oldstring = DATAGRID.Rows[i].Cells[k].Value.ToString();
                                        newstring = oldstring.Replace("True", "新系统");
                                        dr[k - 3] = newstring;
                                    }
                                    if (DATAGRID.Rows[i].Cells[k].Value.ToString() == "False")
                                    {
                                        oldstring = DATAGRID.Rows[i].Cells[k].Value.ToString();
                                        newstring = oldstring.Replace("False", "旧系统");
                                        dr[k - 3] = newstring; ;
                                    }
                                }
                                else
                                    dr[k - 3] = DATAGRID.Rows[i].Cells[k].Value.ToString();
                            }
                        }
                     }
                    dt.Rows.Add(dr);
                }
             }

            if (recordCount > maxRecordCount)
            {
                sheetCount = (int)Math.Ceiling((decimal)recordCount / (decimal)maxRecordCount);
            }

            string date = DateTime.Now.ToString("yyyyMMdd");
            CreateExcel("车辆抵押信息"+date,title, dt);
            this.DATAGRID.ReadOnly = false; Column = 0; recordCount = 0; maxRecordCount = 100;
        }
        #endregion

            #region 导出函数
        public void CreateExcel(string fileName,ArrayList title, DataTable DataSource)
        {
            XlsDocument xls = new XlsDocument();

            #region Sheet标题样式
            XF titleXF = xls.NewXF(); // 为xls生成一个XF实例，XF是单元格格式对象  
            titleXF.HorizontalAlignment = HorizontalAlignments.Centered; // 设定文字居中  
            titleXF.VerticalAlignment = VerticalAlignments.Centered; // 垂直居中  
            titleXF.UseBorder = true; // 使用边框   
            titleXF.BottomLineStyle = 1; // 下边框样式  
            titleXF.BottomLineColor = Colors.Black; // 下边框颜色  
            titleXF.Font.FontName = "宋体"; // 字体  
            titleXF.Font.Bold = true; // 是否加楚  
            titleXF.Font.Height = 12 * 20; // 字大小（字体大小是以 1/20 point 为单位的） 
            #endregion

            #region 列标题样式
            XF columnTitleXF = xls.NewXF(); // 为xls生成一个XF实例，XF是单元格格式对象  
            columnTitleXF.HorizontalAlignment = HorizontalAlignments.Centered; // 设定文字居中  
            columnTitleXF.VerticalAlignment = VerticalAlignments.Centered; // 垂直居中  
            columnTitleXF.UseBorder = true; // 使用边框   
            columnTitleXF.TopLineStyle = 1; // 上边框样式  
            columnTitleXF.TopLineColor = Colors.Black; // 上边框颜色  
            columnTitleXF.BottomLineStyle = 1; // 下边框样式  
            columnTitleXF.BottomLineColor = Colors.Black; // 下边框颜色  
            columnTitleXF.LeftLineStyle = 1; // 左边框样式  
            columnTitleXF.LeftLineColor = Colors.Black; // 左边框颜色  
            columnTitleXF.RightLineStyle = 1; // 右边框样式  
            columnTitleXF.RightLineColor = Colors.Black; // 右边框颜色  
            #endregion

            #region 数据单元格样式
            XF dataXF = xls.NewXF(); // 为xls生成一个XF实例，XF是单元格格式对象  
            dataXF.HorizontalAlignment = HorizontalAlignments.Centered; // 设定文字居中  
            dataXF.VerticalAlignment = VerticalAlignments.Centered; // 垂直居中  
            dataXF.UseBorder = true; // 使用边框   
            dataXF.TopLineStyle = 1; // 上边框样式  
            dataXF.TopLineColor = Colors.Black; // 上边框颜色  
            dataXF.BottomLineStyle = 1; // 下边框样式  
            dataXF.BottomLineColor = Colors.Black; // 下边框颜色  
            dataXF.LeftLineStyle = 1; // 左边框样式  
            dataXF.LeftLineColor = Colors.Black; // 左边框颜色  
            dataXF.RightLineStyle = 1; // 右边框样式  
            dataXF.RightLineColor = Colors.Black; // 右边框颜色  
            dataXF.Font.FontName = "宋体";
            dataXF.Font.Height = 9 * 20; // 设定字大小（字体大小是以 1/20 point 为单位的）  
            dataXF.UseProtection = false; // 默认的就是受保护的，导出后需要启用编辑才可修改  
            dataXF.TextWrapRight = true; // 自动换行  
            #endregion

            #region 数据单元格样式
            XF dataXFchange = xls.NewXF(); // 为xls生成一个XF实例，XF是单元格格式对象  
            dataXFchange.HorizontalAlignment = HorizontalAlignments.Centered; // 设定文字居中  
            dataXFchange.VerticalAlignment = VerticalAlignments.Centered; // 垂直居中  
            dataXFchange.UseBorder = true; // 使用边框   
            dataXFchange.TopLineStyle = 1; // 上边框样式  
            dataXFchange.TopLineColor = Colors.Black; // 上边框颜色  
            dataXFchange.BottomLineStyle = 1; // 下边框样式  
            dataXFchange.BottomLineColor = Colors.Black; // 下边框颜色  
            dataXFchange.LeftLineStyle = 1; // 左边框样式  
            dataXFchange.LeftLineColor = Colors.Black; // 左边框颜色  
            dataXFchange.RightLineStyle = 1; // 右边框样式  
            dataXFchange.RightLineColor = Colors.Black; // 右边框颜色  
            dataXFchange.Font.FontName = "宋体";
            dataXFchange.Font.Height = 9 * 20; // 设定字大小（字体大小是以 1/20 point 为单位的）  
            dataXFchange.UseProtection = false; // 默认的就是受保护的，导出后需要启用编辑才可修改  
            dataXFchange.TextWrapRight = true; // 自动换行 
            dataXFchange.Pattern = 1; // 单元格填充风格。如果设定为0，则是纯色填充(无色)，1代表没有间隙的实色   
            dataXFchange.PatternBackgroundColor = Colors.Red; // 填充的底色   
            dataXFchange.PatternColor = Colors.Default2F; // 填充背景色  
            #endregion

            for (int i = 1; i <= sheetCount; i++)
            {
                // 根据计算出来的Sheet数量，一个个创建  
                // 行和列的设置需要添加到指定的Sheet中，且每个设置对象不能重用（因为可以设置起始和终止行或列，就没有太大必要重用了，这应是一个策略问题）  
                Worksheet sheet;
                if (sheetCount == 1)
                {
                    sheet = xls.Workbook.Worksheets.Add(sheetName);
                }
                else
                {
                    sheet = xls.Workbook.Worksheets.Add(sheetName + " - " + i);
                }

                ColumnInfo columnInfo = new ColumnInfo(xls, sheet);
                columnInfo.ColumnIndexStart = 0;
                columnInfo.ColumnIndexEnd = (ushort)(DataSource.Columns.Count - 1);
                columnInfo.Width = 15 * 330;
                sheet.AddColumnInfo(columnInfo);

                // 合并单元格  
                //sheet.Cells.Merge(1, 1, 1, 4);  
                MergeArea titleArea = new MergeArea(1, 1, 1, Column); // 一个合并单元格实例(合并第1行、第1列 到 第1行、第4列)   
                sheet.AddMergeArea(titleArea); //填加合并单元格   

                // 开始填充数据到单元格  
                Cells cells = sheet.Cells;

                // Sheet标题行，行和列的索引都是从1开始的  
                Cell cell = cells.Add(1, 1, sheetName, titleXF);
                for (i = 2; i <= Column;i++ )
                {
                    cells.Add(1, i, "", titleXF);
                }

                for (int j = 0; j < DataSource.Columns.Count; j++)
                {
                    cells.Add(2, j + 1, DataSource.Columns[j].ColumnName, columnTitleXF);
                }

                if (DataSource.Rows.Count == 0)
                {
                    cells.Add(3, 1, "没有可用数据");
                    return;
                }

                for (int l = 0; l < DataSource.Rows.Count; l++)
                {
                    string datecontrol = DataSource.Rows[l][DataSource.Columns.Count-1].ToString();
                    string morgate = DataSource.Rows[l][DataSource.Columns.Count-3].ToString();
                    string datenow=DateTime.Now.ToString("yyyy/M/d");
                    DateTime dt1 = Convert.ToDateTime(datecontrol);
                    DateTime dt2 = Convert.ToDateTime(datenow);
                    TimeSpan ts = dt1 - dt2;
                    int sub = ts.Days;
                    for (int j = 0; j < DataSource.Columns.Count; j++)

                        if (sub<=0&&morgate=="未抵押")
                        {
                            cells.Add(l + 3, j + 1, DataSource.Rows[l][j].ToString(),dataXFchange);
                        }
                        else
                        {
                            cells.Add(l + 3, j + 1, DataSource.Rows[l][j].ToString(), dataXF);
                        }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "保存车辆抵押信息到Excel文件";
                saveFileDialog.Filter = "Excel文件|*.xls";
                saveFileDialog.FileName = fileName;
                saveFileDialog.RestoreDirectory = true;
                DialogResult result = saveFileDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    int fileI = saveFileDialog.FileName.LastIndexOf("\\");
                    xls.FileName = saveFileDialog.FileName.Substring(fileI + 1, saveFileDialog.FileName.Length - (fileI + 1));
                    string path = saveFileDialog.FileName.Substring(0, fileI + 1);
                    try
                    {
                        xls.Save(path, true);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("未导出成功，文件可能被占用", "提示", MessageBoxButtons.OK);
                    }
                    MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK);
                }
            }
        }
        #endregion
 
        #endregion

        #region 修改

        void ReceiveParam(object sender, object msg)
        {
            Type t = msg.GetType();
            if (t.IsEnum)
            {
                Form.eFrom e = (Form.eFrom)msg;
                switch (e)
                {
                    case Form.eFrom.Show_Change:
                        Showchange(sender as Class1.car_Param);
                        break;
                }
            }
        }

        delegate void ShowMainFrmEventHandler(Class1.car_Param carParam);
        private void Showchange(Class1.car_Param carParam)
        {
            if (this.InvokeRequired)
            {
                ShowMainFrmEventHandler cb = new ShowMainFrmEventHandler(Showchange);
                this.Invoke(cb, new object[] { carParam });
            }
            else
            {
                change change1 = new change(carParam);
                change1.ShowDialog();
            }

        }

        private void DATAGRID_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = DATAGRID.Columns[e.ColumnIndex];
            int i = DATAGRID.CurrentRow.Index;
            if (column is DataGridViewButtonColumn)
            {
                string ID = DATAGRID.Rows[i].Cells[2].Value.ToString();
                carParam = new Class1.car_Param();
                carParam.ID = Convert.ToInt32(ID);
                System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(GetbookLog), ID);
            }
        }

        private void GetbookLog(object o)
        {
            Helper.SendMessage(carParam, Form.eFrom.Show_Change);
        }
        #endregion

        #region 超时未抵押导出
        private void skinButton6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < DATAGRID.RowCount; i++)
            {
                string datecontrol = this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 1].Value.ToString();
                string morgate = this.DATAGRID.Rows[i].Cells[DATAGRID.ColumnCount - 3].Value.ToString();
                string datenow = DateTime.Now.ToString("yyyy/M/d");
                DateTime dt1 = Convert.ToDateTime(datecontrol);
                DateTime dt2 = Convert.ToDateTime(datenow);
                TimeSpan ts = dt1 - dt2;
                int sub = ts.Days;
                if (sub <= 0 && morgate == "False")
                {
                    this.DATAGRID.Rows[i].Cells[0].Value = "True";
                }
            }
            skinButton2_Click(sender,e);
        }
        #endregion
    }
}
