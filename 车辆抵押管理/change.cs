using CCWin;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace 车辆抵押管理
{
    public partial class change : CCSkinMain
    {
        private int ID;

        public change()
        {
            InitializeComponent();
        }

        public change(Class1.car_Param carParam):this()
        {
            //InitializeComponent();
            this.ID = carParam.ID;
            dataFill();
        }

        private void datecontrol()//天数控制
        {
            if (MoneyOut.text == "")
            {
                MoneyOut.text = DateTime.Now.ToString("yyyy-M-d");
            }
            ControlDate.text = DateTime.Parse(MoneyOut.text).AddMonths(1).AddDays(-1).ToString("yyyy-M-d");   
        }

        private void MoneyOut_SelectedValueChange(object sender, string Item)
        {
            datecontrol();
        }

        private void droplist()
        {
            OleDbDataReader bank, car;
            bank = Dbconnect.bankcombobox();
            while (bank.Read())
            {
                bankComboBox.Items.Add(bank[0].ToString());        //循环读区数据
            }
            car = Dbconnect.carcombobox();
            while (car.Read())
            {
                carComboBox.Items.Add(car[0].ToString());        //循环读区数据
            }
            
        }

        private void dataFill()
        {
            OleDbDataReader data;
            data = Dbconnect.changereader(ID);
            nameBox.Text = data["客户姓名"].ToString();
            carComboBox.Text = data["车行"].ToString();
            bankComboBox.Text = data["经办支行"].ToString();
            cardIdBox.Text = data["卡号"].ToString();
            managerBox.Text = data["客户经理姓名"].ToString();
            examineMoneyBox.Text = data["审批金额"].ToString();
            carPriceBox.Text = data["车辆价格"].ToString();
            carTypeBox.Text = data["购买车型"].ToString();
            if (data["抵押情况"].ToString() == "1")
                mortgate.Checked = true;
            else
                nomortgate.Checked = true;
            if (data["系统申请"].ToString() == "1")
                newsystem.Checked = true;
            else
                oldsystem.Checked = true;
            MoneyOut.text = data["放款日期"].ToString().Replace("0:00:00","");
            ControlDate.text = data["车辆抵押日期"].ToString().Replace("0:00:00", "");
            examineDate.text = data["二级分行审核日期"].ToString().Replace("0:00:00", "");
        } 
  
        private void update()
        {
            string name, bank, car, manager, cardid, exampric, carpric, cartype, date, datecontrol, examdate;
            int Mortgate = 0, system = 0;
            name = nameBox.Text; car = carComboBox.Text; bank = bankComboBox.Text; cardid = cardIdBox.Text;
            manager = managerBox.Text; exampric = examineMoneyBox.Text; carpric = carPriceBox.Text; cartype = carTypeBox.Text;
            if (newsystem.Checked == true)
            {
                system = 1;
            }
            if (oldsystem.Checked == true)
            {
                system = 0;
            }
            if (nomortgate.Checked == true)
            {
                Mortgate = 0;
            }
            if (mortgate.Checked == true)
            {
                Mortgate = 1;
            }
            date = MoneyOut.Text; datecontrol = ControlDate.Text; examdate = examineDate.text;
            Dbconnect db = new Dbconnect();
            if (name != "" && bank != "" && car != "" && manager != "" && cartype != "" && date != "" && datecontrol != "" && cardid != "" && exampric != "")
            {
                try
                {
                    db.addcarcombobox(car);
                    db.updatedata(ID,name, bank, car, manager, cardid, exampric, carpric, cartype, Mortgate, date, datecontrol, system, examdate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            update();
            Form1 form = new Form1();
            form.Focus();
            form.dateshow();
            this.Hide();
        }

        private void skinButton3_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1();
            form.Focus();
            form.dateshow();
            this.Hide();
        }

        private void change_Load(object sender, EventArgs e)
        {
            droplist();
            nomortgate.Checked = true;
            try
            {
                this.carComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                this.carComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
            catch { }
        }
    }
}
