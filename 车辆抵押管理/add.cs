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
    public partial class add : CCSkinMain
    {
        public add()
        {
            InitializeComponent();
        }

        private void datecontrol()//天数控制
        {
            try
            {
                if(MoneyOut.text=="")
                {
                    MoneyOut.text = DateTime.Now.ToString("yyyy-M-d");
                }
                ControlDate.text = DateTime.Parse(MoneyOut.text).AddMonths(1).AddDays(-1).ToString("yyyy-M-d");
            }
            catch { }
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
        private void add_Load(object sender, EventArgs e)
        {
            datecontrol();
            droplist();
            this.carComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.carComboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
            examineDate.text = DateTime.Now.ToString("yyyy-M-d");
            nomortgate.Checked = true;
            newsystem.Checked = true;
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            added();
            Form1 form = new Form1();
            form.Focus();
            form.dateshow();
            this.Hide();
        }

        private void added()
        {
            string name,bank,car,manager,cardid,exampric,carpric,cartype,date,datecontrol,examdate;
            int Mortgate = 0,system = 0;
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
            if (name != "" && bank != "" && car != "" && manager != "" && cartype != "" && date != "" && datecontrol != ""&& cardid !=""&& exampric != "")
            {
                try
                {
                    db.addcarcombobox(car);
                    db.addnewdata(name,bank,car,manager,cardid,exampric,carpric,cartype,Mortgate,date,datecontrol,system,examdate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }
        private void skinButton3_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1();
            form.Focus();
            form.dateshow();
            this.Hide();
        }

        private void skinButton2_Click(object sender, EventArgs e)
        {
            added();
            nameBox.Text = "";
            managerBox.Text = ""; carTypeBox.Text = ""; carPriceBox.Text = "";
            MoneyOut.Text = ""; ControlDate.Text = ""; examineMoneyBox.Text = "";
            bankComboBox.Text = ""; cardIdBox.Text = "";
            if (carCheckBox.Checked != true)
            {
                carComboBox.Text = "";
            }
        }
    }
}
