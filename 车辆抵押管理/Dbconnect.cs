using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace 车辆抵押管理
{
    public class Dbconnect
    {
        static string path = Application.StartupPath;

        private static string connection(string path)
        {
            string ConnectionStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + @"\car.mdb;Jet OLEDB:Database Password=jaywatson";
            return ConnectionStr;
         }
        
        public static OleDbDataAdapter datareader()
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co=connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "select ID,客户姓名,经办支行,卡号,审批金额,车辆价格,购买车型,车行,客户经理姓名,系统申请,二级分行审核日期,抵押情况,放款日期,车辆抵押日期 from 车辆抵押信息";//查询student表  
            da.SelectCommand = comm;
            return da;
        }

        public static OleDbDataAdapter datareader(string title,string value)
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "select ID,客户姓名,经办支行,卡号,审批金额,车辆价格,购买车型,车行,客户经理姓名,系统申请,二级分行审核日期,抵押情况,放款日期,车辆抵押日期 from 车辆抵押信息 where " + title + " like '%" + value + "%'";//查询student表  
            da.SelectCommand = comm;
            return da;
        }

        public static OleDbDataReader changereader(int id)
        {
            OleDbCommand comm = new OleDbCommand();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "select 客户姓名,经办支行,卡号,审批金额,车辆价格,购买车型,车行,客户经理姓名,系统申请,二级分行审核日期,抵押情况,放款日期,车辆抵押日期 from 车辆抵押信息 where ID=" + id + "";//查询student表  
            OleDbDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            return sdr;
        }

        public void deletedata(string name,string path)
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "delete from 车辆抵押信息 where ID="+name+"";//删除数据
            int i = comm.ExecuteNonQuery();
            if(i==0)
            {
                throw new Exception("没有删除数据");
            }
            conn.Close();
        }

        public static OleDbDataReader bankcombobox()
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "select 支行名称 from 支行信息";//删除数据
            OleDbDataReader sdr = comm.ExecuteReader();
            return sdr;
        }

        public static OleDbDataReader carcombobox()
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "select 车行名称 from 车行信息";//删除数据
            OleDbDataReader sdr = comm.ExecuteReader();
            return sdr;
        }

        public void addcarcombobox(string car1)
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            int i = 0;
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText ="select * from 车行信息 where 车行名称='"+car1+"'";//删除数据
            OleDbDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            try
            {
                if (sdr["车行名称"].ToString() != "") { }
            }
            catch
            {
                conn.Close();
                conn.Open();
                comm.Connection = conn;
                comm.CommandType = CommandType.Text;
                comm.CommandText = "Insert into  车行信息(车行名称)values('" + car1 + "')";//删除数据
                i = comm.ExecuteNonQuery();
                if (i == 0)
                {
                    throw new Exception("没有更新数据");
                }
            }
            conn.Close();
        }

        public void addnewdata(string name, string bank, string car,string manager,string cardid,string exampric,string carpric,string cartype, int mortgate,string date,string datecontrol,int system,string examdate)
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "insert into 车辆抵押信息(客户姓名,车行,经办支行,客户经理姓名,卡号,审批金额,车辆价格,购买车型,抵押情况,放款日期,车辆抵押日期,系统申请,二级分行审核日期) values('" + name + "','" + car + "','" + bank + "','" + manager + "','" + cardid + "','" + exampric + "','" + carpric + "','" + cartype + "','" + mortgate + "','" + date + "','" + datecontrol + "','" + system + "','" + examdate + "')";//删除数据
            int i = comm.ExecuteNonQuery();
            if (i == 0)
            {
                throw new Exception("没有更新数据");
            }
            conn.Close();
        }

        public void updatedata(int ID, string name, string bank, string car, string manager, string cardid, string exampric, string carpric, string cartype, int mortgate, string date, string datecontrol, int system, string examdate)
        {
            OleDbCommand comm = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //access数据库连接字符串，自行更改数据库路径和名字 
            string co = connection(path);
            OleDbConnection conn = new OleDbConnection(co);
            conn.Open();
            comm.Connection = conn;
            comm.CommandType = CommandType.Text;
            comm.CommandText = "Update 车辆抵押信息 set 客户姓名 ='" + name + "',车行='" + car + "',经办支行='" + bank + "',客户经理姓名='" + manager + "',卡号='" + cardid + "',审批金额='" + exampric + "',车辆价格='" + carpric + "',购买车型='" + cartype + "',抵押情况='" + mortgate + "',放款日期='" + date + "',车辆抵押日期='" + datecontrol + "',系统申请='" + system + "',二级分行审核日期='" + examdate + "' where ID=" + ID + "";//删除数据
            int i = comm.ExecuteNonQuery();
            if (i == 0)
            {
                throw new Exception("没有更新数据");
            }
            conn.Close();
        }
    }
}
