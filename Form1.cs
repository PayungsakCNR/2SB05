using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc; //เรียกใช้ Namespace ชื่อ System.Data.Odbc
using System.Data.SqlClient; //เรียกใช้ System.Data.SqlClient

//สร้าง Namespace ชื่อ _2SB05
namespace _2SB05 
{
    public partial class Form1 : Form
    {
        //กำหนด conString เพื่อเข้าถึง Database โดยใส่ Path ของ .mdb ลงไป
        public string conString = @"Driver={Microsoft Access Driver (*.mdb)}; Dbq=D:\s5810110224\2SB05_Person.mdb";

        //กำหนดค่าตัวแปรข้างล้างนี้เป็น Null

        DataSet dSet = null;
        OdbcDataAdapter dAdapter = null;

        CurrencyManager currManager = null; 

        public Form1()
        {

            InitializeComponent();
        }

        OdbcConnection con = null; 

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ปุ่ม Close
            // กำหนด Event เมื่อกดปุ่มนี้
            //ถ้า con ไม่เท่ากับ null จะเรียก function con.Close ขึ้นมา และจะเเสดง MessageBox ว่า Closed!
            //ซึ่งเป็นการตัดการเชื่อมต่อกับ Database
            if (con != null)
            {
                try
                {
                    con.Close();
                    con = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            MessageBox.Show("Closed!"); 
        }


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //ปุ่ม Connect
            //Event เมื่อคลิกปุ่มนี้
            //คือ จะเป็นการเช็คว่า con เท่ากับ null มั๊ยย
            //ถ้าเท่ากับ null ก็จะไปสั่งให้เชื่อมต่อกับ Database 
            //แต่ถ้าเช็คค่าเเล้วไม่ได้เท่ากับ null หมายถึง มีการเชื่อมต่อ
            //กับฐานข้อมูลอยู่เเล้ว และให้เเสดงข้อความขึ้นมาว่า Already Connected!
            if (con == null)
                {
                        try
                    {
                         con = new OdbcConnection(conString);
                          con.Open();
                }
                 catch (Exception ex)
                    {
                    MessageBox.Show(ex.Message);
                    con.Close();
                    con = null;
                    return;
                    }
                 }
                 else
                 {
                    MessageBox.Show("Already connected!");
                    return;
                 }
                 MessageBox.Show("Connected!"); 
                        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //ปุ่ม Populate
            //Event ของปุ่มนี้ คือ
            //ใช้ DataSet เเละ DataAdapter เพื่อสั่งรัน SQL COMMAND โดยการเลือกทุกๆอย่างใน PersonTable
            //มาเเสดงผลที่ DataGrid
            dSet = new DataSet();
            dAdapter = new OdbcDataAdapter("select * from PersonTable",con);
            dAdapter.Fill(dSet, "PersonTable");
            this.dataGridView1.DataSource = dSet.Tables["PersonTable"];

            //ดึงชื่อ สกุลมาเเสดงที่ Textbox ที่สร้างขึ้น
            txtName.DataBindings.Add("Text", dSet.Tables["PersonTable"],"FirstName");
            txtSurname.DataBindings.Add("Text", dSet.Tables["PersonTable"],"LastName");
            currManager = (CurrencyManager)this.BindingContext[dSet.Tables["PersonTable"]]; 

        }

        private void txtPrevious_Click(object sender, EventArgs e)
        {
            //ปุ่ม Previous 
            //สำหรับเเสดงผล ชื่อ สกุล ก่อนหน้า
            currManager.Position -= 1; 
        }

        private void btNext_Click(object sender, EventArgs e)
        {
            //ปุ่ม Next
            //สำหรับแสดงชื่อ สกุลลำดับต่อไป
            currManager.Position += 1; 
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            try
 {
                //ปุ่ม Add
                //สำหรับ Add Record ชื่อ สกุล ใหม่ลงไปใน PersonalTable
                //เเละจะมี MessageBox ขึ้นมาเเสดงด้วย ว่าได้หรือไม่ได้
                 string strInsert;
                 strInsert = "insert into PersonTable(FirstName, LastName,Title, City, Country)" + " values('" + this.txtName.Text + "', '"
                 + this.txtSurname.Text + "', '"
                 + "none" + "', '"
                 + "none" + "', '"
                 + "none" + "')";
                 if (this.txtName.Text != "" && this.txtSurname.Text != "")
                 {
                 OdbcCommand dbCommand1 = new OdbcCommand(strInsert,
                this.con);
                 dbCommand1.CommandType = CommandType.Text;
                 dbCommand1.ExecuteNonQuery();
                 }
                 else
                 {
                     MessageBox.Show("Error...", "WARNING",
                     MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 }
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show("Error ", " WARNING ", MessageBoxButtons.OK,
                      MessageBoxIcon.Information);
                 } 
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
                 //ปุ่ม Delete
                //สำหรับสั่งลบ Record ที่เราเลือก
            try
 {
             string strDel = "DELETE FROM PersonTable WHERE FirstName = '" + this.txtName.Text + "' "; 
              if (this.txtName.Text != "" && this.txtSurname.Text != "") 
                 {
                     OdbcCommand dbCommand1 = new OdbcCommand(strDel,this.con);
                 dbCommand1.CommandType = CommandType.Text;
                 dbCommand1.ExecuteNonQuery();
                 }
                 else
                 {
                 MessageBox.Show("Error...", "WARNING",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 }
                 }
                 catch (Exception ex)
                 {
                 MessageBox.Show("Error ", " WARNING ", MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
                 } 
        }
                    }
                }
