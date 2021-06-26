using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace Priroda
{
    public partial class Login : Form
    {
        OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataReader dr;
        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() + @"\User.accdb");
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            string str = "SELECT * FROM UserDB where User='" + loginBox.Text + "' AND Pass='" + passwordBox.Text + "'";
            cmd.CommandText = str;
            dr = cmd.ExecuteReader();

            if (dr.Read())
            {
                this.Hide();
                main f = new main();
                f.Show();
                if (loginBox.Text == "dir")
                {
                    f.lv = 3;
                }
                else if (loginBox.Text == "zam")
                {
                    f.lv = 2;
                }
                else if (loginBox.Text == "kons")
                {
                    f.lv = 1;
                }
            }
            else
            {
                MessageBox.Show("Неправильный логин или пароль");
            }

            con.Close();
        }
    }
}
