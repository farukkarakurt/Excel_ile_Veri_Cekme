using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Excel_test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Excel bağlantı dizesi doğru şekilde yapılandırıldı
        OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\faruk\OneDrive\Masaüstü\1.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;'");

        // Excel verilerini listeleme fonksiyonu
        void listele()
        {
            try
            {
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Table 1$]", conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            conn.Open();
            OleDbCommand cmd = new OleDbCommand("insert into [Table 1$] (DERSLER) values (P2)", conn);
            //cmd.Parameters.AddWithValue("P1", textBox5.Text);//ID
            cmd.Parameters.AddWithValue("P2", textBox4.Text);
            //cmd.Parameters.AddWithValue("P3", textBox3.Text);
            //  cmd.Parameters.AddWithValue("P4", textBox2.Text);
            //cmd.Parameters.AddWithValue("P5", textBox1.Text);
            cmd.ExecuteNonQuery();

            conn.Close();
            MessageBox.Show("Yeni Sınav Tarihi Eklendi");
            listele();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
