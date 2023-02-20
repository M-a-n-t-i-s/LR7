using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
//using xlSheet=Microsoft.Office.Interop.Excel.Worksheet;
//using xlSheetRange= Microsoft.Office.Interop.Excel.Range;
//using xlApp = Microsoft.Office.Interop.Excel.Application;
///http://csharpcoding.org/vygruzka-dannyx-iz-sql-v-excel/
namespace CRUD_приложение_для_таблицы_Страны
{
    

    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Worksheet xlSheet;
        Microsoft.Office.Interop.Excel.Range xlSheetRange;
        string constr = "Data Source=dbs; Initial Catalog=Учебная; User ID=test; Password=test";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            string selstr = "SELECT ";
            bool f = false;
            if (checkBox1.Checked)
            {
               selstr += "Название";
               f = true;
            }
            if (checkBox2.Checked)
            {
                if (f)
                {
                    selstr += ",";
                }
                selstr += "Столица";
                f = true;
            }
            if (checkBox3.Checked)
            {
                if (f)
                {
                    selstr += ",";
                }
                selstr += "Континент";
                f = true;
            }
            if (checkBox4.Checked)
            {
                if (f)
                {
                    selstr += ",";
                }
                selstr += "Население";
                f = true;
            }
            if (checkBox5.Checked)
            {
                if (f)
                {
                    selstr += ",";
                }
                selstr += "Площадь";
                f = true;
            }
           
            selstr += " FROM Страны";
           
            using (SqlDataAdapter sda = new SqlDataAdapter(selstr, new SqlConnection(constr)))
            {
                DataTable dt = new DataTable();
                sda.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    {
                        string item = dr[0].ToString();
                        if (dt.Columns.Count > 1)
                        {
                            item = item + " " + dr[1].ToString();
                        }

                        if (dt.Columns.Count > 2)
                        {
                            item = item + " " + dr[2].ToString();
                        }

                        if (dt.Columns.Count > 3)
                        {
                            item = item + " " + dr[3].ToString();
                        }
                        if (dt.Columns.Count > 4)
                        {
                            item = item + " " + dr[4].ToString();
                        }
                        if (dt.Columns.Count > 5)
                        {
                            item = item + " " + dr[5].ToString();
                        }
                        listBox1.Items.Add(item);

                    }
                }
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void textBox8_Click(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                string updstr = "Data Source=dbs; Initial Catalog=Учебная; User ID=test; Password=test";
                //string updstr = "UPDATE Страны Set Страна = '" + textBox1.Text + "' WHERE Страна = '" + listBox1.SelectedItem +"'";
                SqlConnection scon = new SqlConnection(constr);
                scon.Open();
                SqlCommand updcmd = new SqlCommand(updstr, scon);
                updcmd.ExecuteNonQuery();
                scon.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                string updstr = "DELETE FROM Страны  WHERE Страна = '" + listBox1.SelectedItem + "'";
                SqlConnection scon = new SqlConnection(constr);
                scon.Open();
                SqlCommand updcmd = new SqlCommand(updstr, scon);
                updcmd.ExecuteNonQuery();
                scon.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string command = "INSERT INTO Страны SELECT '" + textBox2.Text + "','"
            + textBox3.Text + "','" + textBox5.Text + "',"  + Convert.ToInt32(textBox4.Text) + ","
            + Convert.ToDouble(textBox7.Text);
            SqlConnection scon = new SqlConnection(constr);
            scon.Open();
            SqlCommand updcmd = new SqlCommand(command, scon);
            updcmd.ExecuteNonQuery();
            scon.Close();
        }

        private DataTable GetData()
        {

            string connString = @"Data Source=dbs; Initial Catalog=Учебная; User ID=test; Password=test";

            SqlConnection con = new SqlConnection(connString);

            DataTable dt = new DataTable();
            try
            {
                string query = @"SELECT * from Страны";
                SqlCommand comm = new SqlCommand(query, con);

                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
            return dt;
        }

        private void button5_Click(object sender, EventArgs e)
        {
           

          xlApp = new Excel.Application();
 
    try
    {
        
        xlApp.Workbooks.Add(Type.Missing);
 
        
        xlApp.Interactive = false;
        xlApp.EnableEvents = false;
 
        
        xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
        
        xlSheet.Name = "Данные";
 
        
        DataTable dt = GetData();
 
        int collInd = 0;
        int rowInd = 0;
        string data = "";
 
        
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            data = dt.Columns[i].ColumnName.ToString();
            xlSheet.Cells[1, i + 1] = data;
 
            
            xlSheetRange = xlSheet.get_Range("A1:Z1", Type.Missing);
 
            
            xlSheetRange.WrapText = true;
            xlSheetRange.Font.Bold = true;
        }
 
      
        for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
        {
            for (collInd = 0; collInd < dt.Columns.Count; collInd++)
            {
                data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                xlSheet.Cells[rowInd + 2, collInd + 1] = data;
            }
        }
 
        
        xlSheetRange = xlSheet.UsedRange;
 
        
        xlSheetRange.Columns.AutoFit();
        xlSheetRange.Rows.AutoFit();
    }
    catch (Exception ex)
    {
        MessageBox.Show(ex.ToString());
    }
    finally
    {
        
        xlApp.Visible = true;
 
        xlApp.Interactive = true;
        xlApp.ScreenUpdating = true;
        xlApp.UserControl = true;
 
        
        releaseObject(xlSheetRange);
        releaseObject(xlSheet);
        releaseObject(xlApp);
    }
}

void releaseObject(object obj)
{
    try
    {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
    }
    catch (Exception ex)
    {
        obj = null;
        MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    finally
    {
        GC.Collect();
    }
        
        }

       
    }
}
