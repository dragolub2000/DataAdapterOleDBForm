using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

/*
         try
           {
               OleDbConnection conn = new OleDbConnection();
               conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\testdatabase.xls;Extended Properties=\"Excel 8.0\"";
               String strSQL = "Select * from Radnik2";
               // Create an OleDbDataAdapter object  
               OleDbDataAdapter adapter = new OleDbDataAdapter();
               adapter.SelectCommand = new OleDbCommand(strSQL, conn);
               // Create Data Set object  
               DataSet ds = new DataSet("Radnici");
               // Call DataAdapter's Fill method to fill data from the  
               // DataAdapter to the DataSet  
               adapter.Fill(ds);

               // Bind dataset to a DataGrid control  
               dataGridView1.DataSource = ds;//.DefaultViewManager;
               dataGridView1.Show();

               conn.Close();
           }
           catch (Exception ex)
           {
               //Console.WriteLine("Postoji problem: " + e.Message);
               MessageBox.Show(ex.Message);
           }
        */

namespace DataAdapterOleDBForm
{
   public partial class Form1 : Form
   {
       public Form1()
       {
           InitializeComponent();
       }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\testdatabase.xls;Extended Properties=\"Excel 8.0\"";
                String strSQL = "Select * from Radnik2";
                // Kreiranje  OleDbDataAdapter objekta  
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                adapter.SelectCommand = new OleDbCommand(strSQL, conn);
                // Kreiranje Data Set objekta!
                DataSet ds = new DataSet();
                adapter.Fill(ds,"Radnici");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "Radnici";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Postoji problem: " + ex.Message);
            }
        }
    }
}
