using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DbToExcel
{
    public partial class Form1 : Form
    {

        DataTable dataTable1 = new DataTable();
        List<string> columns = new List<string>();
        List<string> checkedItems = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RemoveCheckBox();
            columns.Clear();
            checkedItems.Clear();
            string table = textBox1.Text;
            string kosul = textBox2.Text;
            var liste = GetTableColumnName(table);
            var Data = GetData(table, kosul);
            dataTable1 = Data;
            dataGridView1.DataSource = Data;
            
            AddCheckBox(liste);
        }
        private void AddCheckBox(List<string> liste)
        {
            int point1 = 0;
            int point2 = 80;
            int i = 1;
            foreach (var item in liste)
            {
                CheckBox checkBox = new CheckBox();
                checkBox.Text = item.Trim().ToString();
                checkBox.Name = item.ToLower().ToString();
                checkBox.CheckedChanged += checkBox_CheckedChanged;
                point1 += point1 == 0 ? 80 : 120;
                checkBox.Location = new Point(point1, point2);
                this.Controls.Add(checkBox);

                if (i % 4 == 0)
                {
                    point1 = 0;
                    point2 += 20; //alt alta sıralıyor.
                }
                i++;

            }
        }
        private void RemoveCheckBox()
        {
            foreach (System.Windows.Forms.Control item in this.Controls.OfType<CheckBox>().ToList())
            {
                this.Controls.Remove(item);
            }
        }
        private List<string> GetTableColumnName(string Table)
        {
            
            string cmdStr = @"SELECT COLUMN_NAME
                            FROM INFORMATION_SCHEMA.COLUMNS
                            WHERE TABLE_NAME = @TABLE";
            using (var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DbContext"].ToString()))
            {
                using (var command = connection.CreateCommand())
                {
                    connection.Open();
                    command.CommandText = cmdStr;
                    command.Parameters.Add("@TABLE", Table);
                    using (var reader = command.ExecuteReader())
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                columns.Add(reader["COLUMN_NAME"].ToString());
                            }
                        } while (reader.NextResult());
                    }
                    return columns;
                }
            }
        }

        private DataTable GetData(string Table,string Kosul)
        {
            string cmdStr = "";
            DataTable dt = new DataTable();
            if (string.IsNullOrEmpty(Kosul))
            {
                cmdStr= string.Format("SELECT * FROM {0}", Table);

            }
            else
            {
                cmdStr = string.Format("SELECT * FROM {0} WHERE {1}", Table,Kosul);
            }
            using (var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DbContext"].ToString()))
            {
                using (var command = connection.CreateCommand())
                {
                    connection.Open();
                    command.CommandText = cmdStr;
                    //command.Parameters.Add("@TABLE", Table);
                    var reader = command.ExecuteReader();
                    dt.Load(reader);
       
                    return dt;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            foreach(var chckItem in checkedItems)
            {
                columns.Remove(chckItem);
            }
            foreach(var item in columns)
            {
                dataTable1.Columns.Remove(item);
            }
            
            ExportExcel(dataTable1);
        }

        private void ExportExcel(DataTable dataTable)
        {
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dataTable, "Sheet1");
            wb.SaveAs("ExcelExport.xlsx");
        }

        private void checkBox_CheckedChanged(object sender, System.EventArgs e)
        {
            CheckBox x = (CheckBox)sender;
            if(x.Checked == true)
            {
                checkedItems.Add(x.Text);
            }
            if(x.Checked == false)
            {
                checkedItems.Remove(x.Text);
            }

        }
    }
}
