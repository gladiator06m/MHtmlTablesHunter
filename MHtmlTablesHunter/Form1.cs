using Itage.MimeHtml2Html;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MHtmlTablesHunter
{
    public partial class Form1 : Form
    {
        private readonly string ConvertedFileName = "page.html";

        private readonly List<string> ColumnsToSeparate = new List<string>() { "life limit", "interval" };

        private readonly List<string> ExtraColumnsToAdd = new List<string>() { "Calendar", "Flight Hours", "Landing" };
        public Form1()
        {
            try
            {
                InitializeComponent();
                this.Text = $"MhtmlTablesHunter v{Application.ProductVersion}";
                pictureBox1.Image = Image.FromFile("MRXSystems.png");
            }
            catch (Exception)
            {

                MessageBox.Show("Error loading picture");
            }

        }

        //browse for specific file type , in this case its .mhtml
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Please choose the MHTML file";
                openFileDialog.Filter = "MHTML files (*.mhtml)|*.mhtml;";  //the file type specified here
                openFileDialog.RestoreDirectory = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {

                    textBoxSourceFile.Text = openFileDialog.FileName;
                    checkAndExtractTable(sender, e);
                }
            }

        }
        private async void checkAndExtractTable(object sender, EventArgs e)
        {
            string sourcePath = textBoxSourceFile.Text;
            if (!string.IsNullOrEmpty(sourcePath)) //check if the input file path is not empty
            {
                if (File.Exists(sourcePath)) //check if the input file path is exists
                {
                    Button btn = sender as Button;
                    btn.Enabled = false;
                    pictureBox1.Visible = true;
                    await Task.Run(async () => await ExtractTable(sourcePath)); //run the extraction process in a thread for the UI to be more responsive
                    pictureBox1.Visible = false;
                    btn.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Source file doesn't exist.");
                }
            }
            else
            {
                MessageBox.Show("Please select the source file.");
            }
        }
        public List<dynamic> datagridList;
        BindingList<dynamic> datagridBindingList;
        DataTable dataTableDynamic;

        public async Task<string> ExtractTable(string sourcePath)
        {
            datagridBindingList = new BindingList<dynamic>();
            datagridList = new List<dynamic>();
            try
            {
                var doc = new HtmlAgilityPack.HtmlDocument(); // HtmlAgilityPack it is a .NET library allow you to parse html file on a winform project, console etc... it is 

                var converter = new MimeConverter();   //converter used to convert mhtml file to html

                if (File.Exists(ConvertedFileName))  //check if previously converted file is exist
                {
                    File.Delete(ConvertedFileName); //delete the file
                }
                using (FileStream sourceStream = File.OpenRead(sourcePath))
                {
                    using (FileStream destinationStream = File.Open("page.html", FileMode.Create))
                    {
                        await converter.Convert(sourceStream, destinationStream);  //convert the file to html, it will be stored in the application folder
                    }
                }

                doc.Load(ConvertedFileName);  //load the html
                var tables = doc.DocumentNode.SelectNodes("//table"); //get all the tables
                HtmlAgilityPack.HtmlNode table = null;


                if (tables.Count > 0)
                {
                    table = tables[tables.Count - 1]; //take the last table
                }


                if (table != null) //if the table exists
                {
                    dataGridView1.Invoke((Action)delegate //we use delegate because we accessing the datagridview from a different thread
                    {
                        if (dataTableDynamic != null)
                        {
                            dataTableDynamic.Rows.Clear();
                            dataTableDynamic.Columns.Clear();
                        }
                        //this.dataGridView1.Rows.Clear();
                        //this.dataGridView1.Columns.Clear();
                    });

                    var rows = table.SelectNodes(".//tr"); //get all the rows

                    var nodes = rows[0].SelectNodes("th|td"); //get the header row values, first item will be the header row always
                    string LifeLimitColumnName = ColumnsToSeparate.Where(c => nodes.Any(n => n.InnerText.ToLower().Contains(c))).FirstOrDefault();
                    if (string.IsNullOrWhiteSpace(LifeLimitColumnName))
                    {
                        LifeLimitColumnName = "Someunknowncolumn";
                    }
                    List<string> headers = new List<string>();
                    List<string> headerProperties = new List<string>();

                    for (int i = 0; i < nodes.Count; i++) //th
                    {
                        headers.Add(nodes[i].InnerText);
                        if (!nodes[i].InnerText.ToLower().Contains(LifeLimitColumnName))
                        {
                            dataGridView1.Invoke((Action)delegate
                            {
                                //dataGridView1.Columns.Add("", nodes[i].InnerText); //add header to the datagridview
                                headerProperties.Add(Regex.Replace(nodes[i].InnerText, @"\s+", ""));
                            });
                        }
                    }

                    int indexOfLifeLimitColumn = headers.FindIndex(h => h.ToLower().Contains(LifeLimitColumnName));
                    if (indexOfLifeLimitColumn > -1)
                    {
                        foreach (var eh in ExtraColumnsToAdd)
                        {
                            dataGridView1.Invoke((Action)delegate
                            {
                                //dataGridView1.Columns.Add("", eh); //add extra header to the datagridview with the variable LifeLimitColumnName
                                headerProperties.Add(Regex.Replace(eh, @"\s+", ""));
                            });
                        }
                    }

                    for (int i = 1; i < rows.Count; i++) ///loop through rest of the rows
                    {
                        var row = rows[i];
                        var nodes2 = row.SelectNodes("th|td"); //get all columns in the current row
                        List<string> values = new List<string>(); //list to store row values
                        for (int x = 0; x < nodes2.Count; x++)
                        {
                            //rowes.Cells[x].Value = nodes2[x].InnerText;
                            string cellText = nodes2[x].InnerText.Replace("&nbsp;", " ");

                            values.Add(cellText); //add the cell value in the list value
                        }


                        if (indexOfLifeLimitColumn > -1)
                        {
                            values.RemoveAt(indexOfLifeLimitColumn);
                            string lifeLimitValue = nodes2[indexOfLifeLimitColumn].InnerText.Replace("&nbsp;", " ");
                            string[] splittedValues = lifeLimitValue.Split(' ');
                            for (int y = 0; y < ExtraColumnsToAdd.Count; y++)
                            {
                                if (ExtraColumnsToAdd[y] == "Calendar")
                                {
                                    string valueToAdd = string.Empty;
                                    string[] times = new string[] { "Years", "Year", "Months", "Month", "Day", "Days" };
                                    if (splittedValues.Any(s => times.Any(t => t == s)))
                                    {
                                        var timeFound = times.Where(t => splittedValues.Any(s => s == t)).FirstOrDefault();
                                        int index = splittedValues.ToList().FindIndex(s => s.Equals(timeFound));
                                        valueToAdd = $"{splittedValues[index - 1]} {timeFound}";
                                    }
                                    values.Add(valueToAdd);
                                }
                                else if (ExtraColumnsToAdd[y] == "Flight Hours")
                                {
                                    string valueToAdd = string.Empty;
                                    if (splittedValues.Any(s => s == "FH"))
                                    {
                                        int index = splittedValues.ToList().FindIndex(s => s.Equals("FH"));
                                        valueToAdd = $"{splittedValues[index - 1]} FH";
                                    }
                                    values.Add(valueToAdd);
                                }
                                else
                                {
                                    string valueToAdd = string.Empty;
                                    if (splittedValues.Any(s => s == "LDG"))
                                    {
                                        int index = splittedValues.ToList().FindIndex(s => s.Equals("LDG"));
                                        valueToAdd = $"{splittedValues[index - 1]} LDG";
                                    }
                                    values.Add(valueToAdd);
                                }
                            }
                        }

                        var rowObject = new ExpandoObject() as IDictionary<string, Object>;
                        for (int x = 0; x < values.Count; x++)
                        {
                            rowObject.Add(headerProperties[x], values[x]);
                        }
                        datagridList.Add(rowObject);

                    }

                    datagridBindingList = new BindingList<dynamic>(datagridList);
                    dataGridView1.Invoke((Action)delegate
                    {
                        dataTableDynamic = ToDataTable(datagridBindingList);

                        //////////////////////////////////////////Delete empty column
                        int[] rowDataCount = new int[dataTableDynamic.Columns.Count];
                        Array.Clear(rowDataCount, 0, rowDataCount.Length);

                        for (int row_i = 0; row_i < this.dataTableDynamic.Rows.Count; row_i++)
                        {
                            for (int col_i = 0; col_i < this.dataTableDynamic.Columns.Count; col_i++)
                            {
                                var cell = this.dataTableDynamic.Rows[row_i][col_i];
                                string cellText = cell.ToString();
                                if (!String.IsNullOrWhiteSpace(cellText))
                                {
                                    rowDataCount[col_i] += 1;
                                }
                            }
                        }

                        int removedCount = 0;
                        for (int index = 0; index < rowDataCount.Length; index++)
                        {
                            if (rowDataCount[index] == 0)
                            {
                                this.dataTableDynamic.Columns.RemoveAt(index - removedCount);
                                removedCount++;
                            }
                        }


                        this.dataGridView1.DataSource = dataTableDynamic;
                    });
                    //////////////////////////////////////////////
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return string.Empty;
        }

        public static DataTable ToDataTable(IEnumerable<dynamic> items)
        {
            var data = items.ToArray();
            if (data.Count() == 0) return null;

            var dt = new DataTable { TableName = "MyTableName" };
            foreach (var key in ((IDictionary<string, object>)data[0]).Keys)
            {
                dt.Columns.Add(key);
            }
            foreach (var d in data)
            {
                dt.Rows.Add(((IDictionary<string, object>)d).Values.ToArray());
            }

            //Export to XML
            DataSet ds = new DataSet();
            dt.WriteXml(File.OpenWrite(@"d:\GridPilatus.xml"));
            return dt;
        }

    }

}
