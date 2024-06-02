using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;
using System.Data;

namespace Excel_Conversion_UI
{
    public partial class Form1 : Form
    {
        List<ExtractedData> extractedData = new List<ExtractedData>();
        List<FilteredData> filteredData = new List<FilteredData>();
        int uni_degree;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "xlsx Files (*.xlsx)|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Clear();
                textBox1.Text = openFileDialog1.FileName;
                textBox3.AppendText("엑셀 파일이 업로드 되었습니다" + Environment.NewLine);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = folderBrowserDialog1.SelectedPath;
                textBox3.AppendText("파일 목적지가 설정 되었습니다" + Environment.NewLine);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox3.AppendText("엑셀 파일을 선택하세요." + Environment.NewLine);
                return;
            }

            if (textBox2.Text == "")
            {
                textBox3.AppendText("엑셀 파일 목적지를 설정하세요." + Environment.NewLine);
                return;
            }

            if(!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked)
            {
                textBox3.AppendText("차수를 선택해주세요." + Environment.NewLine);
                return;
            }

            string file = textBox1.Text;
            try
            {
                using (var workbook = new XLWorkbook(file))
                {
                    
                    var rows = workbook.Worksheet(1).RangeUsed().RowsUsed().Skip(2);
                    foreach (var row in rows)
                    {
                        if (Int32.Parse(row.Cell(3).Value.ToString()) == uni_degree)
                        {
                            if (row.Cell(1).Value.ToString() != "") //check empty row
                            {
                                //removing day in date format 
                                string date = row.Cell(8).Value.ToString();
                                string removeDay = date.Substring(0, date.IndexOf(" "));
                              
                                extractedData.Add(new ExtractedData(
                                    row.Cell(5).Value.ToString(),
                                    row.Cell(6).Value.ToString(),
                                    removeDay,
                                    row.Cell(9).Value.ToString()
                                ));
                            }
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
            }

            foreach (var row in extractedData)
            {
                int idx1 = row.combinedDigits.IndexOf("-");
                int idx2 = row.combinedDigits.LastIndexOf("-");
                string one_digit = row.combinedDigits.Substring(idx1 + 1, 1);
                string three_digit = row.combinedDigits.Substring(idx2 + 1, 3);


                filteredData.Add(new FilteredData(one_digit, three_digit, row.product, row.date, row.name));

            }
           ToExcelFile(ToDataTable(filteredData), textBox2.Text, String.Format("완료파일-{0}{1}{2}-{3}{4}{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second));
        }


        public DataTable ToDataTable(List<FilteredData> list)
        {
            DataTable convertedTable = null;

            DataTable dt = new DataTable();
            dt.Columns.Add("aa");
            dt.Columns.Add("bb");
            dt.Columns.Add("cc");
            dt.Columns.Add("dd");
            dt.Columns.Add("ee");

            foreach (FilteredData row in list)
            {
                DataRow dr = dt.NewRow();
                dr["aa"] = row.one_digit;
                dr["bb"] = row.three_digit;
                dr["cc"] = row.product;
                dr["dd"] = row.date;
                dr["ee"] = row.name;

                dt.Rows.Add(dr);
            }

            dt.AcceptChanges();

            convertedTable = dt;

            return convertedTable;
        }

        public void ToExcelFile(DataTable dt, string path, string filename)
        {
            if (Directory.Exists(path) && path != string.Empty)
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(dt, "Sheet1");
                if (filename.Contains("."))
                {
                    int IndexOfLastFullStop = filename.LastIndexOf('.');
                    filename = filename.Substring(0, IndexOfLastFullStop) + ".xlsx";
                }
                else
                {
                    filename = filename + ".xlsx";
                }

                try
                {
                    wb.SaveAs(path + '\\' + filename);
                    textBox3.AppendText("파일생성이 완료되었습니다" + Environment.NewLine);
                }catch(Exception ex)
                {
                    textBox3.AppendText("파일생성 도중 문제가 생겼습니다. : " + ex.Message + Environment.NewLine);
                }
                
            }
            else
            {
                throw new System.ArgumentException("Error: Output directory does not exist!");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                
               bool isOkay = CheckedValidation(1, textBox1.Text);

             
                if (isOkay == true)
                {
                    textBox3.AppendText("1차를 선택하였습니다" + Environment.NewLine);
                    button3.Enabled = true;
                    uni_degree = 1;
                }
                else
                {
                    textBox3.AppendText("1차가 존재하지 않습니다" + Environment.NewLine);
                    button3.Enabled = false;
                    checkBox1.Checked = false;
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                bool isOkay = CheckedValidation(2, textBox1.Text);

                if (isOkay == true)
                {
                    textBox3.AppendText("2차를 선택하였습니다" + Environment.NewLine);
                    button3.Enabled = true;
                    uni_degree = 2;
                }
                else
                {
                    textBox3.AppendText("2차가 존재하지 않습니다" + Environment.NewLine);
                    button3.Enabled = false;
                    checkBox2.Checked = false;
                }
            }
        }
        

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                bool isOkay = CheckedValidation(3, textBox1.Text);
                if (isOkay == true)
                {
                    textBox3.AppendText("3차를 선택하였습니다" + Environment.NewLine);
                    button3.Enabled = true;
                    uni_degree = 3;
                }
                else
                {
                    textBox3.AppendText("3차가 존재하지 않습니다" + Environment.NewLine);
                    button3.Enabled = false;
                    checkBox3.Checked = false;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button3.Enabled = false;
        }

        public Boolean CheckedValidation(int _degree, string _file)
        {
            int degree = _degree;
            string file = _file;
            bool isOkay = false;
            try
            {
                using (var workbook = new XLWorkbook(file))
                {
                    var rows = workbook.Worksheet(1).RangeUsed().RowsUsed().Skip(2);
                    foreach (var row in rows)
                    {
                        if (Int32.Parse(row.Cell(3).Value.ToString()) == degree)
                        {
                            isOkay = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            return isOkay;
        }
    }

}
