using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Windows.Forms;
// The following to two namespace contains 
// the functions for manipulating the 
// Excel file  
using OfficeOpenXml;
using System.Diagnostics;
using System.Drawing;

namespace Student_Distribution
{
    public partial class Form1 : Form
    {
        ExcelPackage pck;
        ExcelWorksheet worksheet;
        DataTable excelDataTable;
        FileStream objFileStrm;
        int[] has_students = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ,0,0};
        NumericUpDown[] nuds;
        readonly String[] rooms;
        private OpenFileDialog ofd;
        private FolderBrowserDialog fbd;
        String sourceFile = "test_file";
        String AppDirectory = "";
        String DestinationFile = "";
        String templateFile = "";
        String testFileDestination = "";
        private bool serialised = true, randomised = false, balanced = false;

        public Form1()
        {
            InitializeComponent();
            // Removing image margins (space for icons on left) from menubar items:
            foreach (ToolStripMenuItem menuItem in menuStrip1.Items)
                ((ToolStripDropDownMenu)menuItem.DropDown).ShowImageMargin = false;

            //insilize vars...
            rooms = new string[] { label3.Text, label3.Text, label4.Text, label4.Text, label5.Text, label6.Text, label7.Text, label8.Text, label9.Text, label10.Text, label11.Text, label12.Text };
            // file name with .xlsx extension  
            nuds = new NumericUpDown[] { numericUpDown1, numericUpDown2, numericUpDown3, numericUpDown4, numericUpDown5,
                numericUpDown6, numericUpDown7, numericUpDown8, numericUpDown9, numericUpDown10 };
            //if (File.Exists(p_strPath))
            ofd = new OpenFileDialog();
            fbd = new FolderBrowserDialog();
            ofd.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            //  File.Delete(p_strPath);
            AppDirectory = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            DestinationFile = testFileDestination = AppDirectory + "\\Results\\توزيع_" + Path.GetFileNameWithoutExtension(sourceFile)+".xlsx";
            templateFile = AppDirectory + "\\Template\\template.xlsx";
            //for (int i=0;i<nuds.Length;i++)

            label1.Text += " ...";
            label2.Text += " ...";

        }
        
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {

                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        //moddarrag first paper...
        public int getModarrag_p1(int num)
        {
            if (num > 30)
                return 30;
            else
                return num;
        }

        public int left_std(NumericUpDown nud) {
            return excelDataTable.Rows.Count - sum(nud);
        }
        public int sum(NumericUpDown num)
        {
            int sum = 0;
            
            return sum;
        }

        //moddarrag second paper...
        public int getModarrag_p2(int num)
        {
            if (num > 30)
                return num - 30;
            else
                return 0;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (excelDataTable != null)
            {
                testFileDestination = AppDirectory + "\\Results\\" + "توزيع_" + textBox12.Text 
                    + "_" + getYearCod(textBox11.Text) + "_" + getDivCod(textBox13.Text) + ".xlsx";
                SaveExcelFile(testFileDestination);
                Process.Start(testFileDestination);
            }
            
        }

        private void copyFromDtsArray(ExcelWorksheet ws)
        {
            int mod_1_1 = getModarrag_p1((int)numericUpDown1.Value);
            int mod_1_2 = getModarrag_p2((int)numericUpDown1.Value);
            int mod_2_1 = getModarrag_p1((int)numericUpDown2.Value);
            int mod_2_2 = getModarrag_p2((int)numericUpDown2.Value);
            int sumofval = 0, sumof_has_students = 0;
            int[] values = new int[] { mod_1_1, mod_1_2, mod_2_1, mod_2_2, (int)nuds[2].Value, (int)nuds[3].Value,
                (int)nuds[4].Value, (int)nuds[5].Value, (int)nuds[6].Value, (int)nuds[7].Value, (int)nuds[8].Value, (int)nuds[9].Value };
            DataTable[] dts = new DataTable[12];
            //ExcelPackage has a constructor that only requires a stream.

            for (int i = 0; i < 12; i++)
            {
                if (values[i] != 0)
                {
                    has_students[i] = 1;
                    dts[i] = excelDataTable.AsEnumerable().Skip(sumofval).Take(values[i]).CopyToDataTable();
                    //handling paper i;
                    ws.Cells["C" + (int)(sumof_has_students * 43 + 10)].LoadFromDataTable(dts[i], false);
                    ws.Cells["C" + (int)(sumof_has_students * 43 + 7)].Value = rooms[i];
                }
                else has_students[i] = 0;
                sumofval += values[i];
                sumof_has_students += has_students[i];
            }
            ws.Cells["D4"].Value = textBox11.Text;
            ws.Cells["F4"].Value = textBox13.Text;
            ws.Cells["C5"].Value = textBox12.Text;
        }

        private void SaveExcelFile(string destination)
        {
            try
            {
                pck = new OfficeOpenXml.ExcelPackage(new FileInfo(templateFile));

                copyFromDtsArray(pck.Workbook.Worksheets[1]);
                copyFromDtsArray(pck.Workbook.Worksheets[2]);

                // Create excel file on physical disk  
                objFileStrm = File.Create(destination);
                objFileStrm.Close();

                // Write content to excel file  
                File.WriteAllBytes(destination, pck.GetAsByteArray());

                //Close Excel package 
                pck.Dispose();
                Console.Read();
            }
            catch (Exception exception)
            {
                MessageBox.Show("من فضلك, تأكد من إغلاق الملف "+ "توزيع_" + textBox12.Text
                    + "_" + getYearCod(textBox11.Text) + "_" + getDivCod(textBox13.Text) + ".xlsx " + " وأعد المحاولة", "عذرا...",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private int total_stds()
        {
            if (excelDataTable != null)
                return excelDataTable.Rows.Count;
            else return 0;
        }

        private void SetMaximumNum(NumericUpDown[] nud,int k=100)
        {
            for (int i = 0; i < nud.Length; i++)             
                if (nud[i].Maximum > left_students(k)) nud[i].Maximum = left_students(k); 
        }

        private int distributed_stds(int k = 100)
        {
            int sum = 0;
            for (int i = 0; i < nuds.Length; i++)
            {
                if (i != k)
                {
                    sum += (int)nuds[i].Value;
                }
            }

            return sum;
        }

        private int left_students(int k = 100)
        {
            return total_stds() - distributed_stds(k);
        }

        private void NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[0].Value > left_students(0))
                nuds[0].Value = left_students(0);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[0]);
            ColorOfTheGrid();
        }

        private void InitializeValues()
        {
            foreach (NumericUpDown nud in nuds) nud.Value = 0;
            has_students = new int[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ,0,0};
        }

        private void NumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[1].Value > left_students(1))
                nuds[1].Value = left_students(1);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[1]);
            ColorOfTheGrid();
        }

        private void NumericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[2].Value > left_students(2))
                nuds[2].Value = left_students(2);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[2]);
            ColorOfTheGrid();
        }

        private void NumericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            if(nuds[3].Value > left_students(3))
                nuds[3].Value = left_students(3);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[3]);
            ColorOfTheGrid();
        }

        private void NumericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[4].Value > left_students(4))
                nuds[4].Value = left_students(4);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[4]);
            ColorOfTheGrid();
        }

        private void NumericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[5].Value > left_students(5))
                nuds[5].Value = left_students(5);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[5]);
            ColorOfTheGrid();
        }

        private void NumericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[6].Value > left_students(6))
                nuds[6].Value = left_students(6);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[6]);
            ColorOfTheGrid();
        }

        private void NumericUpDown8_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[7].Value > left_students(7))
                nuds[7].Value = left_students(7);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[7]);
            ColorOfTheGrid();
        }

        private void NumericUpDown9_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[8].Value > left_students(8))
                nuds[8].Value = left_students(8);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[8]);
            ColorOfTheGrid();
        }

        private void NumericUpDown10_ValueChanged(object sender, EventArgs e)
        {
            if (nuds[9].Value > left_students(9))
                nuds[9].Value = left_students(9);
            label2.Text = "الطلاب المتبقين : " + left_students();
            checkNullVals(nuds[9]);
            ColorOfTheGrid();
        }

        private string getYear(string abc)
        {
            if (abc.Contains("3") || abc.Contains("ثالث"))
                return "الثالثة";

            else if (abc.Contains("4") || abc.Contains("رابع"))
                return "الرابعة";
            else if (abc.Contains("1") || abc.Contains("أول") || abc.Contains("اول"))
                return "الأولى";
            else if (abc.Contains("2") || abc.Contains("ثاني"))
                return "الثانية";
            else return "";
        }

        private string getDivision(string path)
        {
            if (path.Contains("علم الحيا") || path.Contains("علم حيا"))
                return "علم الحياة";
            else if (path.Contains("قسم الكيما") || path.Contains("كيميا"))
                return "الكيمياء";
            else if (path.Contains("قسم الرياضيات") || path.Contains("رياضيات"))
                return "الرياضيات";
            else if (path.Contains("قسم الفيز") || path.Contains("فيزياء"))
                return "الفيزياء";
            else return "";
        }

        private void add_file()
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                InitializeValues();
                sourceFile = ofd.FileName;
                textBox13.Text = getDivision(Path.GetDirectoryName(sourceFile));
                excelDataTable = GetDataTableFromExcel(sourceFile, true);
                textBox11.Text= getYear(excelDataTable.Rows[1][excelDataTable.Columns.Count - 1].ToString()+ excelDataTable.Rows[1][excelDataTable.Columns.Count-2].ToString());
                excelDataTable.Columns.RemoveAt(0);
                for (int i = excelDataTable.Columns.Count-1; i > 1; i--) 
                    excelDataTable.Columns.RemoveAt(i);


                for (int i = excelDataTable.Rows.Count - 1; i >= 0; i--)
                {
                    if (("" + excelDataTable.Rows[i][0] == "") && ("" + excelDataTable.Rows[i][1] == "")) { excelDataTable.Rows.RemoveAt(i); }
                }
                //excelDataTable.Columns.RemoveAt(4);excelDataTable.Columns.RemoveAt(5);
                dataGridView1.DataSource = excelDataTable;
                dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                label1.Text = "عدد الطلاب الاجمالي : " + total_stds(); //excelDataTable.Rows.Count;
                label2.Text = "الطلاب المتبقين : " + left_students();
                String[] soursestrings = Path.GetFileNameWithoutExtension(sourceFile).Split(' ');
                int k = 1;
                string[] realnamestrings = new string[] { "", "" };
                for (int i = soursestrings.Length - 1; i >= 0; i--)
                {
                    if (k >= 0)
                    {
                        realnamestrings[k] = soursestrings[i];
                    }
                    k--;
                }
                if (realnamestrings[0].Contains("حمل")|| realnamestrings[0].Contains("مقرر")) textBox12.Text =  realnamestrings[1];
                else textBox12.Text = realnamestrings[0] + " " + realnamestrings[1];

            }
            ColorOfTheGrid();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                add_file();
            }catch(Exception exception)
            {
                MessageBox.Show("من فضلك, تأكد من أن الملف الذي تريد اضافته يحوي ورقة عمل واحدة", "عذرا يوجد خطأ",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void فتحToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //try
            //{
                add_file();
            //}
            //catch (Exception exception)
            //{
            //    MessageBox.Show("من فضلك, تأكد من أن الملف الذي تريد اضافته يحوي ورقة عمل واحدة", "عذرا يوجد خطأ",
             //       MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void checkNullVals(NumericUpDown nude)
        {
            if (nude.Text == "") { nude.Value = 0;nude.Text = nude.Value.ToString(); }
        }

        private void NumericUpDown1_Leave(object sender, EventArgs e)
        {
            if (nuds[0].Text == String.Empty)
            {
                nuds[0].Value = nuds[0].Minimum;
                nuds[0].Text = nuds[0].Value.ToString();
            }
        }

        private void NumericUpDown2_Leave(object sender, EventArgs e)
        {
            if (nuds[1].Text == String.Empty)
            {
                nuds[1].Value = nuds[1].Minimum;
                nuds[1].Text = nuds[1].Value.ToString();
            }
        }

        private void NumericUpDown3_Leave(object sender, EventArgs e)
        {
            if (nuds[2].Text == String.Empty)
            {
                nuds[2].Value = nuds[2].Minimum;
                nuds[2].Text = nuds[2].Value.ToString();
            }
        }

        private void NumericUpDown4_Leave(object sender, EventArgs e)
        {
            if (nuds[3].Text == String.Empty)
            {
                nuds[3].Value = nuds[3].Minimum;
                nuds[3].Text = nuds[3].Value.ToString();
            }
        }

        private void NumericUpDown5_Leave(object sender, EventArgs e)
        {
            if (nuds[4].Text == String.Empty)
            {
                nuds[4].Value = nuds[4].Minimum;
                nuds[4].Text = nuds[4].Value.ToString();
            }
        }

        private void NumericUpDown6_Leave(object sender, EventArgs e)
        {
            if (nuds[5].Text == String.Empty)
            {
                nuds[5].Value = nuds[5].Minimum;
                nuds[5].Text = nuds[5].Value.ToString();
            }
        }

        private void NumericUpDown7_Leave(object sender, EventArgs e)
        {
            if (nuds[6].Text == String.Empty)
            {
                nuds[6].Value = nuds[6].Minimum;
                nuds[6].Text = nuds[6].Value.ToString();
            }
        }

        private void NumericUpDown8_Leave(object sender, EventArgs e)
        {
            if (nuds[7].Text == String.Empty)
            {
                nuds[7].Value = nuds[7].Minimum;
                nuds[7].Text = nuds[7].Value.ToString();
            }
        }

        private void NumericUpDown9_Leave(object sender, EventArgs e)
        {
            if (nuds[8].Text == String.Empty)
            {
                nuds[8].Value = nuds[8].Minimum;
                nuds[8].Text = nuds[8].Value.ToString();
            }
        }

        private void NumericUpDown10_Leave(object sender, EventArgs e)
        {
            if (nuds[9].Text == String.Empty)
            {
                nuds[9].Value = nuds[9].Minimum;
                nuds[9].Text = nuds[9].Value.ToString();
            }
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nuds.Length; i++)
                nuds[i].Value = nuds[i].Minimum;
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            if (serialised) SerialDist();
            else if (randomised) RandomDistribution();
        }

        private void متسلسلToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if (متسلسلToolStripMenuItem.Checked)
            {
                عشوائيToolStripMenuItem.Checked = false;
                randomised = false;
                serialised = true;
            }
            else
            {
                عشوائيToolStripMenuItem.Checked = true;
                randomised = true;
                serialised = false;
            }
        }

        private void عشوائيToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if (عشوائيToolStripMenuItem.Checked)
            {
                متسلسلToolStripMenuItem.Checked = false;
                randomised = true;
                serialised = false;
            }else
            {
                متسلسلToolStripMenuItem.Checked = true;
                randomised = false;
                serialised = true;
            }
        }

        private void متوازنToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if (متوازنToolStripMenuItem.Checked) balanced = true; else balanced = false;
        }

        private void حفظباسمToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (excelDataTable != null)
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel xls Files (*.xls)|*.xls";
                saveFileDialog1.Title = "Save File";
                if(textBox13.Text == "")
                saveFileDialog1.FileName = "توزيع_" + Path.GetFileNameWithoutExtension(sourceFile) + ".xlsx";
                else saveFileDialog1.FileName = "توزيع_" + textBox12.Text+"_"+getYearCod(textBox11.Text)+"_"+getDivCod(textBox13.Text) + ".xlsx";
                testFileDestination = AppDirectory + "\\Results\\" + "توزيع_" + textBox12.Text + "_" + getYearCod(textBox11.Text) + "_" + getDivCod(textBox13.Text) + ".xlsx";
                saveFileDialog1.ShowDialog();
                if (saveFileDialog1.FileName != "")
                {
                    DestinationFile = saveFileDialog1.FileName;
                    SaveExcelFile(DestinationFile);
                }
            }
        }

        private string getYearCod(string name)
        {
            if (name.Contains("ثالث"))
                return "س3";
            else if (name.Contains("رابع"))
                return "س4";
            else if (name.Contains("ثاني"))
                return "س2";
            else if (name.Contains("اول") || name.Contains("أول"))
                return "س1";
            else return name;
        }
        private string getDivCod(string name)
        {
            if (name.Contains("فيزي"))
                return "فيزياء";
            else if (name.Contains("كيميا"))
                return "كيمياء";
            else if (name.Contains("رياضيات"))
                return "رياضيات";
            else if (name.Contains("حيا"))
                return "علم حياة";
            else return name;
        }

        private void SerialDist()
        {
            int n;
            if (balanced) n = 2; else n = 1;
            for (int i = 0; i < nuds.Length; i++)
                if (nuds[i].Value == 0)
                    if (nuds[i].Maximum/n > left_students(i))
                        nuds[i].Value = left_students(i);
                    else nuds[i].Value = nuds[i].Maximum/n;
        }

        private void إنهاءToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void تعديلالنموذجToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(templateFile);
        }

        private void RandomDistribution()
        {
            int n;
            if (balanced) n = 2; else n = 1;
            Random r = new Random();
            foreach (int i in Enumerable.Range(0, 10).OrderBy(x => r.Next()))
                if (nuds[i].Value == 0)
                    if (nuds[i].Maximum / n > left_students(i))
                        nuds[i].Value = left_students(i);
                    else nuds[i].Value = nuds[i].Maximum / n;
        }

        private void ColorOfTheGrid()
        {
            Color[] colorswitch = new Color[] { Color.White, Color.AliceBlue};
            int sumofnuds = 0;
            int counter = 0;
            for (int j = 0; j < nuds.Length; j++)
            {
                for (int i = 0; i < (int)nuds[j].Value; i++)
                {
                        dataGridView1.Rows[i+sumofnuds].DefaultCellStyle.BackColor = colorswitch[counter % 2];

                }
                sumofnuds += (int)nuds[j].Value;
                if(nuds[j].Value!=0)
                counter++;
            }
            if(left_students() != 0)
            {
                for(int i = dataGridView1.Rows.Count-2;i> dataGridView1.Rows.Count -2 - left_students(); i--)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Bisque;
                }
            }
        }

    }
}
