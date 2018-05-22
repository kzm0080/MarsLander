using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using System.IO;


namespace MarsLander
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Random random = new Random();
        List<Discrete> lDiscrete = new List<Discrete>();
        double CumX = 0.0;
        double CumY = 0.0;
        double totcount = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            txtXA.Text = "-100";
            txtXB.Text = "100";
            txtYA.Text = "-100";
            txtYB.Text = "100";
            txtFallProb.Text = "0.3";
            txtSimulations.Text = "1000";

            textBox3.Text = "1000";
            textBox4.Text = "0";
            textBox5.Text = "1000";
            textBox6.Text = "0";
            textBox2.Text = "0.3";
            textBox1.Text = "1000";

            panelUniform.Visible = false;
            panelNormal.Visible = false;

        }



        private void btnSimulate_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(txtXA.Text))
            {
                MessageBox.Show("Please enter data for a in X-Movement");
                return;
            }
            else if (string.IsNullOrEmpty(txtXB.Text))
            {
                MessageBox.Show("Please enter data for b in X-Movement");
                return;
            }
            else if (string.IsNullOrEmpty(txtYA.Text))
            {
                MessageBox.Show("Please enter data for a in Y-Movement");
                return;
            }
            else if (string.IsNullOrEmpty(txtYB.Text))
            {
                MessageBox.Show("Please enter data for b in Y-Movement");
                return;
            }
            else if (string.IsNullOrEmpty(txtFallProb.Text))
            {
                MessageBox.Show("Please enter data for FallProbability");
                return;
            }
            else if (string.IsNullOrEmpty(txtSimulations.Text))
            {
                MessageBox.Show("Please enter data for Simulations");
                return;
            }
            else if (Convert.ToDouble(txtFallProb.Text) < 0)
            {
                MessageBox.Show("Please enter correct data for FallProbability");
                return;
            }
            else if (Convert.ToInt32(txtSimulations.Text) < 0)
            {
                MessageBox.Show("Please enter correct data for Simulations [Ex : 1000");
                return;
            }
           


            pbUniform.Value = 10;
            dataGridView1.DataSource = null;
            lDiscrete = new List<Discrete>();

            CumX = 0.0;
            CumY = 0.0;
            string Angle = "", Distance = "";

            totcount = 0;

            int simulations = Convert.ToInt32(txtSimulations.Text);
            double XA = Convert.ToDouble(txtXA.Text);
            double XB = Convert.ToDouble(txtXB.Text);
            double YA = Convert.ToDouble(txtYA.Text);
            double YB = Convert.ToDouble(txtYB.Text);
            double FallProb = Convert.ToDouble(txtFallProb.Text);


            pbUniform.Value = 15;
            for (int i = 1; i <= simulations; i++)
            {
                Discrete discrete = new Discrete();
                discrete.Iteration = i;
                discrete.X = XA + ((XB - XA) * random.NextDouble());
                CumX = CumX + discrete.X;
                discrete.SumX = CumX;
                discrete.Y = YA + ((YB - YA) * random.NextDouble());
                CumY = CumY + discrete.Y;
                discrete.SumY = CumY;
                discrete.FallProbability = random.NextDouble();
                if (discrete.FallProbability < FallProb)
                {
                    discrete.Continue = "Yes";
                    double tan = CumY / CumX;
                    double radians = Math.Atan(tan);
                    double angle = radians * (180 / Math.PI);
                    if (CumX < 0 && CumY > 0)
                    {
                        if (angle < 0)
                            angle = 180 + angle;
                        else
                            angle = 90 + angle;
                    }
                    else if (CumX < 0 && CumY < 0)
                    {
                        if (angle < 0)
                            angle = 270 + angle;
                        else
                            angle = 180 + angle;
                    }
                    else if (CumX > 0 && CumY < 0)
                    {
                        if (angle < 0)
                            angle = 360 + angle;
                    }

                    angle = Convert.ToDouble(angle.ToString("0.000"));
                    discrete.Angle = angle.ToString();
                    totcount = totcount + 1;
                }
                else if (i == simulations)
                {
                    double tan = CumY / CumX;
                    double radians = Math.Atan(tan);
                    double angle = radians * (180 / Math.PI);
                    if (CumX < 0 && CumY > 0)
                    {
                        if (angle < 0)
                            angle = 180 + angle;
                        else
                            angle = 90 + angle;
                    }
                    else if (CumX < 0 && CumY < 0)
                    {
                        if (angle < 0)
                            angle = 270 + angle;
                        else
                            angle = 180 + angle;
                    }
                    else if (CumX > 0 && CumY < 0)
                    {
                        if (angle < 0)
                            angle = 360 + angle;                       
                    }

                    angle = Convert.ToDouble(angle.ToString("0.000"));
                    discrete.Angle = angle.ToString();
                }
                else
                {
                    discrete.Continue = "";
                    discrete.Angle = "";
                }

                if (i == simulations / 2)
                    pbUniform.Value = 55;

                discrete.X = Convert.ToDouble(discrete.X.ToString("0.000"));
                discrete.Y = Convert.ToDouble(discrete.Y.ToString("0.000"));
                discrete.SumX = Convert.ToDouble(discrete.SumX.ToString("0.000"));
                discrete.SumY = Convert.ToDouble(discrete.SumY.ToString("0.000"));
                discrete.FallProbability = Convert.ToDouble(discrete.FallProbability.ToString("0.000"));
                double distanc = (discrete.SumX * discrete.SumX) + (discrete.SumY * discrete.SumY);
                discrete.Distance = Convert.ToDouble(Math.Sqrt(distanc).ToString("0.000"));
                Angle = discrete.Angle;
                Distance = discrete.Distance.ToString();
                lDiscrete.Add(discrete);

            }
            pbUniform.Value = 90;
            dataGridView1.DataSource = lDiscrete;
            lblCountResult.Text = totcount.ToString();
            lblFallResult.Text = (totcount / simulations).ToString();
            lblCountResult.ForeColor = System.Drawing.Color.Red;
            lblFallResult.ForeColor = System.Drawing.Color.Red;
            txtXCordinate.Text = CumX.ToString();
            txtYCordinate.Text = CumY.ToString();
            txtAngle.Text = Angle;
            txtDistance.Text = Distance;
            pbUniform.Value = 100;


        }

        private string MyDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        List<ExcelNormal> lexcelnormal = new List<ExcelNormal>();
        List<Exceldata> lExceldata = new List<Exceldata>();
        private void button1_Click(object sender, EventArgs e)
        {
            string pathN = MyDirectory();
            string filenamelocation = System.IO.Path.Combine(pathN, "MarsLander.txt");
            //MessageBox.Show(filenamelocation);

             if (path.ToString() == "")
            {

                using (System.IO.StreamReader sr = new System.IO.StreamReader(filenamelocation))
                {
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        path = line;
                        //MessageBox.Show(path);
                    }
                }
            }
            if (path.ToString() == "")
            {
                MessageBox.Show("Select browse and create a path");
                return;
            }
            else if(Directory.Exists(path.ToString()))
            {
                MessageBox.Show("Select browse and create a path");
                return;
            }

            lexcelnormal = new List<ExcelNormal>();
            lExceldata = new List<Exceldata>();
            CumX = 0.0;
            CumY = 0.0;

            totcount = 0;

            int simulations = Convert.ToInt32(textBox1.Text);
            double XMean = Convert.ToDouble(textBox6.Text);
            double XVariance = Convert.ToDouble(textBox5.Text);
            double XStanderdDev = Convert.ToDouble(textBox7.Text);
            double YMean = Convert.ToDouble(textBox4.Text);
            double YVariance = Convert.ToDouble(textBox3.Text);
            double YStanderdDev = Convert.ToDouble(textBox8.Text);
            double FallProb = Convert.ToDouble(textBox2.Text);

            pbNormal.Value = 10;
            for (int i = 1; i <= simulations; i++)
            {
                Exceldata objExceldata = new Exceldata();
                objExceldata.X = "=NORM.INV(RAND()," + XMean + ", " + XStanderdDev + ")";
                objExceldata.Y = "=NORM.INV(RAND()," + YMean + ", " + YStanderdDev + ")";
                lExceldata.Add(objExceldata);
            }

            WriteDataTableToExcel("Mars", path, "Details", simulations);


        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            double var = 0;
            if (string.IsNullOrEmpty(textBox3.Text))
            {
                var = 0;
            }
            else
                var = Convert.ToDouble(textBox3.Text);

            textBox8.Text = Math.Sqrt(var).ToString();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            double var = 0;
            if (string.IsNullOrEmpty(textBox5.Text))
            {
                var = 0;
            }
            else
                var = Convert.ToDouble(textBox5.Text);

            textBox7.Text = Math.Sqrt(var).ToString();
        }

        public bool WriteDataTableToExcel(string worksheetName, string saveAsLocation, string ReporType, int simulations)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;

            lDiscrete = new List<Discrete>();
            pbNormal.Value = 15;
            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                excelSheet.Name = worksheetName;

                excelSheet.Cells[1, 1] = "X";
                excelSheet.Cells[1, 2] = "Y";

                for (int i = 2; i <= lExceldata.Count + 1; i++)
                {
                    excelSheet.Cells[i, 1] = lExceldata[i - 2].X.ToString();
                    excelSheet.Cells[i, 2] = lExceldata[i - 2].Y.ToString();

                    if (i == Convert.ToInt32((lExceldata.Count / 10)))
                        pbNormal.Value = 20;
                    else if (i == Convert.ToInt32((lExceldata.Count / 25)))
                        pbNormal.Value = 30;
                    else if (i == Convert.ToInt32((lExceldata.Count / 50)))
                        pbNormal.Value = 40;
                    else if (i == Convert.ToInt32((lExceldata.Count / 75)))
                        pbNormal.Value = 50;
                    else if (i == Convert.ToInt32((lExceldata.Count / 95)))
                        pbNormal.Value = 60;
                }


                excelworkBook.SaveAs(saveAsLocation); ;
                excelworkBook.Close();
                excel.Quit();

                string Angle = "", Distance = "";

                DataTable dtexcel = ReadExcel(saveAsLocation, ".xlsx", worksheetName);


                pbNormal.Value = 80;

                for (int i = 1; i <= dtexcel.Rows.Count; i++)
                {
                    Discrete discrete = new Discrete();
                    discrete.Iteration = i;
                    discrete.X = Convert.ToDouble(dtexcel.Rows[i - 1]["X"].ToString());
                    CumX = CumX + discrete.X;
                    discrete.SumX = CumX;
                    discrete.Y = Convert.ToDouble(dtexcel.Rows[i - 1]["Y"].ToString());
                    CumY = CumY + discrete.Y;
                    discrete.SumY = CumY;
                    discrete.FallProbability = random.NextDouble();
                    if (discrete.FallProbability < Convert.ToDouble(textBox2.Text))
                    {
                        discrete.Continue = "Yes";
                        double tan = CumY / CumX;
                        double radians = Math.Atan(tan);
                        double angle = radians * (180 / Math.PI);
                        if (CumX < 0 && CumY > 0)
                        {
                            if (angle < 0)
                                angle = 180 + angle;
                            else
                                angle = 90 + angle;
                        }
                        else if (CumX < 0 && CumY < 0)
                        {
                            if (angle < 0)
                                angle = 270 + angle;
                            else
                                angle = 180 + angle;
                        }
                        else if (CumX > 0 && CumY < 0)
                        {
                            if (angle < 0)
                                angle = 360 + angle;
                        }

                        angle = Convert.ToDouble(angle.ToString("0.000"));
                        discrete.Angle = angle.ToString();
                        totcount = totcount + 1;
                    }
                    else if (i == dtexcel.Rows.Count)
                    {
                        double tan = CumY / CumX;
                        double radians = Math.Atan(tan);
                        double angle = radians * (180 / Math.PI);
                        if (CumX < 0 && CumY > 0)
                        {
                            if (angle < 0)
                                angle = 180 + angle;
                            else
                                angle = 90 + angle;
                        }
                        else if (CumX < 0 && CumY < 0)
                        {
                            if (angle < 0)
                                angle = 270 + angle;
                            else
                                angle = 180 + angle;
                        }
                        else if (CumX > 0 && CumY < 0)
                        {
                            if (angle < 0)
                                angle = 360 + angle;
                        }

                        angle = Convert.ToDouble(angle.ToString("0.000"));
                        discrete.Angle = angle.ToString();
                    }
                    else
                    {
                        discrete.Continue = "";
                        discrete.Angle = "";
                    }

                    if (i == Convert.ToInt32((dtexcel.Rows.Count / 10)))
                        pbNormal.Value = 83;
                    else if (i == Convert.ToInt32((dtexcel.Rows.Count / 30)))
                        pbNormal.Value = 86;
                    else if (i == Convert.ToInt32((dtexcel.Rows.Count / 50)))
                        pbNormal.Value = 90;
                    else if (i == Convert.ToInt32((dtexcel.Rows.Count / 75)))
                        pbNormal.Value = 93;
                    else if (i == Convert.ToInt32((dtexcel.Rows.Count / 95)))
                        pbNormal.Value = 95;

                    discrete.X = Convert.ToDouble(discrete.X.ToString("0.000"));
                    discrete.Y = Convert.ToDouble(discrete.Y.ToString("0.000"));
                    discrete.SumX = Convert.ToDouble(discrete.SumX.ToString("0.000"));
                    discrete.SumY = Convert.ToDouble(discrete.SumY.ToString("0.000"));
                    discrete.FallProbability = Convert.ToDouble(discrete.FallProbability.ToString("0.000"));
                    double distanc = (discrete.SumX * discrete.SumX) + (discrete.SumY * discrete.SumY);
                    discrete.Distance = Convert.ToDouble(Math.Sqrt(distanc).ToString("0.000"));
                    Angle = discrete.Angle;
                    Distance = discrete.Distance.ToString();
                    lDiscrete.Add(discrete);
                }

                dataGridView2.DataSource = lDiscrete;
                label5.Text = totcount.ToString();
                label6.Text = (totcount / simulations).ToString();
                label5.ForeColor = System.Drawing.Color.Red;
                label6.ForeColor = System.Drawing.Color.Red;
                txtNXFinal.Text = CumX.ToString();
                txtNYFinal.Text = CumY.ToString();
                txtNAngle.Text = Angle;
                txtNDistance.Text = Distance;
                pbNormal.Value = 100;

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excelSheet = null;
                excelworkBook = null;
            }

        }

        public System.Data.DataTable ReadExcel(string fileName, string fileExt, string sheetname)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    //string sheetname = "Test";
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + sheetname + "$]", con); //here we read data from sheet1  
                    pbNormal.Value = 65;
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                    pbNormal.Value = 75;
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            return dtexcel;
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtDisplay.Text))
                txtDisplay.Text = "10";
            else if (string.IsNullOrEmpty(txtSimulations.Text))
            {
                MessageBox.Show("Please enter number of simulations");
                return;
            }
            else if (Convert.ToDouble(txtDisplay.Text) > Convert.ToDouble(txtSimulations.Text))
            {
                MessageBox.Show("Please enter data according to number of simulations used");
                return;
            }
            else if (Convert.ToDouble(txtDisplay.Text) > (Convert.ToDouble(txtSimulations.Text) / 2))
            {
                MessageBox.Show("Please enter data according to number of simulations used");
                return;
            }

            else if (lDiscrete.Count <= 0)
            {
                MessageBox.Show("Please complete the simulate process first");
                return;
            }

            lblDisplayCount.Text = txtDisplay.Text;
            int display = Convert.ToInt32(txtDisplay.Text);
            int simulations = Convert.ToInt32(txtSimulations.Text);
            simulations = simulations - display;
            List<Discrete> lDisplay = new List<Discrete>();
            for (int i = 1; i <= display; i++)
            {
                lDisplay.Add(lDiscrete[i - 1]);
            }
            for (int i = 1; i <= display; i++)
            {
                lDisplay.Add(lDiscrete[simulations]);
                simulations = simulations + 1;
            }
            dataGridView3.DataSource = lDisplay;
            panelUniform.Visible = true;

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            panelUniform.Visible = false;
        }

        private void btnNDisplay_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtNDisplay.Text))
                txtDisplay.Text = "10";
            else if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Please enter number of simulations");
                return;
            }
            else if (Convert.ToDouble(txtNDisplay.Text) > (Convert.ToDouble(textBox1.Text) / 2))
            {
                MessageBox.Show("Please enter data according to number of simulations used");
                return;
            }

            else if (lDiscrete.Count <= 0)
            {
                MessageBox.Show("Please complete the simulate process first");
                return;
            }

            lblNDisplayCount.Text = txtNDisplay.Text;
            int display = Convert.ToInt32(txtNDisplay.Text);
            int simulations = Convert.ToInt32(textBox1.Text);
            simulations = simulations - display;
            List<Discrete> lDisplay = new List<Discrete>();
            for (int i = 1; i <= display; i++)
            {
                lDisplay.Add(lDiscrete[i - 1]);
            }
            for (int i = 1; i <= display; i++)
            {
                lDisplay.Add(lDiscrete[simulations]);
                simulations = simulations + 1;
            }
            dataGridView4.DataSource = lDisplay;
            panelNormal.Visible = true;
        }

        private void btnNClose_Click(object sender, EventArgs e)
        {
            panelNormal.Visible = false;
        }

        private void btnNRefresh_Click(object sender, EventArgs e)
        {
            label5.Text = "0";
            label6.Text = "0";
            txtNXFinal.Text = "";
            txtNYFinal.Text = "";
            txtNAngle.Text = "";
            txtNDistance.Text = "";
            lblNDisplayCount.Text = "0";
            txtNDisplay.Text = "10";
            panelNormal.Visible = false;
            pbNormal.Value = 0;
            dataGridView4.DataSource = null;
            dataGridView2.DataSource = null;


        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            lblCountResult.Text = "0";
            lblFallResult.Text = "0";
            txtXCordinate.Text = "";
            txtYCordinate.Text = "";
            txtAngle.Text = "";
            txtDistance.Text = "";
            lblDisplayCount.Text = "0";
            txtDisplay.Text = "10";
            panelUniform.Visible = false;
            pbUniform.Value = 0;
            dataGridView3.DataSource = null;
            dataGridView1.DataSource = null;
        }
        string path = "";
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
               
                string pathN = MyDirectory();//System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                string filenamelocation = System.IO.Path.Combine(pathN, "MarsLander.txt");

                path = "";
                path = folderBrowserDialog1.SelectedPath;
                path=System.IO.Path.Combine(path, "MarsLander.xlsx");
              

                try
                {

                    //Pass the filepath and filename to the StreamWriter Constructor
                     StreamWriter sw = new StreamWriter(filenamelocation);

                    //Write a line of text
                    sw.WriteLine(path);

                    //Close the file
                    sw.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }
                finally
                {
                    Console.WriteLine("Executing finally block.");
                }

            }
        }

    }
}
