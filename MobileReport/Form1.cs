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
using System.Runtime.InteropServices;
using System.Reflection;

namespace MobileReport
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
            InitializeComponent();
            lblverizonGL.Text = "You must upload GLCODE file first before you upload Verizon Overview invoice";
           
        }

        private static Excel.Application integratedReport = new Excel.Application();
        private static Excel.Workbook workbook = integratedReport.Workbooks.Add(Missing.Value);
        private Excel.Worksheet worksheet;
        private Excel.Worksheet worksheet1;
        private Excel.Worksheet worksheet2;
        private Excel.Range head;
        private Excel.Range range;
        private Excel.Range totalAmount;

        private List<Roam> iocc = new List<Roam>();
        private List<GLcode> glCodes = new List<GLcode>();
        private List<Area> areas = new List<Area>();

        private DataTable dataCCD = new DataTable();
        private DataTable dataIOCC = new DataTable();
        private DataTable dataGLCODE = new DataTable();
        private DataTable dataAREACODE = new DataTable();
        private DataTable dataRogers = new DataTable();
        private DataTable reducedIOCCdata = new DataTable();
        private DataTable rogersInvoice = new DataTable();
        private DataTable verizonInvoice = new DataTable();
        private DataTable bellInvoice = new DataTable();

        private string fileNameCCD = null;
        private string fileNameIOCC = null;
        private string fileNameGLCODE = null;
        private string fileNameAREACODE = null;
        private string fileNameVerizon = null;
        private string fileNameBell = null;
        private string bellmsg = null;
        
        private int tableHeadRow = 8;

        private string Openfile()
        {
            string path = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*xlsx;*xlsm";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog.FileName;
                return path;
            }

            return path;
        }

        private DataTable verizonTable()
        {
            fileNameVerizon = Openfile();
            string path = fileNameVerizon.Split('\\')[fileNameVerizon.Split('\\').Length - 1];
            Connect connection = new Connect();
            DataSet verizonData = new DataSet();
            DataTable verizonDT = new DataTable();

            verizonInvoice.Columns.Add("User Number", typeof(string));
            verizonInvoice.Columns.Add("User Name", typeof(string));
            verizonInvoice.Columns.Add("MRC", typeof(decimal));
            verizonInvoice.Columns.Add("TAX", typeof(decimal));
            verizonInvoice.Columns.Add("Account charges", typeof(decimal));
            verizonInvoice.Columns.Add("Other charges & credits", typeof(decimal));
            verizonInvoice.Columns.Add("TMC", typeof(decimal));
            verizonInvoice.Columns.Add("Equipment charges", typeof(decimal));
            verizonInvoice.Columns.Add("Third party charges", typeof(decimal));
            verizonInvoice.Columns.Add("International charges", typeof(decimal));
            verizonInvoice.Columns.Add("Voice charges", typeof(decimal));

            string[] datasheets = connection.ExcelSheetNames(fileNameVerizon);
            string sheet = datasheets[0];
            verizonData = connection.OpenExcelSheet(fileNameVerizon, sheet);
            verizonDT = verizonData.Tables[0];

            int endRow = verizonDT.Rows.Count;
            Excel.Application originalFile = new Excel.Application();
            Excel.Workbook workbookTheFile = originalFile.Workbooks.Open(fileNameVerizon);
            Excel.Worksheet ws = workbookTheFile.Worksheets.get_Item(1) as Excel.Worksheet;
            Excel.Range rg = ws.Cells[endRow + 1, 1];

            try
            {
                originalFile.Visible = false;
                if (rg.Text != "Total")
                {
                    rg.EntireRow.Delete();
                    workbookTheFile.Save();
                    workbookTheFile.Close();
                }
            }
            finally
            {
                ReleaseExcelObject(rg);
                ReleaseExcelObject(ws);
                ReleaseExcelObject(workbookTheFile);
                ReleaseExcelObject(originalFile);
            }
            string msg = "Re-created the file and has been saved\n";
            lblVerizon.Text = msg;
            foreach (DataRow dr in verizonDT.Rows)
            {
                dr["Wireless number"] = dr["Wireless number"].ToString().Replace("-", string.Empty);
                if (dr["Wireless number"].ToString() == "Total")
                {
                    break;
                }
                else
                {
                    verizonInvoice.Rows.Add(dr["Wireless number"], dr["User name"],
                    dr["Monthly access charges"], dr["Taxes and surcharges"], dr["Account charges"], dr["Other charges & credits"],
                    dr["Total current charges"], dr["Equipment charges"], dr["Third party charges"], dr["International charges"], dr["Voice charges"]);
                }

            }
            verizonInvoice.Columns.Add("GLCODE", typeof(string)).SetOrdinal(0);
            verizonInvoice.Columns.Add("Division", typeof(string)).SetOrdinal(3);
            verizonInvoice.Columns.Add("Position", typeof(string)).SetOrdinal(4);

            foreach (DataRow row in verizonInvoice.Rows)
            {
                foreach (GLcode glcodes in glCodes)
                {
                    if (row["User Number"].ToString() == glcodes.UserNumber.ToString())
                    {
                        row["GLCODE"] = glcodes.GLCODE.ToString();
                        row["Division"] = glcodes.Division.ToString();
                        row["Position"] = glcodes.Position.ToString();
                        break;
                    }
                    else
                    {
                        row["GLCODE"] = "No data";
                        row["Division"] = "No data";
                        row["Position"] = "No data";
                    }
                }
            }

            verizonInvoice.Columns.Add("Subtotal", typeof(decimal)).SetOrdinal(9);
            verizonInvoice.Columns.Add("Overcharges", typeof(decimal)).SetOrdinal(11);
            foreach (DataRow dr in verizonInvoice.Rows)
            {
                dr["Subtotal"] = Convert.ToDouble(dr["MRC"]) + Convert.ToDouble(dr["TAX"]) +
                    Convert.ToDouble(dr["Account charges"]) + Convert.ToDouble(dr["Other charges & credits"]);

                dr["Overcharges"] = Convert.ToDouble(dr["TMC"]) - Convert.ToDouble(dr["Subtotal"]);
            }
            lstStatus.Items.Add(path);
            msg += "Data table has been re-created";
            lblVerizon.Text = msg;

            return verizonInvoice;
        }

        private DataTable CombinedTable()
        {
            DataTable combinedTable = new DataTable();

            foreach (DataColumn ccdCol in dataRogers.Columns)
            {
                combinedTable.Columns.Add(ccdCol.ColumnName);
            }

            for (int i = 2; i < combinedTable.Columns.Count; i++)
            {
                combinedTable.Columns[i].DataType = typeof(decimal);
            }

            int flag = 0;
            foreach (DataColumn rIOCC in reducedIOCCdata.Columns)
            {
                foreach (DataColumn cmbCol in combinedTable.Columns)
                {
                    if (cmbCol.ColumnName == rIOCC.ColumnName)
                    {
                        flag = 1;
                    }
                }
                if (flag == 0)
                {
                    combinedTable.Columns.Add(rIOCC.ColumnName, typeof(decimal));
                }
                flag = 0;
            }

            foreach (DataRow dr in dataRogers.Rows)
            {
                DataRow newData = combinedTable.NewRow();
                for (int i = 0; i < dataRogers.Columns.Count; i++)
                {
                    newData[dataRogers.Columns[i].ColumnName] = dr[dataRogers.Columns[i].ColumnName];
                }
                combinedTable.Rows.Add(newData);
            }

            foreach (DataRow cmbdata in combinedTable.Rows)
            {
                foreach (DataRow iocc in reducedIOCCdata.Rows)
                {
                    if (cmbdata["User Number"].ToString() == iocc["User Number"].ToString())
                    {
                        for (int i = 2; i < reducedIOCCdata.Columns.Count; i++)
                        {
                            cmbdata[reducedIOCCdata.Columns[i].ColumnName] = iocc[reducedIOCCdata.Columns[i].ColumnName];
                        }
                        break;
                    }
                    else
                    {
                        for (int i = 2; i < reducedIOCCdata.Columns.Count; i++)
                        {
                            cmbdata[reducedIOCCdata.Columns[i].ColumnName] = 0.00;
                        }
                        cmbdata["Roam Like Home-All"] = 3.00;
                    }
                }
            }

            combinedTable.Columns.Add("GLCODE", typeof(string)).SetOrdinal(0);
            combinedTable.Columns.Add("Province", typeof(string)).SetOrdinal(1);
            combinedTable.Columns.Add("Division", typeof(string)).SetOrdinal(2);

            foreach (DataRow cmbDtRow in combinedTable.Rows)
            {
                foreach (GLcode glcodes in glCodes)
                {
                    if (cmbDtRow["User Number"].ToString() == glcodes.UserNumber.ToString())
                    {
                        cmbDtRow["GLCODE"] = glcodes.GLCODE.ToString();
                        cmbDtRow["Division"] = glcodes.Division.ToString();
                        break;
                    }
                    else
                    {
                        cmbDtRow["GLCODE"] = "No data";
                        cmbDtRow["Division"] = "No data";
                    }
                }

                foreach (Area ar in areas)
                {
                    string prov = cmbDtRow["User Number"].ToString().Substring(0, 3);
                    if (prov == ar.AreaCode.ToString())
                    {
                        cmbDtRow["Province"] = ar.Province.ToString();
                        break;
                    }
                    else
                    {
                        cmbDtRow["Province"] = "No Data";
                    }
                }
            }

            foreach (DataRow value in combinedTable.Rows)
            {
                value["Roam Like Home-All"] = 3.00;
                value["TAX"] = Convert.ToDouble(value["GST"]) + Convert.ToDouble(value["PST"]) + Convert.ToDouble(value["HST"]) + Convert.ToDouble(value["QST"]);
                value["Subtotal"] = Convert.ToDouble(value["Monthly Service Fee"]) + Convert.ToDouble(value["Credits and Discounts"]) + Convert.ToDouble(value["ROAM LIKE HOME-All"]) + Convert.ToDouble(value["TAX"]);
                value["Overcharges"] = Convert.ToDouble(value["Total Current Charges"]) - Convert.ToDouble(value["Subtotal"]);
            }

            DataTable rogersInvoice = combinedTable.Copy();

            foreach (DataColumn dc in combinedTable.Columns)
            {
                if (dc.ColumnName == "GST")
                {
                    rogersInvoice.Columns.Remove("GST");
                }
                else if (dc.ColumnName == "PST")
                {
                    rogersInvoice.Columns.Remove("PST");
                }
                else if (dc.ColumnName == "HST")
                {
                    rogersInvoice.Columns.Remove("HST");
                }
                else if (dc.ColumnName == "QST")
                {
                    rogersInvoice.Columns.Remove("QST");
                }
                else if (dc.ColumnName == "Total Current Charges Taxable")
                {
                    rogersInvoice.Columns.Remove("Total Current Charges Taxable");
                }
                else if (dc.ColumnName == "Other Charges")
                {
                    rogersInvoice.Columns.Remove("Other Charges");
                }
                else if (dc.ColumnName == "Early Cancellation Payment")
                {
                    rogersInvoice.Columns.Remove("Early Cancellation Payment");
                }
                else if (dc.ColumnName == "Corporate Discount")
                {
                    rogersInvoice.Columns.Remove("Corporate Discount");
                }
                else if (dc.ColumnName == "Business: 10GB Pooled")
                {
                    rogersInvoice.Columns.Remove("Business: 10GB Pooled");
                }
            }

            return rogersInvoice;
        }

        private DataTable bellTable()
        {
            fileNameBell = Openfile();

            Excel.Application originalFile = new Excel.Application();
            Excel.Workbook workbookTheFile = originalFile.Workbooks.Open(fileNameBell);
            Excel.Worksheet ws = workbookTheFile.Worksheets.get_Item(1) as Excel.Worksheet;
            Excel.Range rg = ws.Cells[2, 1];

            try
            {
                originalFile.Visible = false;
                if (rg.Text == "Num compte " || rg.Text == "Num compte")
                {
                    rg.EntireRow.Delete();
                    workbookTheFile.Save();
                    workbookTheFile.Close();
                }
            }
            finally
            {
                ReleaseExcelObject(rg);
                ReleaseExcelObject(ws);
                ReleaseExcelObject(workbookTheFile);
                ReleaseExcelObject(originalFile);
            }

            Connect connection = new Connect();
            DataSet bellData = new DataSet();
            DataTable bellDT = new DataTable();
            DataTable bellDTverified = new DataTable();

            string[] datasheets = connection.ExcelSheetNames(fileNameBell);
            string sheet = datasheets[0];
            bellData = connection.OpenExcelSheet(fileNameBell, sheet);
            bellDT = bellData.Tables[0];

            bellDTverified.Columns.Add("User Number", typeof(string));
            bellDTverified.Columns.Add("User Name", typeof(string));
            bellDTverified.Columns.Add("MRC", typeof(decimal));
            bellDTverified.Columns.Add("Tax", typeof(decimal));
            bellDTverified.Columns.Add("TMC", typeof(decimal));
            bellDTverified.Columns.Add("Ftr Chg Ttl", typeof(decimal));
            bellDTverified.Columns.Add("Txt Msg Amt", typeof(decimal));
            bellDTverified.Columns.Add("Airtime Chg", typeof(decimal));
            bellDTverified.Columns.Add("Data Chg", typeof(decimal));
            bellDTverified.Columns.Add("Roamer Chg", typeof(decimal));
            bellDTverified.Columns.Add("Roamer LD", typeof(decimal));
            bellDTverified.Columns.Add("Other Chgs", typeof(decimal));
            bellDTverified.Columns.Add("Disc Ttl", typeof(decimal));
            bellDTverified.Columns.Add("Rm datachrg", typeof(decimal));
            
            foreach(DataRow dr in bellDT.Rows)
            {
                decimal tax = Convert.ToDecimal(dr["GST        "].ToString()) + Convert.ToDecimal(dr["HST        "].ToString()) + Convert.ToDecimal(dr["HST-PEI Tel"].ToString()) +
                    Convert.ToDecimal(dr["HST-ON Tel "].ToString()) + Convert.ToDecimal(dr["HST-BC Tel "].ToString()) + Convert.ToDecimal(dr["ORST       "].ToString()) +
                    Convert.ToDecimal(dr["QST - Telec"].ToString()) + Convert.ToDecimal(dr["QST - Other"].ToString()) + Convert.ToDecimal(dr["P#E#I# PST "].ToString()) +
                    Convert.ToDecimal(dr["BC PST     "].ToString()) + Convert.ToDecimal(dr["Sask       "].ToString()) + Convert.ToDecimal(dr["Manitoba   "].ToString()) +
                    Convert.ToDecimal(dr["Foreign tax"].ToString());
                string firstname = dr["Surname             "].ToString().Trim();
                string lastname = dr["Given Name   "].ToString().Trim();
                bellDTverified.Rows.Add(dr["Mobile Nbr"], string.Concat(firstname," ", lastname),
                    dr["Mth Chg Ttl"], tax, dr["Ttl Charges"], dr["Ftr Chg Ttl"], dr["Txt Msg Amt"], dr["Airtime Chg"], dr["Data Chg   "], dr["Roamer Chg "],
                    dr["Roamer LD  "], dr["Other Chgs "], dr["Disc Ttl   "], dr["Rm datachrg"]);
            }

            double sum = 0;
            double total = 0;

            for (int i = 2; i < bellDTverified.Columns.Count; i++)
            {
                foreach (DataRow dr in bellDTverified.Rows)
                {
                    if (DBNull.Value.Equals(dr[bellDTverified.Columns[i].ColumnName]))
                    {
                        break;
                    }
                    else
                    {
                        sum = Convert.ToDouble(dr[bellDTverified.Columns[i].ColumnName]);
                        total += sum;
                    }
                }
                if (total != 0)
                {
                    bellInvoice.Columns.Add(bellDTverified.Columns[i].ColumnName, typeof(decimal));
                }
                total = 0;
            }
            bellInvoice.Columns.Add("User Number", typeof(string)).SetOrdinal(0);
            bellInvoice.Columns.Add("User Name", typeof(string)).SetOrdinal(1);

            // Create new DataTable object                
            foreach (DataRow dr in bellDTverified.Rows)
            {
                DataRow newData = bellInvoice.NewRow();
                for (int j = 0; j < bellInvoice.Columns.Count; j++)
                {
                    newData[bellInvoice.Columns[j].ColumnName] = dr[bellInvoice.Columns[j].ColumnName];
                }
                bellInvoice.Rows.Add(newData);
            }

            bellInvoice.Columns.Add("GLCODE", typeof(string)).SetOrdinal(0);
            bellInvoice.Columns.Add("Province", typeof(string)).SetOrdinal(1);
            bellInvoice.Columns.Add("Division", typeof(string)).SetOrdinal(2);
            bellInvoice.Columns.Add("Subtotal", typeof(decimal)).SetOrdinal(7);
            bellInvoice.Columns.Add("Overcharges", typeof(decimal)).SetOrdinal(9);

            foreach(DataRow value in bellInvoice.Rows)
            {
                value["Subtotal"] = Convert.ToDouble(value["MRC"]) + Convert.ToDouble(value["Tax"]);
                value["Overcharges"] = Convert.ToDouble(value["TMC"]) - Convert.ToDouble(value["Subtotal"]);
            }

            foreach (DataRow dr in bellInvoice.Rows)
            {
                foreach (GLcode glcodes in glCodes)
                {
                    if (dr["User Number"].ToString() == glcodes.UserNumber.ToString())
                    {
                        dr["GLCODE"] = glcodes.GLCODE.ToString();
                        dr["Division"] = glcodes.Division.ToString();
                        break;
                    }
                    else
                    {
                        dr["GLCODE"] = "No data";
                        dr["Division"] = "No data";
                    }
                }

                foreach (Area ar in areas)
                {
                    string prov = dr["User Number"].ToString().Substring(0, 3);
                    if (prov == ar.AreaCode.ToString())
                    {
                        dr["Province"] = ar.Province.ToString();
                        break;
                    }
                    else
                    {
                        dr["Province"] = "No Data";
                    }
                }
            }
            return bellInvoice;
        }

        private void HeadTable(Excel.Range head, Excel.Worksheet worksheet, DataTable dt, int row, int col, int columnWidth)
        {
            head = worksheet.Cells[row, col + 1];
            head.Value2 = dt.Columns[col].ColumnName;
            head.ColumnWidth = columnWidth;
            head.WrapText = true;
            head.Font.Bold = true;
            head.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            head.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            head.Interior.Color = Color.FromKnownColor(KnownColor.Black);
            head.Font.Color = Color.FromKnownColor(KnownColor.White);
            head.BorderAround2(Excel.XlLineStyle.xlDash);
            head.Borders.Color = Color.FromKnownColor(KnownColor.White);
        }

        private void GenerateExcelsheet(Excel.Range head, Excel.Worksheet worksheet, DataTable dt, int row, int columnWidth, string vendor)
        {
            int rowIndex = tableHeadRow + 1;
            int totalRows = dt.Rows.Count;
            double total = 0.0;

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                HeadTable(head, worksheet, dt, tableHeadRow, i, 15);

                if (i >= 5)
                {
                    HeadTable(head, worksheet, dt, 1, i, 15);

                    totalAmount = worksheet.Cells[2, i + 1];
                    totalAmount.Value2 = 0.0;
                    totalAmount.NumberFormat = "$#,##0.00";
                }
            }

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    range = worksheet.Cells[rowIndex, i + 1];
                    if (typeof(string) == dr[dt.Columns[i].ColumnName].GetType())
                    {
                        range.Value2 = dr[dt.Columns[i].ColumnName].ToString();
                        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        range.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                    else
                    {
                        range.Value2 = Convert.ToDouble(dr[dt.Columns[i].ColumnName].ToString());
                        if (range.Value2 < 0)
                        {
                            range.NumberFormat = "$#,##0.00";
                            range.Font.Color = Color.Red;
                        }
                        range.NumberFormat = "$#,##0.00";

                        totalAmount = worksheet.Cells[2, i + 1];
                        total = Convert.ToDouble(totalAmount.Value2);
                        total += range.Value2;
                        totalAmount.Value2 = total;
                        if(totalAmount.Value2 < 0)
                        {
                            totalAmount.NumberFormat = "$#,##0.00";
                            totalAmount.Font.Color = Color.Red;
                        }
                        totalAmount.NumberFormat = "$#,##0.00";
                    }
                }
                rowIndex++;
                if(vendor == "Rogers")
                {
                    string txt = "writing {0} rows / {1} rows";
                    string statusTxt = string.Format(txt, rowIndex - tableHeadRow - 1, totalRows);
                    lblDescRog.Text = statusTxt;
                }
                else if(vendor == "Verizon")
                {
                    string txt = "writing {0} rows / {1} rows";
                    string statusTxt = string.Format(txt, rowIndex - tableHeadRow - 1, totalRows);
                    lblVerizon.Text = statusTxt;
                }
                else if(vendor == "Bell")
                {
                    string txt = "writing {0} rows / {1} rows";
                    string statusTxt = string.Format(txt, rowIndex - tableHeadRow - 1, totalRows);
                    lblBell.Text = statusTxt;
                }
                               
            }
        }

        private string DecToAlphabet(int num)
        {
            int rest;
            string alphabet;

            byte[] asciiA = Encoding.ASCII.GetBytes("A");
            rest = num % 26;
            asciiA[0] += (byte)rest;

            alphabet = Encoding.ASCII.GetString(asciiA);

            num = num / 26 - 1;
            if (num > -1)
            {
                alphabet = alphabet.Insert(0, DecToAlphabet(num));
            }

            return alphabet;
        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Reset()
        {
            dataCCD.Clear();
            dataIOCC.Clear();
            dataGLCODE.Clear();
            dataAREACODE.Clear();

            iocc.Clear();
            glCodes.Clear();
            areas.Clear();

            lstReference.Items.Clear();
            lstStatus.Items.Clear();
            fileNameCCD = null;
            fileNameIOCC = null;
            fileNameGLCODE = null;
            fileNameAREACODE = null;
        }

        private void btnGLCODE_Click(object sender, EventArgs e)
        {
            Connect connectExcel = new Connect();
            fileNameGLCODE = Openfile();
            string[] sheetNames = connectExcel.ExcelSheetNames(fileNameGLCODE);
            string sheetName = sheetNames[0];
            string Path = fileNameGLCODE.Split('\\')[fileNameGLCODE.Split('\\').Length - 1];
            lstReference.Items.Add(Path);

            DataSet dataset = connectExcel.OpenExcelSheet(fileNameGLCODE, sheetName);
            dataGLCODE = dataset.Tables[0];

            foreach (DataRow dr in dataGLCODE.Rows)
            {
                GLcode glcode = new GLcode(dr["User Number"], dr["GL Number"], dr["Division"], dr["Position"], dr["UserName"]);
                glCodes.Add(glcode);
            }
            string msgGL = "GLCODE file has been loaded.";
            lstStatus.Items.Add(msgGL);
            lblverizonGL.Text = msgGL;
            bellmsg = msgGL + "\n";
            lblBellGLAR.Text = bellmsg;

            MessageBox.Show(msgGL);
        }

        private void btnAREACODE_Click(object sender, EventArgs e)
        {
            Connect connectExcel = new Connect();
            fileNameAREACODE = Openfile();
            string[] sheetNames = connectExcel.ExcelSheetNames(fileNameAREACODE);
            string sheetName = sheetNames[0];
            string Path = fileNameAREACODE.Split('\\')[fileNameAREACODE.Split('\\').Length - 1];
            lstReference.Items.Add(Path);

            DataSet dataset = connectExcel.OpenExcelSheet(fileNameAREACODE, sheetName);
            dataAREACODE = dataset.Tables[0];

            foreach (DataRow dr2 in dataAREACODE.Rows)
            {
                Area area = new Area(dr2["Area Code"], dr2["Province"], dr2["Tax"]);
                areas.Add(area);
            }
            string msgArea = "AREA CODE file has been loaded.";
            lstStatus.Items.Add(msgArea);
            bellmsg += msgArea;
            lblBellGLAR.Text = bellmsg;
            MessageBox.Show(msgArea);
        }

        private void btnRogers_CCD_MouseHover(object sender, EventArgs e)
        {
            string defaultmsg = "Description\n\n";
            string msg = "CCD file indicates 'Rogers - Monthly Charges breackdown Report'.\n";
            lblDescRog.Text = String.Concat(defaultmsg, msg);
        }

        private void btnRogers_CCD_MouseLeave(object sender, EventArgs e)
        {
            string defaultmsg = "Description\n\n";
            lblDescRog.Text = defaultmsg;
        }

        private void btnRogers_IOCC_MouseHover(object sender, EventArgs e)
        {
            string defaultmsg = "Description\n\n";
            string msg = "IOCC or CCDI file indicates 'Rogers - Roam Like Charges Breakdown'.\n";
            lblDescRog.Text = String.Concat(defaultmsg, msg);
        }

        private void btnRogers_IOCC_MouseLeave(object sender, EventArgs e)
        {
            string defaultmsg = "Description\n\n";
            lblDescRog.Text = defaultmsg;
        }

        private void btnCombinRogers_MouseHover(object sender, EventArgs e)
        {
            string defaultmsg = "Description\n\n";
            string msg = "If you already upload GLCODE and AREA CODE file under 'HOME' tab.\n";
            string msg1 = "Then you can hit this button for creating Rogers report";
            lblDescRog.Text = String.Concat(defaultmsg, msg, msg1);
        }

        private void btnCombinRogers_MouseLeave(object sender, EventArgs e)
        {
            string defaultmsg = "Description\n\n";
            lblDescRog.Text = defaultmsg;
        }

        private void btnRogers_CCD_Click(object sender, EventArgs e)
        {
            fileNameCCD = Openfile();
            string ccdPath = fileNameCCD.Split('\\')[fileNameCCD.Split('\\').Length - 1];
            lstStatus.Items.Add(ccdPath);
            lblDescRog.Text = " ";
            lblDescRog.Text = string.Concat(ccdPath, "\nLoading....");
            string statusMSG = lblDescRog.Text;
            Excel.Application originalFile = new Excel.Application();
            Excel.Workbook workbookTheFile = originalFile.Workbooks.Open(fileNameCCD);
            Excel.Worksheet ws = workbookTheFile.Worksheets.get_Item(1) as Excel.Worksheet;
            Excel.Range rg = ws.Cells[1, 1];
            string statusMSG3 = null;
            string statusMSG4 = null;

            try
            {
                originalFile.Visible = false;
                if (rg.Text != "Billing Account")
                {
                    lblDescRog.Text = statusMSG + "\n invalid data, re-create this file.";
                    string statusMSG2 = lblDescRog.Text;
                    rg.EntireRow.Delete();
                    workbookTheFile.Save();
                    workbookTheFile.Close();
                    lblDescRog.Text = statusMSG2 + "\n The file has been re-created.";
                    statusMSG3 = lblDescRog.Text + "\n";
                }
            }
            finally
            {
                ReleaseExcelObject(rg);
                ReleaseExcelObject(ws);
                ReleaseExcelObject(workbookTheFile);
                ReleaseExcelObject(originalFile);
            }

            // Create a DataTable object for CCD file
            Connect connection = new Connect();
            DataSet rogersData = new DataSet();
            DataTable dataTable = rogersData.Tables.Add("rogersTable");

            try
            {
                string[] datasheets = connection.ExcelSheetNames(fileNameCCD);
                string sheet = datasheets[0];
                rogersData = connection.OpenExcelSheet(fileNameCCD, sheet);
                dataTable = rogersData.Tables[0];

                double sum = 0;
                double total = 0;
                // To find out which sum of values of Column is zero

                dataRogers.Columns.Add("User Number", typeof(string));
                dataRogers.Columns.Add("User Name", typeof(string));

                // Add Columns to 'dataRogers' DataTable
                for (int i = 5; i < dataTable.Columns.Count; i++)
                {
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if (DBNull.Value.Equals(dr[dataTable.Columns[i].ColumnName]))
                        {
                            break;
                        }
                        else
                        {
                            sum = Convert.ToDouble(dr[dataTable.Columns[i].ColumnName]);
                            total += sum;
                        }
                    }
                    if (total != 0)
                    {
                        dataRogers.Columns.Add(dataTable.Columns[i].ColumnName, typeof(decimal));
                    }
                    total = 0;
                }

                lblDescRog.Text = statusMSG3 +"\n" + string.Format("{0}", ccdPath);
                statusMSG4 = lblDescRog.Text;
                // Create new DataTable object                
                foreach (DataRow dr in dataTable.Rows)
                {
                    DataRow newData = dataRogers.NewRow();
                    for (int j = 0; j < dataRogers.Columns.Count; j++)
                    {
                        newData[dataRogers.Columns[j].ColumnName] = dr[dataRogers.Columns[j].ColumnName];
                    }
                    dataRogers.Rows.Add(newData);
                }

                // Reordering Columns
                dataRogers.Columns["Credits and Discounts"].SetOrdinal(3);
                dataRogers.Columns.Add("Roam Like Home-All").SetOrdinal(4);
                dataRogers.Columns.Add("TAX").SetOrdinal(5);
                dataRogers.Columns["Subtotal"].SetOrdinal(6);
                dataRogers.Columns["Total Current Charges"].SetOrdinal(7);
                dataRogers.Columns.Add("Overcharges").SetOrdinal(8);
            }
            finally
            {
                lblDescRog.Text = statusMSG4 + "\n\n" + "Monthly Charges breakdown Report loaded";
                int mcbR = dataRogers.Rows.Count;
                string msg = "{0} Users loaded";
                string msgFormat = string.Format(msg, mcbR);
                MessageBox.Show(msgFormat);
            }
        }

        private void btnRogers_IOCC_Click(object sender, EventArgs e)
        {
            Connect connectIOCC = new Connect();
            fileNameIOCC = Openfile();
            string Path = fileNameIOCC.Split('\\')[fileNameIOCC.Split('\\').Length - 1];
            lstStatus.Items.Add(Path);
            DataSet IoccOriginalaData = new DataSet();
            DataTable dataTable = IoccOriginalaData.Tables.Add("IOCC-Data-Table");
            DataTable rogersIOCCdata = new DataTable();
            lblDescRog.Text = Path + "\n" + "Loading...";
            string msg1 = lblDescRog.Text;
            string msg2 = null;
            string msg3 = null;
           
            Excel.Application originalFile = new Excel.Application();
            Excel.Workbook workbookTheFile = originalFile.Workbooks.Open(fileNameIOCC);
            Excel.Worksheet ws = workbookTheFile.Worksheets.get_Item(1) as Excel.Worksheet;
            Excel.Range rg = ws.Cells[1, 1];
            try
            {
                originalFile.Visible = false;
                if (rg.Text != "Billing Account")
                {
                    rg.EntireRow.Delete();
                    workbookTheFile.Save();
                    workbookTheFile.Close();
                }
            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(workbookTheFile);
                ReleaseExcelObject(originalFile);
            }

            try
            {                
                string[] sheetNames = connectIOCC.ExcelSheetNames(fileNameIOCC);
                string sheetName = sheetNames[0];
                IoccOriginalaData = connectIOCC.OpenExcelSheet(fileNameIOCC, sheetName);
                dataTable = IoccOriginalaData.Tables[0];

                foreach (DataRow row in dataTable.Rows)
                {
                    Roam roam = new Roam(row["User Number"], row["User Name"], row["Charges/Credits Description"], row["Other Charges/Credits Amount"]);
                    iocc.Add(roam);
                }
                lblDescRog.Text = msg1 + "\n" + "has been loaded";
                msg2 = lblDescRog.Text;

                var reformData =
                            from roam in iocc
                            group roam by new
                            {
                                Number = roam.UserNumber,
                                Name = roam.UserName,
                                Description = roam.Description
                            } into numberBy
                            from sortByDescription in (
                                from roam in numberBy
                                group roam by roam.Description
                            )
                            group sortByDescription by numberBy.Key;

                int ioccCount = iocc.Count();
                List<string> columns = new List<string>();

                foreach (var description in iocc)
                {
                    columns.Add(description.Description.ToString());
                }
                List<string> titleColumns = columns.Distinct().ToList();

                rogersIOCCdata.Columns.Add("User Number", typeof(string));
                rogersIOCCdata.Columns.Add("User Name", typeof(string));

                foreach (string name in titleColumns)
                {
                    rogersIOCCdata.Columns.Add(name, typeof(decimal));
                }


                foreach (var data in reformData)
                {
                    foreach (var des in data)
                    {
                        foreach (var number in des)
                        {
                            if (rogersIOCCdata.Rows.Count > 0)
                            {
                                foreach (DataRow dataRow in rogersIOCCdata.Rows)
                                {
                                    if (dataRow["User Number"].ToString() == number.UserNumber.ToString())
                                    {
                                        for (int l = 0; l < titleColumns.Count; l++)
                                        {
                                            double value = 0.0;
                                            string title = titleColumns[l];
                                            if (number.Description.ToString() == titleColumns[l])
                                            {
                                                value = Convert.ToDouble(data.Sum(x => x.Sum(y => (decimal)y.Amount)));
                                                dataRow[title] = value;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            DataRow row = rogersIOCCdata.NewRow();
                            row["User Number"] = number.UserNumber.ToString();
                            row["User Name"] = number.UserName.ToString();

                            for (int l = 0; l < titleColumns.Count; l++)
                            {
                                double value = 0.0;
                                string title = titleColumns[l];
                                if (number.Description.ToString() == titleColumns[l])
                                {
                                    value = Convert.ToDouble(data.Sum(x => x.Sum(y => (decimal)y.Amount)));
                                    row[title] = value;
                                }
                                else
                                {
                                    row[title] = value;
                                }
                            }
                            int flag = 0;
                            foreach (DataRow dr in rogersIOCCdata.Rows)
                            {
                                if (number.UserNumber.ToString() == dr["User Number"].ToString())
                                {
                                    flag = 1;
                                }
                            }
                            if (flag == 0)
                            {
                                rogersIOCCdata.Rows.Add(row);
                                break;
                            }
                            flag = 0;
                        }
                    }
                }

                decimal sum = 0;
                decimal total = 0;

                reducedIOCCdata.Columns.Add("User Number", typeof(string));
                reducedIOCCdata.Columns.Add("User Name", typeof(string));
                lblDescRog.Text = msg2 + "\n" + "The data table has been transformed";
                msg3 = lblDescRog.Text;


                for (int i = 0; i < rogersIOCCdata.Columns.Count; i++)
                {
                    foreach (DataRow dr in rogersIOCCdata.Rows)
                    {
                        if (typeof(string) == dr[rogersIOCCdata.Columns[i].ColumnName].GetType())
                        {
                            break;
                        }
                        else
                        {
                            sum = Convert.ToDecimal(dr[rogersIOCCdata.Columns[i].ColumnName]);
                            total += sum;
                        }
                    }
                    if (total != 0)
                    {
                        reducedIOCCdata.Columns.Add(rogersIOCCdata.Columns[i].ColumnName, typeof(decimal));
                        total = 0;
                    }
                    else
                    {
                        total = 0;
                    }
                }

                foreach (DataRow dr in rogersIOCCdata.Rows)
                {
                    DataRow newData = reducedIOCCdata.NewRow();
                    for (int k = 0; k < reducedIOCCdata.Columns.Count; k++)
                    {
                        newData[reducedIOCCdata.Columns[k].ColumnName] = dr[reducedIOCCdata.Columns[k].ColumnName];
                        if (Convert.ToDouble(dr["Roam Like Home-All"].ToString()) == 0)
                        {
                            newData["Roam Like Home-All"] = 3.00;
                        }
                    }
                    reducedIOCCdata.Rows.Add(newData);
                }
            }
            finally
            {
                lblDescRog.Text = msg3 + "\n\n" + "Roam Like Charges Breakdown loaded";
                int mcbR = reducedIOCCdata.Rows.Count;
                string msg = "{0} Users loaded";
                string msgFormat = string.Format(msg, mcbR);
                MessageBox.Show(msgFormat);
            }
        }

        private void btnCombinRogers_Click(object sender, EventArgs e)
        {
            lblDescRog.Text = "Creating a Rogers report";
            string msg1 = lblDescRog.Text;
            DataTable rogersReport = CombinedTable();

            Excel.Application application = new Excel.Application();
            Excel.Workbook rogersBook = application.Workbooks.Add();
            Excel.Worksheet rogersSheet = rogersBook.Worksheets.get_Item(1) as Excel.Worksheet;
            rogersSheet.Name = "Rogers";

            string msg2 = null;

            int rowIndex = tableHeadRow + 1;
            int totalRows = rogersReport.Rows.Count;

            double total = 0.0;
            try
            {
                GenerateExcelsheet(head, rogersSheet, rogersReport, tableHeadRow, 15, "Rogers");

                lblDescRog.Text = msg1 + "\n" + "all data has been populated in worksheet.";
                msg2 = lblDescRog.Text;

                // Set the width of Division column
                string col1 = "C{0}:C{1}";
                string colRange = String.Format(col1, tableHeadRow + 1, rogersReport.Rows.Count + tableHeadRow);
                range = rogersSheet.Range[colRange];
                range.EntireColumn.AutoFit();
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;                              // 2: left, 3: center, 4: right
                // Set the width of Username column
                string col2 = "E{0}:E{1}";
                string colRange1 = String.Format(col2, tableHeadRow + 1, rogersReport.Rows.Count + tableHeadRow);
                range = rogersSheet.Range[colRange1];
                range.EntireColumn.AutoFit();
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;

                rogersSheet.Application.ActiveWindow.SplitRow = 8;
                rogersSheet.Application.ActiveWindow.SplitColumn = 5;
                rogersSheet.Application.ActiveWindow.FreezePanes = true;

                for (int i = 1; i < rogersReport.Rows.Count + tableHeadRow; i++)
                {
                    range = rogersSheet.Cells[tableHeadRow + i, 12];
                    if (range.Value2 > 0)
                    {
                        range = rogersSheet.Range[rogersSheet.Cells[tableHeadRow + i, 1], rogersSheet.Cells[tableHeadRow + i, rogersReport.Columns.Count]];
                        range.Font.Bold = true;
                        range.Interior.Color = Color.FromArgb(247, 208, 101);
                    }
                }

                // Filter
                range = rogersSheet.Range[rogersSheet.Cells[tableHeadRow, 1], rogersSheet.Cells[rogersReport.Columns.Count, rogersReport.Rows.Count + tableHeadRow]];
                range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                rogersSheet.Name = "Rogers";
                application.Visible = true;
                
                lblDescRog.Text = msg2 + "\n\n" + "Opening Excel file" + "\n\n" + "Report has been creatd";
            }
            finally
            {
                ReleaseExcelObject(rogersSheet);
                ReleaseExcelObject(rogersBook);
                ReleaseExcelObject(application);                
            }
        }

        private void btnVerizon_Click(object sender, EventArgs e)
        {
            if(glCodes.Count == 0)
            {
                btnVerizon.Equals(false);
            }
            else
            {
                verizonTable();
                string msg1 = "Verizon Overview Charges Report has been uploaded";
                string msg2 = "Verizon data has been re-created";
                lblVerizon.Text = string.Concat(msg1, "\n", msg2);

                lstStatus.Items.Add(msg1);
                lstStatus.Items.Add(msg2);
            }
        }

        private void btnCreateVerizonReport_Click(object sender, EventArgs e)
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook verizonBook = application.Workbooks.Add();
            Excel.Worksheet verizonSheet = verizonBook.Worksheets.get_Item(1) as Excel.Worksheet;
            try
            {
                GenerateExcelsheet(head, verizonSheet, verizonInvoice, tableHeadRow, 15, "Verizon");
                string msg = "Excel sheet for Verizon has been created.\n";
                lblVerizon.Text = msg;
                lstStatus.Items.Add(msg);
                // Set the width of User Name column
                string col1 = "C{0}:C{1}";
                string colRange1 = String.Format(col1, tableHeadRow + 1, verizonInvoice.Rows.Count + tableHeadRow);
                range = verizonSheet.Range[colRange1];
                range.EntireColumn.AutoFit();
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;                              // 2: left, 3: center, 4: right

                // Set the width of User Name column
                string col2 = "D{0}:D{1}";
                string colRange2 = String.Format(col2, tableHeadRow + 1, verizonInvoice.Rows.Count + tableHeadRow);
                range = verizonSheet.Range[colRange2];
                range.ColumnWidth = 15;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;                              // 2: left, 3: center, 4: right

                // Set the width of Position column
                string col3 = "E{0}:E{1}";
                string colRange3 = String.Format(col3, tableHeadRow + 1, verizonInvoice.Rows.Count + tableHeadRow);
                range = verizonSheet.Range[colRange3];
                range.EntireColumn.AutoFit();
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;

                verizonSheet.Application.ActiveWindow.SplitRow = 8;
                verizonSheet.Application.ActiveWindow.SplitColumn = 5;
                verizonSheet.Application.ActiveWindow.FreezePanes = true;

                for (int i = 1; i < verizonInvoice.Rows.Count + tableHeadRow; i++)
                {
                    range = verizonSheet.Cells[tableHeadRow + i, 12];
                    if (range.Value2 > 0)
                    {
                        range = verizonSheet.Range[verizonSheet.Cells[tableHeadRow + i, 1], verizonSheet.Cells[tableHeadRow + i, verizonInvoice.Columns.Count]];
                        range.Font.Bold = true;
                        range.Interior.Color = Color.FromArgb(247, 208, 101);
                    }
                }

                // Filter
                range = verizonSheet.Range[verizonSheet.Cells[tableHeadRow, 1], verizonSheet.Cells[verizonInvoice.Columns.Count, verizonInvoice.Rows.Count + tableHeadRow]];
                range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                msg += "Verizon report created";
                lblVerizon.Text = msg;
                lstStatus.Items.Add("Verizon report created");
                verizonSheet.Name = "Verizon";

                application.Visible = true;
            }
            finally
            {
                ReleaseExcelObject(verizonSheet);
                ReleaseExcelObject(verizonBook);
                ReleaseExcelObject(application);
            }
        }

        private void btnBellFile_Click(object sender, EventArgs e)
        {
            if(glCodes.Count == 0 || areas.Count == 0)
            {
                btnBellFile.Enabled = false;
            }
            else
            {
                bellTable();
                string msg = "Bell data table has been created.";
                lblBell.Text = msg;
                lstStatus.Items.Add(msg);
            }
        }

        private void btnBellReport_Click(object sender, EventArgs e)
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook bellBook = application.Workbooks.Add();
            Excel.Worksheet bellSheet = bellBook.Worksheets.get_Item(1) as Excel.Worksheet;
            try
            {
                GenerateExcelsheet(head, bellSheet, bellInvoice, tableHeadRow, 15, "Bell");
                string msg = "All data has been populated on worksheet";
                lblBell.Text = msg;

                // Set the width of User Name column
                string col1 = "C{0}:C{1}";
                string colRange1 = String.Format(col1, tableHeadRow + 1, bellInvoice.Rows.Count + tableHeadRow);
                range = bellSheet.Range[colRange1];
                range.EntireColumn.AutoFit();
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;                              // 2: left, 3: center, 4: right

                // Set the width of User Name column
                string col2 = "D{0}:D{1}";
                string colRange2 = String.Format(col2, tableHeadRow + 1, bellInvoice.Rows.Count + tableHeadRow);
                range = bellSheet.Range[colRange2];
                range.ColumnWidth = 15;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;                              // 2: left, 3: center, 4: right

                // Set the width of Position column
                string col3 = "E{0}:E{1}";
                string colRange3 = String.Format(col3, tableHeadRow + 1, bellInvoice.Rows.Count + tableHeadRow);
                range = bellSheet.Range[colRange3];
                range.EntireColumn.AutoFit();
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.HorizontalAlignment = 2;

                bellSheet.Application.ActiveWindow.SplitRow = 8;
                bellSheet.Application.ActiveWindow.SplitColumn = 5;
                bellSheet.Application.ActiveWindow.FreezePanes = true;

                for (int i = 1; i < bellInvoice.Rows.Count + tableHeadRow; i++)
                {
                    range = bellSheet.Cells[tableHeadRow + i, 10];
                    if (range.Value2 > 0)
                    {
                        range = bellSheet.Range[bellSheet.Cells[tableHeadRow + i, 1], bellSheet.Cells[tableHeadRow + i, bellInvoice.Columns.Count]];
                        range.Font.Bold = true;
                        range.Interior.Color = Color.FromArgb(247, 208, 101);
                    }
                }
                // Filter
                range = bellSheet.Range[bellSheet.Cells[tableHeadRow, 1], bellSheet.Cells[bellInvoice.Columns.Count, bellInvoice.Rows.Count + tableHeadRow]];
                range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                
                msg += "Bell Report created";
                lblBell.Text = msg;
                
                bellSheet.Name = "Bell";
                application.Visible = true;
            }
            finally
            {
                ReleaseExcelObject(bellSheet);
                ReleaseExcelObject(bellBook);
                ReleaseExcelObject(application);
            }
        }

        private void btnVerizon_MouseHover(object sender, EventArgs e)
        {
            lblVerizon.Text = "Please make sure that you uploaded GLCODE file under 'Home' tab\nThis button will be activated once you upload GLCODE file.";
        }

        private void btnVerizon_MouseLeave(object sender, EventArgs e)
        {
            lblVerizon.Text = "";
        }

        private void btnBellFile_MouseHover(object sender, EventArgs e)
        {
            if(glCodes.Count > 0 && areas.Count > 0)
            {
                lblBell.Text = "GLCODE file and AREA CODE file are uploaded.";
            }
            lblBell.Text = "It requirs GLCODE file and AREA CODE file.\nPlease upload both files under 'HOME' tab";
        }

        private void btnBellFile_MouseLeave(object sender, EventArgs e)
        {
            lblBell.Text = "";
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            /*
            if(DBNull.Value.Equals(verizonInvoice) == true || DBNull.Value.Equals(dataRogers) == true || DBNull.Value.Equals(bellInvoice) == true)
            {
                btnCreate.Equals(false);
            }
            else
            {
                btnCreate.Equals(true);
                lstStatus.Items.Add("Data table of Rogers, Verizon, and Bell uploaded");
                workbook.Worksheets.Add(worksheet);
                workbook.Worksheets.Add(worksheet1);
                workbook.Worksheets.Add(worksheet2);
            }
            */
            rogersInvoice = CombinedTable();

            worksheet = workbook.Worksheets.get_Item(1) as Excel.Worksheet;
            worksheet1 = workbook.Worksheets.get_Item(1) as Excel.Worksheet;
            worksheet2 = workbook.Worksheets.get_Item(1) as Excel.Worksheet;

            worksheet = workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
            worksheet.Name = "Rogers";
            worksheet1 = workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
            worksheet1.Name = "Verizon";
            worksheet2 = workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
            worksheet2.Name = "Bell";
           
            try
            {                
                GenerateExcelsheet(head, worksheet, rogersInvoice, tableHeadRow, 15, "Rogers");
                GenerateExcelsheet(head, worksheet1, verizonInvoice, tableHeadRow, 15, "Verizon");
                GenerateExcelsheet(head, worksheet2, bellInvoice, tableHeadRow, 15, "Bell");
                integratedReport.Visible = true;
            }
            finally
            {
                ReleaseExcelObject(worksheet);
                ReleaseExcelObject(worksheet1);
                ReleaseExcelObject(worksheet2);
                ReleaseExcelObject(workbook);
                ReleaseExcelObject(integratedReport);
            }
        }
    }
}
