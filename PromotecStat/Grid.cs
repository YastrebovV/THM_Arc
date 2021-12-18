using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using S7.Net;
using System.Data.SQLite;
using System.Diagnostics;
using System.Runtime.InteropServices;
using MetroFramework;
using MetroFramework.Controls;

namespace THM_Arc
{
    public partial class Grid : MetroFramework.Forms.MetroForm
    {
        public Grid()
        {
            InitializeComponent();
        }

        IniClass ini = new IniClass();

        SQLiteConnection dbConnection = null;
        SQLiteCommand sqlCommand = null;

        Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(int hWnd, ref int lpdwProcessId);

        private bool dragging = false;
        private Point dragCursorPoint;
        private Point dragFormPoint;

        int num_Weld = 0;
        int num_Controllers = 0;
        string[] name_plc;

        string AllTable = "SELECT  [Наименование оборудования],[Номер сварки],Дата,[Ток с учетом обр.связи]," +
                         "[Ток с панели оператора],[Ток - ответ],[Напряжение - ответ],[Напряжение с учетом обр.связи],[Скорость проволоки],[Скорость подачи]" +
                         "FROM  ParamArc";
        string FormatedTable;
        string NumWeld;


        private void KillExcel(Microsoft.Office.Interop.Excel.Application application)
        {
            int excelProcessId = -1;
            GetWindowThreadProcessId(application.Hwnd, ref excelProcessId);

            try
            {
                Process process = Process.GetProcessById(excelProcessId);
                process.Kill();
            }
            finally { }
        }
        private void SaveFile(Microsoft.Office.Interop.Excel.Workbook xlWorkBook)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"C:\";   
            saveFileDialog1.Title = "Save Excel Files";
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                xlWorkBook.SaveAs(saveFileDialog1.FileName);
            }
        }
        private bool FillCombo(out int num_Weld)
        {
             metroComboBox2.Items.Clear();
             NumWeld = "SELECT [Номер сварки] FROM  ParamArc WHERE [Наименование оборудования]='" + metroComboBox1.Text + "' ORDER BY [Номер сварки] DESC";
             SelectFromDB(NumWeld, false, out num_Weld);
            if (num_Weld != 0)
            {
                return true;
            }else return false;              
        }
        private void FilterGrid()
        {
            if (metroCheckBox1.Checked)
            {
                if (metroComboBox1.Text == "" || metroComboBox2.Text == "")
                {
                    MetroMessageBox.Show(this, "Выберите нужные параметры", "Не все параметры выбраны", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    metroCheckBox1.Checked = false;
                }
                else
                {
                    FormatedTable = "SELECT  [Наименование оборудования],[Номер сварки],Дата,[Ток с учетом обр.связи]," +
                                    "[Ток с панели оператора],[Ток - ответ],[Напряжение - ответ],[Напряжение с учетом обр.связи],[Скорость проволоки],[Скорость подачи]" +
                                    "FROM  ParamArc WHERE Дата='" + metroDateTime1.Value.ToString("yyyy-MM-dd") + "' AND [Наименование оборудования]='" + metroComboBox1.Text + "' AND [Номер сварки]=" + metroComboBox2.Text;
                    metroGrid1.DataSource = SelectFromDB(FormatedTable, true, out num_Weld);
                    metroButton4.Enabled = true;
                }
            }
            else
            {
                metroGrid1.DataSource = SelectFromDB(AllTable, true, out num_Weld);
                metroButton4.Enabled = false;
            }
        }
        private DataTable SelectFromDB(string CommandText, bool returnDT, out int numWeld)
        {
            numWeld = 0;
            try
            {
                dbConnection = new SQLiteConnection("Data Source=" + Application.StartupPath + @"\THM_Arc.sqlite ;Version=3;");
                sqlCommand = new SQLiteCommand();
                dbConnection.Open();
                sqlCommand.Connection = dbConnection;

                sqlCommand.CommandText = CommandText;
                
                DataTable tb = new DataTable();
                if (returnDT)
                {
                    SQLiteDataAdapter dap = new SQLiteDataAdapter();                  
                    dap.SelectCommand = sqlCommand;                   
                    dap.Fill(tb);
                }else
                {
                    SQLiteDataReader dr = sqlCommand.ExecuteReader();
                    dr.Read();
                    numWeld = Convert.ToInt32(dr[0]);
                    tb = null;
                }

                return tb;
            }
            catch (Exception ex)
            {
                if (ex.Message != "No current row") { MetroMessageBox.Show(this, ex.Message); }
                else
                {
                    MetroMessageBox.Show(this, "Данное оборудование не выбрано или его нет в базе данных", "Не все параметры выбраны", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    metroCheckBox1.Checked = false;
                }
                return null;
            }
            finally
            {
                if (dbConnection != null)
                {
                    dbConnection.Close();
                }
             
            }
        }

        private void ExportToExcel(DataGridView DGV)
        {
            int index;
            index = 0;

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

         try
          {
            xlWorkBook  = ExcelApp.Application.Workbooks.Add(Type.Missing);
            xlWorkSheet = ExcelApp.Application.Worksheets.Add();
         
            ExcelApp.Sheets[1].Name = "Таблица параметров";
            ExcelApp.Columns.ColumnWidth = 15;
            ExcelApp.Cells[1, 1] = "Наименование оборудование";
            ExcelApp.Cells[1, 2] = "Номер сварки";
            ExcelApp.Cells[1, 3] = "Дата";
            ExcelApp.Cells[1, 4] = "Ток с учетом обратной связи";
            ExcelApp.Cells[1, 5] = "Ток с панели оператора";
            ExcelApp.Cells[1, 6] = "Ток - ответ";
            ExcelApp.Cells[1, 7] = "Напряжение ответ";
            ExcelApp.Cells[1, 8] = "Напряжение с учетом обратной связи";
            ExcelApp.Cells[1, 9] = "Скорость проволоки";
            ExcelApp.Cells[1, 10] = "Скорость подачи";       

            ExcelApp.Rows["1:1"].Columns["A:J"].Select();

            ExcelApp.Selection.ColumnWidth = 15;
            ExcelApp.Selection.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ExcelApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ExcelApp.Selection.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            ExcelApp.Selection.Borders.LineStyle = 1;
            ExcelApp.Selection.WrapText = true;
            ExcelApp.Selection.ShrinkToFit = true;

            ExcelApp.Selection.Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                              
            double jj = Convert.ToDouble(DGV.RowCount-2) / 100.0;              
            for (int j = 0; j < DGV.RowCount-1; j++)
            {
               for (int i = 0; i < DGV.ColumnCount; i++)
                {   
                    object Val = DGV[i, j].Value;
                    ExcelApp.Cells[j + 2, i + 1] = Val.ToString();
                    xlWorkSheet.Cells[j + 2, i + 1] = Val.ToString();
                    Application.DoEvents();
                }
                    metroLabel3.Text = Convert.ToString(Math.Truncate(Convert.ToDouble(j)/ jj)) + " %";
                    Application.DoEvents();
             }

            index = DGV.RowCount;
            ExcelApp.Rows["2:" + Convert.ToString(index)].Columns["A:J"].Select();

            ExcelApp.Selection.RowHeight = 30;
            ExcelApp.Selection.ColumnWidth = 15;
            ExcelApp.Selection.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ExcelApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ExcelApp.Selection.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            ExcelApp.Selection.Borders.LineStyle = 1;
            ExcelApp.Selection.WrapText = true;
            ExcelApp.Selection.ShrinkToFit = true;
            ExcelApp.Selection.Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrange;

            //****************************************************  
            Microsoft.Office.Interop.Excel.ChartObjects chartsobjrcts =
            (Microsoft.Office.Interop.Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Microsoft.Office.Interop.Excel.ChartObject chartsobjrct = chartsobjrcts.Add(900, 5, 1000, 300);
            
            chartsobjrct.Chart.ChartWizard(xlWorkSheet.get_Range("F2", "F"+Convert.ToString(index)),
            Microsoft.Office.Interop.Excel.XlChartType.xlLine, 4, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns, Type.Missing,
              0, true, "Параметры сварки", "Количество замеров", "Значения", Type.Missing);            
         //    chartsobjrct.Chart.ChartStyle = 209;
             chartsobjrct.Chart.SeriesCollection(1).Name = "Ток";
                //***************************************************
                Microsoft.Office.Interop.Excel.ChartObjects chartsobjrcts2 =
                (Microsoft.Office.Interop.Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Microsoft.Office.Interop.Excel.ChartObject chartsobjrct2 = chartsobjrcts2.Add(900, 400, 1000, 300);

                chartsobjrct2.Chart.ChartWizard(xlWorkSheet.get_Range("G2", "G" + Convert.ToString(index)),
                Microsoft.Office.Interop.Excel.XlChartType.xlLine, 4, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns, Type.Missing,
                     0, true, "Параметры сварки", "Количество замеров", "Значения", Type.Missing);
               // chartsobjrct2.Chart.ChartStyle = 209;
                chartsobjrct2.Chart.SeriesCollection(1).Name = "Напряжение";
                //****************************************************
               // xlWorkBook.SaveAs(@"C:\Users\vyastrebov\Desktop\123.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
               //     false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, 
                //    Type.Missing, Type.Missing);
                SaveFile(xlWorkBook);
                KillExcel(ExcelApp);
        
            }
            catch (Exception ee)
            {
               MessageBox.Show(ee.Message);                
            }
        }//функция экспорта в эксел

        private void Main_Load(object sender, EventArgs e)
        {
           
            string path = System.Windows.Forms.Application.StartupPath + "\\IniFile.ini";
            ini.IniFile(path);

            num_Controllers = Convert.ToInt32(ini.IniReadValue("Info", "num_Controllers"));
            name_plc = new string[num_Controllers];

           
            metroGrid1.DataSource = SelectFromDB(AllTable, true, out num_Weld);
            metroGrid1.ReadOnly = true;
            metroGrid1.Font = new Font("Segoe UI", 14f, FontStyle.Regular, GraphicsUnit.Pixel);
            for (int i = 0; i < num_Controllers; i++)
            {
                name_plc[i] = ini.IniReadValue("PLC_Names", "Name" + Convert.ToString(i + 1));
                metroComboBox1.Items.Add(name_plc[i]);
            }
           // metroComboBox1.Text = metroComboBox1.Items[0].ToString();

            //SelectFromDB(NumWeld, false, out num_Weld);
            //for (int i = 1; i < num_Weld; i++)
            //{
            //    metroComboBox2.Items.Add(i);
            //}
            //metroComboBox2.Text = metroComboBox2.Items[0].ToString();
        }

        private void Main_MouseDown(object sender, MouseEventArgs e)
        {
            dragging = true;
            dragCursorPoint = Cursor.Position;
            dragFormPoint = this.Location;
        }
        private void Main_MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point dif = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                this.Location = Point.Add(dragFormPoint, new Size(dif));
            }
        }
        private void Main_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }
        private void Ex_Bt_Click(object sender, EventArgs e)
        {
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            metroLabel2.Visible = true;
            metroLabel3.Visible = true;
            ExportToExcel(metroGrid1);
            metroLabel2.Visible = false;
            metroLabel3.Visible = false;
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }
        private void UploadDataGrid_Click(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (FillCombo(out num_Weld))
            {
                for (int i = 0; i < num_Weld; i++)
                {
                    metroComboBox2.Items.Add(i+1);
                }
                metroComboBox2.Text = metroComboBox2.Items[0].ToString();
            }
            FilterGrid();
        }
        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            FilterGrid();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
