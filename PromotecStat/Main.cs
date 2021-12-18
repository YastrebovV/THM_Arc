using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using S7.Net;
using System.Data.SQLite;
using MetroFramework;
using MetroFramework.Controls;

namespace THM_Arc
{
    public partial class Main : MetroFramework.Forms.MetroForm
    {
        public Main()
        {
            InitializeComponent();
        }
       
        
        MetroCheckBox[] cb;
        MetroLabel[] lb_name;
        MetroLabel[] lb_curr;
        MetroLabel[] lb_volt;
        MetroLabel[] lb_curr_val;
        MetroLabel[] lb_volt_val;
        MetroPanel pn1 = new MetroPanel();
        MetroPanel pn2 = new MetroPanel();
        MetroButton bt_grid = new MetroButton();
        MetroLabel lb_info_plc = new MetroLabel();
        MetroLabel lb_var_val = new MetroLabel();

        int TopCb = 30;
        int LeftCb = 10;
        int TopLbN = 30;
        int LeftLbN = 5;
        int TopLbC = 52;
        int LeftLbC = 5;
        int TopLbV = 70;
        int LeftLbV = 5;

        CpuType[] MyTypePLC;
        ErrorCode[] connectionResult;
        
        IniClass ini = new IniClass();
        Plc[] plc;

        int num_Controllers = 0;   

        string[] ip;
        short[] rack;
        short[] slot;
        string[,] param;
        int[] typePLC;
        bool[] plc_con;
        string[] name_plc;
        int[] count_weld;

        SQLiteConnection dbConnection = null;
        SQLiteCommand sqlCommand = null;

        private bool dragging = false;
        private Point dragCursorPoint;
        private Point dragFormPoint;
        Grid gr_f;

        private void connect(int num)
        {
            _con_plc(out plc[num], MyTypePLC[num], ip[num], rack[num], slot[num], false);

            if (plc[num].IsConnected)
            {            
                cb[num].Style = MetroColorStyle.Green;
                plc_con[num] = true;
            }
            else
            {
                cb[num].Checked = false;
            }
        }//еще одна функция подключения к плс, создана для упрощения кодинга
        private void disconnect(int num)
        {
            plc[num].Close();
            plc_con[num] = false;
            cb[num].Style = MetroColorStyle.Orange;
            cb[num].Checked = false;
        }

        private void cb_checkedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < num_Controllers; i++)
            {
                if (cb[i].Checked && !plc_con[i]) { connect(i); }
                if (!cb[i].Checked && plc_con[i]) { disconnect(i); }
            }
        }//метод вызываемый при чеке любого из checkbox
        private void bt_grid_click(object sender, EventArgs e)
        {
            gr_f = new Grid();
            gr_f.Show();
            this.Visible = false;
            timerGridClosed.Enabled = true;
        }//метод вызываемый при нажатии динамически созданной кнопки, вызывает форму с таблицей
        private void _con_plc(out Plc plc, CpuType MyCpuType, String MyCpuIp, short MyCpuRack, short MyCpuSlot, bool init)
        {
            plc = new Plc(MyCpuType, MyCpuIp, MyCpuRack, MyCpuSlot);

            if (!init) 
            {
                try
                {
                    plc.Open(); //попытка подключения к plc
                }
                catch (PlcException plc_ex)
                {       
                   MetroMessageBox.Show(this, plc_ex.Message, "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }  //функция подключения к плс

        private void IniLoad()
        {
            try
            {
                string path = System.Windows.Forms.Application.StartupPath + "\\IniFile.ini";
                ini.IniFile(path);

                num_Controllers = Convert.ToInt32(ini.IniReadValue("Info", "num_Controllers"));             

                ip = new string[num_Controllers];
                rack = new short[num_Controllers];
                slot = new short[num_Controllers];
                param = new string[num_Controllers, 10];
                typePLC = new int[num_Controllers];
                plc_con = new bool[num_Controllers];
                name_plc = new string[num_Controllers];
                count_weld = new int[num_Controllers];              

                cb = new MetroCheckBox[num_Controllers];
                lb_name = new MetroLabel[num_Controllers];
                lb_curr = new MetroLabel[num_Controllers];
                lb_volt = new MetroLabel[num_Controllers];
                lb_curr_val = new MetroLabel[num_Controllers];
                lb_volt_val = new MetroLabel[num_Controllers];
                plc = new Plc[num_Controllers];

                MyTypePLC = new CpuType[num_Controllers];
                connectionResult = new ErrorCode[num_Controllers];

                for (int i = 0; i < num_Controllers; i++)
                {
                    ip[i] = ini.IniReadValue("PLC" + Convert.ToString(i + 1) + "_Info", "IP");
                    rack[i] = Convert.ToInt16(ini.IniReadValue("PLC" + Convert.ToString(i + 1) + "_Info", "rack"));
                    slot[i] = Convert.ToInt16(ini.IniReadValue("PLC" + Convert.ToString(i + 1) + "_Info", "slot"));

                    for (int j = 0; j < 10; j++)
                    {
                        param[i, j] = ini.IniReadValue("PLC" + Convert.ToString(i + 1) + "_Info", "Param" + Convert.ToString(j + 1));
                    }

                    typePLC[i] = Convert.ToInt32(ini.IniReadValue("PLC" + Convert.ToString(i + 1) + "_Info", "typePLC"));

                    switch (typePLC[i])
                    {
                        case 200: MyTypePLC[i] = CpuType.S7200; break;
                        case 300: MyTypePLC[i] = CpuType.S7300; break;
                        case 400: MyTypePLC[i] = CpuType.S7400; break;
                        case 1200: MyTypePLC[i] = CpuType.S71200; break;
                        case 1500: MyTypePLC[i] = CpuType.S71500; break;
                    }

                    name_plc[i] = ini.IniReadValue("PLC_Names", "Name" + Convert.ToString(i + 1));
                   
                }
            }
            catch (Exception em)
            {
                MetroMessageBox.Show(this, em.Message);
            }
        }//загрузка данных из ини файла

        private void InsertDB(string NamePLC, int CountWeld, DateTime Time, double CurrFeedback, double OperatorPanelCurr,
           double CurrAnswer, double SpeedWire, double VoltFeedback, double VoltAnswer, double SpeedFeed)
        {
            try
            {
                dbConnection = new SQLiteConnection("Data Source=" + Application.StartupPath + @"\THM_Arc.sqlite ;Version=3;");
                sqlCommand = new SQLiteCommand();
                dbConnection.Open();
                sqlCommand.Connection = dbConnection;

                sqlCommand.CommandText = "INSERT INTO ParamArc([Наименование оборудования],[Номер сварки],Дата,[Ток с учетом обр.связи]," +
                         "[Ток с панели оператора],[Ток - ответ],[Напряжение - ответ],[Напряжение с учетом обр.связи],[Скорость проволоки],[Скорость подачи])" +
                    " VALUES('" + NamePLC + "'," + CountWeld + ",'" + Time.ToString("yyyy-MM-dd") + "', " + CurrFeedback +
                    ", " + OperatorPanelCurr + ", " + CurrAnswer + ", " + VoltAnswer + ", " + VoltFeedback + ", " + SpeedWire + ", " +
                    SpeedFeed + ")";
                sqlCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MetroMessageBox.Show(this, ex.Message);
            }
            finally
            {
                if (dbConnection != null)
                {
                    dbConnection.Close();
                }
            }
        }//добавление данных в БД

        private void Main1_Load(object sender, EventArgs e)
        {
            IniLoad();            

            pn1.Location = new Point(4, 55);
            pn1.Size = new Size(305, this.Size.Height - 65);
            pn1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            pn1.Theme = MetroThemeStyle.Dark;
            this.Controls.Add(pn1);

            pn2.Location = new Point(320, 55);
            pn2.Size = new Size(505, this.Size.Height - 65);
            pn2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            pn2.Theme = MetroThemeStyle.Dark;
            this.Controls.Add(pn2);


            bt_grid.FlatStyle = FlatStyle.Flat;
            bt_grid.Location = new Point(350, 490);   
            bt_grid.Size = new Size(150, 50);
            bt_grid.TabIndex = 43;
            bt_grid.Text = "Открыть окно таблицы";
            bt_grid.UseVisualStyleBackColor = true;
            bt_grid.Theme = MetroThemeStyle.Dark;
            bt_grid.Click += new EventHandler(bt_grid_click);
            this.Controls.Add(bt_grid);
            pn2.Controls.Add(bt_grid);

            this.lb_info_plc.AutoSize = true;
            this.lb_info_plc.Font = new Font("Century Gothic", 12f, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            this.lb_info_plc.Location = new Point(5, 5);
            this.lb_info_plc.Size = new Size(90, 20);
            this.lb_info_plc.Text = "Состояние подключения к PLC";
            this.lb_info_plc.Theme = MetroThemeStyle.Dark;
            this.Controls.Add(lb_info_plc);
            pn1.Controls.Add(lb_info_plc);

            this.lb_var_val.AutoSize = true;
            this.lb_var_val.Font = new Font("Century Gothic", 12f, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            this.lb_var_val.Location = new Point(5, 5);
            this.lb_var_val.Size = new Size(90, 20);
            this.lb_var_val.Text = "Мониторинг";
            this.lb_var_val.Theme = MetroThemeStyle.Dark;
            this.Controls.Add(lb_var_val);
            pn2.Controls.Add(lb_var_val);
         
            int countTop = 1;
            int countTopLbN = 1;
            int countTopLbC = 1;
            int countTopLbV = 1;
            for (int i = 0; i < num_Controllers; i++)
            {

                //this.cb[i]= new MetroCheckBox();
                this.cb[i].Text = name_plc[i];
                this.cb[i].Location = new Point(LeftCb, TopCb);
                this.cb[i].Size = new Size(90, 21);
                this.cb[i].Theme = MetroThemeStyle.Dark;
                this.cb[i].Style = MetroColorStyle.Orange;
                this.cb[i].UseStyleColors = true;
                this.cb[i].FontSize = MetroCheckBoxSize.Medium;
                this.Controls.Add(cb[i]);
                LeftCb += 100;
                if (countTop == 3) { countTop = 0; TopCb += 35; LeftCb = 10; }
                   countTop += 1;
                cb[i].CheckedChanged += new EventHandler(cb_checkedChanged);
                pn1.Controls.Add(cb[i]);


                this.lb_name[i] = new MetroLabel();
                this.lb_name[i].AutoSize = true;
                //this.lb_name[i].Font = new Font("Century Gothic", 10, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
                this.lb_name[i].Location = new Point(LeftLbN, TopLbN);
                this.lb_name[i].Size = new Size(90, 20);
                this.lb_name[i].Text = name_plc[i];
                this.lb_name[i].Theme = MetroThemeStyle.Dark;
                this.lb_name[i].Style = MetroColorStyle.Lime;
                this.lb_name[i].UseStyleColors = true;
                this.lb_name[i].FontSize = MetroLabelSize.Medium;
                this.Controls.Add(lb_name[i]);
                pn2.Controls.Add(lb_name[i]);

                LeftLbN += 100;
                if (countTopLbN == 5) { countTopLbN = 0; TopLbN += 65; LeftLbN = 5; }
                   countTopLbN += 1;
                //Label current
                //*********************************
                this.lb_curr[i] = new MetroLabel();
                this.lb_curr[i].AutoSize = true;
                this.lb_curr[i].Font = new Font("Century Gothic", 9, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
                this.lb_curr[i].Location = new Point(LeftLbC, TopLbC);
                this.lb_curr[i].Size = new Size(90, 20);
                this.lb_curr[i].Text = "Ток:";
                this.lb_curr[i].Theme = MetroThemeStyle.Dark;
                this.Controls.Add(lb_curr[i]);
                pn2.Controls.Add(lb_curr[i]);

                this.lb_curr_val[i] = new MetroLabel();
                this.lb_curr_val[i].AutoSize = true;
                this.lb_curr_val[i].Font = new Font("Century Gothic", 9, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
                this.lb_curr_val[i].Location = new Point(LeftLbC+50, TopLbC);
                this.lb_curr_val[i].Size = new Size(90, 20);
                this.lb_curr_val[i].Text = "0";
                this.lb_curr_val[i].Theme = MetroThemeStyle.Dark;
                this.Controls.Add(lb_curr_val[i]);
                pn2.Controls.Add(lb_curr_val[i]);

                LeftLbC += 100;
                if (countTopLbC == 5) { countTopLbC = 0; TopLbC += 65; LeftLbC = 5; }
                countTopLbC += 1;
                //*******************************************
                //Label volt
                //*******************************************
                this.lb_volt[i] = new MetroLabel();
                this.lb_volt[i].AutoSize = true;
                this.lb_volt[i].Font = new Font("Century Gothic", 9, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
                this.lb_volt[i].Location = new Point(LeftLbV, TopLbV);
                this.lb_volt[i].Size = new Size(90, 20);
                this.lb_volt[i].Text = "Вольт:";
                this.lb_volt[i].Theme = MetroThemeStyle.Dark;
                this.Controls.Add(lb_volt[i]);
                pn2.Controls.Add(lb_volt[i]);

                this.lb_volt_val[i] = new MetroLabel();
                this.lb_volt_val[i].AutoSize = true;
                this.lb_volt_val[i].Font = new Font("Century Gothic", 9, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
                this.lb_volt_val[i].Location = new Point(LeftLbV+50, TopLbV);
                this.lb_volt_val[i].Size = new Size(90, 20);
                this.lb_volt_val[i].Text = "0";
                this.lb_volt_val[i].Theme = MetroThemeStyle.Dark;
                this.Controls.Add(lb_volt_val[i]);
                pn2.Controls.Add(lb_volt_val[i]);

                LeftLbV += 100;
                if (countTopLbV == 5) { countTopLbV = 0; TopLbV += 65; LeftLbV = 5; }
                countTopLbV += 1;
                //************************************

            }
            //инициализирует объекты plc
            for (int j = 0; j < num_Controllers; j++)
            {
                count_weld[j] = 0;
                _con_plc(out plc[j], MyTypePLC[j], ip[j], rack[j], slot[j], true);
            }

            timer_Read_PLC.Enabled = true;
            checkPlc.Enabled = true;          
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
        private void timer_Read_PLC_Tick(object sender, EventArgs e)
        {
            timer_Read_PLC.Enabled = false;

           try { 
               for (int i = 0; i < num_Controllers; i++)
               {
                if (plc[i].IsConnected)
                {
                    bool start = Convert.ToBoolean(plc[i].Read(param[i, 0]));
                    bool startOK = Convert.ToBoolean(plc[i].Read(param[i, 1]));
                    if (start && startOK)
                    {
                        double CurrFeedBack = Math.Round(((uint)plc[i].Read(param[i, 2])).ConvertToDouble(), 2);
                        double OperatorPanel = Math.Round(((uint)plc[i].Read(param[i, 3])).ConvertToDouble(), 2);
                        double CurrAnswer = Math.Round(((uint)plc[i].Read(param[i, 4])).ConvertToDouble(), 2);
                        lb_curr_val[i].Text = Convert.ToString(CurrAnswer);
                        double SpeedWire = Convert.ToDouble(plc[i].Read(param[i, 5]));
                        double VoltFeedBack = Math.Round(((uint)plc[i].Read(param[i, 6])).ConvertToDouble(), 2);
                        double VoltAnswer = Math.Round(((uint)plc[i].Read(param[i, 7])).ConvertToDouble(), 2);
                        lb_volt_val[i].Text = Convert.ToString(VoltAnswer);
                        double SpeedFeed = Math.Round(((uint)plc[i].Read(param[i, 8])).ConvertToDouble(), 2);//Convert.ToDouble(plc[i].Read(param[i, 9]));
                        count_weld[i] = Convert.ToInt32(plc[i].Read(param[i, 9]));

                        InsertDB(name_plc[i], count_weld[i], DateTime.Now, CurrFeedBack, OperatorPanel, CurrAnswer, SpeedWire, VoltFeedBack, VoltAnswer, SpeedFeed);
                    }
                }
            }
            }
            catch
            {

            }
            finally
            {
                timer_Read_PLC.Enabled = true;
            }
          
            timer_Read_PLC.Enabled = true;
        }//таймер проверяющий состояние сварки. Если сварка активна происходит запись параметров
        private void checkPlc_Tick(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < num_Controllers; i++)
                {
                    if (!plc[i].IsConnected)
                    {
                        plc_con[i] = false;
                        cb[i].Style = MetroColorStyle.Orange;
                        cb[i].Checked = false;
                    }
                }
            }
            catch
            {

            }
            finally
            {
                timer_Read_PLC.Enabled = true;
            }

        }//проверка состояния PLC

        private void metroButton1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void timerGridClosed_Tick(object sender, EventArgs e)
        {
            if (gr_f.IsDisposed)
            {
                this.Visible = true;
                timerGridClosed.Enabled = false;
            }
        }
    }
}
