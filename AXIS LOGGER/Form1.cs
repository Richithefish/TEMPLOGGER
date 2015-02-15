using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Globalization;
using System.Threading;
using System.Net.Sockets;
using System.Net;
using System.Timers;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

//using System.Net.Mail;
//using System.Net.Mime;
//using Outlook = Microsoft.Office.Interop.Outlook;

namespace AXIS_LOGGER
{
    public partial class MAIN_FORM : Form
    {
        
        //Main Form Variables
        string[] COM_Portar;
        long exposure_Time;
        decimal Timer_Counter;
        bool turn_off = false, turn_off_chamber = false, temp_Reached = false, temp_Reached1 = false, temp_Reached2 = false, temp_Reached3 = false, temp_Reached4 = false, continuous = false, time1_manual_reached = false, time2_manual_reached = false, time3_manual_reached = false, time4_manual_reached = false, time_60_reached = false, timer_start = false, cycle_Comp = false, negativ1 = false, negativ2 = false;
        public Thread PICO1_Thread, PICO2_Thread;
        public short status1, status2;
        public float[] tempbuffer1 = new float[9], tempbuffer2 = new float[9], TempChannel1 = new float[10], TempChannel2 = new float[10], TempChannel3 = new float[10], TempChannel4 = new float[10], TempChannel5 = new float[10], TempChannel6 = new float[10], TempChannel7 = new float[10], TempChannel8 = new float[10], TempChannel9 = new float[10], TempChannel10 = new float[10], TempChannel11 = new float[10], TempChannel12 = new float[10], TempChannel13 = new float[10], TempChannel14 = new float[10], TempChannel15 = new float[10], TempChannel16 = new float[10];
        public float[] Tempcheck1 = new float[2], Tempcheck2 = new float[2], Tempcheck3 = new float[2], Tempcheck4 = new float[2], Tempcheck5 = new float[2], Tempcheck6 = new float[2], Tempcheck7 = new float[2], Tempcheck8 = new float[2], Tempcheck9 = new float[2], Tempcheck10 = new float[2], Tempcheck11 = new float[2], Tempcheck12 = new float[2], Tempcheck13 = new float[2], Tempcheck14 = new float[2], Tempcheck15 = new float[2], Tempcheck16 = new float[2];
        public short overflow;
        public int _step = 1, _procent = 0, stable_counter = 0, Chan_Temp_Count = 0, get_10_count = 0, comp_i = 0;
        public static string IPnr1, IPnr2;

        //Report Variables
        public Excel.Application excelApp = null;
        public Excel.Workbook newBook;

        //CTS Variables
        System.Windows.Forms.DataVisualization.Charting.Series ChartSeries1 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series ChartSeries2 = new System.Windows.Forms.DataVisualization.Charting.Series();
        DataTable Data_Table_CTS = new DataTable();
        DataColumn Time_Column_CTS = new DataColumn();
        DataColumn act_Column = new DataColumn();
        DataColumn set_Column = new DataColumn();
        DataColumn[] PrimaryKeyColumn_CTS = new DataColumn[1];
        DataRow row_CTS;
        byte[] data_Tx = new byte[64];
        byte[] data_Rx = new byte[64];
        int Rx_buf = 0;
        public double[] ActTemp = new double[864000];
        public double[] SetTemp = new double[864000];
        public int check_Set;

        //PICO variables
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel1 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel2 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel3 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel4 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel5 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel6 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel7 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel8 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel9 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel10 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel11 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel12 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel13 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel14 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel15 = new System.Windows.Forms.DataVisualization.Charting.Series();
        System.Windows.Forms.DataVisualization.Charting.Series series_Channel16 = new System.Windows.Forms.DataVisualization.Charting.Series();
        DataTable Data_Table_PICO_Comp = new DataTable();
        DataTable Data_Table_PICO_Manual = new DataTable();
        DataColumn Time_Column_PICO_Comp = new DataColumn();
        DataColumn Time_Column_PICO_Manual = new DataColumn();
        DataColumn Channel1_Column = new DataColumn();
        DataColumn Channel2_Column = new DataColumn();
        DataColumn Channel3_Column = new DataColumn();
        DataColumn Channel4_Column = new DataColumn();
        DataColumn Channel5_Column = new DataColumn();
        DataColumn Channel6_Column = new DataColumn();
        DataColumn Channel7_Column = new DataColumn();
        DataColumn Channel8_Column = new DataColumn();
        DataColumn Channel9_Column = new DataColumn();
        DataColumn Channel10_Column = new DataColumn();
        DataColumn Channel11_Column = new DataColumn();
        DataColumn Channel12_Column = new DataColumn();
        DataColumn Channel13_Column = new DataColumn();
        DataColumn Channel14_Column = new DataColumn();
        DataColumn Channel15_Column = new DataColumn();
        DataColumn Channel16_Column = new DataColumn();
        DataColumn[] PrimaryKeyColumn_PICO = new DataColumn[1];
        DataRow row_PICO;
        public short handle1;
        public short handle2;
        int[] handle_array = new int[8];
        private short _handle1;
        private short _handle2;
        public const int USBTC08_MAX_CHANNELS = 8;
        public const char TC_TYPE_K = 'K';
        public const int PICO_OK = 1;
        string no_of_units;
        int PICO_Counter = 0;

        //Test variables
        
        
        //Timer Variables
        System.Timers.Timer timer_expose = new System.Timers.Timer();
        System.Timers.Timer timer_expose_60 = new System.Timers.Timer(60000);
        System.Timers.Timer timer_expose_manual1 = new System.Timers.Timer();
        System.Timers.Timer timer_expose_manual2 = new System.Timers.Timer();
        System.Timers.Timer timer_expose_manual3 = new System.Timers.Timer();
        System.Timers.Timer timer_expose_manual4 = new System.Timers.Timer();
        System.Timers.Timer Time_Logger = new System.Timers.Timer(1000);
        Stopwatch stopWatch = new Stopwatch();
        Stopwatch logger_watch = new Stopwatch();
        Stopwatch Stop_CTS_watch = new Stopwatch();
        
        //Init Main Form
        public MAIN_FORM()
        {
            InitializeComponent();
        }

        //On Main Form load, set Communication Ports, static CTS&PICO values and chart info. Create new elapsed events on timers and get portNames.
        private void MAIN_FORM_Load(object sender, EventArgs e)
        {
            COM_Portar = SerialPort.GetPortNames();
            ComPort_Box.DataSource = COM_Portar;
            Get_PICO_Unit();

            //CTS Static values
            Time_Column_CTS.DataType = Type.GetType("System.Double");
            Time_Column_CTS.ColumnName = "Time";
            Data_Table_CTS.Columns.Add(Time_Column_CTS);
            act_Column.DataType = Type.GetType("System.Double");
            act_Column.ColumnName = "Actual";
            Data_Table_CTS.Columns.Add(act_Column);
            set_Column.DataType = Type.GetType("System.Double");
            set_Column.ColumnName = "Set";
            Data_Table_CTS.Columns.Add(set_Column);
            Data_Table_PICO_Comp.Columns.Add(Time_Column_PICO_Comp);
            Data_Table_PICO_Manual.Columns.Add(Time_Column_PICO_Manual);
            chart_CTS_Comp.ChartAreas[0].AxisX.Title = "Time (Sec)";
            chart_CTS_Comp.ChartAreas[0].AxisY.Title = "Temperature (°C)";
            chart_CTS_Comp.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_CTS_Comp.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_CTS_Manual.ChartAreas[0].AxisX.Title = "Time (Sec)";
            chart_CTS_Manual.ChartAreas[0].AxisY.Title = "Temperature (°C)";
            chart_CTS_Manual.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_CTS_Manual.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;

            //PICO static values
            Time_Column_PICO_Comp.DataType = Type.GetType("System.Double");
            Time_Column_PICO_Comp.ColumnName = "Time";
            Time_Column_PICO_Manual.DataType = Type.GetType("System.Double");
            Time_Column_PICO_Manual.ColumnName = "Time";
            Channel1_Column.DataType = Type.GetType("System.Double");
            Channel2_Column.DataType = Type.GetType("System.Double");
            Channel3_Column.DataType = Type.GetType("System.Double");
            Channel4_Column.DataType = Type.GetType("System.Double");
            Channel5_Column.DataType = Type.GetType("System.Double");
            Channel6_Column.DataType = Type.GetType("System.Double");
            Channel7_Column.DataType = Type.GetType("System.Double");
            Channel8_Column.DataType = Type.GetType("System.Double");
            Channel9_Column.DataType = Type.GetType("System.Double");
            Channel10_Column.DataType = Type.GetType("System.Double");
            Channel11_Column.DataType = Type.GetType("System.Double");
            Channel12_Column.DataType = Type.GetType("System.Double");
            Channel13_Column.DataType = Type.GetType("System.Double");
            Channel14_Column.DataType = Type.GetType("System.Double");
            Channel15_Column.DataType = Type.GetType("System.Double");
            Channel16_Column.DataType = Type.GetType("System.Double");
            chart_PICO_Comp.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_PICO_Comp.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_PICO_Comp.ChartAreas[0].AxisX.Title = "Time (Sec)";
            chart_PICO_Comp.ChartAreas[0].AxisY.Title = "Temperature (°C)";
            chart_PICO_Manual.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_PICO_Manual.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_PICO_Manual.ChartAreas[0].AxisX.Title = "Time (Sec)";
            chart_PICO_Manual.ChartAreas[0].AxisY.Title = "Temperature (°C)";

            //Timer elapsed events
            timer_expose.Elapsed += timer_expose_Elapsed;
            timer_expose_manual1.Elapsed += timer_expose_manual1_Elapsed;
            timer_expose_manual2.Elapsed += timer_expose_manual2_Elapsed;
            timer_expose_manual3.Elapsed += timer_expose_manual3_Elapsed;
            timer_expose_manual4.Elapsed += timer_expose_manual4_Elapsed;
            timer_expose_60.Elapsed += timer_expose_60_Elapsed;
        }

        //Timer_expose_60_Elapsed = On enabled: After 60 sec, set bool values for controling statements in timer_Comp and disable timer after time elapsed.
        void timer_expose_60_Elapsed(object sender, ElapsedEventArgs e)
        {
            time_60_reached = true;
            timer_start = false;
            timer_expose_60.Enabled = false;
        }

        //timer_expose_manual4_Elapsed = On enabled: Exposure_time_box_4_Manual sets the timer interval. @ timer elapsed event, set bool value for controling statements in main timer
        void timer_expose_manual4_Elapsed(object sender, ElapsedEventArgs e)
        {
            time4_manual_reached = true;
        }

        //timer_expose_manual3_Elapsed = On enabled: Exposure_time_box_3_Manual sets the timer interval. @ timer elapsed event, set bool value for controling statements in main timer
        void timer_expose_manual3_Elapsed(object sender, ElapsedEventArgs e)
        {
            time3_manual_reached = true;
        }

        //timer_expose_manual2_Elapsed = On enabled: Exposure_time_box_2_Manual sets the timer interval. @ timer elapsed event, set bool value for controling statements in main timer
        void timer_expose_manual2_Elapsed(object sender, ElapsedEventArgs e)
        {
            time2_manual_reached = true;
        }

        //timer_expose_manual1_Elapsed = On enabled: Exposure_time_box_1_Manual sets the timer interval. @ timer elapsed event, set bool value for controling statements in main timer
        void timer_expose_manual1_Elapsed(object sender, ElapsedEventArgs e)
        {
            time1_manual_reached = true;
        }

        //timer_expose_Elapsed = On enabled: Exposure_time1_Comp box and Exposure_time2_Comp box sets the timer interval. @ timer elapsed event, set bool value for controling statements in main timer. Disable timer after time elapsed.
        void timer_expose_Elapsed(object sender, ElapsedEventArgs e)
        {
            turn_off_chamber = true;
            timer_expose.Enabled = false;
        }

        //Start Temperature on Components button
        private void Start_Btn_Comp_Click(object sender, EventArgs e)
        {
            string start = string.Empty;
            string set_etemp = string.Empty;
            bool empty_string = false;
            //Clear Rx/Tx data
            for (int i = 0; i < data_Rx.Length; i++)
            {
                data_Rx[i] = 0;
                data_Tx[i] = 0;
            }
            //Check if any program is running
            if (timer_Manual.Enabled == false && timer_Comp_Exposure.Enabled == false)
            {
                if (Set_Analog_Box2_Comp.Visible == true && Set_Analog_Box2_Comp.Text == string.Empty)
                    empty_string = true;
                if (Set_Analog_Box1_Comp.Text != string.Empty && empty_string != true)
                {
                    Notify_Label.Text = "-";
                    //Set exposure time if exposure time radiobutton is selected
                    if (Automatic_ON_Comp.Checked == false)
                        exposure_Time = long.Parse(Exposure_Time1_Comp.Text);
                    //Check if ethernet is not connected to CTS chamber, use serialport instead
                    if (Ethernet_label.Text != "Connected")
                    {
                        if (COM_Portar.Length > 0)//Check if there are any serialports
                        {
                            if (chart_CTS_Comp.Series.Count < 1)//Check if CTS chart has any series data. If not, add chartseries with defined color.
                            {
                                this.chart_CTS_Comp.Series.Add(ChartSeries1);
                                this.chart_CTS_Comp.Series.Add(ChartSeries2);
                                ChartSeries1.Color = System.Drawing.Color.Green;
                                ChartSeries2.Color = System.Drawing.Color.Red;
                            }
                            serialPort_CTS.PortName = ComPort_Box.SelectedItem.ToString();//Select serialport
                            serialPort_CTS.Open();                                        //And open it
                            if (serialPort_CTS.IsOpen)
                            {
                                //Set Tempeture on Chamber
                                check_Set = SET_Temp_CTS(Set_Analog_Box1_Comp.Text);
                                //Start chamber
                                Start_CTS();
                            }


                        }
                        else
                            Notify_Label.Text = "No COM Port! ";
                    }
                    //Check if ethernet is connected to CTS chamber, use ethernet communication
                    if (Ethernet_label.Text == "Connected")
                    {
                        if (chart_CTS_Comp.Series.Count < 1)
                        {
                            this.chart_CTS_Comp.Series.Add(ChartSeries1);
                            this.chart_CTS_Comp.Series.Add(ChartSeries2);
                            ChartSeries1.Color = System.Drawing.Color.Green;
                            ChartSeries2.Color = System.Drawing.Color.Red;
                        }
                        if (Set_Analog_Box1_Comp.Text.Length == 2)
                        {
                            set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box1_Comp.Text + ".0");
                        }
                        else
                            set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box1_Comp.Text + ".0");
                        Thread.Sleep(100);
                        start = Read_Write_Ethernet(Ethernet_Box.Text, "s1 1");
                    }
                    //Predefine variables
                    label_Comp_Finished.Visible = false;
                    stopWatch.Reset();
                    PICO_Counter = 0;
                    Timer_Counter = 0;
                    stable_counter = 0;
                    Exposure_Time_Label1.Text = "-";
                    Set_analog_minus_10_Label1.Text = "-";
                    Exposure_Time_Label2.Text = "-";
                    Set_analog_minus_10_Label2.Text = "-";
                    turn_off = false;
                    turn_off_chamber = false;
                    temp_Reached = false;
                    cycle_Comp = false;
                    timer_start = false;
                    empty_string = false;
                    //disable selections (fault protection)
                    disable_selections();
                    Set_Analog_Box1_Comp.ReadOnly = true;
                    Set_Analog_Box2_Comp.ReadOnly = true;
                    Exposure_Time1_Comp.ReadOnly = true;
                    Exposure_Time2_Comp.ReadOnly = true;
                    //Set CTS/PICO chart parameters
                    SET_Chart_Grid();
                    //Set PICO channels to log
                    SetChannels();
                    //Check if Automatic or Exposure radiobutton is selected and start Main timer
                    if (Automatic_ON_Comp.Checked == true)
                        timer_Comp_Automatic.Enabled = true;
                    else if (Exposure_ON_Comp.Checked == true)
                        timer_Comp_Exposure.Enabled = true;
                }
                else
                    Notify_Label.Text = "Empty SET temperature value";
            }
            else
                Notify_Label.Text = "Another test is running!";
        }
        //Stop Temperature on Components button
        private void Stop_Btn_Comp_Click(object sender, EventArgs e)
        {
            string stop = string.Empty;
            timer_Comp_Exposure.Enabled = false;//Disable Main timer
            timer_Comp_Automatic.Enabled = false;//Disable Main timer
            //Check if ethernet is not connected to CTS chamber, use serialport to stop chamber 
            if (Ethernet_label.Text != "Connected")
            {
                //Send serialport commands to stop CTS chamber and close serialport
                serialPort_CTS.ReceivedBytesThreshold = 8;
                data_Tx[0] = Convert.ToByte("0x02", 16);
                data_Tx[1] = Convert.ToByte("0x81", 16);
                data_Tx[2] = Convert.ToByte("0xF0", 16);
                data_Tx[3] = Convert.ToByte("0xB0", 16);
                data_Tx[4] = Convert.ToByte("0xB0", 16);
                data_Tx[5] = Convert.ToByte("0xB0", 16);
                data_Tx[6] = Convert.ToByte("0xC1", 16);
                data_Tx[7] = Convert.ToByte("0x03", 16);
                serialPort_CTS.Write(data_Tx, 0, 8);
                serialPort_CTS.Close();
            }
            //Check if ethernet is connected to CTS chamber, use ethernet communication to stop chamber
            if(Ethernet_label.Text == "Connected")
                stop = Read_Write_Ethernet(Ethernet_Box.Text, "s1 0");
            timer_expose.Enabled = false;//Disable timer
            timer_expose_60.Enabled = false;//Disable timer
            //enable_ChannelBox();//Enable PICO channel checkboxes
            enable_selections();
            Set_Analog_Box1_Comp.ReadOnly = false;
            Set_Analog_Box2_Comp.ReadOnly = false;
            Exposure_Time1_Comp.ReadOnly = false;
            Exposure_Time2_Comp.ReadOnly = false;
        }
        //Get all PICO loggers (1 or 2). Using PICO technologies dll file for handling units, function example by PICO.
        private void Get_PICO_Unit()
        {
            System.Text.StringBuilder line1 = new System.Text.StringBuilder(256);
            System.Text.StringBuilder line2 = new System.Text.StringBuilder(256);
            int i, new_handle;
            for (i = 0; (new_handle = TC08example.Imports.TC08OpenUnit()) > 0; i++)
            {
                // store the handle in an array
                handle_array[i] = new_handle;
               
            }
            handle1 = (short)handle_array[0];
            no_of_units = i.ToString();
            Nbr_Pico_Units.Text = no_of_units;
            if (i > 0)
            {
                PICO_Unit_1_gb.Visible = true;
                textBox_Unit1.Visible = true;
                handle1 = (short)handle_array[0];
                _handle1_to_handle1(handle1);
                TC08example.Imports.TC08GetFormattedInfo(handle1, line1, 256);
                textBox_Unit1.Text = line1.ToString().Substring(101, 18).TrimStart();
            }
            if (i > 1)
            {
                PICO_Unit_2_gb.Visible = true;
                textBox_Unit2.Visible = true;
                handle2 = (short)handle_array[1];
                _handle2_to_handle2(handle2);
                TC08example.Imports.TC08GetFormattedInfo(handle2, line2, 256);
                textBox_Unit2.Text = line2.ToString().Substring(101, 18).TrimStart();
            }
        }
        //Save handle to public variable _handle1
        public void _handle1_to_handle1(short handle)
        {
            _handle1 = handle;
        }
        //Save handle to public variable _handle2
        public void _handle2_to_handle2(short handle)
        {
            _handle2 = handle;
        }
        //GetChannel_value_Unit1 = Gets all the checked PICO channels values from checkbox menu (Unit 1) and adds the values to the PICO data table and chart series. Tempbuffer[0] is the embedded temp in the logger unit (Not used)
        unsafe void GetChannel_Value_Unit1()
        {
            short status;
            float[] tempbuffer = new float[9];
            short overflow;
            
            status = TC08example.Imports.TC08GetSingle(_handle1, tempbuffer, &overflow, TC08example.Imports.TempUnit.USBTC08_UNITS_CENTIGRADE);
            if (status == PICO_OK)
            {
                if (Channel1_Box.CheckState == CheckState.Checked)
                {
                    series_Channel1.Points.AddXY(Timer_Counter, tempbuffer[1]);
                    row_PICO[series_Channel1.Name] = (decimal)tempbuffer[1];
                }
                if (Channel2_Box.CheckState == CheckState.Checked)
                {
                    series_Channel2.Points.AddXY(Timer_Counter, tempbuffer[2]);
                    row_PICO[series_Channel2.Name] = (decimal)tempbuffer[2];
                }
                if (Channel3_Box.CheckState == CheckState.Checked)
                {
                    series_Channel3.Points.AddXY(Timer_Counter, tempbuffer[3]);
                    row_PICO[series_Channel3.Name] = (decimal)tempbuffer[3];
                }
                if (Channel4_Box.CheckState == CheckState.Checked)
                {
                    series_Channel4.Points.AddXY(Timer_Counter, tempbuffer[4]);
                    row_PICO[series_Channel4.Name] = (decimal)tempbuffer[4];
                }
                if (Channel5_Box.CheckState == CheckState.Checked)
                {
                    series_Channel5.Points.AddXY(Timer_Counter, tempbuffer[5]);
                    row_PICO[series_Channel5.Name] = (decimal)tempbuffer[5];
                }
                if (Channel6_Box.CheckState == CheckState.Checked)
                {
                    series_Channel6.Points.AddXY(Timer_Counter, tempbuffer[6]);
                    row_PICO[series_Channel6.Name] = (decimal)tempbuffer[6];
                }
                if (Channel7_Box.CheckState == CheckState.Checked)
                {
                    series_Channel7.Points.AddXY(Timer_Counter, tempbuffer[7]);
                    row_PICO[series_Channel7.Name] = (decimal)tempbuffer[7];
                }
                if (Channel8_Box.CheckState == CheckState.Checked)
                {
                    series_Channel8.Points.AddXY(Timer_Counter, tempbuffer[8]);
                    row_PICO[series_Channel8.Name] = (decimal)tempbuffer[8];
                }
                
            }
            else
                Notify_Label.Text = "Status: " + status;
        }
        //GetChannel_value_Unit2 = Gets all the checked PICO channels values from checkbox menu (Unit 2) and adds the values to the PICO data table and chart series. Tempbuffer[0] is the embedded temp in the logger unit (Not used)
        unsafe void GetChannel_Value_Unit2()
        {
            short status;
            float[] tempbuffer = new float[9];
            short overflow;

            status = TC08example.Imports.TC08GetSingle(_handle2, tempbuffer, &overflow, TC08example.Imports.TempUnit.USBTC08_UNITS_CENTIGRADE);
            if (status == PICO_OK)
            {
               
                if (Channel9_Box.CheckState == CheckState.Checked)
                {
                    series_Channel9.Points.AddXY(Timer_Counter, tempbuffer[1]);
                    row_PICO[series_Channel9.Name] = (decimal)tempbuffer[1];
                }
                if (Channel10_Box.CheckState == CheckState.Checked)
                {
                    series_Channel10.Points.AddXY(Timer_Counter, tempbuffer[2]);
                    row_PICO[series_Channel10.Name] = (decimal)tempbuffer[2];
                }
                if (Channel11_Box.CheckState == CheckState.Checked)
                {
                    series_Channel11.Points.AddXY(Timer_Counter, tempbuffer[3]);
                    row_PICO[series_Channel11.Name] = (decimal)tempbuffer[3];
                }
                if (Channel12_Box.CheckState == CheckState.Checked)
                {
                    series_Channel12.Points.AddXY(Timer_Counter, tempbuffer[4]);
                    row_PICO[series_Channel12.Name] = (decimal)tempbuffer[4];
                }
                if (Channel13_Box.CheckState == CheckState.Checked)
                {
                    series_Channel13.Points.AddXY(Timer_Counter, tempbuffer[5]);
                    row_PICO[series_Channel13.Name] = (decimal)tempbuffer[5];
                }
                if (Channel14_Box.CheckState == CheckState.Checked)
                {
                    series_Channel14.Points.AddXY(Timer_Counter, tempbuffer[6]);
                    row_PICO[series_Channel14.Name] = (decimal)tempbuffer[6];
                }
                if (Channel15_Box.CheckState == CheckState.Checked)
                {
                    series_Channel15.Points.AddXY(Timer_Counter, tempbuffer[7]);
                    row_PICO[series_Channel15.Name] = (decimal)tempbuffer[7];
                }
                if (Channel16_Box.CheckState == CheckState.Checked)
                {
                    series_Channel16.Points.AddXY(Timer_Counter, tempbuffer[8]);
                    row_PICO[series_Channel16.Name] = (decimal)tempbuffer[8];
                }
                
            }
            else
                Notify_Label.Text = "Status: " + status;
        }
        //SetChannels = Sets all the checked PICO channels values from checkbox menu (Unit 2) and adds the channel name to the PICO data table and chart series + tooltip in chart.
        void SetChannels()
        {
            short channel;
            short ok;

            if (Channel1_Box.CheckState == CheckState.Checked)
            {
                channel = 1;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel1.Name = textBox_Channel1.Text;
                if (series_Channel1.Name == "")
                    series_Channel1.Name = "Channel1";
                series_Channel1.ToolTip = series_Channel1.Name;
                Channel1_Column.ColumnName = series_Channel1.Name;
            }

            if (Channel2_Box.CheckState == CheckState.Checked)
            {
                channel = 2;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel2.Name = textBox_Channel2.Text;
                if (series_Channel2.Name == "")
                    series_Channel2.Name = "Channel2";
                series_Channel2.ToolTip = series_Channel2.Name;
                Channel2_Column.ColumnName = series_Channel2.Name;
            }

            if (Channel3_Box.CheckState == CheckState.Checked)
            {
                channel = 3;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel3.Name = textBox_Channel3.Text;
                if (series_Channel3.Name == "")
                    series_Channel3.Name = "Channel3";
                series_Channel3.ToolTip = series_Channel3.Name;
                Channel3_Column.ColumnName = series_Channel3.Name;
            }

            if (Channel4_Box.CheckState == CheckState.Checked)
            {
                channel = 4;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel4.Name = textBox_Channel4.Text;
                if (series_Channel4.Name == "")
                    series_Channel4.Name = "Channel4";
                series_Channel4.ToolTip = series_Channel4.Name;
                Channel4_Column.ColumnName = series_Channel4.Name;
            }

            if (Channel5_Box.CheckState == CheckState.Checked)
            {
                channel = 5;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel5.Name = textBox_Channel5.Text;
                if (series_Channel5.Name == "")
                    series_Channel5.Name = "Channel5";
                series_Channel5.ToolTip = series_Channel5.Name;
                Channel5_Column.ColumnName = series_Channel5.Name;
            }

            if (Channel6_Box.CheckState == CheckState.Checked)
            {
                channel = 6;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel6.Name = textBox_Channel6.Text;
                if (series_Channel6.Name == "")
                    series_Channel6.Name = "Channel6";
                series_Channel6.ToolTip = series_Channel6.Name;
                Channel6_Column.ColumnName = series_Channel6.Name;
            }

            if (Channel7_Box.CheckState == CheckState.Checked)
            {
                channel = 7;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel7.Name = textBox_Channel7.Text;
                if (series_Channel7.Name == "")
                    series_Channel7.Name = "Channel7";
                series_Channel7.ToolTip = series_Channel7.Name;
                Channel7_Column.ColumnName = series_Channel7.Name;
            }

            if (Channel8_Box.CheckState == CheckState.Checked)
            {
                channel = 8;
                ok = TC08example.Imports.TC08SetChannel(_handle1, channel, TC_TYPE_K);
                series_Channel8.Name = textBox_Channel8.Text;
                if (series_Channel8.Name == "")
                    series_Channel8.Name = "Channel8";
                series_Channel8.ToolTip = series_Channel8.Name;
                Channel8_Column.ColumnName = series_Channel8.Name;
            }

            if (Channel9_Box.CheckState == CheckState.Checked)
            {
                channel = 1;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel9.Name = textBox_Channel9.Text;
                if (series_Channel9.Name == "")
                    series_Channel9.Name = "Channel9";
                series_Channel9.ToolTip = series_Channel9.Name;
                
                Channel9_Column.ColumnName = series_Channel9.Name;
            }

            if (Channel10_Box.CheckState == CheckState.Checked)
            {
                channel = 2;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel10.Name = textBox_Channel10.Text;
                if (series_Channel10.Name == "")
                    series_Channel10.Name = "Channel10";
                series_Channel10.ToolTip = series_Channel10.Name;
                Channel10_Column.ColumnName = series_Channel10.Name;
            }

            if (Channel11_Box.CheckState == CheckState.Checked)
            {
                channel = 3;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel11.Name = textBox_Channel11.Text;
                if (series_Channel11.Name == "")
                    series_Channel11.Name = "Channel11";
                series_Channel11.ToolTip = series_Channel11.Name;
                Channel11_Column.ColumnName = series_Channel11.Name;
            }

            if (Channel12_Box.CheckState == CheckState.Checked)
            {
                channel = 4;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel12.Name = textBox_Channel12.Text;
                if (series_Channel12.Name == "")
                    series_Channel12.Name = "Channel12";
                series_Channel12.ToolTip = series_Channel12.Name;
                Channel12_Column.ColumnName = series_Channel12.Name;
            }

            if (Channel13_Box.CheckState == CheckState.Checked)
            {
                channel = 5;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel13.Name = textBox_Channel13.Text;
                if (series_Channel13.Name == "")
                    series_Channel13.Name = "Channel13";
                series_Channel13.ToolTip = series_Channel13.Name;
                Channel13_Column.ColumnName = series_Channel13.Name;
            }

            if (Channel14_Box.CheckState == CheckState.Checked)
            {
                channel = 6;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel14.Name = textBox_Channel14.Text;
                if (series_Channel14.Name == "")
                    series_Channel14.Name = "Channel14";
                series_Channel14.ToolTip = series_Channel14.Name;
                Channel14_Column.ColumnName = series_Channel14.Name;
            }

            if (Channel15_Box.CheckState == CheckState.Checked)
            {
                channel = 7;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel15.Name = textBox_Channel15.Text;
                if (series_Channel15.Name == "")
                    series_Channel15.Name = "Channel15";
                series_Channel15.ToolTip = series_Channel15.Name;
                Channel15_Column.ColumnName = series_Channel15.Name;
            }

            if (Channel16_Box.CheckState == CheckState.Checked)
            {
                channel = 8;
                ok = TC08example.Imports.TC08SetChannel(_handle2, channel, TC_TYPE_K);
                series_Channel16.Name = textBox_Channel16.Text;
                if (series_Channel16.Name == "")
                    series_Channel16.Name = "Channel16";
                series_Channel16.ToolTip = series_Channel16.Name;
                Channel16_Column.ColumnName = series_Channel16.Name;
            }
        }
        //Set_Chart_Grid = Sets the charts in a predefined structure and clears prior data if there are any.
        void SET_Chart_Grid()
        {

            chart_PICO_Comp.ChartAreas[0].AxisX.Minimum = 0;
            chart_CTS_Comp.ChartAreas[0].AxisX.Minimum = 0;
            chart_PICO_Manual.ChartAreas[0].AxisX.Minimum = 0;
            chart_CTS_Manual.ChartAreas[0].AxisX.Minimum = 0;
            

            ChartSeries1.ChartArea = "ChartArea_CTS";
            ChartSeries1.Legend = "Legend1";
            ChartSeries1.Name = "Actual_value";
            ChartSeries2.ChartArea = "ChartArea_CTS";
            ChartSeries2.Legend = "Legend1";
            ChartSeries2.Name = "Set_value";
            ChartSeries1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            ChartSeries2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel5.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel6.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel7.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel8.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel9.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel10.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel11.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel12.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel13.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel14.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel15.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel16.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;
            series_Channel1.ChartArea = "ChartArea_PICO";
            series_Channel1.Legend = "Legend1";

            series_Channel2.ChartArea = "ChartArea_PICO";
            series_Channel2.Legend = "Legend1";

            series_Channel3.ChartArea = "ChartArea_PICO";
            series_Channel3.Legend = "Legend1";

            series_Channel4.ChartArea = "ChartArea_PICO";
            series_Channel4.Legend = "Legend1";

            series_Channel5.ChartArea = "ChartArea_PICO";
            series_Channel5.Legend = "Legend1";

            series_Channel6.ChartArea = "ChartArea_PICO";
            series_Channel6.Legend = "Legend1";

            series_Channel7.ChartArea = "ChartArea_PICO";
            series_Channel7.Legend = "Legend1";

            series_Channel8.ChartArea = "ChartArea_PICO";
            series_Channel8.Legend = "Legend1";

            series_Channel9.ChartArea = "ChartArea_PICO";
            series_Channel9.Legend = "Legend1";

            series_Channel10.ChartArea = "ChartArea_PICO";
            series_Channel10.Legend = "Legend1";

            series_Channel11.ChartArea = "ChartArea_PICO";
            series_Channel11.Legend = "Legend1";

            series_Channel12.ChartArea = "ChartArea_PICO";
            series_Channel12.Legend = "Legend1";

            series_Channel13.ChartArea = "ChartArea_PICO";
            series_Channel13.Legend = "Legend1";

            series_Channel14.ChartArea = "ChartArea_PICO";
            series_Channel14.Legend = "Legend1";

            series_Channel15.ChartArea = "ChartArea_PICO";
            series_Channel15.Legend = "Legend1";

            series_Channel16.ChartArea = "ChartArea_PICO";
            series_Channel16.Legend = "Legend1";


            int j1, r1, n1;
            for (j1 = Data_Table_PICO_Comp.Rows.Count - 1; j1 >= 0; j1--)
            {
                Data_Table_PICO_Comp.Rows[j1].Delete();
            }
            for (r1 = Data_Table_CTS.Rows.Count - 1; r1 >= 0; r1--)
            {
                Data_Table_CTS.Rows[r1].Delete();
            }
            int j2, r2, n2;
            for (j2 = Data_Table_PICO_Manual.Rows.Count - 1; j2 >= 0; j2--)
            {
                Data_Table_PICO_Manual.Rows[j2].Delete();
            }
            for (r2 = Data_Table_CTS.Rows.Count - 1; r2 >= 0; r2--)
            {
                Data_Table_CTS.Rows[r2].Delete();
            }
            //Index of Temperature on Components
            if (tabControl1.SelectedIndex == 0)
            {
                if (chart_PICO_Comp.Series[0].Points.Count > 0 || chart_CTS_Comp.Series[0].Points.Count > 0)
                {
                    chart_CTS_Comp.Series[0].Points.Clear();
                    chart_CTS_Comp.Series[1].Points.Clear();
                    for (n1 = chart_PICO_Comp.Series.Count - 1; n1 >= 0; n1--)
                        chart_PICO_Comp.Series[n1].Points.Clear();
                }
            }
            //Index of manual testing
            if (tabControl1.SelectedIndex == 1)
            {
                if (chart_PICO_Manual.Series[0].Points.Count > 0 || chart_CTS_Manual.Series[0].Points.Count > 0)
                {
                    chart_CTS_Manual.Series[0].Points.Clear();
                    chart_CTS_Manual.Series[1].Points.Clear();
                    for (n2 = chart_PICO_Manual.Series.Count - 1; n2 >= 0; n2--)
                        chart_PICO_Manual.Series[n2].Points.Clear();
                }
            }
        }
        //Thread DoWork1 gets all PICO channel values from unit1 and stores them in a tempbuffer
        unsafe void DoWork1()
        {
            short overflow;
            status1 = TC08example.Imports.TC08GetSingle(_handle1, tempbuffer1, &overflow, TC08example.Imports.TempUnit.USBTC08_UNITS_CENTIGRADE);
        }
        //Thread DoWork2 gets all PICO channel values from unit2 and stores them in a tempbuffer
        unsafe void DoWork2()
        {
            short overflow;
            status2 = TC08example.Imports.TC08GetSingle(_handle2, tempbuffer2, &overflow, TC08example.Imports.TempUnit.USBTC08_UNITS_CENTIGRADE);
        }
        //ThreadFinished1 = Stores the temperature values from Unit1 in the PICO data table and chart series
        void ThreadFinished1(float[] tempb1, short stat1)
        {
            if (stat1 == PICO_OK)
            {
                if (Channel1_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel1.Points.AddXY(Timer_Counter, tempb1[1]);
                    row_PICO[series_Channel1.Name] = Math.Round( (decimal)tempb1[1], 2);
                }
                if (Channel2_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel2.Points.AddXY(Timer_Counter, tempb1[2]);
                    row_PICO[series_Channel2.Name] = Math.Round( (decimal)tempb1[2], 2);
                }
                if (Channel3_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel3.Points.AddXY(Timer_Counter, tempb1[3]);
                    row_PICO[series_Channel3.Name] = Math.Round( (decimal)tempb1[3], 2);
                }
                if (Channel4_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel4.Points.AddXY(Timer_Counter, tempb1[4]);
                    row_PICO[series_Channel4.Name] = Math.Round( (decimal)tempb1[4], 2);
                }
                if (Channel5_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel5.Points.AddXY(Timer_Counter, tempb1[5]);
                    row_PICO[series_Channel5.Name] = Math.Round( (decimal)tempb1[5], 2);
                }
                if (Channel6_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel6.Points.AddXY(Timer_Counter, tempb1[6]);
                    row_PICO[series_Channel6.Name] = Math.Round( (decimal)tempb1[6], 2);
                }
                if (Channel7_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel7.Points.AddXY(Timer_Counter, tempb1[7]);
                    row_PICO[series_Channel7.Name] = Math.Round( (decimal)tempb1[7], 2);
                }
                if (Channel8_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel8.Points.AddXY(Timer_Counter, tempb1[8]);
                    row_PICO[series_Channel8.Name] = Math.Round( (decimal)tempb1[8], 2);
                }

            }
            
        }
        //ThreadFinished2 = Stores the temperature values from Unit2 in the PICO data table and chart series
        void ThreadFinished2(float[] tempb2, short stat2)
        {
            if (stat2 == PICO_OK)
            {

                if (Channel9_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel9.Points.AddXY(Timer_Counter, tempb2[1]);
                    row_PICO[series_Channel9.Name] = Math.Round( (decimal)tempb2[1], 2);
                }
                if (Channel10_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel10.Points.AddXY(Timer_Counter, tempb2[2]);
                    row_PICO[series_Channel10.Name] = Math.Round( (decimal)tempb2[2], 2);
                }
                if (Channel11_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel11.Points.AddXY(Timer_Counter, tempb2[3]);
                    row_PICO[series_Channel11.Name] = Math.Round( (decimal)tempb2[3], 2);
                }
                if (Channel12_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel12.Points.AddXY(Timer_Counter, tempb2[4]);
                    row_PICO[series_Channel12.Name] = Math.Round( (decimal)tempb2[4], 2);
                }
                if (Channel13_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel13.Points.AddXY(Timer_Counter, tempb2[5]);
                    row_PICO[series_Channel13.Name] = Math.Round( (decimal)tempb2[5], 2);
                }
                if (Channel14_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel14.Points.AddXY(Timer_Counter, tempb2[6]);
                    row_PICO[series_Channel14.Name] = Math.Round( (decimal)tempb2[6], 2);
                }
                if (Channel15_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel15.Points.AddXY(Timer_Counter, tempb2[7]);
                    row_PICO[series_Channel15.Name] = Math.Round( (decimal)tempb2[7], 2);
                }
                if (Channel16_Box.CheckState == CheckState.Checked)
                {
                    this.series_Channel16.Points.AddXY(Timer_Counter, tempb2[8]);
                    row_PICO[series_Channel16.Name] = Math.Round( (decimal)tempb2[8], 2);
                }

            }
            
        }
        //Main timer - timer_Comp_Exposure = Temperature on Components timer when Exposure radiobutton is selected
        private void timer_Comp_Exposure_Tick(object sender, EventArgs e)
        {
            string stop_Chamber = string.Empty, start = string.Empty, set_etemp = string.Empty;
            stopWatch.Stop();
            long duration = stopWatch.ElapsedMilliseconds;
            stopWatch.Reset();
            Timer_Counter += Math.Round((decimal)duration / 1000, 1);
            Timing_Box1.Text = duration.ToString() + " turn_off:" + turn_off.ToString();
            stopWatch.Start();                                                                      //Stopwatch to determine exact time for the timer loop, timers can be dynamic depending on the computers activity
            int Selected = 1;
            row_PICO = Data_Table_PICO_Comp.NewRow();                                               //Create new row in PICO data table
            row_PICO["Time"] = Timer_Counter;                                                       //Set Column "Time" value
            this.PICO1_Thread = new Thread(new ThreadStart(this.DoWork1));
            this.PICO1_Thread.IsBackground = true;                                                  // this prevents the extra thread from blocking an application shutdown
            this.PICO1_Thread.Start();
            this.PICO2_Thread = new Thread(new ThreadStart(this.DoWork2));
            this.PICO2_Thread.IsBackground = true;                                                  // this prevents the extra thread from blocking an application shutdown
            this.PICO2_Thread.Start();
            //Wait for collection of PICO values
            while (PICO1_Thread.IsAlive||PICO2_Thread.IsAlive)
            {
                ;
            }
            //Store PICO values in data table and chart series
            ThreadFinished1(tempbuffer1, status1);
            ThreadFinished2(tempbuffer2, status2);
            //If a CTS chamber is present, log temperatures "Actual" and "SET". If not, continue with only PICO temperatures
            if (serialPort_CTS.IsOpen || Ethernet_label.Text == "Connected")
            {
                logger_watch.Start();                                                               //watches to determine time spent to read and write data through serialport
                Stop_CTS_watch.Start();
                Logg_temp(Selected);                                                                //logger_watch stop time is in the serialPort_CTS_DataReceived event
            }
            //Add row to data table and add data table to gridview datasource. Scroll datagridview to the last indexed row.
            Data_Table_PICO_Comp.Rows.Add(row_PICO);
            dataGridView_PICO_Comp.DataSource = Data_Table_PICO_Comp;
            dataGridView_PICO_Comp.FirstDisplayedScrollingRowIndex = dataGridView_PICO_Comp.Rows.Count -1;
            //if CTS is running and temperature is stable (SET = ACTUAL +/- 0.1degC) start timer for exposure, print out Start/End time on label (/labels if cycle value is 2)
            if (turn_off == false && temp_Reached == false)
            {
                if (stabled() == true)
                {
                    timer_expose.Interval = exposure_Time * 1000;
                    timer_expose.Enabled = true;
                    DateTime expose = DateTime.Now;                                                 //Get computer time as "Start"
                    DateTime exposed = expose.AddSeconds(exposure_Time);                            //Calculate time as "End"
                    if(Exposure_Time_Label1.Text == "-")
                        Exposure_Time_Label1.Text = "Start: " + expose.ToString() + "\nEnd: " + exposed.ToString();
                    else
                        Exposure_Time_Label2.Text = "Start: " + expose.ToString() + "\nEnd: " + exposed.ToString();
                    chart_CTS_Comp.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                    chart_CTS_Comp.ChartAreas[0].CursorX.Interval = timer_expose.Interval;
                    chart_PICO_Comp.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                    temp_Reached = true;
                }
            }
            
            Stop_CTS_watch.Stop();
            long logger_time = logger_watch.ElapsedMilliseconds;
            long stop_CTS_time = Stop_CTS_watch.ElapsedMilliseconds;
            Timing_Box1.Text += " logg: " + logger_time.ToString() + " Stop: " + stop_CTS_time.ToString();
            logger_watch.Reset();
            Stop_CTS_watch.Reset();
            //Turn off the CTS chamber after exposure time has elapsed
            if (turn_off_chamber == true && turn_off == false)
            {
                //Turn off chamber using serialport or Ethernet
                if (Ethernet_label.Text != "Connected")
                {
                    serialPort_CTS.ReceivedBytesThreshold = 8;
                    data_Tx[0] = Convert.ToByte("0x02", 16);
                    data_Tx[1] = Convert.ToByte("0x81", 16);
                    data_Tx[2] = Convert.ToByte("0xF0", 16);
                    data_Tx[3] = Convert.ToByte("0xB0", 16);
                    data_Tx[4] = Convert.ToByte("0xB0", 16);
                    data_Tx[5] = Convert.ToByte("0xB0", 16);
                    data_Tx[6] = Convert.ToByte("0xC1", 16);
                    data_Tx[7] = Convert.ToByte("0x03", 16);
                    serialPort_CTS.Write(data_Tx, 0, 8);
                    Thread.Sleep(100);
                }
                
                if(Ethernet_label.Text == "Connected")
                    stop_Chamber = Read_Write_Ethernet(Ethernet_Box.Text, "s1 0");
                turn_off = true;
                turn_off_chamber = false;
            }
            //When the CTS chamber has turned off, SET labels and wait for CTS temperature to reach 10 degrees C below SET value, stop the measurements or cycle to next SET value
            if (turn_off == true)
            {
                if(Set_analog_minus_10_Label1.Text == "-")
                    Set_analog_minus_10_Label1.Text = "Measure until tempchamber reaches: " + (double.Parse(Set_Analog_Box1_Comp.Text) - 10).ToString();
                else if(Exposure_Time_Label2.Text != "-")
                    Set_analog_minus_10_Label2.Text = "Measure until tempchamber reaches: " + (double.Parse(Set_Analog_Box2_Comp.Text) - 10).ToString();
                //Wait for CTS temperature to reach 10 degrees C below SET value and check if it has cycled
                if (ActTemp[PICO_Counter] <= double.Parse(Set_Analog_Box1_Comp.Text) - 10 && cycle_Comp == false)
                {
                    //Stop measurement if cycle value is 1
                    if (comboBox_Comp.SelectedIndex == 0)
                    {
                        Stop_Btn_Comp_Click(null, null);
                        label_Comp_Finished.Visible = true;
                        //Send email if the radiobutton is checked and email textbox isnt empty
                        if (Email_ON.Checked == true && Email_Box.Text != "")
                        {
                            Send_Cycle_Finished();
                        }
                    }
                    //Or continue to next SET value
                    else if (comboBox_Comp.SelectedIndex == 1)
                    {
                        turn_off = false;
                        temp_Reached = false;
                        turn_off_chamber = false;
                        exposure_Time = long.Parse(Exposure_Time2_Comp.Text);
                        if (Ethernet_label.Text != "Connected")
                        {
                            //Set Tempeture on Chamber
                            check_Set = SET_Temp_CTS(Set_Analog_Box2_Comp.Text);
                            //Start chamber
                            Start_CTS();
                            cycle_Comp = true;
                        }
                        if (Ethernet_label.Text == "Connected")
                        {
                            //Set Tempeture on Chamber
                            if (Set_Analog_Box1_Comp.Text.Length == 2)
                            {
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box2_Comp.Text + ".0");
                            }
                            else
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box2_Comp.Text + ".0");
                            Thread.Sleep(100);
                            //Start chamber
                            start = Read_Write_Ethernet(Ethernet_Box.Text, "s1 1");
                            cycle_Comp = true;
                        }
                    }
                }
                //or if CTS temperature reached 10 degrees C below SET value and it has cycled, stop measurements
                else if (Set_Analog_Box2_Comp.Visible == true)
                {
                    if (ActTemp[PICO_Counter] <= double.Parse(Set_Analog_Box2_Comp.Text) - 10 && cycle_Comp == true)
                    {
                        Stop_Btn_Comp_Click(null, null);
                        label_Comp_Finished.Visible = true;
                        //Send email if the radiobutton is checked and email textbox isnt empty
                        if (Email_ON.Checked == true && Email_Box.Text != "")
                        {
                            Send_Cycle_Finished();
                        }
                    }
                }
            }
            PICO_Counter += 1;                                                                                  //Increase the PICO counter
        }

        private void Channel1_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel1_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel1.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel1.Name = "Channel1";
                    Channel1_Column.ColumnName = "Channel1";
                    this.chart_PICO_Comp.Series.Add(series_Channel1);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel1_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel1.Name = "Channel1";
                    Channel1_Column.ColumnName = "Channel1";
                    this.chart_PICO_Manual.Series.Add(series_Channel1);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel1_Column);
                }
            }
            else
            {
                textBox_Channel1.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel1.Name = "Channel1";
                    Channel1_Column.ColumnName = "Channel1";
                    this.chart_PICO_Comp.Series.Remove(series_Channel1);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel1_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel1.Name = "Channel1";
                    Channel1_Column.ColumnName = "Channel1";
                    this.chart_PICO_Manual.Series.Remove(series_Channel1);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel1_Column);
                }
            }
        }

        private void Channel2_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel2_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel2.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel2.Name = "Channel2";
                    Channel2_Column.ColumnName = "Channel2";
                    this.chart_PICO_Comp.Series.Add(series_Channel2);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel2_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel2.Name = "Channel2";
                    Channel2_Column.ColumnName = "Channel2";
                    this.chart_PICO_Manual.Series.Add(series_Channel2);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel2_Column);
                }
            }
            else
            {
                textBox_Channel2.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel2.Name = "Channel2";
                    Channel2_Column.ColumnName = "Channel2";
                    this.chart_PICO_Comp.Series.Remove(series_Channel2);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel2_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel2.Name = "Channel2";
                    Channel2_Column.ColumnName = "Channel2";
                    this.chart_PICO_Manual.Series.Remove(series_Channel2);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel2_Column);
                }
            }
        }

        private void Channel3_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel3_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel3.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel3.Name = "Channel3";
                    Channel3_Column.ColumnName = "Channel3";
                    this.chart_PICO_Comp.Series.Add(series_Channel3);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel3_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel3.Name = "Channel3";
                    Channel3_Column.ColumnName = "Channel3";
                    this.chart_PICO_Manual.Series.Add(series_Channel3);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel3_Column);
                }
            }
            else
            {
                textBox_Channel3.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel3.Name = "Channel3";
                    Channel3_Column.ColumnName = "Channel3";
                    this.chart_PICO_Comp.Series.Remove(series_Channel3);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel3_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel3.Name = "Channel3";
                    Channel3_Column.ColumnName = "Channel3";
                    this.chart_PICO_Manual.Series.Remove(series_Channel3);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel3_Column);
                }
            }
        }

        private void Channel4_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel4_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel4.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel4.Name = "Channel4";
                    Channel4_Column.ColumnName = "Channel4";
                    this.chart_PICO_Comp.Series.Add(series_Channel4);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel4_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel4.Name = "Channel4";
                    Channel4_Column.ColumnName = "Channel4";
                    this.chart_PICO_Manual.Series.Add(series_Channel4);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel4_Column);
                }
            }
            else
            {
                textBox_Channel4.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel4.Name = "Channel4";
                    Channel4_Column.ColumnName = "Channel4";
                    this.chart_PICO_Comp.Series.Remove(series_Channel4);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel4_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel4.Name = "Channel4";
                    Channel4_Column.ColumnName = "Channel4";
                    this.chart_PICO_Manual.Series.Remove(series_Channel4);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel4_Column);
                }
            }
        }

        private void Channel5_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel5_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel5.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel5.Name = "Channel5";
                    Channel5_Column.ColumnName = "Channel5";
                    this.chart_PICO_Comp.Series.Add(series_Channel5);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel5_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel5.Name = "Channel5";
                    Channel5_Column.ColumnName = "Channel5";
                    this.chart_PICO_Manual.Series.Add(series_Channel5);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel5_Column);
                }
            }
            else
            {
                textBox_Channel5.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel5.Name = "Channel5";
                    Channel5_Column.ColumnName = "Channel5";
                    this.chart_PICO_Comp.Series.Remove(series_Channel5);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel5_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel5.Name = "Channel5";
                    Channel5_Column.ColumnName = "Channel5";
                    this.chart_PICO_Manual.Series.Remove(series_Channel5);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel5_Column);
                }
            }
        }

        private void Channel6_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel6_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel6.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel6.Name = "Channel6";
                    Channel6_Column.ColumnName = "Channel6";
                    this.chart_PICO_Comp.Series.Add(series_Channel6);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel6_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel6.Name = "Channel6";
                    Channel6_Column.ColumnName = "Channel6";
                    this.chart_PICO_Manual.Series.Add(series_Channel6);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel6_Column);
                }
            }
            else
            {
                textBox_Channel6.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel6.Name = "Channel6";
                    Channel6_Column.ColumnName = "Channel6";
                    this.chart_PICO_Comp.Series.Remove(series_Channel6);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel6_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel6.Name = "Channel6";
                    Channel6_Column.ColumnName = "Channel6";
                    this.chart_PICO_Manual.Series.Remove(series_Channel6);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel6_Column);
                }
            }
        }

        private void Channel7_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel7_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel7.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel7.Name = "Channel7";
                    Channel7_Column.ColumnName = "Channel7";
                    this.chart_PICO_Comp.Series.Add(series_Channel7);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel7_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel7.Name = "Channel7";
                    Channel7_Column.ColumnName = "Channel7";
                    this.chart_PICO_Manual.Series.Add(series_Channel7);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel7_Column);
                }
            }
            else
            {
                textBox_Channel7.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel7.Name = "Channel7";
                    Channel7_Column.ColumnName = "Channel7";
                    this.chart_PICO_Comp.Series.Remove(series_Channel7);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel7_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel7.Name = "Channel7";
                    Channel7_Column.ColumnName = "Channel7";
                    this.chart_PICO_Manual.Series.Remove(series_Channel7);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel7_Column);
                }
            }
        }

        private void Channel8_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel8_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel8.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel8.Name = "Channel8";
                    Channel8_Column.ColumnName = "Channel8";
                    this.chart_PICO_Comp.Series.Add(series_Channel8);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel8_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel8.Name = "Channel8";
                    Channel8_Column.ColumnName = "Channel8";
                    this.chart_PICO_Manual.Series.Add(series_Channel8);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel8_Column);
                }
            }
            else
            {
                textBox_Channel8.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel8.Name = "Channel8";
                    Channel8_Column.ColumnName = "Channel8";
                    this.chart_PICO_Comp.Series.Remove(series_Channel8);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel8_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel8.Name = "Channel8";
                    Channel8_Column.ColumnName = "Channel8";
                    this.chart_PICO_Manual.Series.Remove(series_Channel8);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel8_Column);
                }
            }
        }

        private void Channel9_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel9_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel9.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel9.Name = "Channel9";
                    Channel9_Column.ColumnName = "Channel9";
                    this.chart_PICO_Comp.Series.Add(series_Channel9);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel9_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel9.Name = "Channel9";
                    Channel9_Column.ColumnName = "Channel9";
                    this.chart_PICO_Manual.Series.Add(series_Channel9);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel9_Column);
                }
            }
            else
            {
                textBox_Channel9.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel9.Name = "Channel9";
                    Channel9_Column.ColumnName = "Channel9";
                    this.chart_PICO_Comp.Series.Remove(series_Channel9);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel9_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel9.Name = "Channel9";
                    Channel9_Column.ColumnName = "Channel9";
                    this.chart_PICO_Manual.Series.Remove(series_Channel9);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel9_Column);
                }
            }
        }

        private void Channel10_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel10_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel10.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel10.Name = "Channel10";
                    Channel10_Column.ColumnName = "Channel10";
                    this.chart_PICO_Comp.Series.Add(series_Channel10);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel10_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel10.Name = "Channel10";
                    Channel10_Column.ColumnName = "Channel10";
                    this.chart_PICO_Manual.Series.Add(series_Channel10);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel10_Column);
                }
            }
            else
            {
                textBox_Channel10.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel10.Name = "Channel10";
                    Channel10_Column.ColumnName = "Channel10";
                    this.chart_PICO_Comp.Series.Remove(series_Channel10);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel10_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel10.Name = "Channel10";
                    Channel10_Column.ColumnName = "Channel10";
                    this.chart_PICO_Manual.Series.Remove(series_Channel10);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel10_Column);
                }
            }
        }

        private void Channel11_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel11_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel11.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel11.Name = "Channel11";
                    Channel11_Column.ColumnName = "Channel11";
                    this.chart_PICO_Comp.Series.Add(series_Channel11);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel11_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel11.Name = "Channel11";
                    Channel11_Column.ColumnName = "Channel11";
                    this.chart_PICO_Manual.Series.Add(series_Channel11);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel11_Column);
                }
            }
            else
            {
                textBox_Channel11.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel11.Name = "Channel11";
                    Channel11_Column.ColumnName = "Channel11";
                    this.chart_PICO_Comp.Series.Remove(series_Channel11);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel11_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel11.Name = "Channel11";
                    Channel11_Column.ColumnName = "Channel11";
                    this.chart_PICO_Manual.Series.Remove(series_Channel11);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel11_Column);
                }
            }
        }

        private void Channel12_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel12_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel12.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel12.Name = "Channel12";
                    Channel12_Column.ColumnName = "Channel12";
                    this.chart_PICO_Comp.Series.Add(series_Channel12);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel12_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel12.Name = "Channel12";
                    Channel12_Column.ColumnName = "Channel12";
                    this.chart_PICO_Manual.Series.Add(series_Channel12);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel12_Column);
                }
            }
            else
            {
                textBox_Channel12.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel12.Name = "Channel12";
                    Channel12_Column.ColumnName = "Channel12";
                    this.chart_PICO_Comp.Series.Remove(series_Channel12);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel12_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel12.Name = "Channel12";
                    Channel12_Column.ColumnName = "Channel12";
                    this.chart_PICO_Manual.Series.Remove(series_Channel12);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel12_Column);
                }
            }
        }

        private void Channel13_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel13_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel13.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel13.Name = "Channel13";
                    Channel13_Column.ColumnName = "Channel13";
                    this.chart_PICO_Comp.Series.Add(series_Channel13);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel13_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel13.Name = "Channel13";
                    Channel13_Column.ColumnName = "Channel13";
                    this.chart_PICO_Manual.Series.Add(series_Channel13);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel13_Column);
                }
            }
            else
            {
                textBox_Channel13.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel13.Name = "Channel13";
                    Channel13_Column.ColumnName = "Channel13";
                    this.chart_PICO_Comp.Series.Remove(series_Channel13);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel13_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel13.Name = "Channel13";
                    Channel13_Column.ColumnName = "Channel13";
                    this.chart_PICO_Manual.Series.Remove(series_Channel13);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel13_Column);
                }
            }
        }

        private void Channel14_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel14_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel14.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel14.Name = "Channel14";
                    Channel14_Column.ColumnName = "Channel14";
                    this.chart_PICO_Comp.Series.Add(series_Channel14);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel14_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel14.Name = "Channel14";
                    Channel14_Column.ColumnName = "Channel14";
                    this.chart_PICO_Manual.Series.Add(series_Channel14);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel14_Column);
                }
            }
            else
            {
                textBox_Channel14.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel14.Name = "Channel14";
                    Channel14_Column.ColumnName = "Channel14";
                    this.chart_PICO_Comp.Series.Remove(series_Channel14);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel14_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel14.Name = "Channel14";
                    Channel14_Column.ColumnName = "Channel14";
                    this.chart_PICO_Manual.Series.Remove(series_Channel14);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel14_Column);
                }
            }
        }

        private void Channel15_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel15_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel15.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel15.Name = "Channel15";
                    Channel15_Column.ColumnName = "Channel15";
                    this.chart_PICO_Comp.Series.Add(series_Channel15);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel15_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel15.Name = "Channel15";
                    Channel15_Column.ColumnName = "Channel15";
                    this.chart_PICO_Manual.Series.Add(series_Channel15);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel15_Column);
                }
            }
            else
            {
                textBox_Channel15.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel15.Name = "Channel15";
                    Channel15_Column.ColumnName = "Channel15";
                    this.chart_PICO_Comp.Series.Remove(series_Channel15);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel15_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel15.Name = "Channel15";
                    Channel15_Column.ColumnName = "Channel15";
                    this.chart_PICO_Manual.Series.Remove(series_Channel15);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel15_Column);
                }
            }
        }

        private void Channel16_Box_CheckedChanged(object sender, EventArgs e)
        {
            if (Channel16_Box.CheckState == CheckState.Checked)
            {
                textBox_Channel16.Visible = true;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel16.Name = "Channel16";
                    Channel16_Column.ColumnName = "Channel16";
                    this.chart_PICO_Comp.Series.Add(series_Channel16);
                    this.Data_Table_PICO_Comp.Columns.Add(Channel16_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel16.Name = "Channel16";
                    Channel16_Column.ColumnName = "Channel16";
                    this.chart_PICO_Manual.Series.Add(series_Channel16);
                    this.Data_Table_PICO_Manual.Columns.Add(Channel16_Column);
                }
            }
            else
            {
                textBox_Channel16.Visible = false;
                if (tabControl1.SelectedIndex == 0)
                {
                    series_Channel16.Name = "Channel16";
                    Channel16_Column.ColumnName = "Channel16";
                    this.chart_PICO_Comp.Series.Remove(series_Channel16);
                    this.Data_Table_PICO_Comp.Columns.Remove(Channel16_Column);
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    series_Channel16.Name = "Channel16";
                    Channel16_Column.ColumnName = "Channel16";
                    this.chart_PICO_Manual.Series.Remove(series_Channel16);
                    this.Data_Table_PICO_Manual.Columns.Remove(Channel16_Column);
                }
            }
        }

        private void serialPort_CTS_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Rx_buf = serialPort_CTS.Read(data_Rx, 0, 32);
            
            logger_watch.Stop();
            
            Thread.Sleep(100);
            
            //Exposure_Time_Page1.Text = Rx_tid.ToString();
        }

        public int SET_Temp_CTS(string Set_Value)
        {
            int Return_Set_Value = int.Parse(Set_Value);
            
            serialPort_CTS.ReceivedBytesThreshold = 5;
            string Set1, Set2, Set3, Set4 = "0xB0";
            byte CHK = 0;
            Set1 = Set_Value.Substring(0, 1);
            Set2 = Set_Value.Substring(1, 1);
            data_Tx[0] = Convert.ToByte("0x02", 16);
            data_Tx[1] = Convert.ToByte("0x81", 16);
            data_Tx[2] = Convert.ToByte("0xE1", 16);
            data_Tx[3] = Convert.ToByte("0xB0", 16);
            data_Tx[4] = Convert.ToByte("0xA0", 16);
            if (Set1 == "-")
            {
                Set1 = "0xAD";
                Set3 = Set_Value.Substring(2, 1);
                data_Tx[5] = Convert.ToByte(Set1, 16);
                data_Tx[6] = Convert.ToByte("0xB" + Set2, 16);
                data_Tx[7] = Convert.ToByte("0xB" + Set3, 16);
                data_Tx[8] = Convert.ToByte("0xAE", 16);
                data_Tx[9] = Convert.ToByte(Set4, 16);
            }
            else
            {
                data_Tx[5] = Convert.ToByte("0xB0", 16);
                data_Tx[6] = Convert.ToByte("0xB" + Set1, 16);
                data_Tx[7] = Convert.ToByte("0xB" + Set2, 16);
                data_Tx[8] = Convert.ToByte("0xAE", 16);
                data_Tx[9] = Convert.ToByte(Set4, 16);
            }
            for (int i = 1; i < 10; i++)
                CHK ^= data_Tx[i];
            if (CHK < 128)
                CHK += 128;
            data_Tx[10] = CHK;
            data_Tx[11] = Convert.ToByte("0x03", 16);
            serialPort_CTS.Write(data_Tx, 0, 12);
            Thread.Sleep(100);
            return Return_Set_Value;
        }

        public void Start_CTS()
        {
            serialPort_CTS.ReceivedBytesThreshold = 6;
            data_Tx[0] = Convert.ToByte("0x02", 16);
            data_Tx[1] = Convert.ToByte("0x81", 16);
            data_Tx[2] = Convert.ToByte("0xF3", 16);
            data_Tx[3] = Convert.ToByte("0xB1", 16);
            data_Tx[4] = Convert.ToByte("0xA0", 16);
            data_Tx[5] = Convert.ToByte("0xB1", 16);
            data_Tx[6] = Convert.ToByte("0xD2", 16);
            data_Tx[7] = Convert.ToByte("0x03", 16);
            serialPort_CTS.Write(data_Tx, 0, 8);
            Thread.Sleep(100);
        }

        public string Return_temp()
        {
            string act_set = "0";
            
            serialPort_CTS.ReceivedBytesThreshold = 18;
            data_Tx[0] = Convert.ToByte("0x02", 16);
            data_Tx[1] = Convert.ToByte("0x81", 16);
            data_Tx[2] = Convert.ToByte("0xC1", 16);
            data_Tx[3] = Convert.ToByte("0xB0", 16);
            data_Tx[4] = Convert.ToByte("0xF0", 16);
            data_Tx[5] = Convert.ToByte("0x03", 16);
            serialPort_CTS.Write(data_Tx, 0, 6);
            Thread.Sleep(200);
            /*
            if (Rx_tid < 15)
                Thread.Sleep(50);
            else
                Thread.Sleep((int)Rx_tid);
            */
            act_set = data_Rx[5].ToString("x").ToUpper();
            act_set += data_Rx[6].ToString("x").ToUpper();
            act_set += data_Rx[7].ToString("x").ToUpper();
            act_set += data_Rx[8].ToString("x").ToUpper();
            act_set += data_Rx[9].ToString("x").ToUpper();
            act_set += data_Rx[11].ToString("x").ToUpper();
            act_set += data_Rx[12].ToString("x").ToUpper();
            act_set += data_Rx[13].ToString("x").ToUpper();
            act_set += data_Rx[14].ToString("x").ToUpper();
            act_set += data_Rx[15].ToString("x").ToUpper();
            //Thread.Sleep(100);
            //Exposure_Time_Page1.Text = act_set;
            return act_set;
        }

        public void Logg_temp(int select_Program)
        {
            string B1, B2, B3, B4, B5, B6, B7, B8, B9, B10;
            string Returned_Value = "";
            try
            {
                if (Ethernet_label.Text != "Connected")
                {
                    Returned_Value = Return_temp();
                    if (Returned_Value != "0000000000")
                    {
                        B1 = Returned_Value.Substring(0, 2); // Kan vara "-" (AD) eller "0" (B0)
                        B2 = Returned_Value.Substring(3, 1);
                        B3 = Returned_Value.Substring(5, 1);
                        //B4 = Returned_Value.Substring(6, 2); = "AE" decimal tecken
                        B4 = ".";
                        B5 = Returned_Value.Substring(9, 1);
                        B6 = Returned_Value.Substring(10, 2); // Kan vara "-" (AD) eller "0" (B0)
                        B7 = Returned_Value.Substring(13, 1);
                        B8 = Returned_Value.Substring(15, 1);
                        //B9 = Returned_Value.Substring(16, 2);  = "AE" decimal tecken
                        B9 = ".";
                        B10 = Returned_Value.Substring(19, 1);
                        if (B1 == "AD")
                            B1 = "-";
                        else
                            B1 = "";
                        if (B6 == "AD")
                            B6 = "-";
                        else
                            B6 = "";
                        string actOut = string.Concat(B1, B2, B3, B4, B5);
                        string setOut = string.Concat(B6, B7, B8, B9, B10);

                        ActTemp[PICO_Counter] = double.Parse(actOut, NumberStyles.Number, NumberFormatInfo.InvariantInfo);
                        SetTemp[PICO_Counter] = double.Parse(setOut, NumberStyles.Number, NumberFormatInfo.InvariantInfo);
                        //Ful fix nr1 (a problem have been encountered with a tempchamber that doesnt send the data B8 in a correct way and I get a blank value for example 40 I get 4, it happens sometimes often and sometimes not at all, depends on its mood ;)
                        if (SetTemp[PICO_Counter] != (double)check_Set)
                            SetTemp[PICO_Counter] = (double)check_Set;
                        row_CTS = Data_Table_CTS.NewRow();
                        row_CTS["Time"] = Timer_Counter;
                        row_CTS["Actual"] = ActTemp[PICO_Counter];
                        row_CTS["Set"] = SetTemp[PICO_Counter];
                        Data_Table_CTS.Rows.Add(row_CTS);
                        if (select_Program == 1)
                        {
                            dataGridView_CTS_Comp.DataSource = Data_Table_CTS;
                            dataGridView_CTS_Comp.FirstDisplayedScrollingRowIndex = dataGridView_CTS_Comp.Rows.Count - 1;
                        }
                        if (select_Program == 2)
                        {
                            dataGridView_CTS_Manual.DataSource = Data_Table_CTS;
                            dataGridView_CTS_Manual.FirstDisplayedScrollingRowIndex = dataGridView_CTS_Manual.Rows.Count - 1;
                        }
                        ChartSeries1.Points.AddXY(Timer_Counter, ActTemp[PICO_Counter]);
                        ChartSeries2.Points.AddXY(Timer_Counter, SetTemp[PICO_Counter]);


                    }
                    else
                        Notify_Label.Text = "No communication with chamber!";
                }
                if(Ethernet_label.Text == "Connected")
                {
                    Returned_Value = Read_Write_Ethernet(Ethernet_Box.Text, "A0");
                    richTextBox1.Text = Returned_Value;
                    B1 = Returned_Value.Substring(3, 5); 
                    B2 = Returned_Value.Substring(9, 5);
                    richTextBox1.Text = B1 + "; " + B2;
                    ActTemp[PICO_Counter] = double.Parse(B1, NumberStyles.Number, NumberFormatInfo.InvariantInfo);
                    SetTemp[PICO_Counter] = double.Parse(B2, NumberStyles.Number, NumberFormatInfo.InvariantInfo);

                    row_CTS = Data_Table_CTS.NewRow();
                    row_CTS["Time"] = Timer_Counter;
                    row_CTS["Actual"] = ActTemp[PICO_Counter];
                    row_CTS["Set"] = SetTemp[PICO_Counter];
                    Data_Table_CTS.Rows.Add(row_CTS);
                    if (select_Program == 1)
                    {
                        dataGridView_CTS_Comp.DataSource = Data_Table_CTS;
                        dataGridView_CTS_Comp.FirstDisplayedScrollingRowIndex = dataGridView_CTS_Comp.Rows.Count - 1;
                    }
                    if (select_Program == 2)
                    {
                        dataGridView_CTS_Manual.DataSource = Data_Table_CTS;
                        dataGridView_CTS_Manual.FirstDisplayedScrollingRowIndex = dataGridView_CTS_Manual.Rows.Count - 1;
                    }
                    ChartSeries1.Points.AddXY(Timer_Counter, ActTemp[PICO_Counter]);
                    ChartSeries2.Points.AddXY(Timer_Counter, SetTemp[PICO_Counter]);
                }
            }
            catch (Exception ex)
            {
                string logg_exeption = ex.Message;
                Notify_Label.Text = "Temperature chamber: Logg_temp collision @ measurement:" + Timer_Counter.ToString() + " ActTemp: " + ActTemp[PICO_Counter].ToString() + " SetTemp: " + SetTemp[PICO_Counter].ToString() + " Returned: " + Returned_Value;
            }
        }

        private void CHART_ON_CheckedChanged(object sender, EventArgs e)
        {
            if (CHART_ON_Comp.Checked == true)
            {
                chart_CTS_Comp.Visible = true;
                chart_PICO_Comp.Visible = true;
            }
            dataGridView_PICO_Comp.Visible = false;
            dataGridView_CTS_Comp.Visible = false;
        }

        private void GRID_ON_CheckedChanged(object sender, EventArgs e)
        {
            if (GRID_ON_Comp.Checked == true)
            {
                dataGridView_CTS_Comp.Visible = true;
                dataGridView_PICO_Comp.Visible = true;
            }
            chart_CTS_Comp.Visible = false;
            chart_PICO_Comp.Visible = false;
        }

        public void Send_Cycle_Finished()
        {
            try
            {
                string email_to = Email_Box.Text;
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = "Hello,\r\nTemperature measurements are finished ";
                //Add an attachment.
                //String sDisplayName = "MyAttachment";
                //int iPosition = (int)oMsg.Body.Length + 1;
                //int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                //Outlook.Attachment oAttach = oMsg.Attachments.Add(@"C:\\fileName.jpg", iAttachType, iPosition, sDisplayName);
                //Subject line
                oMsg.Subject = "Temperature Chamber info";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(email_to);
                oRecip.Resolve();
                // Send.
                oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception ex)
            {
                Notify_Label.Text = ex.ToString();
            }//end of catch
        }

        private void Email_ON_CheckedChanged(object sender, EventArgs e)
        {
            if (Email_ON.Checked == true)
            {
                Email_Box.Visible = true;
                Email_example_label.Visible = true;
            }
        }

        private void Email_OFF_CheckedChanged(object sender, EventArgs e)
        {
            if (Email_OFF.Checked == true)
            {
                Email_Box.Visible = false;
                Email_example_label.Visible = false;
            }
        }

        private void IPnr_1_btn_Click(object sender, EventArgs e)
        {
            if (Camera1_Box.Text != "")
            {
                IPnr1 = Camera1_Box.Text;
                ProcessStartInfo startInfo = new ProcessStartInfo("IExplore.exe");
                //startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                startInfo.Arguments = IPnr1.Replace(",", "."); ;
                Process.Start(startInfo);
                /*if (IPnr1.ToUpper().Contains("HTTP:"))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo("IExplore.exe");
                    //startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    startInfo.Arguments = IPnr1;
                    Process.Start(startInfo);
                }
                else if (IPnr1.ToUpper().Contains("C:"))
                    Process.Start(IPnr1);
                else
                {
                    Device_1_Dialog Dev_1_Dialog = new Device_1_Dialog();
                    Dev_1_Dialog.Show();
                }*/
            }
        }

        private void IPnr_2_btn_Click(object sender, EventArgs e)
        {
            if (Camera2_Box.Text != "")
            {
                IPnr2 = Camera2_Box.Text;
                ProcessStartInfo startInfo = new ProcessStartInfo("IExplore.exe");
                //startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                startInfo.Arguments = IPnr2.Replace(",", ".");
                Process.Start(startInfo);
                /*if (IPnr2.ToUpper().Contains("HTTP:"))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo("IExplore.exe");
                    //startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    startInfo.Arguments = IPnr2;
                    Process.Start(startInfo);
                }
                else if (IPnr1.ToUpper().Contains("C:"))
                    Process.Start(IPnr2);
                else
                {
                    Device_2_Dialog Dev_2_Dialog = new Device_2_Dialog();
                    Dev_2_Dialog.Show();
                }*/
            }
        }

        public delegate void InvokeDelegate();

        public void ExportToExcel(DataGridView dgView)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                // instantiating the excel application class
                object misValue = System.Reflection.Missing.Value;
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook currentWorkbook = excelApp.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet currentWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)currentWorkbook.ActiveSheet;
                currentWorksheet.Name = "Measurements";
                Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)currentWorksheet.ChartObjects(Type.Missing);
                currentWorksheet.Columns.ColumnWidth = 18;
                string column_A_Q = "";
                if (dgView.Columns.Count == 2)
                    column_A_Q = "B";
                else if (dgView.Columns.Count == 3)
                    column_A_Q = "C";
                else if (dgView.Columns.Count == 4)
                    column_A_Q = "D";
                else if (dgView.Columns.Count == 5)
                    column_A_Q = "E";
                else if (dgView.Columns.Count == 6)
                    column_A_Q = "F";
                else if (dgView.Columns.Count == 7)
                    column_A_Q = "G";
                else if (dgView.Columns.Count == 8)
                    column_A_Q = "H";
                else if (dgView.Columns.Count == 9)
                    column_A_Q = "I";
                else if (dgView.Columns.Count == 10)
                    column_A_Q = "J";
                else if (dgView.Columns.Count == 11)
                    column_A_Q = "K";
                else if (dgView.Columns.Count == 12)
                    column_A_Q = "L";
                else if (dgView.Columns.Count == 13)
                    column_A_Q = "M";
                else if (dgView.Columns.Count == 14)
                    column_A_Q = "N";
                else if (dgView.Columns.Count == 15)
                    column_A_Q = "O";
                else if (dgView.Columns.Count == 16)
                    column_A_Q = "P";
                else if (dgView.Columns.Count == 17)
                    column_A_Q = "Q";

                if (dgView.Rows.Count > 0)
                {
                    string timestamp = DateTime.Now.ToString("s");
                    timestamp = timestamp.Replace("T", "\nT");
                    currentWorksheet.Cells[1, 1] = timestamp;
                    int i = 1;
                    foreach (DataGridViewColumn dgviewColumn in dgView.Columns)
                    {
                        // Excel work sheet indexing starts with 1
                        currentWorksheet.Cells[2, i] = dgviewColumn.Name;
                        ++i;
                    }
                    Microsoft.Office.Interop.Excel.Range headerColumnRange = currentWorksheet.get_Range("A2", "Q2");
                    headerColumnRange.Font.Bold = true;
                    headerColumnRange.Font.Color = 0xFF0000;
                    //headerColumnRange.EntireColumn.AutoFit();
                    int rowIndex = 0;
                    for (rowIndex = 0; rowIndex < dgView.Rows.Count; rowIndex++)
                    {
                        DataGridViewRow dgRow = dgView.Rows[rowIndex];
                        for (int cellIndex = 0; cellIndex < dgRow.Cells.Count; cellIndex++)
                        {
                            currentWorksheet.Cells[rowIndex + 3, cellIndex + 1] = dgRow.Cells[cellIndex].Value;
                        }
                        _procent = (int)(_step / dgView.Rows.Count) * 100;
                        _step++;
                    }
                    
                    Microsoft.Office.Interop.Excel.Range fullTextRange = currentWorksheet.get_Range("A1", column_A_Q + (rowIndex + 1).ToString());
                    fullTextRange.WrapText = true;
                    fullTextRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    Microsoft.Office.Interop.Excel.ChartObject myChart = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                    Microsoft.Office.Interop.Excel.Chart chartPage = myChart.Chart;
                    Microsoft.Office.Interop.Excel.Range chartRange = currentWorksheet.get_Range("B2", column_A_Q + (rowIndex + 1).ToString());
                    chartPage.ChartType = Excel.XlChartType.xlLineMarkers;
                    chartPage.HasTitle = true;
                    chartPage.ChartTitle.Text = "Temperature on Components";
                    chartPage.HasLegend = true;
                    chartPage.SetSourceData(chartRange, Excel.XlRowCol.xlColumns);
                    var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    yAxis.HasTitle = true;
                    yAxis.AxisTitle.Text = "Temperature (°C)";
                    //yAxis.MaximumScale = 20;
                    yAxis.AxisTitle.Orientation = Excel.XlOrientation.xlUpward;

                    //Excel.Range Data_Range = currentWorksheet.get_Range("B3", "Q10"); //+ (rowIndex + 1).ToString());//Data to be plotted in chart
                    Excel.Range XVal_Range = currentWorksheet.get_Range("A3", "A" + (rowIndex + 1).ToString());//Catagory Names I want on X-Axis as range

                    Excel.Axis xAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    xAxis.CategoryNames = XVal_Range;
                    xAxis.HasTitle = true;
                    xAxis.AxisTitle.Text = "Seconds";
                    //chartPage.PlotBy = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns;

                    //chartPage.ChartWizard(chartRange, Excel.XlChartType.xlLineMarkers, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing);

                    chartPage.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet, "Chart");
                }
                
                using (SaveFileDialog exportSaveFileDialog = new SaveFileDialog())
                {
                    exportSaveFileDialog.Title = "Select Excel File";
                    exportSaveFileDialog.Filter = "Microsoft Office Excel Workbook(*.xlsx)|*.xlsx";

                    if (DialogResult.OK == exportSaveFileDialog.ShowDialog())
                    {
                        string fullFileName = exportSaveFileDialog.FileName;
                        // currentWorkbook.SaveCopyAs(fullFileName);
                        // indicating that we already saved the workbook, otherwise call to Quit() will pop up
                        // the save file dialogue box

                        currentWorkbook.SaveAs(fullFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, misValue, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
                        currentWorkbook.Saved = true;
                        MessageBox.Show("Exported successfully", "Exported to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        public Excel.Workbook ExportToExcel2(DataGridView dgView, string text_input)
        {
            // instantiating the excel application class
            object misValue = System.Reflection.Missing.Value;
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook currentWorkbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet currentWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)currentWorkbook.ActiveSheet;
            currentWorksheet.Name = "Measurements";
            Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)currentWorksheet.ChartObjects(Type.Missing);
            currentWorksheet.Columns.ColumnWidth = 18;
            string column_A_Q = "";
            if (dgView.Columns.Count == 2)
                column_A_Q = "B";
            else if (dgView.Columns.Count == 3)
                column_A_Q = "C";
            else if (dgView.Columns.Count == 4)
                column_A_Q = "D";
            else if (dgView.Columns.Count == 5)
                column_A_Q = "E";
            else if (dgView.Columns.Count == 6)
                column_A_Q = "F";
            else if (dgView.Columns.Count == 7)
                column_A_Q = "G";
            else if (dgView.Columns.Count == 8)
                column_A_Q = "H";
            else if (dgView.Columns.Count == 9)
                column_A_Q = "I";
            else if (dgView.Columns.Count == 10)
                column_A_Q = "J";
            else if (dgView.Columns.Count == 11)
                column_A_Q = "K";
            else if (dgView.Columns.Count == 12)
                column_A_Q = "L";
            else if (dgView.Columns.Count == 13)
                column_A_Q = "M";
            else if (dgView.Columns.Count == 14)
                column_A_Q = "N";
            else if (dgView.Columns.Count == 15)
                column_A_Q = "O";
            else if (dgView.Columns.Count == 16)
                column_A_Q = "P";
            else if (dgView.Columns.Count == 17)
                column_A_Q = "Q";

            if (dgView.Rows.Count > 0)
            {
                string timestamp = DateTime.Now.ToString("s");
                timestamp = timestamp.Replace("T", "\nT");
                currentWorksheet.Cells[1, 1] = timestamp;
                int i = 1;
                foreach (DataGridViewColumn dgviewColumn in dgView.Columns)
                {
                    // Excel work sheet indexing starts with 1
                    currentWorksheet.Cells[2, i] = dgviewColumn.Name;
                    ++i;
                }
                Microsoft.Office.Interop.Excel.Range headerColumnRange = currentWorksheet.get_Range("A2", "Q2");
                headerColumnRange.Font.Bold = true;
                headerColumnRange.Font.Color = 0xFF0000;
                //headerColumnRange.EntireColumn.AutoFit();
                int rowIndex = 0;
                for (rowIndex = 0; rowIndex < dgView.Rows.Count; rowIndex++)
                {
                    DataGridViewRow dgRow = dgView.Rows[rowIndex];
                    for (int cellIndex = 0; cellIndex < dgRow.Cells.Count; cellIndex++)
                    {
                        currentWorksheet.Cells[rowIndex + 3, cellIndex + 1] = dgRow.Cells[cellIndex].Value;
                    }
                    _procent = (int)(((double)_step / dgView.Rows.Count) * 100)%101;
                    if (progressBar_Save_Comp.Visible == true)
                    {
                        //bgw_File_Save_Comp.ReportProgress(_procent);
                        progressBar_Save_Comp.Value = _procent;
                    }
                    else if (progressBar_Save_Manual.Visible == true)
                    {
                        //bgw_File_Save_Manual.ReportProgress(_procent);
                        progressBar_Save_Comp.Value = _procent;
                    }
                    Notify_Label.Text = _procent.ToString()+"%";
                    _step++;
                }

                Microsoft.Office.Interop.Excel.Range fullTextRange = currentWorksheet.get_Range("A1", column_A_Q + (rowIndex + 1).ToString());
                fullTextRange.WrapText = true;
                fullTextRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                Microsoft.Office.Interop.Excel.ChartObject myChart = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                Microsoft.Office.Interop.Excel.Chart chartPage = myChart.Chart;
                Microsoft.Office.Interop.Excel.Range chartRange = currentWorksheet.get_Range("B2", column_A_Q + (rowIndex + 1).ToString());
                chartPage.ChartType = Excel.XlChartType.xlLineMarkers;
                chartPage.HasTitle = true;
                chartPage.ChartTitle.Text = text_input;
                chartPage.HasLegend = true;
                chartPage.SetSourceData(chartRange, Excel.XlRowCol.xlColumns);
                var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                yAxis.HasTitle = true;
                yAxis.AxisTitle.Text = "Temperature (°C)";
                //yAxis.MaximumScale = 20;
                yAxis.AxisTitle.Orientation = Excel.XlOrientation.xlUpward;

                //Excel.Range Data_Range = currentWorksheet.get_Range("B3", "Q10"); //+ (rowIndex + 1).ToString());//Data to be plotted in chart
                Excel.Range XVal_Range = currentWorksheet.get_Range("A3", "A" + (rowIndex + 1).ToString());//Catagory Names I want on X-Axis as range

                Excel.Axis xAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                xAxis.CategoryNames = XVal_Range;
                xAxis.HasTitle = true;
                xAxis.AxisTitle.Text = "Seconds";
                //chartPage.PlotBy = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns;

                //chartPage.ChartWizard(chartRange, Excel.XlChartType.xlLineMarkers, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing);

                chartPage.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet, "Chart");

                return currentWorkbook;
            }
            else
                return null;
        }

        public void create_File(Excel.Workbook work)
        {
            using (SaveFileDialog exportSaveFileDialog = new SaveFileDialog())
            {
                object misValue = System.Reflection.Missing.Value;
                exportSaveFileDialog.Title = "Select Excel File";
                exportSaveFileDialog.Filter = "Microsoft Office Excel Workbook(*.xlsx)|*.xlsx";

                if (DialogResult.OK == exportSaveFileDialog.ShowDialog())
                {
                    string fullFileName = exportSaveFileDialog.FileName;
                    // currentWorkbook.SaveCopyAs(fullFileName);
                    // indicating that we already saved the workbook, otherwise call to Quit() will pop up
                    // the save file dialogue box

                    work.SaveAs(fullFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, misValue, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
                    work.Saved = true;
                    MessageBox.Show("Exported successfully", "Exported to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        private void Save_btn_Click(object sender, EventArgs e)
        {
            _step = 1;
            _procent = 0;
            progressBar_Save_Comp.Value = 0;
            progressBar_Save_Comp.Visible = true;
            
            bgw_File_Save_Comp.RunWorkerAsync();
            
        }

        private void comboBox_Set_Analog_Box_Manual_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_Set_Analog_Box_Manual.SelectedIndex == 0)
            {
                Set_Analog_Box_Manual_1.Visible = true;
                Exposure_Time_Box_1_Manual.Visible = true;
                Set_Analog_Box_Manual_2.Visible = false;
                Set_Analog_Box_Manual_2.Text = "";
                Exposure_Time_Box_2_Manual.Visible = false;
                Exposure_Time_Box_2_Manual.Text = "";
                Set_Analog_Box_Manual_3.Visible = false;
                Set_Analog_Box_Manual_3.Text = "";
                Exposure_Time_Box_3_Manual.Visible = false;
                Exposure_Time_Box_3_Manual.Text = "";
                Set_Analog_Box_Manual_4.Visible = false;
                Set_Analog_Box_Manual_4.Text = "";
                Exposure_Time_Box_4_Manual.Visible = false;
                Exposure_Time_Box_4_Manual.Text = "";
            }
            else if (comboBox_Set_Analog_Box_Manual.SelectedIndex == 1)
            {
                Set_Analog_Box_Manual_1.Visible = true;
                Exposure_Time_Box_1_Manual.Visible = true;
                Set_Analog_Box_Manual_2.Visible = true;
                Exposure_Time_Box_2_Manual.Visible = true;
                Set_Analog_Box_Manual_3.Visible = false;
                Set_Analog_Box_Manual_3.Text = "";
                Exposure_Time_Box_3_Manual.Visible = false;
                Exposure_Time_Box_3_Manual.Text = "";
                Set_Analog_Box_Manual_4.Visible = false;
                Set_Analog_Box_Manual_4.Text = "";
                Exposure_Time_Box_4_Manual.Visible = false;
                Exposure_Time_Box_4_Manual.Text = "";
            }
            else if(comboBox_Set_Analog_Box_Manual.SelectedIndex == 2)
            {
                Set_Analog_Box_Manual_1.Visible = true;
                Exposure_Time_Box_1_Manual.Visible = true;
                Set_Analog_Box_Manual_2.Visible = true;
                Exposure_Time_Box_2_Manual.Visible = true;
                Set_Analog_Box_Manual_3.Visible = true;
                Exposure_Time_Box_3_Manual.Visible = true;
                Set_Analog_Box_Manual_4.Visible = false;
                Set_Analog_Box_Manual_4.Text = "";
                Exposure_Time_Box_4_Manual.Visible = false;
                Exposure_Time_Box_4_Manual.Text = "";
            }
            else if (comboBox_Set_Analog_Box_Manual.SelectedIndex == 3)
            {
                Set_Analog_Box_Manual_1.Visible = true;
                Exposure_Time_Box_1_Manual.Visible = true;
                Set_Analog_Box_Manual_2.Visible = true;
                Exposure_Time_Box_2_Manual.Visible = true;
                Set_Analog_Box_Manual_3.Visible = true;
                Exposure_Time_Box_3_Manual.Visible = true;
                Set_Analog_Box_Manual_4.Visible = true;
                Exposure_Time_Box_4_Manual.Visible = true;
            }
        }

        private void CHART_ON_Manual_CheckedChanged(object sender, EventArgs e)
        {
            if (CHART_ON_Manual.Checked == true)
            {
                chart_CTS_Manual.Visible = true;
                chart_PICO_Manual.Visible = true;
            }
            dataGridView_PICO_Manual.Visible = false;
            dataGridView_CTS_Manual.Visible = false;
        }

        private void GRID_ON_Manual_CheckedChanged(object sender, EventArgs e)
        {
            if (GRID_ON_Manual.Checked == true)
            {
                dataGridView_PICO_Manual.Visible = true;
                dataGridView_CTS_Manual.Visible = true;
            }
            chart_CTS_Manual.Visible = false;
            chart_PICO_Manual.Visible = false;
        }

        private void Start_Btn_Manual_Click(object sender, EventArgs e)
        {
            string set_etemp = string.Empty;
            string start = string.Empty;
            if (timer_Comp_Exposure.Enabled == false && timer_Manual.Enabled == false)
            {
                Notify_Label.Text = "-";
                label_manual_Finished.Visible = false;
                if (Ethernet_label.Text != "Connected")
                {
                    if (COM_Portar.Length > 0)
                    {
                        if (chart_CTS_Manual.Series.Count < 1)
                        {
                            this.chart_CTS_Manual.Series.Add(ChartSeries1);
                            this.chart_CTS_Manual.Series.Add(ChartSeries2);
                            ChartSeries1.Color = System.Drawing.Color.Green;
                            ChartSeries2.Color = System.Drawing.Color.Red;
                        }
                        if (Exposure_Time_Box_1_Manual.Text == "0")
                        {
                            continuous = true;
                            Notify_Label.Text = "Continuous measurements";
                        }
                        else
                            timer_expose_manual1.Interval = (int.Parse(Exposure_Time_Box_1_Manual.Text) * 1000);
                        if (Exposure_Time_Box_2_Manual.Visible)
                            timer_expose_manual2.Interval = (int.Parse(Exposure_Time_Box_2_Manual.Text) * 1000);
                        if (Exposure_Time_Box_3_Manual.Visible)
                            timer_expose_manual3.Interval = (int.Parse(Exposure_Time_Box_2_Manual.Text) * 1000);
                        if (Exposure_Time_Box_4_Manual.Visible)
                            timer_expose_manual4.Interval = (int.Parse(Exposure_Time_Box_2_Manual.Text) * 1000);
                        serialPort_CTS.PortName = ComPort_Box.SelectedItem.ToString();
                        exposure_Time = long.Parse(Exposure_Time_Box_1_Manual.Text);
                        serialPort_CTS.Open();
                        if (serialPort_CTS.IsOpen)
                        {
                            //Set Tempeture on Chamber
                            check_Set = SET_Temp_CTS(Set_Analog_Box_Manual_1.Text);
                            //Start chamber
                            Start_CTS();
                        }
                    }
                    else
                        Notify_Label.Text = "No COM Port! ";
                }
                if (Ethernet_label.Text == "Connected")
                {
                    if (chart_CTS_Manual.Series.Count < 1)
                    {
                        this.chart_CTS_Manual.Series.Add(ChartSeries1);
                        this.chart_CTS_Manual.Series.Add(ChartSeries2);
                        ChartSeries1.Color = System.Drawing.Color.Green;
                        ChartSeries2.Color = System.Drawing.Color.Red;
                    }
                    if (Exposure_Time_Box_1_Manual.Text == "0")
                    {
                        continuous = true;
                        Notify_Label.Text = "Continuous measurements";
                    }
                    else
                        timer_expose_manual1.Interval = (int.Parse(Exposure_Time_Box_1_Manual.Text) * 1000);
                    if (Exposure_Time_Box_2_Manual.Visible)
                        timer_expose_manual2.Interval = (int.Parse(Exposure_Time_Box_2_Manual.Text) * 1000);
                    if (Exposure_Time_Box_3_Manual.Visible)
                        timer_expose_manual3.Interval = (int.Parse(Exposure_Time_Box_2_Manual.Text) * 1000);
                    if (Exposure_Time_Box_4_Manual.Visible)
                        timer_expose_manual4.Interval = (int.Parse(Exposure_Time_Box_2_Manual.Text) * 1000);
                    exposure_Time = long.Parse(Exposure_Time_Box_1_Manual.Text);
                    if (Set_Analog_Box_Manual_1.Text.Length == 2)
                    {
                        set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box_Manual_1.Text + ".0");
                    }
                    else
                        set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box_Manual_1.Text + ".0");
                    Thread.Sleep(100);
                    start = Read_Write_Ethernet(Ethernet_Box.Text, "s1 1");
                }
                PICO_Counter = 0;
                Timer_Counter = 0;
                stable_counter = 0;
                stopWatch.Reset();
                temp_Reached1 = false;
                temp_Reached2 = false;
                temp_Reached3 = false;
                temp_Reached4 = false;
                time1_manual_reached = false;
                time2_manual_reached = false;
                time3_manual_reached = false;
                time4_manual_reached = false;
                disable_selections();
                Set_Analog_Box_Manual_1.ReadOnly = true;
                Set_Analog_Box_Manual_2.ReadOnly = true;
                Set_Analog_Box_Manual_3.ReadOnly = true;
                Set_Analog_Box_Manual_4.ReadOnly = true;
                Exposure_Time_Box_1_Manual.ReadOnly = true;
                Exposure_Time_Box_2_Manual.ReadOnly = true;
                Exposure_Time_Box_3_Manual.ReadOnly = true;
                Exposure_Time_Box_4_Manual.ReadOnly = true;
                SET_Chart_Grid();
                SetChannels();

                timer_Manual.Enabled = true;
            }
            else
                Notify_Label.Text = "Another test is running!";
        }

        private void Stop_Btn_Manual_Click(object sender, EventArgs e)
        {
            string stop = string.Empty;
            if (Ethernet_label.Text != "Connected")
            {
                timer_Manual.Enabled = false;
                timer_expose_manual1.Enabled = false;
                timer_expose_manual2.Enabled = false;
                timer_expose_manual3.Enabled = false;
                timer_expose_manual4.Enabled = false;
                serialPort_CTS.ReceivedBytesThreshold = 8;
                data_Tx[0] = Convert.ToByte("0x02", 16);
                data_Tx[1] = Convert.ToByte("0x81", 16);
                data_Tx[2] = Convert.ToByte("0xF0", 16);
                data_Tx[3] = Convert.ToByte("0xB0", 16);
                data_Tx[4] = Convert.ToByte("0xB0", 16);
                data_Tx[5] = Convert.ToByte("0xB0", 16);
                data_Tx[6] = Convert.ToByte("0xC1", 16);
                data_Tx[7] = Convert.ToByte("0x03", 16);
                serialPort_CTS.Write(data_Tx, 0, 8);
                serialPort_CTS.Close();
            }
            if (Ethernet_label.Text == "Connected")
            {
                stop = Read_Write_Ethernet(Ethernet_Box.Text, "s1 0");
            }
            timer_Manual.Enabled = false;
            enable_selections();
            Set_Analog_Box_Manual_1.ReadOnly = false;
            Set_Analog_Box_Manual_2.ReadOnly = false;
            Set_Analog_Box_Manual_3.ReadOnly = false;
            Set_Analog_Box_Manual_4.ReadOnly = false;
            Exposure_Time_Box_1_Manual.ReadOnly = false;
            Exposure_Time_Box_2_Manual.ReadOnly = false;
            Exposure_Time_Box_3_Manual.ReadOnly = false;
            Exposure_Time_Box_4_Manual.ReadOnly = false;
        }

        private void timer_Manual_Tick(object sender, EventArgs e)
        {
            string set_etemp = string.Empty;
            stopWatch.Stop();
            long duration = stopWatch.ElapsedMilliseconds;
            stopWatch.Reset();
            Timer_Counter += Math.Round((decimal)duration / 1000, 1);
            
            stopWatch.Start();
            int Selected = 2;
            row_PICO = Data_Table_PICO_Manual.NewRow();
            row_PICO["Time"] = Timer_Counter;
            
            this.PICO1_Thread = new Thread(new ThreadStart(this.DoWork1));
            this.PICO1_Thread.IsBackground = true; // this prevents the extra thread from blocking an application shutdown
            this.PICO1_Thread.Start();
            this.PICO2_Thread = new Thread(new ThreadStart(this.DoWork2));
            this.PICO2_Thread.IsBackground = true; // this prevents the extra thread from blocking an application shutdown
            this.PICO2_Thread.Start();

            while (PICO1_Thread.IsAlive || PICO2_Thread.IsAlive)
            {
                ;
            }
            ThreadFinished1(tempbuffer1, status1);
            ThreadFinished2(tempbuffer2, status2);
            if (serialPort_CTS.IsOpen || Ethernet_label.Text == "Connected")
                Logg_temp(Selected);

            Data_Table_PICO_Manual.Rows.Add(row_PICO);
            dataGridView_PICO_Manual.DataSource = Data_Table_PICO_Manual;
            dataGridView_PICO_Manual.FirstDisplayedScrollingRowIndex = dataGridView_PICO_Manual.Rows.Count - 1;
            //Start conditions
            if (continuous == false)
            {
                //Step1 = Temperature1 reached and start Timer1
                if (temp_Reached1 == false && temp_Reached2 == false && temp_Reached3 == false && temp_Reached4 == false && time1_manual_reached == false)
                {
                    if (stabled() == true) //Check if temperature is stable
                    {
                        timer_expose_manual1.Enabled = true;
                        temp_Reached1 = true;
                        chart_CTS_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        chart_PICO_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        stable_counter = 0;
                    }
                }
                //Step2 = Timer1 finished and SET new temperature
                if (time1_manual_reached == true && time2_manual_reached == false && time3_manual_reached == false && time4_manual_reached == false)
                {
                    timer_expose_manual1.Enabled = false;
                    if (Set_Analog_Box_Manual_2.Visible)
                    {
                        if (Ethernet_label.Text != "Connected")
                        {
                            check_Set = SET_Temp_CTS(Set_Analog_Box_Manual_2.Text);
                        }
                        else if (Ethernet_label.Text == "Connected")
                        {
                            if (Set_Analog_Box_Manual_2.Text.Length == 2)
                            {
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box_Manual_2.Text + ".0");
                            }
                            else
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box_Manual_2.Text + ".0");
                        }
                    }
                    else
                    {
                        label_manual_Finished.Visible = true;
                        Stop_Btn_Manual_Click(null, null);
                    }

                }
                //Step3 Temperature2 reached and start Timer2
                if (timer_expose_manual1.Enabled == false && temp_Reached1 == true && temp_Reached2 == false && temp_Reached3 == false && temp_Reached4 == false && time1_manual_reached == true && time2_manual_reached == false)
                {
                    if (stabled() == true) //Check if temperature is stable
                    {
                        timer_expose_manual2.Enabled = true;
                        temp_Reached2 = true;
                        chart_CTS_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        chart_PICO_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        stable_counter = 0;
                    }
                }
                //Step4 = Timer2 finished and SET new temperature
                if (time1_manual_reached == true && time2_manual_reached == true && time3_manual_reached == false && time4_manual_reached == false)
                {
                    timer_expose_manual2.Enabled = false;
                    if (Set_Analog_Box_Manual_3.Visible)
                    {
                        if (Ethernet_label.Text != "Connected")
                        {
                            check_Set = SET_Temp_CTS(Set_Analog_Box_Manual_3.Text);
                        }
                        else if (Ethernet_label.Text == "Connected")
                        {
                            if (Set_Analog_Box_Manual_3.Text.Length == 2)
                            {
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box_Manual_3.Text + ".0");
                            }
                            else
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box_Manual_3.Text + ".0");
                        }
                    }
                    else
                    {
                        label_manual_Finished.Visible = true;
                        Stop_Btn_Manual_Click(null, null);
                    }
                }
                //Step5 Temperature3 reached and start Timer3
                if (timer_expose_manual2.Enabled == false && temp_Reached1 == true && temp_Reached2 == true && temp_Reached3 == false && temp_Reached4 == false && time1_manual_reached == true && time2_manual_reached == true && time3_manual_reached == false)
                {
                    if (stabled() == true) //Check if temperature is stable
                    {
                        timer_expose_manual3.Enabled = true;
                        temp_Reached3 = true;
                        chart_CTS_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        chart_PICO_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        stable_counter = 0;
                    }
                }
                //Step6 = Timer3 finished and SET new temperature
                if (time1_manual_reached == true && time2_manual_reached == true && time3_manual_reached == true && time4_manual_reached == false)
                {
                    timer_expose_manual3.Enabled = false;
                    if (Set_Analog_Box_Manual_4.Visible)
                    {
                        if (Ethernet_label.Text != "Connected")
                        {
                            check_Set = SET_Temp_CTS(Set_Analog_Box_Manual_4.Text);
                        }
                        else if (Ethernet_label.Text == "Connected")
                        {
                            if (Set_Analog_Box_Manual_4.Text.Length == 2)
                            {
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box_Manual_4.Text + ".0");
                            }
                            else
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box_Manual_4.Text + ".0");
                        }
                    }
                    else
                    {
                        label_manual_Finished.Visible = true;
                        Stop_Btn_Manual_Click(null, null);
                    }
                }
                //Step7 Temperature4 reached and start Timer4
                if (timer_expose_manual3.Enabled == false && temp_Reached1 == true && temp_Reached2 == true && temp_Reached3 == true && temp_Reached4 == false && time1_manual_reached == true && time2_manual_reached == true && time3_manual_reached == true && time4_manual_reached == false)
                {
                    if (stabled() == true) //Check if temperature is stable
                    {
                        timer_expose_manual4.Enabled = true;
                        temp_Reached4 = true;
                        chart_CTS_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        chart_PICO_Manual.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                        stable_counter = 0;
                    }
                }
                //Step8 = Timer4 finished, stop chamber and measurements
                if (time1_manual_reached == true && time2_manual_reached == true && time3_manual_reached == true && time4_manual_reached == true)
                {
                    timer_expose_manual4.Enabled = false;
                    label_manual_Finished.Visible = true;
                    Stop_Btn_Manual_Click(null, null);
                }
            }
            richTextBox_Manual.Text = "Continuous? " + continuous.ToString() + "\n" + "Temp1 reached? " + temp_Reached1.ToString() + "\n" + "Time1 reached? " + time1_manual_reached.ToString() + "\n" + "Temp2 reached? " + temp_Reached2.ToString() + "\n" + "Time2 reached? " + time2_manual_reached.ToString() + "\n" + "Temp3 reached? " + temp_Reached3.ToString() + "\n" + "Time3 reached? " + time3_manual_reached.ToString() + "\n" + "Temp4 reached? " + temp_Reached4.ToString() +"\n"+ "Time4 reached? " + time4_manual_reached.ToString();

            PICO_Counter += 1;
        }

        private void Clear_All_Click(object sender, EventArgs e)
        {
            
            if (timer_Comp_Exposure.Enabled == false && timer_Manual.Enabled == false)
            {
                int n1, n2;
                label_Comp_Finished.Visible = false;
                label_manual_Finished.Visible = false;
                Exposure_Time_Label1.Text = "-";
                Set_analog_minus_10_Label1.Text = "-";
                Notify_Label.Text = "-";
                dataGridView_CTS_Comp.DataSource = null;
                dataGridView_CTS_Manual.DataSource = null;
                dataGridView_PICO_Comp.DataSource = null;
                dataGridView_PICO_Manual.DataSource = null;
                
                if (chart_PICO_Comp.Series.Count > 0)
                {
                    for (n1 = chart_PICO_Comp.Series.Count - 1; n1 >= 0; n1--)
                        chart_PICO_Comp.Series[n1].Points.Clear();
                }

                if (chart_PICO_Manual.Series.Count > 0)
                {
                    for (n2 = chart_PICO_Manual.Series.Count - 1; n2 >= 0; n2--)
                        chart_PICO_Manual.Series[n2].Points.Clear();
                }
                //Ingen snygg lösning!!
                if (Channel1_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel1.Text = "";
                    Channel1_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel2_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel2.Text = "";
                    Channel2_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel3_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel3.Text = "";
                    Channel3_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel4_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel4.Text = "";
                    Channel4_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel5_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel5.Text = "";
                    Channel5_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel6_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel6.Text = "";
                    Channel6_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel7_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel7.Text = "";
                    Channel7_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel8_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel8.Text = "";
                    Channel8_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel9_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel9.Text = "";
                    Channel9_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel10_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel10.Text = "";
                    Channel10_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel11_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel11.Text = "";
                    Channel11_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel12_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel12.Text = "";
                    Channel12_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel13_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel13.Text = "";
                    Channel13_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel14_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel14.Text = "";
                    Channel14_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel15_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel15.Text = "";
                    Channel15_Box.CheckState = CheckState.Unchecked;
                }
                if (Channel16_Box.CheckState == CheckState.Checked)
                {
                    textBox_Channel16.Text = "";
                    Channel16_Box.CheckState = CheckState.Unchecked;
                }
                if (chart_CTS_Comp.Series.Count > 0)
                {
                    chart_CTS_Comp.Series[0].Points.Clear();
                    chart_CTS_Comp.Series[1].Points.Clear();
                    chart_CTS_Comp.Series.Remove(ChartSeries1);
                    chart_CTS_Comp.Series.Remove(ChartSeries2);
                }
                if (chart_CTS_Manual.Series.Count > 0)
                {
                    chart_CTS_Manual.Series[0].Points.Clear();
                    chart_CTS_Manual.Series[1].Points.Clear();
                    chart_CTS_Manual.Series.Remove(ChartSeries1);
                    chart_CTS_Manual.Series.Remove(ChartSeries2);
                }
            }
            else 
            {
                Notify_Label.Text = "Cant clear inputs during measurements!";
            }
        }

        private void Save_Btn_Manual_Click(object sender, EventArgs e)
        {
            _step = 1;
            _procent = 0;
            progressBar_Save_Manual.Value = 0;
            progressBar_Save_Manual.Visible = true;
            bgw_File_Save_Manual.RunWorkerAsync();
        }

        private void bgw_File_Save_Comp_DoWork(object sender, DoWorkEventArgs e)
        {
            string tab_text = tabControl1.SelectedTab.Text;
            newBook = ExportToExcel2(dataGridView_PICO_Comp, tab_text);
            //bgw_File_Save_Comp.ReportProgress(_procent);
        }

        private void bgw_File_Save_Comp_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            create_File(newBook);
            progressBar_Save_Comp.Visible = false;
        }
        
        public void ImportFile()
        {
            try
            {

                Excel.Workbook ExWorkbook;
                Excel.Worksheet ExWorksheet;
                Excel.Range ExRange;
                Excel.Application ExObj = new Excel.Application();

                DataTable dt = new DataTable("dataTable");
                DataSet dsSource = new DataSet("dataSet");
                dt.Reset();

                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                DialogResult result = openFileDialog1.ShowDialog();

                if (result == DialogResult.OK) // Test result.
                {
                    ExWorkbook = ExObj.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ExWorksheet = (Excel.Worksheet)ExWorkbook.Sheets.get_Item(2);
                    ExRange = ExWorksheet.UsedRange;

                    for (int Cnum = 1; Cnum <= ExRange.Columns.Count; Cnum++)
                    {
                        dt.Columns.Add(new DataColumn((ExRange.Cells[2, Cnum] as Excel.Range).Value2.ToString()));
                    }
                    dt.AcceptChanges();

                    string[] columnNames = new String[dt.Columns.Count];
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        columnNames[0] = dt.Columns[i].ColumnName;
                    }

                    for (int Rnum = 3; Rnum <= ExRange.Rows.Count; Rnum++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int Cnum = 1; Cnum <= ExRange.Columns.Count; Cnum++)
                        {
                            if ((ExRange.Cells[Rnum, Cnum] as Excel.Range).Value2 != null)
                            {
                                dr[Cnum - 1] = (ExRange.Cells[Rnum, Cnum] as Excel.Range).Value2.ToString();
                            }
                        }
                        dt.Rows.Add(dr);
                        dt.AcceptChanges();
                        progressBar_Save_Comp.Value = (int)((Rnum / ExRange.Rows.Count) * 100) % 101;
                    }
                    ExWorkbook.Close(true, Type.Missing, Type.Missing);
                    ExObj.Quit();

                    dataGridView_PICO_Comp.DataSource = dt;
                }
            }
            catch(Exception ex)
            {
                string _exeption_read = ex.Message;
            }
        }

        private void Reset_btn_Click(object sender, EventArgs e)
        {
            if (timer_Comp_Exposure.Enabled == false && timer_Manual.Enabled == false)
            {
                Application.Restart();
            }
            else
                Notify_Label.Text = "Measurements are running!";
        }

        public bool stabled()
        {
            double[] control_values = new double[10];
            int i, j;
            for (i = 0; i < control_values.Length; i++)
                control_values[i] = 0.5;
            if ((ActTemp[PICO_Counter] - 0.1) == SetTemp[PICO_Counter] || (ActTemp[PICO_Counter] + 0.1) == SetTemp[PICO_Counter] || ActTemp[PICO_Counter] == SetTemp[PICO_Counter])
            {
                control_values[stable_counter] = ActTemp[PICO_Counter];
                stable_counter++;
                richTextBox1.Clear();
                Temp_Rich.Clear();
                for (j = 0; j < control_values.Length; j++)
                {
                    richTextBox1.Text += control_values[j].ToString() + "\n";
                    Temp_Rich.Text += control_values[j].ToString() + "\n";
                }
                if (stable_counter == control_values.Length)
                    stable_counter = 0;
                if (control_values[9] == SetTemp[PICO_Counter] || (control_values[9] - 0.1)== SetTemp[PICO_Counter] || (control_values[9] + 0.1) == SetTemp[PICO_Counter])
                {
                    return true;
                }
                else
                    return false;
            }
            else
                return false;
        }

        private void Notify_Label_Click(object sender, EventArgs e)
        {
            Notify_Label.Text = "-";
        }

        private void MAIN_FORM_Click(object sender, EventArgs e)
        {
            this.TopLevel = true;
        }

        private void Email_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Email_Box.Text == "HELP")
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    Timing_Box1.Visible = true;
                    richTextBox_Manual.Visible = true;
                    richTextBox1.Visible = true;
                    Temp_Rich.Visible = true;
                }
            }
            else
            {
                Timing_Box1.Visible = false;
                richTextBox_Manual.Visible = false;
                richTextBox1.Visible = false;
                Temp_Rich.Visible = false;
            }

        }

        public void Connect_to(String server)
        {
            try
            {
                int port = 1080;
                TcpClient client = new TcpClient(server, port);
                
                if (client.Connected == true)
                    Ethernet_label.Text = "Connected";
                
                client.Close();
            }
            catch (ArgumentNullException e)
            {
                MessageBox.Show(e.Message);
                Ethernet_label.Text = "Not Connected";
            }
            catch (SocketException e)
            {
                MessageBox.Show(e.Message);
                Ethernet_label.Text = "Not Connected";
            }
        }
            public string Read_Write_Ethernet(String server, String message)
        {
            try
            {
                // Create a TcpClient. 
                // Note, for this client to work you need to have a TcpServer  
                // connected to the same address as specified by the server, port 
                // combination.
                int port = 1080;
                TcpClient client = new TcpClient(server, port);
                bool connected_client = client.Connected;
                if (connected_client == true)
                    Ethernet_label.Text = "Connected";
                else
                    Ethernet_label.Text = "Not Connected";
                // Translate the passed message into ASCII and store it as a Byte array.
                Byte[] data = System.Text.Encoding.ASCII.GetBytes(message);

                // Get a client stream for reading and writing. 
                //  Stream stream = client.GetStream();

                NetworkStream stream = client.GetStream();

                // Send the message to the connected TcpServer. 
                stream.Write(data, 0, data.Length);
                Thread.Sleep(100);
                // Receive the TcpServer.response. 

                // Buffer to store the response bytes.
                data = new Byte[256];

                // String to store the response ASCII representation.
                String responseData = String.Empty;

                // Read the first batch of the TcpServer response bytes.
                Int32 bytes = stream.Read(data, 0, data.Length);
                responseData = System.Text.Encoding.ASCII.GetString(data, 0, bytes);
                //richTextBox1.Text = responseData;
                
                // Close everything.
                stream.Close();
                client.Close();
                return responseData;
            }
            catch (ArgumentNullException e)
            {
                string nullex = e.Message;
                Ethernet_label.Text = "Not Connected";
                return "0";
            }
            catch (SocketException e)
            {
                string sockex = e.Message;
                Ethernet_label.Text = "Not Connected";
                return "0";
            }
        }

        private void Ethernet_btn_Click(object sender, EventArgs e)
        {
            if (Ethernet_label.Text != "Connected")
            {
                Connect_to(Ethernet_Box.Text);
            }
            else
                Ethernet_label.Text = "Not Connected";
        }

        private void bgw_File_Save_Manual_DoWork(object sender, DoWorkEventArgs e)
        {
            string tab_text = tabControl1.SelectedTab.Text;
            newBook = ExportToExcel2(dataGridView_PICO_Manual, tab_text);
        }

        

        private void bgw_File_Save_Manual_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            create_File(newBook);
            progressBar_Save_Manual.Visible = false;
        }

        private void timer_Comp_Automatic_Tick(object sender, EventArgs e)
        {
            string stop_Chamber = string.Empty, start = string.Empty, set_etemp = string.Empty;
            stopWatch.Stop();
            long duration = stopWatch.ElapsedMilliseconds;
            stopWatch.Reset();
            Timer_Counter += Math.Round((decimal)duration / 1000, 1);
            Timing_Box1.Text = duration.ToString() + " turn_off:" + turn_off.ToString();
            stopWatch.Start();
            int Selected = 1;
            row_PICO = Data_Table_PICO_Comp.NewRow();
            row_PICO["Time"] = Timer_Counter;
            this.PICO1_Thread = new Thread(new ThreadStart(this.DoWork1));
            this.PICO1_Thread.IsBackground = true; // this prevents the extra thread from blocking an application shutdown
            this.PICO1_Thread.Start();
            this.PICO2_Thread = new Thread(new ThreadStart(this.DoWork2));
            this.PICO2_Thread.IsBackground = true; // this prevents the extra thread from blocking an application shutdown
            this.PICO2_Thread.Start();

            while (PICO1_Thread.IsAlive || PICO2_Thread.IsAlive)
            {
                ;
            }
            ThreadFinished1(tempbuffer1, status1);
            ThreadFinished2(tempbuffer2, status2);
            if (serialPort_CTS.IsOpen || Ethernet_label.Text == "Connected")
            {
                logger_watch.Start();
                Stop_CTS_watch.Start();
                Logg_temp(Selected);
            }

            Data_Table_PICO_Comp.Rows.Add(row_PICO);
            dataGridView_PICO_Comp.DataSource = Data_Table_PICO_Comp;
            dataGridView_PICO_Comp.FirstDisplayedScrollingRowIndex = dataGridView_PICO_Comp.Rows.Count - 1;//(int)test_Counter;

            if (turn_off == false && temp_Reached == false)
            {
                if (stabled() == true)
                {

                    /*timer_expose.Interval = exposure_Time * 1000;
                    timer_expose.Enabled = true;*/
                    DateTime expose = DateTime.Now;
                    //DateTime exposed = expose.AddSeconds(exposure_Time);
                    if(Exposure_Time_Label1.Text == "-")
                        Exposure_Time_Label1.Text = "Temperature reached " + Set_Analog_Box1_Comp.Text + "\nStarted: " + expose.ToString();
                    else
                        Exposure_Time_Label2.Text = "Temperature reached " + Set_Analog_Box2_Comp.Text + "\nStarted: " + expose.ToString();
                    chart_CTS_Comp.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                    chart_PICO_Comp.ChartAreas[0].CursorX.Position = (double)Timer_Counter;
                    temp_Reached = true;
                }
            }
            if (temp_Reached == true && timer_start == false)
            {
                timer_expose_60.Enabled = true;
                timer_start = true;
            }
            if (time_60_reached == true && turn_off == false)
            {
                if (get_10_count < 10)
                {
                    get_all_Temp_Meas(tempbuffer1, tempbuffer2);
                    get_10_count++;
                }
                else
                {
                    
                    Tempcheck1[comp_i] = calc_temp(TempChannel1);
                    Tempcheck2[comp_i] = calc_temp(TempChannel2);
                    Tempcheck3[comp_i] = calc_temp(TempChannel3);
                    Tempcheck4[comp_i] = calc_temp(TempChannel4);
                    Tempcheck5[comp_i] = calc_temp(TempChannel5);
                    Tempcheck6[comp_i] = calc_temp(TempChannel6);
                    Tempcheck7[comp_i] = calc_temp(TempChannel7);
                    Tempcheck8[comp_i] = calc_temp(TempChannel8);
                    Tempcheck9[comp_i] = calc_temp(TempChannel9);
                    Tempcheck10[comp_i] = calc_temp(TempChannel10);
                    Tempcheck11[comp_i] = calc_temp(TempChannel11);
                    Tempcheck12[comp_i] = calc_temp(TempChannel12);
                    Tempcheck13[comp_i] = calc_temp(TempChannel13);
                    Tempcheck14[comp_i] = calc_temp(TempChannel14);
                    Tempcheck15[comp_i] = calc_temp(TempChannel15);
                    Tempcheck16[comp_i] = calc_temp(TempChannel16);
                    Temp_Rich.Text = Tempcheck1[comp_i].ToString() + " " + Tempcheck2[comp_i].ToString() + " " + Tempcheck3[comp_i].ToString() + " " + Tempcheck4[comp_i].ToString() + " " + Tempcheck5[comp_i].ToString() + " " + Tempcheck6[comp_i].ToString() + " " + Tempcheck7[comp_i].ToString() + " " + Tempcheck8[comp_i].ToString() + " " + Tempcheck9[comp_i].ToString() + " " + Tempcheck10[comp_i].ToString() + " " + Tempcheck11[comp_i].ToString() + " " + Tempcheck12[comp_i].ToString() + " " + Tempcheck13[comp_i].ToString() + " " + Tempcheck14[comp_i].ToString() + " " + Tempcheck15[comp_i].ToString() + " " + Tempcheck16[comp_i].ToString();
                    if (comp_i == 1)
                    {
                        if (compare_meas(Tempcheck1) == true && compare_meas(Tempcheck2) == true && compare_meas(Tempcheck3) == true && compare_meas(Tempcheck4) == true && compare_meas(Tempcheck5) == true && compare_meas(Tempcheck6) == true && compare_meas(Tempcheck7) == true && compare_meas(Tempcheck8) == true && compare_meas(Tempcheck9) == true && compare_meas(Tempcheck10) == true && compare_meas(Tempcheck11) == true && compare_meas(Tempcheck12) == true && compare_meas(Tempcheck13) == true && compare_meas(Tempcheck14) == true && compare_meas(Tempcheck15) == true && compare_meas(Tempcheck16) == true)
                        {
                            turn_off_chamber = true;
                        }
                    }
                    time_60_reached = false;
                    get_10_count = 0;
                    comp_i = (comp_i + 1) % 2;
                    
                }
            }
            Stop_CTS_watch.Stop();
            long logger_time = logger_watch.ElapsedMilliseconds;
            long stop_CTS_time = Stop_CTS_watch.ElapsedMilliseconds;
            Timing_Box1.Text += " logg: " + logger_time.ToString() + " Stop: " + stop_CTS_time.ToString();
            logger_watch.Reset();
            Stop_CTS_watch.Reset();
            if (Set_Analog_Box1_Comp.Text.Contains("-"))
            {
                negativ1 = true;
            }
            else
                negativ1 = false;
            if (Set_Analog_Box2_Comp.Text.Contains("-"))
            {
                negativ2 = true;
            }
            else
                negativ2 = false;
            if ((turn_off_chamber == true && turn_off == false && negativ1 == false) || (turn_off_chamber == true && turn_off == false && negativ2 == false && cycle_Comp == true))//Set_Analog_Box1.Text.Substring(0,1) is to see if a negativ number has been entered
            {
                if (Ethernet_label.Text != "Connected")
                {
                    serialPort_CTS.ReceivedBytesThreshold = 8;
                    data_Tx[0] = Convert.ToByte("0x02", 16);
                    data_Tx[1] = Convert.ToByte("0x81", 16);
                    data_Tx[2] = Convert.ToByte("0xF0", 16);
                    data_Tx[3] = Convert.ToByte("0xB0", 16);
                    data_Tx[4] = Convert.ToByte("0xB0", 16);
                    data_Tx[5] = Convert.ToByte("0xB0", 16);
                    data_Tx[6] = Convert.ToByte("0xC1", 16);
                    data_Tx[7] = Convert.ToByte("0x03", 16);
                    serialPort_CTS.Write(data_Tx, 0, 8);
                    Thread.Sleep(100);
                }
                
                if (Ethernet_label.Text == "Connected")
                    stop_Chamber = Read_Write_Ethernet(Ethernet_Box.Text, "s1 0");
                turn_off = true;
                turn_off_chamber = false;
            }
            //If a negative number has been entered in Set_Analog_Box1_Comp or Set_Analog_Box2_Comp, check if it shall cycle to next POSITIV value or stop measuring.
            else if ((turn_off_chamber == true && turn_off == false && Set_Analog_Box1_Comp.Text.Contains("-")) || (turn_off_chamber == true && turn_off == false && Set_Analog_Box2_Comp.Text.Contains("-")))
            {
                if ((comboBox_Comp.SelectedIndex == 0) || (cycle_Comp == true && Set_Analog_Box2_Comp.Text.Contains("-")))
                {
                    if (comboBox_Comp.SelectedIndex == 0)
                    {
                        DateTime exposed1 = DateTime.Now;
                        Exposure_Time_Label1.Text += "\nEnded: " + exposed1;
                    }
                    if (cycle_Comp == true && Set_Analog_Box2_Comp.Text.Contains("-"))
                    {
                        DateTime exposed1 = DateTime.Now;
                        Exposure_Time_Label2.Text += "\nEnded: " + exposed1;
                    }
                    Stop_Btn_Comp_Click(null, null);
                    label_Comp_Finished.Visible = true;
                    if (Email_ON.Checked == true && Email_Box.Text != "")
                    {
                        Send_Cycle_Finished();
                    }
                }
                else if (comboBox_Comp.SelectedIndex == 1)
                {
                    turn_off = false;
                    temp_Reached = false;
                    turn_off_chamber = false;
                    if (Ethernet_label.Text != "Connected")
                    {
                        //Set Tempeture on Chamber
                        check_Set = SET_Temp_CTS(Set_Analog_Box2_Comp.Text);
                        //Start chamber
                        Start_CTS();
                        cycle_Comp = true;
                    }
                    if (Ethernet_label.Text == "Connected")
                    {
                        if (Set_Analog_Box1_Comp.Text.Length == 2)
                        {
                            set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box2_Comp.Text + ".0");
                        }
                        else
                            set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box2_Comp.Text + ".0");
                        Thread.Sleep(100);
                        start = Read_Write_Ethernet(Ethernet_Box.Text, "s1 1");
                        cycle_Comp = true;
                    }
                }
            }
            if (turn_off == true)
            {
                if (Set_analog_minus_10_Label1.Text == "-" && negativ1 == false)
                {
                    Set_analog_minus_10_Label1.Text = "Measure until tempchamber reaches: " + (double.Parse(Set_Analog_Box1_Comp.Text) - 10).ToString();
                }
                else if (Exposure_Time_Label2.Text != "-" && !(Exposure_Time_Label2.Text.Contains("Ended")) && negativ2 == false)
                {
                    Set_analog_minus_10_Label2.Text = "Measure until tempchamber reaches: " + (double.Parse(Set_Analog_Box2_Comp.Text) - 10).ToString();
                }
                if (ActTemp[PICO_Counter] <= double.Parse(Set_Analog_Box1_Comp.Text) - 10 && cycle_Comp == false)
                {
                    //Stop measurement if cycle value is 1
                    if (comboBox_Comp.SelectedIndex == 0)
                    {
                        DateTime exposed1 = DateTime.Now;
                        Exposure_Time_Label1.Text += "\nEnded: " + exposed1;
                        Stop_Btn_Comp_Click(null, null);
                        label_Comp_Finished.Visible = true;
                        if (Email_ON.Checked == true && Email_Box.Text != "")
                        {
                            Send_Cycle_Finished();
                        }
                    }
                    else if (comboBox_Comp.SelectedIndex == 1)
                    {
                        turn_off = false;
                        temp_Reached = false;
                        turn_off_chamber = false;
                        if (Ethernet_label.Text != "Connected")
                        {
                            //Set Tempeture on Chamber
                            check_Set = SET_Temp_CTS(Set_Analog_Box2_Comp.Text);
                            //Start chamber
                            Start_CTS();
                            cycle_Comp = true;
                        }
                        if (Ethernet_label.Text == "Connected")
                        {
                            if (Set_Analog_Box1_Comp.Text.Length == 2)
                            {
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 0" + Set_Analog_Box2_Comp.Text + ".0");
                            }
                            else
                                set_etemp = Read_Write_Ethernet(Ethernet_Box.Text, "a0 " + Set_Analog_Box2_Comp.Text + ".0");
                            Thread.Sleep(100);
                            start = Read_Write_Ethernet(Ethernet_Box.Text, "s1 1");
                            cycle_Comp = true;
                        }
                    }
                }
                else if (Set_Analog_Box2_Comp.Visible == true)
                {
                    if (ActTemp[PICO_Counter] <= double.Parse(Set_Analog_Box2_Comp.Text) - 10 && cycle_Comp == true)
                    {
                        DateTime exposed2 = DateTime.Now;
                        Exposure_Time_Label2.Text += "\nEnded: " + exposed2;
                        Stop_Btn_Comp_Click(null, null);
                        label_Comp_Finished.Visible = true;
                        if (Email_ON.Checked == true && Email_Box.Text != "")
                        {
                            Send_Cycle_Finished();
                        }
                    }
                }
            }
            PICO_Counter += 1;
        }

        public float calc_temp(float[] temp)
        {
            int i;
            float send_temp = 0;
            for (i = 0; i < temp.Length; i++)
                send_temp += temp[i];

            send_temp /= temp.Length;

            return send_temp;
        }

        public void get_all_Temp_Meas(float[] meas_temp1, float[] meas_temp2)
        {
            if(Channel1_Box.Checked == true)
                TempChannel1[Chan_Temp_Count] = meas_temp1[1];
            else
                TempChannel1[Chan_Temp_Count] = 0;
            if (Channel2_Box.Checked == true)
                TempChannel2[Chan_Temp_Count] = meas_temp1[2];
            else
                TempChannel2[Chan_Temp_Count] = 0;
            if (Channel3_Box.Checked == true)
                TempChannel3[Chan_Temp_Count] = meas_temp1[3];
            else
                TempChannel3[Chan_Temp_Count] = 0;
            if (Channel4_Box.Checked == true)
                TempChannel4[Chan_Temp_Count] = meas_temp1[4];
            else
                TempChannel4[Chan_Temp_Count] = 0;
            if (Channel5_Box.Checked == true)
                TempChannel5[Chan_Temp_Count] = meas_temp1[5];
            else
                TempChannel5[Chan_Temp_Count] = 0;
            if (Channel6_Box.Checked == true)
                TempChannel6[Chan_Temp_Count] = meas_temp1[6];
            else
                TempChannel6[Chan_Temp_Count] = 0;
            if (Channel7_Box.Checked == true)
                TempChannel7[Chan_Temp_Count] = meas_temp1[7];
            else
                TempChannel7[Chan_Temp_Count] = 0;
            if (Channel8_Box.Checked == true)
                TempChannel8[Chan_Temp_Count] = meas_temp1[8];
            else
                TempChannel8[Chan_Temp_Count] = 0;
            if (Channel9_Box.Checked == true)
                TempChannel9[Chan_Temp_Count] = meas_temp2[1];
            else
                TempChannel9[Chan_Temp_Count] = 0;
            if (Channel10_Box.Checked == true)
                TempChannel10[Chan_Temp_Count] = meas_temp2[2];
            else
                TempChannel10[Chan_Temp_Count] = 0;
            if (Channel11_Box.Checked == true)
                TempChannel11[Chan_Temp_Count] = meas_temp2[3];
            else
                TempChannel11[Chan_Temp_Count] = 0;
            if (Channel12_Box.Checked == true)
                TempChannel12[Chan_Temp_Count] = meas_temp2[4];
            else
                TempChannel12[Chan_Temp_Count] = 0;
            if (Channel13_Box.Checked == true)
                TempChannel13[Chan_Temp_Count] = meas_temp2[5];
            else
                TempChannel13[Chan_Temp_Count] = 0;
            if (Channel14_Box.Checked == true)
                TempChannel14[Chan_Temp_Count] = meas_temp2[6];
            else
                TempChannel14[Chan_Temp_Count] = 0;
            if (Channel15_Box.Checked == true)
                TempChannel15[Chan_Temp_Count] = meas_temp2[7];
            else
                TempChannel15[Chan_Temp_Count] = 0;
            if (Channel16_Box.Checked == true)
                TempChannel16[Chan_Temp_Count] = meas_temp2[8];
            else
                TempChannel16[Chan_Temp_Count] = 0;
            Chan_Temp_Count = (Chan_Temp_Count + 1) % 10;
        }

        public bool compare_meas(float[] temp)
        {

            if ((double)(temp[1] - temp[0]) <= 0.1 && (double)(temp[0] - temp[1]) <= 0.1)
            {
                Temp_Rich.Text += "\nTrue" + (temp[1] - temp[0]).ToString();
                return true;
            }
            Temp_Rich.Text += "\nFalse: "+(temp[1] - temp[0]).ToString();
            return false;
        }

        public void enable_ChannelBox()
        {
            Channel1_Box.Enabled = true;
            Channel2_Box.Enabled = true;
            Channel3_Box.Enabled = true;
            Channel4_Box.Enabled = true;
            Channel5_Box.Enabled = true;
            Channel6_Box.Enabled = true;
            Channel7_Box.Enabled = true;
            Channel8_Box.Enabled = true;
            Channel9_Box.Enabled = true;
            Channel10_Box.Enabled = true;
            Channel11_Box.Enabled = true;
            Channel12_Box.Enabled = true;
            Channel13_Box.Enabled = true;
            Channel14_Box.Enabled = true;
            Channel15_Box.Enabled = true;
            Channel16_Box.Enabled = true;
        }

        public void disable_ChannelBox()
        {
            Channel1_Box.Enabled = false;
            Channel2_Box.Enabled = false;
            Channel3_Box.Enabled = false;
            Channel4_Box.Enabled = false;
            Channel5_Box.Enabled = false;
            Channel6_Box.Enabled = false;
            Channel7_Box.Enabled = false;
            Channel8_Box.Enabled = false;
            Channel9_Box.Enabled = false;
            Channel10_Box.Enabled = false;
            Channel11_Box.Enabled = false;
            Channel12_Box.Enabled = false;
            Channel13_Box.Enabled = false;
            Channel14_Box.Enabled = false;
            Channel15_Box.Enabled = false;
            Channel16_Box.Enabled = false;
        }

        public void disable_selections()
        {
            Save_Btn_Comp.Enabled = false;
            Save_Btn_Manual.Enabled = false;
            Clear_All.Enabled = false;
            groupBox_Communication.Enabled = false;
            groupBox_Use.Enabled = false;
            PICO_Unit_1_gb.Enabled = false;
            PICO_Unit_2_gb.Enabled = false;
            comboBox_Comp.Enabled = false;
            comboBox_Set_Analog_Box_Manual.Enabled = false;
        }

        public void enable_selections()
        {
            Save_Btn_Comp.Enabled = true;
            Save_Btn_Manual.Enabled = true;
            Clear_All.Enabled = true;
            groupBox_Communication.Enabled = true;
            groupBox_Use.Enabled = true;
            PICO_Unit_1_gb.Enabled = true;
            PICO_Unit_2_gb.Enabled = true;
            comboBox_Comp.Enabled = true;
            comboBox_Set_Analog_Box_Manual.Enabled = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_Comp.SelectedIndex == 0)
            {
                Set_Analog_Box2_Comp.Visible = false;
                Exposure_Time2_Comp.Visible = false;
                label_Set_Temp_Comp2.Visible = false;
                label_Exposure_Comp2.Visible = false;
                Exposure_Time_Label2.Visible = false;
                Set_analog_minus_10_Label2.Visible = false;
            }
            if (comboBox_Comp.SelectedIndex == 1)
            {
                Set_Analog_Box2_Comp.Visible = true;
                label_Set_Temp_Comp2.Visible = true;
                Set_analog_minus_10_Label2.Visible = true;
                Exposure_Time_Label2.Visible = true;
                if (Exposure_ON_Comp.Checked == true)
                {
                    Exposure_Time2_Comp.Visible = true;
                    label_Exposure_Comp2.Visible = true;
                }
            }
        }

        private void Automatic_ON_Comp_CheckedChanged(object sender, EventArgs e)
        {
            if (Automatic_ON_Comp.Checked == true)
            {
                Exposure_Time1_Comp.Visible = false;
                label_Exposure_Comp1.Visible = false;
                Exposure_Time_Label1.Visible = true;
                Exposure_Time2_Comp.Visible = false;
                label_Exposure_Comp2.Visible = false;
                if(comboBox_Comp.SelectedIndex == 1)
                    Exposure_Time_Label2.Visible = true;
            }
        }

        private void Exposure_ON_Comp_CheckedChanged(object sender, EventArgs e)
        {
            if (Exposure_ON_Comp.Checked == true)
            {
                Exposure_Time1_Comp.Visible = true;
                label_Exposure_Comp1.Visible = true;
                Exposure_Time_Label1.Visible = true;
                if (comboBox_Comp.SelectedIndex == 1)
                {
                    Exposure_Time2_Comp.Visible = true;
                    label_Exposure_Comp2.Visible = true;
                    Exposure_Time_Label2.Visible = true;
                }
            }
        }

        public Device_1_Dialog Device_1_Dialog
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        public Device_2_Dialog Device_2_Dialog
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }
    }
}
