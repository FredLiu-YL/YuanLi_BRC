using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using Automation.BDaq;
using System.Diagnostics;
using YuanLi_Logger;
using System.Media;

namespace BRC
{
    public partial class Form1 : Form
    {

        private NIHighFrequencyCutting nIHighFrequencyCutting;
        private ZaberMotion zaberMotion;


        public Form1()
        {
            InitializeComponent();
        }

        #region Var
        SerialPort Motion_sp = new SerialPort();
        double Movement_Ratio;
        string Sp1_Terminator;
        bool connect_Motion_OK;
        bool Cut_And_Scan_Finish;
        int Now_Read_Status = 0;
        bool Get_Motion_Feedback = false;
        bool X_Busy, Y_Busy, Z_Busy;
        const int Get_Axis_Position = 1; //"Q:"
        const int Get_Axis_Busy_Ready = 2; //"!:"
        const int Get_Axis_Status = 3; //"Q:S"
        int now_step = 0;
        int wait_delay = 0;
        const int wait_second = 20;
        const int Position_Range = 5;//0.1um
        bool IO_Can_Cut = false;
        string andor_file_address;
        string Andor_Error_Meaasge;
        string Protocal_name;
        string Process_Nmae_File_Address;
        bool need_scan = false;
        Logger logger = new Logger("BRC_Cutting_Log");
        Logger logger_Motion = new Logger("Motion_Log");
        int motion_move_Z = 0;
        bool move_z_Step_ok = true;
        #endregion

        #region Icon
        private void Form1_Load(object sender, EventArgs e)
        {
            button_Start.Text = "掃  描  開  始";
            logger.Write_Logger("Start Program");
            foreach (string portlist in System.IO.Ports.SerialPort.GetPortNames())
                comboBox_COM_Port.Items.Add(portlist);
            if (comboBox_COM_Port.Items.Count > 0)
                comboBox_COM_Port.SelectedIndex = 0;
            //
            logger.Write_Logger("Initial Parameter");
            Form_Value_Initial(System.Windows.Forms.Application.StartupPath + "\\Setup\\Initial_Value.xml");
            groupBox_Position_Setup.Enabled = false;
            Process_Nmae_File_Address = System.Windows.Forms.Application.StartupPath + "\\Setup\\ProcessName.txt";
            //
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Setup\\LoadFileAddress.txt")) {
                StreamReader sr_ = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\Setup\\LoadFileAddress.txt");
                string read_old_file_path = sr_.ReadLine();
                sr_.Close();
                textBox_Andor_Exe_Address.Text = read_old_file_path;
            }
            if (File.Exists(Process_Nmae_File_Address)) {
                StreamReader sr_ = new StreamReader(Process_Nmae_File_Address);
                string read_dat = sr_.ReadLine();
                while ((read_dat != null)) {
                    if (read_dat != "")
                        comboBox_Process_Name.Items.Add(read_dat);
                    read_dat = sr_.ReadLine();
                }
                sr_.Close();
            }
            textBox_Now_layer.Text = "0";
            if (!File.Exists(System.Windows.Forms.Application.StartupPath + "\\Setup\\Test_Mode.txt"))
                button_Auto_Click(sender, e);
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Setup\\Last_Process_Name.txt")) {
                StreamReader sr_ = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\Setup\\Last_Process_Name.txt");
                string process_name = sr_.ReadLine();
                sr_.Close();
                for (int i = 0; i < comboBox_Process_Name.Items.Count; i++) {
                    if (comboBox_Process_Name.Items[i].ToString() == process_name)
                        comboBox_Process_Name.SelectedIndex = i;
                }
            }
            //載入數據到即時顯示textbox
            Imshow_Real_Time_value("Total_Scan_Layer", (double)numericUpDown_Total_Scan_Layer.Value);
            Imshow_Real_Time_value("Cut", (double)numericUpDown_Cut.Value);
            Imshow_Real_Time_value("Cut_Dis_M", (double)numericUpDown_Cut_Dis_M.Value);
            Imshow_Real_Time_value("Scan_Dis_M", (double)numericUpDown_Scan_Dis_M.Value);
            //載入間距到textbox
            textBox_Top_Bottom_Diff.Text = Convert.ToInt32(Math.Abs(Convert.ToDouble(textBox_Z_Up.Text) - Convert.ToDouble(textBox_Z_Down.Text))).ToString();


            //初始化震動刀功能
            nIHighFrequencyCutting = new NIHighFrequencyCutting();
        }
        private void button_Save_Parameter_Click(object sender, EventArgs e)
        {
            try {
                logger.Write_Logger("Save Parameter : " + textBox_ProcessName.Text);
                DateTime Now_ = DateTime.Now;
                String Today_ = "_" +
                    Convert.ToString(Now_.Year) + "_" +
                    Convert.ToString(Now_.Month) + "_" +
                    Convert.ToString(Now_.Day) + "_" +
                    Convert.ToString(Now_.Hour) + "_" +
                    Convert.ToString(Now_.Minute) + "_" +
                    Convert.ToString(Now_.Second);

                File.Move(
                    System.Windows.Forms.Application.StartupPath + "\\Setup\\Initial_Value.xml",
                    System.Windows.Forms.Application.StartupPath + "\\Setup\\Backup\\Initial_Value" + Today_ + ".xml");
                StreamWriter SW_ = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\Setup\\Initial_Value.xml");
                SW_.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");//<?xml version="1.0" encoding="utf-8" ?>
                SW_.WriteLine("<BRC_Program_Setup>");//<BRC_Program_Setup>
                SW_.WriteLine("  <Setup_Part Setup_Part=\"Motion\">");//<Setup_Part Setup_Part="Motion">
                SW_.WriteLine("    <Setup Setup=\"COM_PORT\">" + Convert.ToString(comboBox_COM_Port.Text) + "</Setup>");//<Setup Setup="COM_PORT">COM1</Setup>
                SW_.WriteLine("    <Setup Setup=\"X_Inverse\">" + Convert.ToString(checkBox_Inverse_X.Checked) + "</Setup>");//<Setup Setup="X_Inverse">False</Setup>
                SW_.WriteLine("    <Setup Setup=\"Y_Inverse\">" + Convert.ToString(checkBox_Inverse_Y.Checked) + "</Setup>");//<Setup Setup="Y_Inverse">False</Setup>
                SW_.WriteLine("    <Setup Setup=\"Z_Inverse\">" + Convert.ToString(checkBox_Inverse_Z.Checked) + "</Setup>");//<Setup Setup="Z_Inverse">False</Setup>
                SW_.WriteLine("    <Setup Setup=\"XY_Inverse\">" + Convert.ToString(checkBox_Inverse_XY.Checked) + "</Setup>");//<Setup Setup="XY_Inverse">False</Setup>
                SW_.WriteLine("    <Setup Setup=\"X_Number\">" + Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) + "</Setup>");//<Setup Setup="X_Number">1</Setup >
                SW_.WriteLine("    <Setup Setup=\"Y_Number\">" + Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) + "</Setup>");//<Setup Setup="Y_Number">2</Setup >
                SW_.WriteLine("    <Setup Setup=\"Z_Number\">" + Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) + "</Setup>");//<Setup Setup="Z_Number">0</Setup >
                SW_.WriteLine("    <Setup Setup=\"Ratio\">" + Convert.ToString(Movement_Ratio) + "</Setup>");//<Setup Setup="Ratio">50</Setup >
                SW_.WriteLine("  </Setup_Part>");//</Setup_Part>
                SW_.WriteLine("  <Setup_Part Setup_Part=\"Cut_Data\">");//<Setup_Part Setup_Part="Cut_Data">
                SW_.WriteLine("    <Setup Setup=\"Cut_Start_X\">" + Convert.ToString(textBox_Cut_Start_X.Text) + "</Setup >");//<Setup Setup="Cut_Start_X">1000</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Start_Y\">" + Convert.ToString(textBox_Cut_Start_Y.Text) + "</Setup >");//<Setup Setup="Cut_Start_Y">1001</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Start_Z\">" + Convert.ToString(textBox_Cut_Start_Z.Text) + "</Setup >");//<Setup Setup="Cut_Start_Z">1002</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_End_X\">" + Convert.ToString(textBox_Cut_End_X.Text) + "</Setup >");//<Setup Setup="Cut_End_X">1010</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_End_Y\">" + Convert.ToString(textBox_Cut_End_Y.Text) + "</Setup >");//<Setup Setup="Cut_End_Y">1011</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Speed_X\">" + Convert.ToString(textBox_Cut_Speed_X.Text) + "</Setup >");//<Setup Setup="Cut_Speed_X">1020</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Speed_Y\">" + Convert.ToString(textBox_Cut_Speed_Y.Text) + "</Setup >");//<Setup Setup="Cut_Speed_Y">1021</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Layer\">" + Convert.ToString(numericUpDown_Total_Cut_Layer.Text) + "</Setup >");//<Setup Setup="Cut_Layer">20</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Distance\">" + Convert.ToString(numericUpDown_Cut_Distance.Text) + "</Setup >");//<Setup Setup="Cut_Distance">1.5</Setup >
                SW_.WriteLine("    <Setup Setup=\"Cut_Frequency\">" + Convert.ToString(numericUpDown_Cut_Frequency.Text) + "</Setup >");//<Setup Setup="Cut_Frequency">80</Setup >
                SW_.WriteLine("    <Setup Setup=\"FrequencyMinV\">" + Convert.ToString(textBox_FrequencyMinV.Text) + "</Setup >");//<Setup Setup="FrequencyMinV">1</Setup >
                SW_.WriteLine("    <Setup Setup=\"FrequencyMaxV\">" + Convert.ToString(textBox_FrequencyMaxV.Text) + "</Setup >");//<Setup Setup="FrequencyMaxV">10</Setup >
                SW_.WriteLine("    <Setup Setup=\"FrequencyOutputV\">" + Convert.ToString(textBox_FrequencyOutputV.Text) + "</Setup >");//<Setup Setup="FrequencyOutputV">5</Setup >

                SW_.WriteLine("  </Setup_Part>");//</Setup_Part>
                SW_.WriteLine("  <Setup_Part Setup_Part=\"Scan_Data\">");//<Setup_Part Setup_Part="Scan_Data">
                SW_.WriteLine("    <Setup Setup=\"Scan_Start_X\">" + Convert.ToString(textBox_Scan_Start_X.Text) + "</Setup >");//<Setup Setup="Scan_Start_X">1030</Setup >
                SW_.WriteLine("    <Setup Setup=\"Scan_Start_Y\">" + Convert.ToString(textBox_Scan_Start_Y.Text) + "</Setup >");//<Setup Setup="Scan_Start_Y">1031</Setup >
                SW_.WriteLine("    <Setup Setup=\"Scan_Start_Z\">" + Convert.ToString(textBox_Scan_Start_Z.Text) + "</Setup >");//<Setup Setup="Scan_Start_Z">1032</Setup >
                SW_.WriteLine("    <Setup Setup=\"Scan_Distance_Z\">" + Convert.ToString(numericUpDown_Scan_Distance_Z.Text) + "</Setup >");//<Setup Setup="Scan_Distance_Z">1.2</Setup >
                SW_.WriteLine("    <Setup Setup=\"Start_Scan_Z\">" + Convert.ToString(textBox_Z_Down.Text) + "</Setup >");//<Setup Setup="Start_Scan_Z">30000</Setup>
                SW_.WriteLine("    <Setup Setup=\"End_Scan_Z\">" + Convert.ToString(textBox_Z_Up.Text) + "</Setup >");//<Setup Setup="End_Scan_Z">50000</Setup>
                SW_.WriteLine("  </Setup_Part>");//</Setup_Part>
                SW_.WriteLine("  <Setup_Part Setup_Part=\"Move\">");//<Setup_Part Setup_Part="Move">
                SW_.WriteLine("    <Setup Setup=\"Safety_Hight\">" + Convert.ToString(textBox_Safety_Hight.Text) + "</Setup >");//<Setup Setup="Safety_Hight">1060</Setup >
                SW_.WriteLine("    <Setup Setup=\"Move_Speed_X\">" + Convert.ToString(textBox_Move_Speed_X.Text) + "</Setup >");//<Setup Setup="Move_Speed_X">1070</Setup >
                SW_.WriteLine("    <Setup Setup=\"Move_Speed_Y\">" + Convert.ToString(textBox_Move_Speed_Y.Text) + "</Setup >");//<Setup Setup="Move_Speed_Y">1071</Setup >
                SW_.WriteLine("    <Setup Setup=\"Move_Speed_Z\">" + Convert.ToString(textBox_Move_Speed_Z.Text) + "</Setup >");//<Setup Setup="Move_Speed_Z">1072</Setup >
                SW_.WriteLine("    <Setup Setup=\"StandBy_X\">" + Convert.ToString(textBox_Standby_X.Text) + "</Setup >");//<Setup Setup="StandBy_X">1080</Setup >
                SW_.WriteLine("    <Setup Setup=\"StandBy_Y\">" + Convert.ToString(textBox_Standby_Y.Text) + "</Setup >");//<Setup Setup="StandBy_Y">1081</Setup >
                SW_.WriteLine("    <Setup Setup=\"StandBy_Z\">" + Convert.ToString(textBox_Standby_Z.Text) + "</Setup >");//<Setup Setup="StandBy_Z">1082</Setup >
                SW_.WriteLine("    <Setup Setup=\"Micro_X\">" + Convert.ToString(textBox_Micro_X.Text) + "</Setup >");//<Setup Setup="Micro_X">1080</Setup >
                SW_.WriteLine("    <Setup Setup=\"Micro_Y\">" + Convert.ToString(textBox_Micro_Y.Text) + "</Setup >");//<Setup Setup="Micro_Y">1081</Setup >
                SW_.WriteLine("    <Setup Setup=\"Micro_Z\">" + Convert.ToString(textBox_Micro_Z.Text) + "</Setup >");//<Setup Setup="Micro_Z">1082</Setup >

                SW_.WriteLine("    <Setup Setup=\"CVfirstDelaytime\">" + Convert.ToString(textBox_CVfirstDelaytime.Text) + "</Setup >");//<Setup Setup="CVfirstDelaytime">1</Setup >
                SW_.WriteLine("    <Setup Setup=\"CVsecondDelaytime\">" + Convert.ToString(textBox_CVsecondDelaytime.Text) + "</Setup >");//<Setup Setup="CVsecondDelaytime">10</Setup >
                SW_.WriteLine("    <Setup Setup=\"CVfirstVelocity\">" + Convert.ToString(textBox_CVfirstVelocity.Text) + "</Setup >");//<Setup Setup="CVfirstVelocity">5</Setup >
                SW_.WriteLine("    <Setup Setup=\"CVsecondVelocity\">" + Convert.ToString(textBox_CVsecondVelocity.Text) + "</Setup >");//<Setup Setup="CVsecondVelocity">5</Setup >
                SW_.WriteLine("    <Setup Setup=\"CVZaberComPort\">" + Convert.ToString(textBox_ZaberComPort.Text) + "</Setup >");//<Setup Setup="CVsecondVelocity">5</Setup >



                SW_.WriteLine("  </Setup_Part>");//</Setup_Part>
                SW_.WriteLine("</BRC_Program_Setup>");//</BRC_Program_Setup>
                SW_.Close();
                //
                string new_process_Address = System.Windows.Forms.Application.StartupPath + "\\Setup\\Process\\" + textBox_ProcessName.Text + ".xml";
                if (!File.Exists(new_process_Address)) {
                    //ProcessName
                    StreamWriter sw_process = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\Setup\\ProcessName.txt", true);
                    sw_process.Write(textBox_ProcessName.Text + "\r\n");
                    sw_process.Close();
                    //
                    if (File.Exists(Process_Nmae_File_Address)) {
                        comboBox_Process_Name.Items.Clear();
                        StreamReader sr_ = new StreamReader(Process_Nmae_File_Address);
                        string read_dat = sr_.ReadLine();
                        while ((read_dat != null)) {
                            if (read_dat != "")
                                comboBox_Process_Name.Items.Add(read_dat);
                            read_dat = sr_.ReadLine();
                        }
                        sr_.Close();
                    }
                }
                else
                    File.Delete(new_process_Address);
                File.Copy(
                    System.Windows.Forms.Application.StartupPath + "\\Setup\\Initial_Value.xml",
                    new_process_Address);
                logger.Write_Logger("Save OK");
                MessageBox.Show("儲存成功");
            }
            catch (Exception error) {
                logger.Write_Logger("Save Error " + Convert.ToString(error));
                MessageBox.Show("儲存失敗\n" + Convert.ToString(error));
            }
        }
        private void button_Move_To_Standby_All_Click(object sender, EventArgs e)
        {
            now_step = 50;
        }
        private void timer_IO_Tick(object sender, EventArgs e)
        {
            string Read_Data = "";
            try {
                #region Read Serial Port

                try//Read Motion
                {
                    bool Can_Read = Motion_sp.CtsHolding;
                    int lll = Motion_sp.BytesToRead;
                    if (lll != 0 && Can_Read) {
                        char[] abc = new char[lll];
                        Motion_sp.Read(abc, 0, lll);
                        Motion_sp.DiscardInBuffer();
                        for (int i = 0; i < lll; i++) {
                            Read_Data = Read_Data + abc[i] + " ";
                        }
                    }
                }
                catch (Exception error) {
                    logger.Write_Error_Logger(error.ToString());
                }
                if (Read_Data.Length <= 2) {
                    if (Read_Data == "OK" || Read_Data == "Ok" || Read_Data == "ok" || Read_Data == "oK") {
                    }
                    else {
                    }
                }
                else {
                    string[] Calculate_Array = Read_Data.Split(',');
                    string X_Data;
                    string Y_Data;
                    string Z_Data;
                    if (Now_Read_Status == Get_Axis_Status) {
                        X_Data = Calculate_Array[comboBox_Axis_Num_X.SelectedIndex + 1];
                        Y_Data = Calculate_Array[comboBox_Axis_Num_Y.SelectedIndex + 1];
                        Z_Data = Calculate_Array[comboBox_Axis_Num_Z.SelectedIndex + 1];
                    }
                    else {
                        X_Data = Calculate_Array[comboBox_Axis_Num_X.SelectedIndex];
                        Y_Data = Calculate_Array[comboBox_Axis_Num_Y.SelectedIndex];
                        Z_Data = Calculate_Array[comboBox_Axis_Num_Z.SelectedIndex];
                    }
                    //
                    if (X_Data != null)
                        X_Data = X_Data.Replace(" ", "");
                    if (Y_Data != null)
                        Y_Data = Y_Data.Replace(" ", "");
                    if (Z_Data != null)
                        Z_Data = Z_Data.Replace(" ", "");
                    //
                    switch (Now_Read_Status) {
                        case Get_Axis_Position:  //"Q:"
                            textBox_Now_Position_X.Text = Convert.ToString(Convert.ToInt32(Convert.ToInt32(X_Data) / Movement_Ratio));
                            textBox_Now_Position_Y.Text = Convert.ToString(Convert.ToInt32(Convert.ToInt32(Y_Data) / Movement_Ratio));
                            textBox_Now_Position_Z.Text = Convert.ToString(Convert.ToInt32(Convert.ToInt32(Z_Data) / Movement_Ratio));
                            textBox_Now_Position_X_Main.Text = Convert.ToString(Convert.ToInt32(Convert.ToInt32(X_Data) / Movement_Ratio));
                            textBox_Now_Position_Y_Main.Text = Convert.ToString(Convert.ToInt32(Convert.ToInt32(Y_Data) / Movement_Ratio));
                            textBox_Now_Position_Z_Main.Text = Convert.ToString(Convert.ToInt32(Convert.ToInt32(Z_Data) / Movement_Ratio));
                            Get_Motion_Feedback = true;
                            logger_Motion.Write_Logger(
                                "Position X :" + Convert.ToString(X_Data) +
                                " Position Y :" + Convert.ToString(Y_Data) +
                                " Position Z :" + Convert.ToString(Z_Data));
                            if (!move_z_Step_ok) {
                                if (motion_move_Z == Convert.ToInt32(Z_Data))
                                    move_z_Step_ok = true;
                            }
                            break;
                        case Get_Axis_Busy_Ready: //"!:"
                            if (X_Data == "0") {
                                pictureBox_Axis_X_Busy.Image = Properties.Resources.Green;
                                pictureBox_Moving_X_Main.Image = Properties.Resources.Green;
                                X_Busy = false;
                            }
                            else {
                                pictureBox_Axis_X_Busy.Image = Properties.Resources.Red;
                                pictureBox_Moving_X_Main.Image = Properties.Resources.Red;
                                X_Busy = true;
                            }
                            if (Y_Data == "0") {
                                pictureBox_Axis_Y_Busy.Image = Properties.Resources.Green;
                                pictureBox_Moving_Y_Main.Image = Properties.Resources.Green;
                                Y_Busy = false;
                            }
                            else {
                                pictureBox_Axis_Y_Busy.Image = Properties.Resources.Red;
                                pictureBox_Moving_Y_Main.Image = Properties.Resources.Red;
                                Y_Busy = true;
                            }
                            if (Z_Data == "0") {
                                pictureBox_Axis_Z_Busy.Image = Properties.Resources.Green;
                                pictureBox_Moving_Z_Main.Image = Properties.Resources.Green;
                                Z_Busy = false;
                            }
                            else {
                                pictureBox_Axis_Z_Busy.Image = Properties.Resources.Red;
                                pictureBox_Moving_Z_Main.Image = Properties.Resources.Red;
                                Z_Busy = true;
                            }
                            Get_Motion_Feedback = true;
                            logger_Motion.Write_Logger(
                                "Busy X :" + Convert.ToString(X_Busy) +
                                " Busy Y :" + Convert.ToString(Y_Busy) +
                                " Busy Z :" + Convert.ToString(Z_Busy));
                            break;
                        case Get_Axis_Status:     //"Q:S"
                            #region Calculate Data
                            #region X
                            int X_Number = Convert.ToInt32(X_Data);
                            int[] X_Calculate = new int[8];
                            X_Calculate[7] = X_Number / 128;
                            X_Number = X_Number - X_Calculate[7] * 128;
                            X_Calculate[6] = X_Number / 64;
                            X_Number = X_Number - X_Calculate[6] * 64;
                            X_Calculate[5] = X_Number / 32;
                            X_Number = X_Number - X_Calculate[5] * 32;
                            X_Calculate[4] = X_Number / 16;
                            X_Number = X_Number - X_Calculate[4] * 16;
                            X_Calculate[3] = X_Number / 8;
                            X_Number = X_Number - X_Calculate[3] * 8;
                            X_Calculate[2] = X_Number / 4;
                            X_Number = X_Number - X_Calculate[2] * 4;
                            X_Calculate[1] = X_Number / 2;
                            X_Calculate[0] = X_Number % 2;
                            #endregion
                            #region Y
                            int Y_Number = Convert.ToInt32(Y_Data);
                            int[] Y_Calculate = new int[8];
                            Y_Calculate[7] = Y_Number / 128;
                            Y_Number = Y_Number - Y_Calculate[7] * 128;
                            Y_Calculate[6] = Y_Number / 64;
                            Y_Number = Y_Number - Y_Calculate[6] * 64;
                            Y_Calculate[5] = Y_Number / 32;
                            Y_Number = Y_Number - Y_Calculate[5] * 32;
                            Y_Calculate[4] = Y_Number / 16;
                            Y_Number = Y_Number - Y_Calculate[4] * 16;
                            Y_Calculate[3] = Y_Number / 8;
                            Y_Number = Y_Number - Y_Calculate[3] * 8;
                            Y_Calculate[2] = Y_Number / 4;
                            Y_Number = Y_Number - Y_Calculate[2] * 4;
                            Y_Calculate[1] = Y_Number / 2;
                            Y_Calculate[0] = Y_Number % 2;
                            #endregion
                            #region Z
                            int Z_Number = Convert.ToInt32(Z_Data);
                            int[] Z_Calculate = new int[8];
                            Z_Calculate[7] = Z_Number / 128;
                            Z_Number = Z_Number - Z_Calculate[7] * 128;
                            Z_Calculate[6] = Z_Number / 64;
                            Z_Number = Z_Number - Z_Calculate[6] * 64;
                            Z_Calculate[5] = Z_Number / 32;
                            Z_Number = Z_Number - Z_Calculate[5] * 32;
                            Z_Calculate[4] = Z_Number / 16;
                            Z_Number = Z_Number - Z_Calculate[4] * 16;
                            Z_Calculate[3] = Z_Number / 8;
                            Z_Number = Z_Number - Z_Calculate[3] * 8;
                            Z_Calculate[2] = Z_Number / 4;
                            Z_Number = Z_Number - Z_Calculate[2] * 4;
                            Z_Calculate[1] = Z_Number / 2;
                            Z_Calculate[0] = Z_Number % 2;
                            #endregion
                            #endregion
                            #region Show Data
                            #region X
                            if (X_Calculate[6] == 0 && X_Calculate[5] == 0 && X_Calculate[4] == 0)
                                pictureBox_Axis_X_Alarm.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_X_Alarm.Image = Properties.Resources.Red;
                            if (X_Calculate[2] == 0)
                                pictureBox_Axis_X_ORG.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_X_ORG.Image = Properties.Resources.Red;
                            if (X_Calculate[1] == 0)
                                pictureBox_Axis_X_Limit_P.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_X_Limit_P.Image = Properties.Resources.Red;
                            if (X_Calculate[0] == 0)
                                pictureBox_Axis_X_Limit_N.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_X_Limit_N.Image = Properties.Resources.Red;
                            #endregion
                            #region Y
                            if (Y_Calculate[6] == 0 && Y_Calculate[5] == 0 && Y_Calculate[4] == 0)
                                pictureBox_Axis_Y_Alarm.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Y_Alarm.Image = Properties.Resources.Red;
                            if (Y_Calculate[2] == 0)
                                pictureBox_Axis_Y_ORG.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Y_ORG.Image = Properties.Resources.Red;
                            if (Y_Calculate[1] == 0)
                                pictureBox_Axis_Y_Limit_P.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Y_Limit_P.Image = Properties.Resources.Red;
                            if (Y_Calculate[0] == 0)
                                pictureBox_Axis_Y_Limit_N.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Y_Limit_N.Image = Properties.Resources.Red;
                            #endregion
                            #region Z
                            if (Z_Calculate[6] == 0 && Z_Calculate[5] == 0 && Z_Calculate[4] == 0)
                                pictureBox_Axis_Z_Alarm.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Z_Alarm.Image = Properties.Resources.Red;
                            if (Z_Calculate[2] == 0)
                                pictureBox_Axis_Z_ORG.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Z_ORG.Image = Properties.Resources.Red;
                            if (Z_Calculate[1] == 0)
                                pictureBox_Axis_Z_Limit_P.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Z_Limit_P.Image = Properties.Resources.Red;
                            if (Z_Calculate[0] == 0)
                                pictureBox_Axis_Z_Limit_N.Image = Properties.Resources.Green;
                            else
                                pictureBox_Axis_Z_Limit_N.Image = Properties.Resources.Red;
                            #endregion
                            #endregion
                            Get_Motion_Feedback = true;
                            break;
                    }
                }
                #endregion

                #region Send Serial Port
                if (Now_Read_Status == 0) {
                    Write_Motion("Q:");
                    Now_Read_Status = Get_Axis_Position;
                }
                else if (Now_Read_Status == Get_Axis_Position && Get_Motion_Feedback == true) {
                    Write_Motion("!:");
                    Now_Read_Status = Get_Axis_Busy_Ready;
                }
                else if (Now_Read_Status == Get_Axis_Busy_Ready && Get_Motion_Feedback == true) {
                    Write_Motion("Q:S");
                    Now_Read_Status = Get_Axis_Status;
                }
                else if (Now_Read_Status == Get_Axis_Status && Get_Motion_Feedback == true) {
                    Write_Motion("Q:");
                    Now_Read_Status = Get_Axis_Position;
                }
                #endregion
            }
            catch (Exception error) {
                logger.Write_Error_Logger(Read_Data);
                logger.Write_Error_Logger(error.ToString());
            }
        }
        private void button_Load_Andor_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_ = new OpenFileDialog();
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Setup\\LoadFilePath.txt")) {
                StreamReader sr_ = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\Setup\\LoadFilePath.txt");
                string read_old_file_path = sr_.ReadLine();
                sr_.Close();
                open_.InitialDirectory = read_old_file_path;
            }
            else
                open_.InitialDirectory = "c:\\";
            open_.Filter = "exe files (*.exe)|*.exe|All files (*.*)|*.*";
            open_.FilterIndex = 1;
            open_.RestoreDirectory = true;
            open_.Multiselect = false;
            if (open_.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                if (open_.CheckFileExists) {
                    string need_read_File = open_.FileName;
                    try {
                        andor_file_address = need_read_File;
                        int last_ = andor_file_address.LastIndexOf("\\");
                        string need_file_path = andor_file_address.Substring(0, last_) + "\\";
                        StreamWriter sw = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\Setup\\LoadFilePath.txt");
                        sw.WriteLine(need_file_path);
                        sw.Close();
                        textBox_Andor_Exe_Address.Text = andor_file_address;
                        //
                        StreamWriter sw_2 = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\Setup\\LoadFileAddress.txt");
                        sw_2.WriteLine(andor_file_address);
                        sw_2.Close();

                        logger.Write_Logger("Load Andor Address : " + need_file_path);
                    }
                    catch (Exception error) {
                        logger.Write_Error_Logger("Load File Address Error " + error.ToString());
                        MessageBox.Show("Load File Address Error " + error.ToString());
                    }
                }
            }
        }
        private void button_Move_Micro_Click(object sender, EventArgs e)
        {
            now_step = 60;
        }
        private void numericUpDown_Cut_Frequency_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown_Cut.Value = numericUpDown_Cut_Frequency.Value;
            textBox_Cut_Frequency.Text = Convert.ToString(numericUpDown_Cut_Frequency.Value);
        }
        private void numericUpDown_Cut_ValueChanged(object sender, EventArgs e)
        {
            Imshow_Real_Time_value("Cut", (double)numericUpDown_Cut.Value);
            numericUpDown_Cut_Frequency.Value = numericUpDown_Cut.Value;
            textBox_Cut_Frequency.Text = Convert.ToString(numericUpDown_Cut_Frequency.Value);
        }
        private void button_Set_Z_Up_Click(object sender, EventArgs e)
        {
            textBox_Z_Up.Text = textBox_Now_Position_Z.Text;
            if (textBox_Z_Up.Text != "" && textBox_Z_Down.Text != "" && numericUpDown_Cut.Value != 0)
                Cal_Data();
        }
        private void button_Set_Z_Down_Click(object sender, EventArgs e)
        {
            textBox_Z_Down.Text = textBox_Now_Position_Z.Text;
            if (textBox_Z_Up.Text != "" && textBox_Z_Down.Text != "" && numericUpDown_Cut.Value != 0)
                Cal_Data();
        }
        private void numericUpDown_Total_Scan_Layer_ValueChanged(object sender, EventArgs e)
        {
            Imshow_Real_Time_value("Total_Scan_Layer", (double)numericUpDown_Total_Scan_Layer.Value);
            numericUpDown_Total_Cut_Layer.Value = numericUpDown_Total_Scan_Layer.Value - 1;
            if (textBox_Z_Up.Text != "" && textBox_Z_Down.Text != "" && numericUpDown_Cut.Value != 0)
                Cal_Data();
        }
        private void textBox_Z_Speed_M_KeyDown(object sender, KeyEventArgs e)
        {
            textBox_Step_Speed_Z.Text = textBox_Z_Speed_M.Text;
            button_Set_Speed_Click(sender, e);
        }
        private void numericUpDown_Cut_Dis_M_ValueChanged(object sender, EventArgs e)
        {
            Imshow_Real_Time_value("Cut_Dis_M", (double)numericUpDown_Cut_Dis_M.Value);
            numericUpDown_Cut_Distance.Value = numericUpDown_Cut_Dis_M.Value;
        }
        private void numericUpDown_Scan_Dis_M_ValueChanged(object sender, EventArgs e)
        {
            Imshow_Real_Time_value("Scan_Dis_M", (double)numericUpDown_Scan_Dis_M.Value);
            numericUpDown_Scan_Distance_Z.Value = numericUpDown_Scan_Dis_M.Value;
        }

        #region Motion
        private void button_Set_Speed_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(textBox_Step_Speed_X.Text);
                int Y_Speed = Convert.ToInt32(textBox_Step_Speed_Y.Text);
                int Z_Speed = Convert.ToInt32(textBox_Step_Speed_Z.Text);
                String Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(Movement_Ratio * X_Speed / 4) +
                    "," + Convert.ToString(Movement_Ratio * X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Movement_Ratio * Y_Speed / 4) +
                    "," + Convert.ToString(Movement_Ratio * Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Movement_Ratio * Z_Speed / 4) +
                    "," + Convert.ToString(Movement_Ratio * Z_Speed) + ",500";
                logger.Write_Logger("Set Speed X : " + textBox_Step_Speed_X.Text + " Y : " + textBox_Step_Speed_Y.Text + " Z : " + textBox_Step_Speed_Z.Text);
                Write_Motion(Write_Data);
                //

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Set Speed Error! " + Convert.ToString(error));
                MessageBox.Show("Set Speed Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_X_P_MouseDown(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                if (checkBox_Inverse_X.Checked)
                    Data[comboBox_Axis_Num_X.SelectedIndex] = "999999999";
                else
                    Data[comboBox_Axis_Num_X.SelectedIndex] = "-10000000";
                String Write_Data = "";
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move X P");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move X P error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_X_P_MouseUp(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_X.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("X Stop");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("X Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_X_N_MouseDown(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                if (checkBox_Inverse_X.Checked)
                    Data[comboBox_Axis_Num_X.SelectedIndex] = "-10000000";
                else
                    Data[comboBox_Axis_Num_X.SelectedIndex] = "999999999";
                String Write_Data = "";
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move X N");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move X N error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_X_N_MouseUp(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_X.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("X Stop");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("X Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Y_P_MouseDown(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                if (checkBox_Inverse_Y.Checked)
                    Data[comboBox_Axis_Num_Y.SelectedIndex] = "999999999";
                else
                    Data[comboBox_Axis_Num_Y.SelectedIndex] = "-10000000";
                String Write_Data = "";
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Y P");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Y P error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Y_P_MouseUp(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_Y.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Y Stop");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Y Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Y_N_MouseDown(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                if (checkBox_Inverse_Y.Checked)
                    Data[comboBox_Axis_Num_Y.SelectedIndex] = "-10000000";
                else
                    Data[comboBox_Axis_Num_Y.SelectedIndex] = "999999999";
                String Write_Data = "";
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Y N");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Y N error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Y_N_MouseUp(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_Y.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Y Stop");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Y Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Z_P_MouseDown(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                if (checkBox_Inverse_Z.Checked)
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = "999999999";
                else
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = "-10000000";
                String Write_Data = "";
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Z P");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Z P error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Z_P_MouseUp(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_Z.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Z Stop");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Z Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Z_N_MouseDown(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                if (checkBox_Inverse_Z.Checked)
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = "-10000000";
                else
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = "999999999";
                String Write_Data = "";
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Z N");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Z N error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Z_N_MouseUp(object sender, MouseEventArgs e)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_Z.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Z Stop");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Z Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        #endregion

        #region Set Position
        private void checkBox_Position_Setup_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Position_Setup.Checked)
                groupBox_Position_Setup.Enabled = true;
            else
                groupBox_Position_Setup.Enabled = false;
        }
        private void button_Set_Cut_Start_Click(object sender, EventArgs e)
        {
            textBox_Now_Position_X.Text = "15000";
            textBox_Now_Position_Y.Text = "12000";
            textBox_Cut_Start_X.Text = Convert.ToInt32(textBox_Now_Position_X.Text).ToString();
            textBox_Cut_Start_Y.Text = Convert.ToInt32(textBox_Now_Position_Y.Text).ToString();
            //textBox_Cut_Start_Z.Text = textBox_Now_Position_Z.Text;
        }
        private void button_Set_Cut_End_Click(object sender, EventArgs e)
        {
            textBox_Now_Position_X.Text = "28000";
            textBox_Now_Position_Y.Text = "30000";
            textBox_Cut_End_X.Text = Convert.ToInt32(textBox_Now_Position_X.Text).ToString();
            textBox_Cut_End_Y.Text = Convert.ToInt32(textBox_Now_Position_Y.Text).ToString();
        }
        private void button_Set_Safety_Position_Click(object sender, EventArgs e)
        {
            textBox_Safety_Hight.Text = textBox_Now_Position_Z.Text;
        }
        private void button_Standby_Click(object sender, EventArgs e)
        {
            textBox_Standby_X.Text = textBox_Now_Position_X.Text;
            textBox_Standby_Y.Text = textBox_Now_Position_Y.Text;
            //textBox_Standby_Z.Text = textBox_Now_Position_Z.Text;
        }
        private void button_Set_Scan_Start_Click(object sender, EventArgs e)
        {
            textBox_Scan_Start_X.Text = textBox_Now_Position_X.Text;
            textBox_Scan_Start_Y.Text = textBox_Now_Position_Y.Text;
            textBox_Scan_Start_Z.Text = textBox_Now_Position_Z.Text;
        }
        private void button_Set_Micro_Click(object sender, EventArgs e)
        {
            textBox_Micro_X.Text = textBox_Now_Position_X.Text;
            textBox_Micro_Y.Text = textBox_Now_Position_Y.Text;
            textBox_Micro_Z.Text = textBox_Now_Position_Z.Text;
        }
        #endregion

        #region Startup Step
        private void button_Connect_Motion_Click(object sender, EventArgs e)
        {
            logger.Write_Logger("Connect Motion");
            Motion_sp.PortName = comboBox_COM_Port.Text;
            Motion_sp.BaudRate = 38400;
            Motion_sp.DataBits = 8;
            Motion_sp.Parity = System.IO.Ports.Parity.None;
            Motion_sp.StopBits = System.IO.Ports.StopBits.One;
            Sp1_Terminator = "\r\n";
            Motion_sp.RtsEnable = true;
            Motion_sp.Open();
            timer_IO.Enabled = true;
            connect_Motion_OK = true;
            button_Connect_Motion.Enabled = false;
        }
        private void button_Z_ORG_Click(object sender, EventArgs e)
        {
            try {
                SetMotionSpeed(Convert.ToInt32(textBox_Move_Speed_X.Text) / 2, Convert.ToInt32(textBox_Move_Speed_Y.Text) / 2, Convert.ToInt32(textBox_Move_Speed_Z.Text) / 2);
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_Z.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "H:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Z ORG");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Z ORG Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_X_ORG_Click(object sender, EventArgs e)
        {
            try {
                SetMotionSpeed(Convert.ToInt32(textBox_Move_Speed_X.Text) / 2, Convert.ToInt32(textBox_Move_Speed_Y.Text) / 2, Convert.ToInt32(textBox_Move_Speed_Z.Text) / 2);
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_X.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "H:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("X ORG");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("X ORG Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Y_ORG_Click(object sender, EventArgs e)
        {
            try {
                SetMotionSpeed(Convert.ToInt32(textBox_Move_Speed_X.Text) / 2, Convert.ToInt32(textBox_Move_Speed_Y.Text) / 2, Convert.ToInt32(textBox_Move_Speed_Z.Text) / 2);
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[comboBox_Axis_Num_Y.SelectedIndex] = "1";
                String Write_Data = "";
                Write_Data = "H:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Y ORG");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Y ORG Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Standby_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Standby_X = Convert.ToInt32(Convert.ToInt32(textBox_Standby_X.Text) * Movement_Ratio);
                int Standby_Y = Convert.ToInt32(Convert.ToInt32(textBox_Standby_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Standby_X);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Standby_Y);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Standby XY : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Standby XY Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Safe_High_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Safety_Hight = Convert.ToInt32(Convert.ToInt32(textBox_Safety_Hight.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Safety_Hight);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Safe High : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Safe High Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void SetMotionSpeed(double xSpeed, double ySpeed, double zSpeed)
        {

            int X_Speed = Convert.ToInt32(xSpeed * Movement_Ratio);
            int Y_Speed = Convert.ToInt32(ySpeed * Movement_Ratio);
            int Z_Speed = Convert.ToInt32(zSpeed * Movement_Ratio);
            string Write_Data = "";
            //X
            Write_Data =
                "D:" +
                Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                "," + Convert.ToString(X_Speed / 4) +
                "," + Convert.ToString(X_Speed) + ",500";
            Write_Motion(Write_Data);
            //Y
            Write_Data =
                "D:" +
                Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                "," + Convert.ToString(Y_Speed / 4) +
                "," + Convert.ToString(Y_Speed) + ",500";
            Write_Motion(Write_Data);
            //Z
            Write_Data =
                "D:" +
                Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                "," + Convert.ToString(Z_Speed / 4) +
                "," + Convert.ToString(Z_Speed) + ",500";
            Write_Motion(Write_Data);

        }

        private void Move_Back1mm()
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Cut_End_X = Convert.ToInt32((Convert.ToInt32(textBox_Cut_End_X.Text) + 1000) * Movement_Ratio);
                int Cut_End_Y = Convert.ToInt32(Convert.ToInt32(textBox_Cut_End_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Cut_End_X);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Cut_End_Y);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                Write_Motion(Write_Data);
                logger.Write_Logger("Move Cut End XY : " + Write_Data);
            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Cut End XY error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        #endregion

        #region Process Step
        private void button_Move_XY_Micro_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Scan_Start_X = Convert.ToInt32(Convert.ToInt32(textBox_Scan_Start_X.Text) * Movement_Ratio);
                int Scan_Start_Y = Convert.ToInt32(Convert.ToInt32(textBox_Scan_Start_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Scan_Start_X);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Scan_Start_Y);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Micro XY : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Micro XY Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_Z_Micro_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                double Next_Layer = Convert.ToDouble(textBox_Now_layer.Text) + 1;
                double Next_Layer_Position =
                    Convert.ToDouble(textBox_Z_Down.Text) +
                    Convert.ToDouble(numericUpDown_Scan_Distance_Z.Text) * (Next_Layer - 1) * 1000;
                Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Next_Layer_Position * Movement_Ratio);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Micro Z : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Micro Z error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Start_Scan_Click(object sender, EventArgs e)
        {
            logger.Write_Logger("Start Scan");
            IO_Can_Cut = false;
            Andor_Error_Meaasge = "";
            andor_file_address = textBox_Andor_Exe_Address.Text;
            Protocal_name = comboBox_Process_Name.Text;
            backgroundWorker_Andor.RunWorkerAsync();
            timer_Andor.Enabled = true;
        }
        private void button_Move_XY_Cut_Start_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Cut_Start_X = Convert.ToInt32(Convert.ToInt32(textBox_Cut_Start_X.Text) * Movement_Ratio);
                int Cut_Start_Y = Convert.ToInt32(Convert.ToInt32(textBox_Cut_Start_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Cut_Start_X);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Cut_Start_Y);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Cut Start XY : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Cut Start XY error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Start_Hz_Click(object sender, EventArgs e)
        {
            logger.Write_Logger("Start Hz");
            //Start Hz
            /* timerPulseCtrl_Frequency.ChannelStart = 0;
             timerPulseCtrl_Frequency.ChannelCount = 1;
             timerPulseCtrl_Frequency.Frequency = Convert.ToInt32(numericUpDown_Cut_Frequency.Value);*/


            double minV = Convert.ToDouble(textBox_FrequencyMinV.Text);
            double maxV = Convert.ToDouble(textBox_FrequencyMaxV.Text);
            double outputV = Convert.ToDouble(textBox_FrequencyOutputV.Text);
            //nIHighFrequencyCutting.Run(outputV, minV, maxV);

            nIHighFrequencyCutting.On();

            //    timerPulseCtrl_Frequency.Enabled = true;
        }
        private void button_Move_Z_Cut_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                Cut_And_Scan_Finish = false;
                double Next_Layer = Convert.ToDouble(textBox_Now_layer.Text) + 1;
                if (Next_Layer > Convert.ToDouble(numericUpDown_Total_Cut_Layer.Text) + 1) {
                    Next_Layer = 1;
                    Cut_And_Scan_Finish = true;
                }
                if (!Cut_And_Scan_Finish) {
                    string[] Data = new string[3];
                    Data[0] = "";
                    Data[1] = "";
                    Data[2] = "";
                    double need_cal_dis = Convert.ToDouble(textBox_Z_Down.Text) - Convert.ToDouble(textBox_Scan_Start_Z.Text);
                    double Next_Layer_Position =
                        Convert.ToDouble(textBox_Cut_Start_Z.Text) + need_cal_dis +
                        Convert.ToDouble(numericUpDown_Cut_Distance.Text) * (Next_Layer - 1) * 1000;
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Next_Layer_Position * Movement_Ratio);
                    Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                    logger.Write_Logger("Move Cut Z : " + Write_Data);
                    Write_Motion(Write_Data);

                }
                else
                    button_Move_Safe_High_Click(sender, e);
            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Cut Z error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Move_XY_Cut_End_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Cut_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Cut_Speed_Y.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Cut_End_X = Convert.ToInt32(Convert.ToInt32(textBox_Cut_End_X.Text) * Movement_Ratio);
                int Cut_End_Y = Convert.ToInt32(Convert.ToInt32(textBox_Cut_End_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Cut_End_X);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Cut_End_Y);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Cut End XY : " + Write_Data);
                Write_Motion(Write_Data);
            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Cut End XY error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }
        private void button_Close_Hz_Click(object sender, EventArgs e)
        {
            double minV = Convert.ToDouble(textBox_FrequencyMinV.Text);
            double maxV = Convert.ToDouble(textBox_FrequencyMaxV.Text);
            logger.Write_Logger("Close Hz");
            //Close Hz
            //    nIHighFrequencyCutting.Run(0, minV, maxV);


            nIHighFrequencyCutting.Off();
            //    timerPulseCtrl_Frequency.Enabled = false;

        }
        #endregion

        #region Auto
        private void button_Auto_Click(object sender, EventArgs e)
        {
            if (timer_Step.Enabled) {
                logger.Write_Logger("Close Auto");
                //
                timer_Step.Enabled = false;
                button_Auto.BackColor = System.Drawing.Color.Transparent;
                button_Start.BackColor = System.Drawing.Color.Transparent;
                groupBox_Start_Setup.Enabled = true;
                groupBox_Process_Step.Enabled = true;
            }
            else {
                if (comboBox_Process_Name.Items.Count == 0) {
                    string now_process_name = comboBox_Process_Name.Text;
                    StreamWriter sw = new StreamWriter(Process_Nmae_File_Address);
                    sw.WriteLine(now_process_name);
                    sw.Close();
                }
                else {
                    bool same = false;
                    string now_process_name = comboBox_Process_Name.Text;
                    for (int i = 0; i < comboBox_Process_Name.Items.Count; i++) {
                        if (now_process_name == comboBox_Process_Name.Items[i].ToString())
                            same = true;
                    }
                    if (!same) {
                        StreamWriter sw = new StreamWriter(Process_Nmae_File_Address);
                        sw.WriteLine(now_process_name);
                        for (int i = 0; i < comboBox_Process_Name.Items.Count; i++)
                            sw.WriteLine(comboBox_Process_Name.Items[i].ToString());
                        sw.Close();
                    }
                }
                //
                timer_Step.Enabled = true;
                button_Auto.BackColor = Color.LightGreen;
                button_Start.BackColor = Color.LightGreen;
                groupBox_Start_Setup.Enabled = false;
                groupBox_Process_Step.Enabled = false;
                logger.Write_Logger("Start Auto");
            }
        }
        private void timer_Step_Tick(object sender, EventArgs e)
        {
            if (!X_Busy && !Y_Busy && !Z_Busy) {
                wait_delay++;
            }

            textBox_now_Step.Text = Convert.ToString(now_step);
            #region Start Setup
            if (now_step < 10) {
                if (now_step == 0) {
                    connect_Motion_OK = false;
                    button_Connect_Motion_Click(sender, e);
                    now_step = 1;
                }
                else if (now_step == 1 && connect_Motion_OK) {
                    wait_delay = 0;

                    button_Z_ORG_Click(sender, e);
                    now_step = 2;
                }
                else if (now_step == 2 &&
                    wait_delay >= wait_second &&
                    !Z_Busy &&
                    Convert.ToDouble(textBox_Now_Position_Z.Text) <= (0 + Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_Z.Text) >= (0 - Position_Range)) {
                    wait_delay = 0;
                    button_Y_ORG_Click(sender, e);
                    now_step = 3;
                }
                else if (now_step == 3 &&
                    wait_delay >= wait_second &&
                    !Y_Busy &&
                    Convert.ToDouble(textBox_Now_Position_Y.Text) <= (0 + Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_Y.Text) >= (0 - Position_Range)) {
                    wait_delay = 0;

                    button_X_ORG_Click(sender, e);
                    now_step = 4;
                }
                else if (now_step == 4 &&
                    wait_delay >= wait_second &&
                    !X_Busy &&
                    Convert.ToDouble(textBox_Now_Position_X.Text) <= (0 + Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_X.Text) >= (0 - Position_Range)) {
                    wait_delay = 0;
                    button_Move_Safe_High_Click(sender, e);
                    now_step = 5;
                }
                else if (now_step == 5 &&
                    wait_delay >= wait_second &&
                    !Z_Busy &&
                    Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Safety_Hight.Text) + Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Safety_Hight.Text) - Position_Range)) {
                    wait_delay = 0;
                    button_Move_Standby_Click(sender, e);
                    now_step = 6;
                }
                else if (now_step == 6 &&
                    wait_delay >= wait_second &&
                    !X_Busy && !Y_Busy &&
                    Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Standby_X.Text) + Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Standby_X.Text) - Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Standby_Y.Text) + Position_Range) &&
                    Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Standby_Y.Text) - Position_Range)) {
                    wait_delay = 0;
                    now_step = 10;
                }
            }
            #endregion

            #region Process Step
            else {
                if (need_scan) {
                    if (now_step == 10 &&
                        wait_delay >= wait_second &&
                        !X_Busy && !Y_Busy && !Z_Busy) {
                        wait_delay = 0;
                        button_Z_ORG_Click(sender, e);//Z軸原點復歸

                        now_step = 11;
                    }
                    else if (now_step == 11 &&
                        wait_delay >= wait_second &&
                        !Z_Busy &&
                        (Convert.ToDouble(textBox_Now_Position_Z.Text) == 0 || Convert.ToDouble(textBox_Now_Position_Z.Text) == Convert.ToDouble(textBox_Safety_Hight.Text))) {
                        button_Move_Safe_High_Click(sender, e);
                        now_step = 12;
                    }
                    else if (now_step == 12 &&
                        wait_delay >= wait_second &&
                        !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Safety_Hight.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Safety_Hight.Text) - Position_Range)) {
                        wait_delay = 0;
                        button_Move_XY_Micro_Click(sender, e);
                        now_step = 13;
                    }
                    else if (now_step == 13 &&
                        wait_delay >= wait_second &&
                        !X_Busy && !Y_Busy &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Scan_Start_X.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Scan_Start_X.Text) - Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Scan_Start_Y.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Scan_Start_Y.Text) - Position_Range)) {
                        wait_delay = 0;
                        button_Move_Z_Micro_Click(sender, e);
                        now_step = 14;
                    }
                    else if (now_step == 14 &&
                        wait_delay >= wait_second * 8 &&
                        !Z_Busy) {
                        wait_delay = 0;
                        IO_Can_Cut = false;
                        pictureBox_Scanning.Image = Properties.Resources.Green;
                        button_Start_Scan_Click(sender, e);
                        now_step = 15;
                    }
                    else if (now_step == 15 &&
                        wait_delay >= wait_second &&
                        IO_Can_Cut && Andor_Error_Meaasge.IndexOf("success") >= 0)//7.移動到安全高度
                    {
                        pictureBox_Scanning.Image = Properties.Resources.Red;
                        IO_Can_Cut = false;
                        wait_delay = 0;
                        button_Move_Safe_High_Click(sender, e);
                        if (Convert.ToInt32(textBox_Now_layer.Text) >= Convert.ToInt32(textBox_Cut_Layer.Text))
                            now_step = 32;
                        else
                            now_step = 16;
                    }
                    else if (now_step == 16 && wait_delay >= wait_second && !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Safety_Hight.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Safety_Hight.Text) - Position_Range))//8. 移動到切割起點XY
                    {
                        wait_delay = 0;
                        button_Move_XY_Cut_Start_Click(sender, e);
                        now_step = 17;
                    }
                    else if (now_step == 17 &&
                        wait_delay >= wait_second / 2 &&
                        !X_Busy && !Y_Busy &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Cut_Start_X.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Cut_Start_X.Text) - Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Cut_Start_Y.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Cut_Start_Y.Text) - Position_Range))//9.移動到切割區Z
                    {
                        wait_delay = 0;
                        int now_layer = Convert.ToInt32(textBox_Now_layer.Text);
                        textBox_Now_layer.Text = Convert.ToString(now_layer + 1);
                        button_Move_Z_Cut_Click(sender, e);

                        if (zaberMotion == null)//如果輸送帶沒有初始化過 ， 就連線
                            zaberMotion = new ZaberMotion(textBox_ZaberComPort.Text);
                        now_step = 18;


                    }
                    else if (now_step == 18 &&
                        wait_delay >= wait_second / 2 &&
                        !Z_Busy)  //10.開啟震動刀
                    {
                        pictureBox_Cut.Image = Properties.Resources.Green;
                        wait_delay = 0;
                        button_Start_Hz_Click(sender, e);
                        zaberMotion.Velocity = Convert.ToDouble(textBox_CVfirstVelocity.Text);  //輸送帶第一段速度
                        zaberMotion.MoveMax();

                        now_step = 19;
                    }
                    else if (now_step == 19 &&
                        wait_delay >= wait_second * 2 &&
                        !Z_Busy) //11.移動到切割xy終點
                    {
                        wait_delay = 0;
                        if (!Cut_And_Scan_Finish) {
                            button_Move_XY_Cut_End_Click(sender, e);
                            now_step = 20;
                        }
                        else
                            now_step = 30;
                    }
                    else if (now_step == 20 &&
                        wait_delay >= wait_second &&
                        !X_Busy && !Y_Busy &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Cut_End_X.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Cut_End_X.Text) - Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Cut_End_Y.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Cut_End_Y.Text) - Position_Range))//12.關閉震動刀  ， 退後 
                    {
                        wait_delay = 0;
                        pictureBox_Cut.Image = Properties.Resources.Red;
                        //button_Move_Safe_High_Click(sender, e);
                        button_Close_Hz_Click(sender, e);
                        double firstDelaytime = Convert.ToDouble(textBox_CVfirstDelaytime.Text);
                        double firstVelocity = Convert.ToDouble(textBox_CVfirstVelocity.Text);
                        double secondDelaytime = Convert.ToDouble(textBox_CVsecondDelaytime.Text);
                        double secondVelocity = Convert.ToDouble(textBox_CVsecondDelaytime.Text);
                        //關閉震動刀後 等待第一段延遲時間 ；  第一段延遲時間結束後 切換第二段速度與延遲滾動時間
                        zaberMotion.CVRun(firstVelocity, firstDelaytime, secondVelocity, secondDelaytime);

                        now_step = 21;
                    }
                    else if (now_step == 21 &&
                       wait_delay >= wait_second &&
                       !X_Busy && !Y_Busy) {
                        wait_delay = 0;
                        Move_Back1mm();
                        now_step = 22;
                    }
                    else if (now_step == 22 &&
                        wait_delay >= wait_second &&
                        !X_Busy && !Y_Busy) {

                        wait_delay = 0;
                        button_Z_ORG_Click(sender, e);
                        now_step = 23;
                    }
                    else if (now_step == 23 &&
                        wait_delay >= wait_second &&
                        !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (0 + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (0 - Position_Range)) {
                        wait_delay = 0;
                        //   button_Close_Hz_Click(sender, e);
                        now_step = 11;
                    }
                    else if (now_step == 30 &&
                        wait_delay >= wait_second / 2) {
                        wait_delay = 0;
                        button_Close_Hz_Click(sender, e);
                        now_step = 31;
                    }
                    else if (now_step == 31 &&
                        wait_delay >= wait_second / 2) {
                        wait_delay = 0;
                        button_Move_Safe_High_Click(sender, e);
                        now_step = 32;
                    }
                    else if (now_step == 32 &&
                        wait_delay >= wait_second &&
                        !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Safety_Hight.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Safety_Hight.Text) - Position_Range)) {
                        wait_delay = 0;
                        button_Move_Standby_Click(sender, e);
                        now_step = 33;
                    }
                    else if (now_step == 33 &&
                        wait_delay >= wait_second &&
                        !X_Busy && !Y_Busy &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Standby_X.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Standby_X.Text) - Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Standby_Y.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Standby_Y.Text) - Position_Range)) {
                        button_Start_Click(sender, e);
                        textBox_Now_layer.Text = "0";
                        now_step = 10;
                        wait_delay = 0;
                        logger.Write_Logger("Scan Finish");
                        MessageBox.Show("掃描完畢!");
                    }
                }
                else {
                    if (now_step == 50 &&
                        wait_delay >= wait_second &&
                        !X_Busy && !Y_Busy && !Z_Busy) {
                        wait_delay = 0;
                        now_step = 51;
                        button_Move_Safe_High_Click(sender, e);

                    }
                    else if (now_step == 51 &&
                        wait_delay >= wait_second / 2 &&
                        !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Safety_Hight.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Safety_Hight.Text) - Position_Range)) {
                        wait_delay = 0;
                        button_Move_Standby_Click(sender, e);
                        now_step = 52;
                    }
                    else if (now_step == 52 &&
                        wait_delay >= wait_second / 2 &&
                        !X_Busy && !Y_Busy &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Standby_X.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Standby_X.Text) - Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Standby_Y.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Standby_Y.Text) - Position_Range)) {
                        wait_delay = 0;
                        now_step = 10;
                    }
                    else if (now_step == 60 &&
                        wait_delay >= wait_second / 2 &&
                        !X_Busy && !Y_Busy && !Z_Busy) {
                        wait_delay = 0;
                        button_Move_Safe_High_Click(sender, e);
                        now_step = 61;

                    }
                    else if (now_step == 61 &&
                        wait_delay >= wait_second / 2 &&
                        !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Safety_Hight.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Safety_Hight.Text) - Position_Range)) {
                        wait_delay = 0;
                        Move_Micro_XY();
                        now_step = 62;
                    }
                    else if (now_step == 62 &&
                        wait_delay >= wait_second / 2 &&
                        !X_Busy && !Y_Busy &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) <= (Convert.ToInt32(textBox_Micro_X.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_X.Text) >= (Convert.ToInt32(textBox_Micro_X.Text) - Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) <= (Convert.ToInt32(textBox_Micro_Y.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Y.Text) >= (Convert.ToInt32(textBox_Micro_Y.Text) - Position_Range)) {
                        wait_delay = 0;
                        //Move_Micro_Z();
                        now_step = 63;
                    }
                    else if (now_step == 63 &&
                        wait_delay >= wait_second / 2 &&
                        !Z_Busy &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) <= (Convert.ToInt32(textBox_Micro_Z.Text) + Position_Range) &&
                        Convert.ToDouble(textBox_Now_Position_Z.Text) >= (Convert.ToInt32(textBox_Micro_Z.Text) - Position_Range)) {
                        wait_delay = 0;
                        now_step = 10;
                    }
                }
            }
            #endregion
        }
        private void backgroundWorker_Andor_DoWork(object sender, DoWorkEventArgs e)
        {
            try {
                string output;
                //
                Process p = new Process();
                p.StartInfo.FileName = andor_file_address;//需要執行的檔案路徑
                p.StartInfo.UseShellExecute = false; //必需
                p.StartInfo.RedirectStandardOutput = true;//輸出引數設定
                p.StartInfo.RedirectStandardInput = true;//傳入引數設定
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.Arguments = Protocal_name;//引數以空格分隔，如果某個引數為空，可以傳入””
                p.Start();
                output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();//關鍵，等待外部程式退出後才能往下執行}
                p.Close();
                //
                Andor_Error_Meaasge = output;
                IO_Can_Cut = true;
            }
            catch (Exception error) {

                //
                Andor_Error_Meaasge = error.ToString();
                IO_Can_Cut = true;
            }

        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (IO_Can_Cut) {
                timer_Andor.Enabled = false;
                logger.Write_Logger("Andoe say " + Andor_Error_Meaasge);
                //MessageBox.Show(Protocal_name + "\r\nAndoe say " + Andor_Error_Meaasge);
            }

        }
        private void button_Start_Click(object sender, EventArgs e)
        {
            if (comboBox_Process_Name.Text != "" && textBox_Andor_Exe_Address.Text != "") {
                if (need_scan) {
                    logger.Write_Logger("Stop Process");
                    button_Move_To_Standby_All.Enabled = true;
                    comboBox_Process_Name.Enabled = true;
                    button_Move_Micro.Enabled = true;
                    button_Load_Andor.Enabled = true;
                    need_scan = false;
                    panel_Motion.Enabled = true;
                    panel_Parameter.Enabled = true;
                    button_Step_Reset.Enabled = true;
                    button_Start.Text = "掃  描  開  始";
                    button_Start.BackColor = Color.LightGreen;
                    button_Auto.Enabled = true;
                    panel_Set_Scan_Data.Enabled = true;
                }
                else {
                    if (File.Exists(textBox_Andor_Exe_Address.Text)) {
                        logger.Write_Logger("Start Process");
                        //textBox_Scan_Start_Z.Text = textBox_Z_Down.Text;
                        need_scan = true;
                        button_Move_To_Standby_All.Enabled = false;
                        comboBox_Process_Name.Enabled = false;
                        button_Move_Micro.Enabled = false;
                        button_Load_Andor.Enabled = false;
                        panel_Motion.Enabled = false;
                        panel_Parameter.Enabled = false;
                        button_Step_Reset.Enabled = false;
                        StreamWriter sw_ = new StreamWriter(System.Windows.Forms.Application.StartupPath + "\\Setup\\Last_Process_Name.txt");
                        sw_.WriteLine(comboBox_Process_Name.Text);
                        sw_.Close();
                        button_Start.Text = "掃  描  停  止";
                        button_Start.BackColor = Color.Red;
                        button_Auto.Enabled = false;
                        panel_Set_Scan_Data.Enabled = false;
                    }
                    else {
                        MessageBox.Show("請先設定顯微鏡執行檔路徑");
                    }
                }
            }
        }
        private void comboBox_Process_Name_SelectedIndexChanged(object sender, EventArgs e)
        {
            logger.Write_Logger("Load Parameter " + comboBox_Process_Name.Text);
            Form_Value_Initial(System.Windows.Forms.Application.StartupPath + "\\Setup\\Process\\" + comboBox_Process_Name.Text + ".xml");
            textBox_ProcessName.Text = comboBox_Process_Name.Text;
        }
        private void textBox_Safety_Hight_TextChanged(object sender, EventArgs e)
        {
            textBox_Standby_Z.Text = textBox_Safety_Hight.Text;
        }
        private void numericUpDown_Total_Cut_Layer_ValueChanged(object sender, EventArgs e)
        {
            textBox_Total_Layer.Text = Convert.ToString(numericUpDown_Total_Cut_Layer.Value);
            textBox_Cut_Layer.Text = Convert.ToString(numericUpDown_Total_Cut_Layer.Value);
            numericUpDown_Total_Scan_Layer.Value = numericUpDown_Total_Cut_Layer.Value + 1;
        }
        private void button_Step_Reset_Click(object sender, EventArgs e)
        {
            textBox_Now_layer.Text = "0";
            now_step = 10;
        }

        #endregion

        #endregion

        #region Subfunction
        private void Form_Value_Initial(string Load_File_Address)
        {
            try {
                Variable Initial_Value = new Variable(Load_File_Address);
                //Main
                textBox_Total_Layer.Text = Convert.ToString(Initial_Value.Cut_Layer);
                //Scan Setup . Cut
                textBox_Cut_Start_X.Text = Convert.ToString(Initial_Value.Cut_Start_X);
                textBox_Cut_Start_Y.Text = Convert.ToString(Initial_Value.Cut_Start_Y);
                textBox_Cut_Start_Z.Text = Convert.ToString(Initial_Value.Cut_Start_Z);
                textBox_Cut_End_X.Text = Convert.ToString(Initial_Value.Cut_End_X);
                textBox_Cut_End_Y.Text = Convert.ToString(Initial_Value.Cut_End_Y);
                textBox_Cut_Speed_X.Text = Convert.ToString(Initial_Value.Cut_Speed_X);
                textBox_Cut_Speed_Y.Text = Convert.ToString(Initial_Value.Cut_Speed_Y);
                numericUpDown_Total_Cut_Layer.Value = Convert.ToDecimal(Initial_Value.Cut_Layer);
                numericUpDown_Total_Scan_Layer.Value = Convert.ToDecimal(Initial_Value.Cut_Layer + 1);
                numericUpDown_Cut_Distance.Value = Convert.ToDecimal(Initial_Value.Cut_Distance);
                numericUpDown_Cut_Frequency.Value = Convert.ToDecimal(Initial_Value.Cut_Frequency);
                numericUpDown_Cut.Value = numericUpDown_Cut_Frequency.Value;
                numericUpDown_Cut_Dis_M.Value = numericUpDown_Cut_Distance.Value;
                numericUpDown_Scan_Dis_M.Value = numericUpDown_Cut_Dis_M.Value;
                //Scan Setup . Move
                textBox_Safety_Hight.Text = Convert.ToString(Initial_Value.Safety_Hight);
                textBox_Move_Speed_X.Text = Convert.ToString(Initial_Value.Move_Speed_X);
                textBox_Move_Speed_Y.Text = Convert.ToString(Initial_Value.Move_Speed_Y);
                textBox_Move_Speed_Z.Text = Convert.ToString(Initial_Value.Move_Speed_Z);
                textBox_Standby_X.Text = Convert.ToString(Initial_Value.StandBy_X);
                textBox_Standby_Y.Text = Convert.ToString(Initial_Value.StandBy_Y);
                textBox_Standby_Z.Text = Convert.ToString(Initial_Value.StandBy_Z);
                textBox_Z_Up.Text = Convert.ToString(Initial_Value.End_Scan_Z);
                textBox_Z_Down.Text = Convert.ToString(Initial_Value.Start_Scan_Z);
                //Scan Setup . Scan
                textBox_Scan_Start_X.Text = Convert.ToString(Initial_Value.Scan_Start_X);
                textBox_Scan_Start_Y.Text = Convert.ToString(Initial_Value.Scan_Start_Y);
                textBox_Scan_Start_Z.Text = Convert.ToString(Initial_Value.Scan_Start_Z);
                numericUpDown_Scan_Distance_Z.Value = Convert.ToDecimal(Initial_Value.Scan_Distance_Z);
                textBox_Cut_Layer.Text = Convert.ToString(Initial_Value.Cut_Layer);
                //Motion Setup
                for (int i = 0; i < comboBox_COM_Port.Items.Count; i++) {
                    if (Convert.ToString(comboBox_COM_Port.Items[i]) == Initial_Value.COM_PORT)
                        comboBox_COM_Port.SelectedIndex = i;
                }
                checkBox_Inverse_X.Checked = Initial_Value.X_Inverse;
                checkBox_Inverse_Y.Checked = Initial_Value.Y_Inverse;
                checkBox_Inverse_Z.Checked = Initial_Value.Z_Inverse;
                checkBox_Inverse_XY.Checked = Initial_Value.XY_Inverse;
                comboBox_Axis_Num_X.SelectedIndex = Initial_Value.X_Number;
                comboBox_Axis_Num_Y.SelectedIndex = Initial_Value.Y_Number;
                comboBox_Axis_Num_Z.SelectedIndex = Initial_Value.Z_Number;
                textBox_Step_Speed_X.Text = Convert.ToString(Initial_Value.Move_Speed_X);
                textBox_Step_Speed_Y.Text = Convert.ToString(Initial_Value.Move_Speed_Y);
                textBox_Step_Speed_Z.Text = Convert.ToString(Initial_Value.Move_Speed_Z);
                textBox_Z_Speed_M.Text = textBox_Step_Speed_Z.Text;
                textBox_Cut_Frequency.Text = Convert.ToString(numericUpDown_Cut_Frequency.Value);
                textBox_Micro_X.Text = Convert.ToString(Initial_Value.Micro_X);
                textBox_Micro_Y.Text = Convert.ToString(Initial_Value.Micro_Y);
                textBox_Micro_Z.Text = Convert.ToString(Initial_Value.Micro_Z);
                Movement_Ratio = Initial_Value.Motion_Ratio;

                textBox_CVfirstDelaytime.Text = Convert.ToString(Initial_Value.CVfirstDelaytime);
                textBox_CVfirstVelocity.Text = Convert.ToString(Initial_Value.CVfirstVelocity);
                textBox_CVsecondDelaytime.Text = Convert.ToString(Initial_Value.CVsecondDelaytime);
                textBox_CVsecondVelocity.Text = Convert.ToString(Initial_Value.CVsecondVelocity);
                textBox_ZaberComPort.Text = Convert.ToString(Initial_Value.CVZaberComPort);
                textBox_FrequencyMinV.Text = Convert.ToString(Initial_Value.FrequencyMinV);
                textBox_FrequencyMaxV.Text = Convert.ToString(Initial_Value.FrequencyMaxV);

            }
            catch (Exception error) {
                logger.Write_Error_Logger(error.ToString());
                textBox_Total_Layer.Text = Convert.ToString(50);
                //Scan Setup . Cut
                textBox_Cut_Start_X.Text = Convert.ToString(100);
                textBox_Cut_Start_Y.Text = Convert.ToString(100);
                textBox_Cut_Start_Z.Text = Convert.ToString(100);
                textBox_Cut_End_X.Text = Convert.ToString(100);
                textBox_Cut_End_Y.Text = Convert.ToString(100);
                textBox_Cut_Speed_X.Text = Convert.ToString(100);
                textBox_Cut_Speed_Y.Text = Convert.ToString(100);
                numericUpDown_Total_Cut_Layer.Value = Convert.ToDecimal(50);
                numericUpDown_Cut_Distance.Value = Convert.ToDecimal(1.1);
                numericUpDown_Cut_Frequency.Value = Convert.ToDecimal(60);
                //Scan Setup . Move
                textBox_Safety_Hight.Text = Convert.ToString(100);
                textBox_Move_Speed_X.Text = Convert.ToString(100);
                textBox_Move_Speed_Y.Text = Convert.ToString(100);
                textBox_Move_Speed_Z.Text = Convert.ToString(100);
                //Scan Setup . Scan
                textBox_Scan_Start_X.Text = Convert.ToString(100);
                textBox_Scan_Start_Y.Text = Convert.ToString(100);
                textBox_Scan_Start_Z.Text = Convert.ToString(100);
                numericUpDown_Scan_Distance_Z.Value = Convert.ToDecimal(100);
                textBox_Cut_Layer.Text = Convert.ToString(100);
                //Motion Setup
                for (int i = 0; i < comboBox_COM_Port.Items.Count; i++) {
                    comboBox_COM_Port.SelectedIndex = 0;
                }
                checkBox_Inverse_X.Checked = false;
                checkBox_Inverse_Y.Checked = false;
                checkBox_Inverse_Z.Checked = false;
                checkBox_Inverse_XY.Checked = false;
                comboBox_Axis_Num_X.SelectedIndex = 1;
                comboBox_Axis_Num_Y.SelectedIndex = 2;
                comboBox_Axis_Num_Z.SelectedIndex = 0;
                textBox_Step_Speed_X.Text = Convert.ToString(100);
                textBox_Step_Speed_Y.Text = Convert.ToString(100);
                textBox_Step_Speed_Z.Text = Convert.ToString(100);
                textBox_Micro_X.Text = Convert.ToString(100);
                textBox_Micro_Y.Text = Convert.ToString(100);
                textBox_Micro_Z.Text = Convert.ToString(100);
            }
        }
        private void Write_Motion(string Write_Data)
        {
            Motion_sp.Write(Write_Data + Sp1_Terminator);
        }
        private void Move_Micro_XY()
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Micro_X = Convert.ToInt32(Convert.ToInt32(textBox_Micro_X.Text) * Movement_Ratio);
                int Micro_Y = Convert.ToInt32(Convert.ToInt32(textBox_Micro_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Micro_X);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Micro_Y);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Micro XY");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Micro XY Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void textBox_Scan_Start_Z_TextChanged(object sender, EventArgs e)
        {
            textBox_Z_Down.Text = textBox_Scan_Start_Z.Text;
            textBox_Z_Up.Text = Convert.ToString(Convert.ToDouble(textBox_Z_Down.Text) + Convert.ToDouble(numericUpDown_Total_Scan_Layer.Value) * Convert.ToDouble(numericUpDown_Scan_Dis_M.Value));
        }

        private void Move_Micro_Z()
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Micro_Z = Convert.ToInt32(Convert.ToInt32(textBox_Micro_Z.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Micro_Z);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Micro Z");
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Micro Z error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Safety_Hight = Convert.ToInt32(Convert.ToInt32(textBox_Z_Up.Text) * Movement_Ratio);//改變移動高度
                Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Safety_Hight);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Safe High : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Safe High Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button_Move_Z_P_MouseDown(sender, (MouseEventArgs)e);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button_Move_Z_N_MouseDown(sender, (MouseEventArgs)e);
        }

        private void button_Move_Pos_X_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Safety_X = Convert.ToInt32(Convert.ToInt32(textBox_Step_Pos_X.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_X.SelectedIndex] = Convert.ToString(Safety_X);//移動至X位置
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Safe High : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Safe High Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void button_Move_Pos_Y_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Safety_Y = Convert.ToInt32(Convert.ToInt32(textBox_Step_Pos_Y.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_Y.SelectedIndex] = Convert.ToString(Safety_Y);//改變Y位置
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Safe High : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Safe High Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void button_Move_Pos_Z_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Safety_Z = Convert.ToInt32(Convert.ToInt32(textBox_Step_Pos_Z.Text) * Movement_Ratio);
                Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Safety_Z);//改變Z位置
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Safe High : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Safe High Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void button_Move_Z_Down_Click(object sender, EventArgs e)
        {
            try {
                int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                string Write_Data = "";
                //X
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                    "," + Convert.ToString(X_Speed / 4) +
                    "," + Convert.ToString(X_Speed) + ",500";
                Write_Motion(Write_Data);
                //Y
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                    "," + Convert.ToString(Y_Speed / 4) +
                    "," + Convert.ToString(Y_Speed) + ",500";
                Write_Motion(Write_Data);
                //Z
                Write_Data =
                    "D:" +
                    Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                    "," + Convert.ToString(Z_Speed / 4) +
                    "," + Convert.ToString(Z_Speed) + ",500";
                Write_Motion(Write_Data);
                //
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                int Safety_Hight = Convert.ToInt32(Convert.ToInt32(textBox_Z_Down.Text) * Movement_Ratio);//改變移動高度
                Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(Safety_Hight);
                Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                logger.Write_Logger("Move Safe High : " + Write_Data);
                Write_Motion(Write_Data);

            }
            catch (Exception error) {
                logger.Write_Error_Logger("Move Safe High Error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }
        }

        private void button_Move_Z_Pos_Down_Click(object sender, EventArgs e)
        {
            if (!Z_Busy && move_z_Step_ok && (Convert.ToInt32(textBox_Now_Position_Z_Main.Text) - Convert.ToInt32(textBox_Z_Step_M.Text)) > 0) {
                move_z_Step_ok = false;
                int move_step = Convert.ToInt32(Convert.ToInt32(textBox_Z_Step_M.Text) * Movement_Ratio);
                int now_Z = Convert.ToInt32(Convert.ToInt32(textBox_Now_Position_Z_Main.Text) * Movement_Ratio);
                int move_Z = now_Z - move_step;
                motion_move_Z = move_Z;
                try {
                    int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                    int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                    int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                    string Write_Data = "";
                    //X
                    Write_Data =
                        "D:" +
                        Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                        "," + Convert.ToString(X_Speed / 4) +
                        "," + Convert.ToString(X_Speed) + ",500";
                    Write_Motion(Write_Data);
                    //Y
                    Write_Data =
                        "D:" +
                        Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                        "," + Convert.ToString(Y_Speed / 4) +
                        "," + Convert.ToString(Y_Speed) + ",500";
                    Write_Motion(Write_Data);
                    //Z
                    Write_Data =
                        "D:" +
                        Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                        "," + Convert.ToString(Z_Speed / 4) +
                        "," + Convert.ToString(Z_Speed) + ",500";
                    Write_Motion(Write_Data);
                    //
                    string[] Data = new string[3];
                    Data[0] = "";
                    Data[1] = "";
                    Data[2] = "";
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(move_Z);
                    Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                    logger.Write_Logger("Move Step High : " + Write_Data);
                    Write_Motion(Write_Data);

                }
                catch (Exception error) {
                    logger.Write_Error_Logger("Move Step High Error " + Convert.ToString(error));
                    MessageBox.Show("Move Error!\n" + Convert.ToString(error));
                }
            }
        }

        private void button_Move_Z_Pos_Up_Click(object sender, EventArgs e)
        {
            if (!Z_Busy && move_z_Step_ok) {
                move_z_Step_ok = false;
                int move_step = Convert.ToInt32(Convert.ToInt32(textBox_Z_Step_M.Text) * Movement_Ratio);
                int now_Z = Convert.ToInt32(Convert.ToInt32(textBox_Now_Position_Z_Main.Text) * Movement_Ratio);
                int move_Z = now_Z + move_step;
                motion_move_Z = move_Z;
                try {
                    int X_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_X.Text) * Movement_Ratio);
                    int Y_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Y.Text) * Movement_Ratio);
                    int Z_Speed = Convert.ToInt32(Convert.ToInt32(textBox_Move_Speed_Z.Text) * Movement_Ratio);
                    string Write_Data = "";
                    //X
                    Write_Data =
                        "D:" +
                        Convert.ToString(comboBox_Axis_Num_X.SelectedIndex) +
                        "," + Convert.ToString(X_Speed / 4) +
                        "," + Convert.ToString(X_Speed) + ",500";
                    Write_Motion(Write_Data);
                    //Y
                    Write_Data =
                        "D:" +
                        Convert.ToString(comboBox_Axis_Num_Y.SelectedIndex) +
                        "," + Convert.ToString(Y_Speed / 4) +
                        "," + Convert.ToString(Y_Speed) + ",500";
                    Write_Motion(Write_Data);
                    //Z
                    Write_Data =
                        "D:" +
                        Convert.ToString(comboBox_Axis_Num_Z.SelectedIndex) +
                        "," + Convert.ToString(Z_Speed / 4) +
                        "," + Convert.ToString(Z_Speed) + ",500";
                    Write_Motion(Write_Data);
                    //
                    string[] Data = new string[3];
                    Data[0] = "";
                    Data[1] = "";
                    Data[2] = "";
                    Data[comboBox_Axis_Num_Z.SelectedIndex] = Convert.ToString(move_Z);
                    Write_Data = "A:" + Data[0] + "," + Data[1] + "," + Data[2];
                    logger.Write_Logger("Move Step High : " + Write_Data);
                    Write_Motion(Write_Data);

                }
                catch (Exception error) {
                    logger.Write_Error_Logger("Move Step High Error " + Convert.ToString(error));
                    MessageBox.Show("Move Error!\n" + Convert.ToString(error));
                }
            }
        }

        private void btn_Stop_Click(object sender, EventArgs e)
        {
            AxisStop(comboBox_Axis_Num_X.SelectedIndex);
            AxisStop(comboBox_Axis_Num_Y.SelectedIndex);
            AxisStop(comboBox_Axis_Num_Z.SelectedIndex);
        }
        private void AxisStop(int selectAxis)
        {
            try {
                string[] Data = new string[3];
                Data[0] = "";
                Data[1] = "";
                Data[2] = "";
                Data[selectAxis] = "1";
                String Write_Data = "";
                Write_Data = "L:" + Data[0] + "," + Data[1] + "," + Data[2];
                Write_Motion(Write_Data);
                logger.Write_Logger(" Stop");
            }
            catch (Exception error) {
                logger.Write_Error_Logger("  Stop error " + Convert.ToString(error));
                MessageBox.Show("Move Error!\n" + Convert.ToString(error));
            }

        }

        private async void btn_CVRun_Click(object sender, EventArgs e)
        {
            try {
                if (zaberMotion == null)
                    zaberMotion = new ZaberMotion(textBox_ZaberComPort.Text);

                double firstDelaytime = Convert.ToDouble(textBox_CVfirstDelaytime.Text);
                double firstVelocity = Convert.ToDouble(textBox_CVfirstVelocity.Text);
                double secondDelaytime = Convert.ToDouble(textBox_CVsecondDelaytime.Text);
                double secondVelocity = Convert.ToDouble(textBox_CVsecondVelocity.Text);
                //    zaberMotion.Home();
                zaberMotion.MoveMax();
                await zaberMotion.CVRun(firstVelocity, firstDelaytime, secondVelocity, secondDelaytime);
                zaberMotion.Stop();

            }
            catch (Exception ex) {

                MessageBox.Show(ex.Message);
            }

        }

        private void textBox_Cut_Speed_X_TextChanged(object sender, EventArgs e)
        {
            double speed;
            if (Double.TryParse(textBox_Cut_Speed_X.Text, out speed)) {
                var arcSpeed = speed * 360 / (1000 * 48 * Math.PI);
                textBox_CVfirstVelocity.Text = arcSpeed.ToString("0.000");
            }


        }

        private void Cal_Data()
        {
            double Z_Down = Convert.ToDouble(textBox_Z_Down.Text);
            double Z_Up = Convert.ToDouble(textBox_Z_Up.Text);
            if (Z_Up > Z_Down) {
                double Z_Length = Z_Up - Z_Down;
                double Z_Distance = Z_Length / Convert.ToDouble(numericUpDown_Total_Scan_Layer.Value);
                double Z_Di_mm = Z_Distance / 1000.0;
                double Z_Di_Display = Convert.ToDouble(Convert.ToInt32(Z_Di_mm * 100.0)) / 100.0;
                numericUpDown_Cut_Dis_M.Value = Convert.ToDecimal(Z_Di_Display);
                numericUpDown_Scan_Dis_M.Value = Convert.ToDecimal(Z_Di_Display);

                textBox_Top_Bottom_Diff.Text = ((int)Z_Length).ToString();//把高度間距顯示在畫面中

            }
            else {
                DialogResult result = MessageBox.Show("高度設定相反\n是否互換?", "錯誤!", MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.Yes) {
                    textBox_Z_Down.Text = Convert.ToString(Z_Up);
                    textBox_Z_Up.Text = Convert.ToString(Z_Down);
                    Cal_Data();
                }
            }
        }

        private void Imshow_Real_Time_value(string key, double value)
        {
            switch (key) {
                case "Total_Scan_Layer":
                    textBox_Total_Scan_Layer.Text = value.ToString();
                    break;
                case "Cut":
                    textBox_Cut_Feq.Text = value.ToString();
                    break;
                case "Cut_Dis_M":
                    textBox_Cut_Dis_M.Text = value.ToString();
                    break;
                case "Scan_Dis_M":
                    textBox_Scan_Dis_M.Text = value.ToString();
                    break;
            }
        }
        #endregion

    }
}
