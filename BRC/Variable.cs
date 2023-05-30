using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows.Forms;

namespace BRC
{
    class Variable
    {
        public string COM_PORT;
        public bool X_Inverse, Y_Inverse, Z_Inverse, XY_Inverse;
        public int X_Number, Y_Number, Z_Number;
        public double Motion_Ratio;
        public double Cut_Start_X, Cut_Start_Y, Cut_Start_Z;
        public double Cut_End_X, Cut_End_Y;
        public double Cut_Speed_X, Cut_Speed_Y;
        public int Cut_Layer, Cut_Frequency;
        public double Cut_Distance;
        public double Scan_Start_X, Scan_Start_Y, Scan_Start_Z;
        public double Scan_End_X, Scan_End_Y;
        public double Scan_Speed_X, Scan_Speed_Y;
        public double Scan_Point_X, Scan_Point_Y;
        public double Scan_Distance_X, Scan_Distance_Y, Scan_Distance_Z;
        public int Scan_Type;
        public double Safety_Hight;
        public double Move_Speed_X, Move_Speed_Y, Move_Speed_Z;
        public double StandBy_X, StandBy_Y, StandBy_Z;
        public int Can_Scan_Port, Can_Scan_Channel;
        public int Can_Move_Port, Can_Move_Channel;
        public int Can_Next_Port, Can_Next_Channel;
        public double Micro_X, Micro_Y, Micro_Z;
        public double Start_Scan_Z,End_Scan_Z;

        public double FrequencyMinV { get; set; }
        public double FrequencyMaxV { get; set; }
        public double FrequencyOutputV { get; set; }
        public double CVfirstDelaytime { get; set; }
        public double CVsecondDelaytime { get; set; }
        public double CVfirstVelocity { get; set; }
        public double CVsecondVelocity { get; set; }
        public string CVZaberComPort { get; set; }

        public Variable(string Load_File_Path)
        {
            try
            {
                //String Load_File_Path = System.Windows.Forms.Application.StartupPath + "\\Setup\\Initial_Value.xml";
                //
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Load_File_Path);//載入xml檔
                                            //
                XmlNode Fist_Node = xmlDoc.SelectSingleNode("BRC_Program_Setup");
                //顯示根目錄(A層)下第一層(B層)的所有屬性值
                XmlNodeList Second_Node = Fist_Node.ChildNodes;
                //
                foreach(XmlNode  Second_Node_Each in Second_Node)

                {
                    XmlElement Second_Node_XmlElement = (XmlElement)Second_Node_Each;
                    String Second_Node_Data = Second_Node_XmlElement.GetAttribute("Setup_Part");
                    //顯示根目錄(B層)下第一層(C層)的所有屬性值
                    XmlNodeList Third_Node = Second_Node_XmlElement.ChildNodes;
                    foreach(XmlNode  Third_Node_Each in Third_Node)

                    {
                        XmlElement Third_Node_Each_XmlElement = (XmlElement)Third_Node_Each;
                        String Third_Node_Data_1 = Third_Node_Each_XmlElement.GetAttribute("Setup");
                        if (Second_Node_Data == "Motion")
                        {
                            if (Third_Node_Data_1 == "COM_PORT")
                                COM_PORT = Convert.ToString(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "X_Inverse")
                                X_Inverse = Convert.ToBoolean(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Y_Inverse")
                                Y_Inverse = Convert.ToBoolean(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Z_Inverse")
                                Z_Inverse = Convert.ToBoolean(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "XY_Inverse")
                                XY_Inverse = Convert.ToBoolean(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "X_Number")
                                X_Number = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Y_Number")
                                Y_Number = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Z_Number")
                                Z_Number = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Ratio")
                                Motion_Ratio = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                        }
                        else if (Second_Node_Data == "Cut_Data")
                        {
                            if (Third_Node_Data_1 == "Cut_Start_X")
                                Cut_Start_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Start_Y")
                                Cut_Start_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Start_Z")
                                Cut_Start_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_End_X")
                                Cut_End_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_End_Y")
                                Cut_End_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Speed_X")
                                Cut_Speed_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Speed_Y")
                                Cut_Speed_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Layer")
                                Cut_Layer = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Distance")
                                Cut_Distance = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Cut_Frequency")
                                Cut_Frequency = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "FrequencyMinV")
                                FrequencyMinV = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "FrequencyMaxV")
                                FrequencyMaxV = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "FrequencyOutputV")
                                FrequencyOutputV = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                        }
                        else if (Second_Node_Data == "Scan_Data")
                        {
                            if (Third_Node_Data_1 == "Scan_Start_X")
                                Scan_Start_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Start_Y")
                                Scan_Start_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Start_Z")
                                Scan_Start_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_End_X")
                                Scan_End_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_End_Y")
                                Scan_End_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Speed_X")
                                Scan_Speed_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Speed_Y")
                                Scan_Speed_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Point_X")
                                Scan_Point_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Point_Y")
                                Scan_Point_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Distance_X")
                                Scan_Distance_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Distance_Y")
                                Scan_Distance_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Distance_Z")
                                Scan_Distance_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Scan_Type")
                                Scan_Type = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Start_Scan_Z")
                                Start_Scan_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "End_Scan_Z")
                                End_Scan_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                        }
                        else if (Second_Node_Data == "Move")
                        {
                            if (Third_Node_Data_1 == "Safety_Hight")
                                Safety_Hight = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Move_Speed_X")
                                Move_Speed_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Move_Speed_Y")
                                Move_Speed_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Move_Speed_Z")
                                Move_Speed_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "StandBy_X")
                                StandBy_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "StandBy_Y")
                                StandBy_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "StandBy_Z")
                                StandBy_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Micro_X")
                                Micro_X = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Micro_Y")
                                Micro_Y = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Micro_Z")
                                Micro_Z = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);

                            else if (Third_Node_Data_1 == "CVfirstDelaytime")
                                CVfirstDelaytime = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "CVsecondDelaytime")
                                CVsecondDelaytime = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "CVfirstVelocity")
                                CVfirstVelocity = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "CVsecondVelocity")
                                CVsecondVelocity = Convert.ToDouble(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "CVZaberComPort")
                                CVZaberComPort = Convert.ToString(Third_Node_Each_XmlElement.InnerText);
                        }
                        else if (Second_Node_Data == "IO")
                        {
                            if (Third_Node_Data_1 == "Can_Scan_Port")
                                Can_Scan_Port = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Can_Scan_Channel")
                                Can_Scan_Channel = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Can_Move_Port")
                                Can_Move_Port = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Can_Move_Channel")
                                Can_Move_Channel = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Can_Next_Port")
                                Can_Next_Port = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                            else if (Third_Node_Data_1 == "Can_Next_Channel")
                                Can_Next_Channel = Convert.ToInt32(Third_Node_Each_XmlElement.InnerText);
                        }
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(Convert.ToString(error));
            }
        }
    }
}
