using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace BRC.Component
{
    public class Rest_Fusion
    {

        //string fusion_state = "/v1/protocol/state";
        //string fusion_current = "/v1/protocol/current";
        //string fusion_progress = "/v1/protocol/progress";
        private string m_host = "localhost";
        private string m_port = "15120";
        private string m_Devicename = "xyz-stage";//"xyz-stage";//"ASI-Motorized-Stage";
                                                       //"dummy-xy-stage";
        private string m_Xfeaturename = "xposition";
        private string m_Yfeaturename = "yposition";

        private string baseUrl = "http://localhost:15120"; // 修改成你的 API 網址


        public string m_NowX;
        public string m_NowY;

        private string m_TargetX;
        private string m_TargetY;

        private string url;
        public string State;
        public string ErrorMessage = "";
        Fusion_ProgressData Progress;
        public string Name;

        object RestLock = new object();

        public Rest_Fusion(string url)
        {
            try
            {
                this.url = url;

                //m_MoveAbsX = m_NowX;
                //m_MoveAbsY = m_NowY;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public async Task Initial()
        {
            try
            {
                await Task.Run(async () =>
               {
                   await Sample();
                   m_TargetX = m_NowX;
                   m_TargetY = m_NowY;
               });
            }
            catch (Exception)
            {
                throw;
            }
        }

        #region Fusion_python轉C#
        public class Fusion_API
        {
            /// <summary>
            /// Gives the name of the API endpoint for which the error happened.
            /// </summary>
            public string Endpoint { get; set; }

            /// <summary>
            /// Gives the HTTP response code for the error, as returned by the API.
            /// Also see `.reason()` for a more readable description of the problem.
            /// </summary>
            public string Code { get; set; }

            /// <summary>
            /// Gives the reason for the error, as returned by the API. (a string)
            /// </summary>
            public string Reason { get; set; }
        }

        public class Fusion_Data
        {
            public string Title { get; set; }//三種 Name State Progress
            public string Value { get; set; }

        }

        class Fusion_StateData
        {
            public string State { get; set; }
        }
        class Fusion_NameData
        {
            public string Name { get; set; }
        }
        class Fusion_ValueData
        {
            public string Value { get; set; }
        }
        class Fusion_DeviceData
        {
            public string Name { get; set; }
            public string Value { get; set; }
            public string Min { get; set; }
            public string Max { get; set; }

        }
        class Fusion_ProgressData
        {
            public string StartTime { get; set; }
            public string ElapsedTime { get; set; }
            public string RemainingTime { get; set; }
            public string EstimatedTimeOfCompletion { get; set; }
            public string Progress { get; set; }

        }



        private string Make_Address(string endPoint)
        {
            try
            {
                return $"http://{m_host}:{m_port}{endPoint}";
            }
            catch (Exception ex)
            {

                throw ex;
                //return "";
            }

        }




        public async Task Get_Value(string endpoint, string key)
        {
            try
            {
                await Task.Run(() =>
               {
                   try
                   {
                       baseUrl = $"http://{m_host}:{m_port}";
                       string sendMessage = Make_Address(endpoint);
                       var restClient = new RestClient(baseUrl);
                       var restRequest = new RestRequest(endpoint, Method.Get);
                       var response = restClient.ExecuteAsync(restRequest);

                       while (response.IsCompleted == false)
                       {
                       }
                       if (response.Result.IsSuccessful)
                       {
                           var content = response.Result.Content;

                           if (key == "Name")
                           {
                               var data = JsonConvert.DeserializeObject<Fusion_NameData>(content);
                               this.Name = data.Name;
                           }
                           else if (key == "Progress")
                           {
                               var data = JsonConvert.DeserializeObject<Fusion_ProgressData>(content);
                               this.Progress = data;
                           }
                           else if (key == "xposition" || key == "yposition")
                           {
                               var data = JsonConvert.DeserializeObject<Fusion_ValueData>(content);
                               if (key == "xposition")
                               {
                                   m_NowX = data.Value;
                               }
                               else
                               {
                                   m_NowY = data.Value;
                               }
                           }
                           else if (key == "device")
                           {
                               var data = JsonConvert.DeserializeObject<List<Fusion_DeviceData>>(content);
                           }
                           else
                           {
                               var data = JsonConvert.DeserializeObject<Fusion_StateData>(content);
                               this.State = data.State;
                           }
                       }
                       else
                       {
                           this.State = "Error";
                           // 處理錯誤回應
                       }
                       //lock (RestLock)
                       //{

                       //}
                   }
                   catch (Exception ex)
                   {
                       this.State = "Error";
                       throw ex;
                   }
               });
            }
            catch (Exception ex)
            {
                this.State = "Error";
                throw ex;
            }
        }
        private async Task Get_ValueOLD(string endpoint, string key)//Get 獲取所有用戶
        {
            try
            {
                await Task.Run(async () =>
                {
                    try
                    {
                        string sendMessage = Make_Address(endpoint);
                        using (var httpClient = new HttpClient())
                        {
                            HttpResponseMessage response = await httpClient.GetAsync(sendMessage);

                            if (response.IsSuccessStatusCode)
                            {
                                var content = response.Content.ReadAsStringAsync().Result;

                                var data2 = JsonConvert.DeserializeObject(content);

                                if (key == "Name")
                                {
                                    var data = JsonConvert.DeserializeObject<Fusion_NameData>(content);
                                    this.Name = data.Name;
                                }
                                else if (key == "Progress")
                                {
                                    var data = JsonConvert.DeserializeObject<Fusion_ProgressData>(content);
                                    this.Progress = data;
                                }
                                else if (key == "xposition" || key == "yposition")
                                {
                                    var data = JsonConvert.DeserializeObject<Fusion_ValueData>(content);
                                    if (key == "xposition")
                                    {
                                        m_NowX = data.Value;
                                    }
                                    else
                                    {
                                        m_NowY = data.Value;
                                    }
                                }
                                else if (key == "device")
                                {
                                    var data = JsonConvert.DeserializeObject<List<Fusion_DeviceData>>(content);

                                }
                                else
                                {
                                    var data = JsonConvert.DeserializeObject<Fusion_StateData>(content);
                                    this.State = data.State;
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }
                });



            }
            catch (Exception ex)
            {

                throw ex;
            }

        }





        private async Task Put_Value(string endpoint, string key, string value)
        {
            try
            {
                await Task.Run(() =>
               {
                   lock (RestLock)
                   {
                       try
                       {
                           var client = new RestClient(baseUrl);
                           var request = new RestRequest(endpoint, Method.Put);
                           string json;

                           if (key == "Name")
                           {
                               var obj = new { Name = value };
                               json = JsonConvert.SerializeObject(obj);
                           }
                           else if (key == "Value")
                           {
                               var obj = new { Value = value };
                               json = JsonConvert.SerializeObject(obj);
                           }
                           else
                           {
                               var obj = new { State = value };
                               json = JsonConvert.SerializeObject(obj);
                           }

                           request.AddParameter("application/json", json, ParameterType.RequestBody);
                           var response = client.ExecuteAsync(request);
                           while (response.IsCompleted == false)
                           {
                           }
                           if (response.Result.IsSuccessful)
                           {
                               ErrorMessage = "";
                           }
                           else
                           {
                               ErrorMessage = "Error";
                           }
                       }
                       catch (Exception ex)
                       {
                           ErrorMessage = "Error";
                           throw ex;
                       }
                   }
               });
            }
            catch (Exception ex)
            {
                ErrorMessage = "Error";
                throw ex;
            }
        }
        private async Task Put_ValueOLD(string endpoint, string key, string value)//Put 更新特定用戶
        {
            try
            {
                await Task.Run(async () =>
                {
                    try
                    {
                        using (var httpClient = new HttpClient())
                        {
                            string sendMessage = Make_Address(endpoint);

                            HttpResponseMessage response = new HttpResponseMessage();
                            string body;
                            string json;
                            if (key == "Name")
                            {
                                List<Fusion_NameData> obj = new List<Fusion_NameData>();
                                obj.Add(new Fusion_NameData() { Name = value });
                                body = JsonConvert.SerializeObject(obj);
                                var data = new { Name = value };
                                json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                            }
                            else if (key == "Value")
                            {
                                var data = new { Value = value };
                                json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                            }
                            else
                            {
                                var data = new { State = value };
                                json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                            }

                            var content = new StringContent(json, Encoding.UTF8, "application/json");
                            response = await httpClient.PutAsync(sendMessage, content);
                            if (response.IsSuccessStatusCode)
                            {
                                State = "Ok";
                            }
                            else
                            {
                                State = "Error";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex.InnerException;
                    }
                });
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }





        public async Task Get_State()
        {
            try
            {
                //Returns the current run state of the protocol.

                //Always returns one of the following strings:
                //          *Idle:      The protocol is not running.
                //          *Waiting:   User requested protocol run(transitional state).
                //          *Running:   Protocol is running.
                //          *Paused:    Protocol was running and is now paused.
                //          *Aborting:  User has requested protocol stop(transitional state).
                //          * Aborted:  The protocol has stopped(transitional state, will become Idle).

                await Get_Value("/v1/protocol/state", "State");

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }


        public async Task Get_Device_List()
        {
            try
            {
                await Get_Value("/v1/devices/", "device");
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public async Task Get_Device_Feature(string device)
        {
            try
            {
                device = m_Devicename;
                await Get_Value("/v1/devices/" + device, "device");
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        //
        private async Task Get_Device_Feature_Value(string device, string feature)
        {
            try
            {
                await Get_Value("/v1/devices/" + device + "/" + feature, feature);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public async Task Sample()
        {
            try
            {
                await GetX();
                await GetY();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public async Task MoveAbs(string x_mm, string y_mm)
        {
            try
            {
                m_TargetX = x_mm;
                m_TargetY = y_mm;
                await SetX(x_mm);
                await SetY(y_mm);
                await Sample();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public async Task MoveRel(double x_mm, double y_mm)
        {
            try
            {
                await Sample();
                string newX = (Convert.ToDouble(m_NowX) + Convert.ToDouble(x_mm)).ToString();
                string newY = (Convert.ToDouble(m_NowY) + Convert.ToDouble(y_mm)).ToString();
                m_TargetX = newX;
                m_TargetY = newY;
                await SetX(newX);
                await SetY(newY);
                await Sample();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public double Tolerance = 0.05;
        public async Task ChkRunOver()
        {
            try
            {
                await Wait_until_state("Idle", 1);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public async Task ChkRunStop()
        {
            try
            {
                await Task.Run(async () =>
                {
                    try
                    {
                        while (State == "Running")
                        {
                            await Task.Delay(Convert.ToInt32(0.1 * 1000));
                            await Get_State();
                            await Sample();
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                });
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public async Task ChkMoveStop()
        {
            try
            {
                await Sample();
                double pnowX = (Convert.ToDouble(m_NowX));
                double pnowY = (Convert.ToDouble(m_NowY));
                double ptargetX = (Convert.ToDouble(m_TargetX));
                double ptargetY = (Convert.ToDouble(m_TargetY));
                while (pnowX > ptargetX + Tolerance ||
                    pnowX < ptargetX - Tolerance ||
                    pnowY > ptargetY + Tolerance ||
                    pnowY < ptargetY - Tolerance)
                {
                    await Sample();
                    pnowX = (Convert.ToDouble(m_NowX));
                    pnowY = (Convert.ToDouble(m_NowY));
                    //ptargetX = Math.Abs(Convert.ToDouble(m_TargetX));
                    //ptargetY = Math.Abs(Convert.ToDouble(m_TargetY));
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }



        private async Task GetX()
        {
            try
            {
                await Get_Device_Feature_Value(m_Devicename, m_Xfeaturename);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        private async Task GetY()
        {
            try
            {
                await Get_Device_Feature_Value(m_Devicename, m_Yfeaturename);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        private async Task SetX(string value)
        {
            try
            {
                await Set_Device_Feature_Value(m_Devicename, m_Xfeaturename, value);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        private async Task SetY(string value)
        {
            try
            {
                await Set_Device_Feature_Value(m_Devicename, m_Yfeaturename, value);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        public async Task Set_Device_Feature_Value(string device, string feature, string value)
        {
            try
            {
                await Put_Value("/v1/devices/" + device + "/" + feature, "Value", value);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }






        private async Task Set_State(string value)
        {
            try
            {
                await Put_Value("/v1/protocol/state", "State", value);
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public async Task Get_Selected_Protocol()
        {
            try
            {
                await Get_Value("/v1/protocol/current", "Name");
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public async Task Set_Selected_Protocol(string value)
        {
            try
            {
                await Put_Value("/v1/protocol/current", "Name", value);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public async Task Get_Protocol_Progress()
        {
            try
            {
                await Get_Value("/v1/protocol/progress", "Progress");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public async Task Run()
        {
            try
            {
                await Set_State("Running");
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public async Task Pause()
        {
            try
            {
                await Set_State("Paused");
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public async Task Stop()
        {
            try
            {
                await Set_State("Aborted");
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public async Task Wait_until_state(String target_state, double check_interval_secs)
        {
            try
            {
                await Task.Run(async () =>
                {
                    try
                    {
                        while (State != target_state)
                        {
                            await Task.Delay(Convert.ToInt32(check_interval_secs * 1000));
                            await Get_State();
                            await Sample();
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                });
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        public async Task completion_percentage()
        {
            try
            {
                //Returns the current protocol completion percentage, as a number ranging from 0 to 100.
                //If called after the protocol has stopped, this function will return whatever the final completion percentage was.
                //This may be less than 100 if the protocol was manually stopped early.
                await Get_Protocol_Progress();
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }


        public async Task Run_protocol_completely(string protocol_name)
        {
            //Tells Fusion to run the named protocol, and waits for it to complete.
            //This call will block until the protocol has finished.
            try
            {
                await Set_Selected_Protocol(protocol_name);
                if (ErrorMessage != "Error")
                {
                    await Set_State("Running");
                }
                if (ErrorMessage != "Error")
                {
                    await Task.Run(async () =>
                    {
                        try
                        {
                            while (State != "Running")
                            {
                                await Task.Delay(Convert.ToInt32(0.1 * 1000));
                                await Get_State();
                            }
                        }
                        catch (Exception ex)
                        {

                            throw ex;
                        }
                    });
                    //await Wait_until_state("Idle", 1);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        #endregion








    }
}
