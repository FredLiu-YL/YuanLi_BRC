using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Zaber.Motion;
using Zaber.Motion.Ascii;

namespace BRC
{
    public class ZaberMotion
    {
        private Connection connection;
        private Axis axis;
        public ZaberMotion(string comPort)
        {
               Zaber.Motion.Library.SetDeviceDbSource(DeviceDbSourceType.File, "C:\\ZaberDataBase\\devices-public.sqlite");
            //  Zaber.Motion.Library.SetDeviceDbSource(DeviceDbSourceType.File, "path_to_the_folder/devices-public.sqlite");
            connection = Connection.OpenSerialPort(comPort);
            connection.EnableAlerts();
            Devices = connection.DetectDevices();

            axis = Devices[0].GetAxis(1);  //第一個裝置的第一支軸

        }
        public ZaberMotion()
        {
            Library.EnableDeviceDbStore();
            var comm = Connection.OpenIot("f4fe6a3d-2f64-4195-8e11-8b81a11b782a");
            Devices = comm.DetectDevices();

            axis = Devices[0].GetAxis(1);  //第一個裝置的第一支軸

        }

        public Device[] Devices { get; }

        public double Velocity { get => GetSpeed(); set { SetSpeed(value); } }
        public double Acceleration { get; set; } = 620.722;

        public void Home()
        {
            if (!axis.IsHomed()) {
                axis.Home();
            }

        }


        public Task CVRun(double firstSpeed, double firstDelayTime, double secondSpeed, double secondDelayTime)
        {
            return Task.Run(() =>
          {
              Velocity = firstSpeed;
              int firstdelay = (int)firstDelayTime * 1000;
              Task.Delay(firstdelay).Wait();
              Velocity = secondSpeed;
              int seconddelay = (int)secondDelayTime * 1000;
              Task.Delay(seconddelay).Wait();

              Velocity = firstSpeed;//改回初始速度
              Stop();
            });
        }
        public void MoveMax()
        {
            axis.MoveMaxAsync();
        }
        public void Move(double distance)
        {
            //        if (!axis.IsHomed()) throw new Exception("Home not completed");
            axis.MoveRelative(distance, Units.Length_Millimetres);
        }

        public void MoveTo(double pos)
        {
            //        if (!axis.IsHomed()) throw new Exception("Home not completed");
            // await axis.MoveAbsoluteAsync(pos, Units.Length_Millimetres, true, Velocity , Units.Velocity_MillimetresPerSecond);
            axis.MoveAbsoluteAsync(pos, Units.Length_Millimetres);
        }
        public void SetSpeed(double velocity)
        {
            axis.Settings.Set("maxspeed", velocity, Units.AngularVelocity_DegreesPerSecond);
        }
        public double GetSpeed()
        {
            double vec = axis.Settings.Get("maxspeed", Units.AngularVelocity_DegreesPerSecond); 
            return vec;
        }
        public void Stop()
        {
            axis.Stop();
        }


        public void test()
        {

            using (var connection = Connection.OpenSerialPort("COM4")) {
                connection.EnableAlerts();

                var deviceList = connection.DetectDevices();
                Console.WriteLine($"Found {deviceList.Length} devices.");

                var device = deviceList[0];

                var axis = device.GetAxis(1);
                if (!axis.IsHomed()) {
                    axis.Home();
                }

                // Move to 10mm
                axis.MoveAbsolute(10, Units.Length_Millimetres);

                // Move by an additional 5mm
                axis.MoveRelative(5, Units.Length_Millimetres);
            }
        }



    }
}
