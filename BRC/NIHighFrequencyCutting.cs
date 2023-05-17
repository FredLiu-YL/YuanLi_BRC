using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NationalInstruments.DAQmx;


namespace BRC
{
    public class NIHighFrequencyCutting
    {

        private string[] channels;

        private double minVoltage, maxVoltage;

        public NIHighFrequencyCutting(double minVoltage = 0, double maxVoltage = 5)
        {

            channels = DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.AO, PhysicalChannelAccess.External);
            this.minVoltage = minVoltage;
            this.maxVoltage = maxVoltage;
        }

        public void On()
        {
            Run(5, minVoltage, maxVoltage);

        }
        public void Off()
        {
            Run(0, minVoltage, maxVoltage);

        }
        public void Run(double voltage, double minValue, double maxValue)
        {
            try {
                using (NationalInstruments.DAQmx.Task myTask = new NationalInstruments.DAQmx.Task()) {
                    myTask.AOChannels.CreateVoltageChannel(channels[0], "aoChannel", minValue, maxValue, AOVoltageUnits.Volts);
                    AnalogSingleChannelWriter writer = new AnalogSingleChannelWriter(myTask.Stream);
                    writer.WriteSingleSample(true, voltage);
                }
            }
            catch (DaqException ex) {
                throw ex;
            }
        }
    }
}
