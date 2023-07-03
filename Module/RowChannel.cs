using DBClasses.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBClasses.Module
{
    class RowChannel : RowAutoID
    {
        private static int _counter;
        private double _frequency;
        public int IdChannel { get; set; }
        public string Name { get; set; }
        public TypeChannel ChannelType { get; set; }

        public double Frequency
        {
            get => _frequency;
            set
            {
                if (_frequency > 1000.0 || _frequency < 9999.9)
                    _frequency = value;
            }
        }

        public RowChannel(string name, TypeChannel channelType, double frequency)
        {
            IdChannel = _counter++;
            Name = name;
            ChannelType = channelType;
            Frequency = frequency;

            if (Frequency < 1000.0 || Frequency > 9999.9)
                throw new ArgumentException("Frequency must be < 1000 and > 9999");

        }
        public RowChannel()
        {
            IdChannel = _counter++;
        }

        public override string ToString()
        {
            return $"Name {Name}" +
                   $"ChannelType {ChannelType}";
        }
    }
}
