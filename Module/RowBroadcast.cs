using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBClasses.Module
{
    class RowBroadcast: RowAutoID
    {
        private static int _counter;
        public int IdBroadcast { get; }
        public DateTime Beginning { get; }
        public DateTime Ending { get; }
        public int IdChannel { get; }
        public int IdShow { get; }
        //public RowChannel MyRowChannel { get; }
        //public RowShow MyRowShow { get; }

        public RowBroadcast()
        {
            IdBroadcast = _counter++;
        }

        public RowBroadcast(DateTime beginning, DateTime ending, int idChannel, int idShow)
        {
            IdBroadcast = _counter++;
            Beginning = beginning;
            Ending = ending;
            IdChannel = idChannel;
            IdShow = idShow;

            if (ending <= beginning)
                throw new Exception("Ending mast be >= Beginning");
        }

        private bool Equals(RowBroadcast other)
        {
            return Beginning.Equals(other.Beginning) && Ending.Equals(other.Ending) && IdChannel == other.IdChannel && IdShow == other.IdShow;
        }

        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((RowBroadcast)obj);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Beginning, Ending, IdChannel, IdShow);
        }


    }
}
