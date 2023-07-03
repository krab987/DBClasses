using DBClasses.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBClasses.Module
{
    public class RowShow: RowAutoID
    {
        private static int counter;
        public int IdShow { get; set; }
        public string Name { get; set; }
        public TypeShow TypeShow { get; set; }
        public uint Duration { get; set; }
        public CategoryShow ShowCategory { get; set; }

        public RowShow(string name, TypeShow typeShow, uint duration, CategoryShow categoryShow)
        {
            IdShow = ++counter;
            Name = name;
            TypeShow = typeShow;
            Duration = duration;
            ShowCategory = categoryShow;
        }
        public RowShow()
        {
            IdShow = ++counter;
        }

        public override bool Equals(object? obj)
        {
            return obj is RowShow show &&
                   Name == show.Name;
        }
        public override int GetHashCode()
        {
            return HashCode.Combine(Name);
        }
    }
}
