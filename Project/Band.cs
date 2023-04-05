using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project
{
    internal class Band : Product
    {
        private string color;

        public Band(string name, string price, string color, string quantity, string size, string length) : base(price,name,quantity)
        {
            this.color = color;
        }

        public string getColor()
        {
            return color;
        }
    }
}
