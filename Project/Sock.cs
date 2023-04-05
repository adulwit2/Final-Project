using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project
{
    internal class Sock : Product
    {
        private string length;

        public Sock(string name, string price, string length, string quantity, string size) : base(price,name,quantity)
        {
            this.length = length;
        }

        public string getLength()
        {
            return length;
        }
    }
}
