using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project
{
    internal class Shoe : Product
    {
        private string size;

        public Shoe(string name, string price, string size, string quantity) : base(price ,name,quantity) 
        {
            this.size = size;
        }

        public string getSize()
        {
            return size;
        }
    }
}
