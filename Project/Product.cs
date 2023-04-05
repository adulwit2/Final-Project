using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project
{
    internal class Product
    {
        private string price;
        private string name;
        private string quantity;

        public Product(string price, string name, string quantity)
        {
            this.price = price;
            this.name = name;
            this.quantity = quantity;
        }
        public string getPrice()
        {
            return price;
        }
        public string getName()
        {
            return name;
        }
        public string getQuantity()
        {
            return quantity;
        }
    }
}
