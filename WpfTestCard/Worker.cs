using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfTestCard
{
    public class Worker
    {
        public Worker(string name,string room,string unit,string id)
        {
            this.name = name;
            this.room = room;
            this.unit = unit;
            this.id = id;
        }
        private string name;
        private string room;
        private string unit;
        private string id;

        public string Name
        {
            get
            {
                return name;
            }

            //set
            //{
            //    name = value;
            //}
        }

        public string Room
        {
            get
            {
                return room;
            }

            //set
            //{
            //    room = value;
            //}
        }

        public string Unit
        {
            get
            {
                return unit;
            }

            //set
            //{
            //    unit = value;
            //}
        }

        public string ID 
        {
            get
            {
                return id;
            }

            //set
            //{
            //    id = value;
            //}
        }
    }
}
