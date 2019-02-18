using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    class Student
    {
        protected string id;
        protected string minGrade;
        protected string finGrade;
        protected string set;
        protected string ws;
        protected string Tstart;
        protected string Tmid;
        protected string Tfinal;
        protected string TlastEdit;

        public Student() { }

        public string id_p {
            get { return id; }
            set { id = value; }
        }

        public string minGrade_p
        {
            get { return minGrade; }
            set { minGrade = value; }
        }

        public string FinGrade_p
        {
            get { return finGrade; }
            set { finGrade = value; }
        }

        public string set_p
        {
            get { return set; }
            set { set = value; }
        }

        public string ws_p
        {
            get { return ws; }
            set { ws = value; }
        }

        public string Tstart_p
        {
            get { return Tstart; }
            set { Tstart = value; }
        }

        public string Tmid_p
        {
            get { return Tmid; }
            set { Tmid = value; }
        }

        public string Tfinal_p
        {
            get { return Tfinal; }
            set { Tfinal = value; }
        }

        public string TlastEdit_p
        {
            get { return TlastEdit; }
            set { TlastEdit = value; }
        }
    }
}
