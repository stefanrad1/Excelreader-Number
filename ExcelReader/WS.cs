using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    class WS:Student
    {
        protected string WSid;
        protected string WSorigin;
        protected string Tarival;
        protected string Tleave;
        protected string Treturn;
        protected string Tleave2;
        protected Student korisnik;

        public WS() { }
        
        public string WSid_p
        {
            get { return WSid; }
            set { WSid = value; }
        }

        public string WSorigin_p
        {
            get { return WSorigin; }
            set { WSorigin = value; }
        }

        public string Tarival_p
        {
            get { return Tarival; }
            set { Tarival = value; }
        }

        public string Tleave_p
        {
            get { return Tleave; }
            set { Tleave = value; }
        }

        public string Treturn_p
        {
            get { return Treturn; }
            set { Treturn = value; }
        }

        public string Tleave2_2
        {
            get { return Tleave2; }
            set { Tleave2 = value; }
        }

        public string studentID
        {
            get { return korisnik.id_p; }
        }

        public string midGrade
        {
            get { return korisnik.minGrade_p; }
        }

        public string finalGrade
        {
            get { return korisnik.FinGrade_p; }
        }

        public string Tstart
        {
            get { return korisnik.Tstart_p; }
        }

        public string Tmid
        {
            get { return korisnik.Tmid_p; }
        }

        public string Tfinal
        {
            get { return korisnik.Tfinal_p; }
        }

        public string TlastEdit
        {
            get { return korisnik.TlastEdit_p; }
        }

        public Student korisnik_p
        {
            get { return korisnik; }
            set { korisnik = value; }
        }

    }
}
