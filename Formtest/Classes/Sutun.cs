using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Text;
using Excel2Json.Abstract;

namespace Formtest.Classes
{
    public class Sutun : IQuery
    {

        public object ad { get; set; }
        public object soyad { get; set; }
        public object gorev { get; set; }
        public object brans { get; set; }
        public object telno { get; set; }
        public object sifre { get; set; }
        public object kadi { get; set; }

        public IQuery GetObjects()
        {
            Sutun sutun = new Sutun();
            sutun = this;
            return sutun;
        }

        public IQuery InitObjects(DbDataRecord row)
        {
            SetObjects(row);
            return GetObjects();
        }

        public void SetObjects(DbDataRecord row)
        {
            ad = row[0];
            soyad = row[1];
            gorev = row[2];
            brans = row[3];
            telno = row[4];
            sifre = row[5];
            kadi = row[6];
        }

        
    }
}
