using Excel2Json.Abstract;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace Excel2Json.Concrate
{
    public class Excel12 : IExcel
    {
        public OleDbConnection Connection(string filepath)
        {

            return new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filepath + "';Extended Properties='Excel 12.0 Xml; HDR = YES';");
        }
    }
}
