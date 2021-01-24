using Excel2Json.Abstract;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace Excel2Json.Concrate
{
    public class Excel8 : IExcel
    {
        public OleDbConnection Connection(string filepath)
        {
            return new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filepath + "';Extended Properties='Excel 8.0; HDR = YES';");
        }
    }
}
