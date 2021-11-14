using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace Excel2Json.Abstract
{
    internal interface IExcel2JsonDal
    {
        OleDbDataReader Read(string sheet);
        string ConvertJson<T>(OleDbDataReader readexcel ) where T : new();
        
    }
}
