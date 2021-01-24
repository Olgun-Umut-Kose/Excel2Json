using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace Excel2Json.Abstract
{
    public interface IExcel
    {
        OleDbConnection Connection(string filepath);
    }
}
