using Excel2Json.Abstract;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace Excel2Json.Concrate
{
    public class Excel2JsonDal : IExcel2JsonDal
    {
        private string _filePath;

        private OleDbConnection _connection;
        private OleDbCommand cmd;
        private IExcel _excel;


        public Excel2JsonDal(string filepath, IExcel excel)
        {
            _filePath = filepath;
            _excel = excel;

        }

        private void StartConnectionAndReader()
        {
            _connection = _excel.Connection(_filePath);
            _connection.Open();
            cmd = _connection.CreateCommand();
            cmd.Connection = _connection;
        }

        public OleDbDataReader Read(string sheet)
        {
            StartConnectionAndReader();

            cmd.CommandText = "SELECT * FROM [" + sheet.ToUpper() + "$]";

            return cmd.ExecuteReader();

        }

        public string ConvertJson(OleDbDataReader readexcel, IQuery queryobjets)
        {

            IEnumerable<IQuery> query = readexcel.Cast<DbDataRecord>().Select(row => queryobjets.InitObjects(row));

            return JsonConvert.SerializeObject(query);
        }
   
        }


}






    

