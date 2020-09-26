using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExeclConvertSQLScript
{
    class Program
    {

        static void Main(string[] args)
        {

            string input = @"D:\uscities.xlsx";
            string output = @"D:\uscities.sql";
            string tablename = "__table__";
            SqlExecl sql = new SqlExecl(5, tablename);
            sql.Read(input);
            sql.WriteToSql(output);
            Console.Write("completed");

        }

    }
}
