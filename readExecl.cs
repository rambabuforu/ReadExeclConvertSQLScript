using LinqToExcel;
using System.Linq;
using System;
using System.Text;
using System.Collections.Generic;

namespace ReadExeclConvertSQLScript
{
    public class SqlExecl
    {
        int coloum;
        string TableColoumn;
        List<string> _SqlList;

        System.Collections.Generic.List<Cell> _header;
        System.Collections.Generic.List<Cell> _row;


        string tablename;
        public SqlExecl(int Coloum, string TableName)
        {
            coloum = Coloum;
            tablename = TableName;
            _SqlList = new List<string>();
        }

        public void Read(string directorio)
        {
            var book = new ExcelQueryFactory(directorio);
            var rows = from c in book.WorksheetNoHeader()
                       select c;
            try { 
            int i = 1;
                foreach (var row in rows)
                {
                    if (i == 1)
                    {
                        this._header = row;
                        this.TableColoumn = this.PrepareTableColoumn(row);
                    }

                    if (i > 1)
                    {
                        this._row = row;

                        List<string> _whereCasueList = new List<string>();
                        int j = 0;
                        foreach (var coloumn in this._header)
                        {
                            string col = "(" + coloumn.Value.ToString() + " = '" + this._row[j].ToString() + "')";
                            _whereCasueList.Add(col);
                            j++;
                        }

                        string _whereCasue = string.Join(" and ", _whereCasueList.ToArray());

                        string _sqlcondition = string.Format("if not exists(select 1 from {0} where {1}) then", tablename, _whereCasue);
                        _SqlList.Add(_sqlcondition);

                        _SqlList.Add("begin");
                        string _values = this.PrepareTableValues(row);
                        string sql = string.Format("\t\t insert into __table__ ({0}) values ({1})", this.TableColoumn, _values);

                        _SqlList.Add(sql);
                        _SqlList.Add("end");
                        _SqlList.Add(null);
                        _SqlList.Add(null);
                    }
                    i++;
                }
            }
            catch(Exception ex)
            {
                ex.ToString();

            }
        }

        public void WriteToSql(string path)
        {
            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(path))
            {
                foreach (string line in _SqlList)
                {
                    file.WriteLine(line);
                }
            }
        }

        public string PrepareTableColoumn(System.Collections.Generic.List<Cell> LinqToExcel)
        {
            var HeaderString = LinqToExcel.ConvertAll<string>(new Converter<Cell, string>(CelltoString));
            return string.Join(",", HeaderString.ToArray());
        }

        public string PrepareTableValues(System.Collections.Generic.List<Cell> LinqToExcel)
        {
            var HeaderString = LinqToExcel.ConvertAll<string>(new Converter<Cell, string>(CelltoString));
            return "'" + string.Join("', '", HeaderString.ToArray()) + "'";
        }

        public static string CelltoString(Cell pf)
        {
            return pf.ToString();
        }

        public static string ClasstoString(StringClass obj)
        {
            return obj.ToString();
        }
    }

    public class StringClass
    {
        public string value { get; set; }
    }

}
