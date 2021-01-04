using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADOMD.netTest
{
    public class TestDemo
    {
        public static readonly string dbkey = "Data Source=DESKTOP-EON4Q93;Initial Catalog=Analysis Services Tutorial";

        public static readonly AdomdConnection _connection;

        static TestDemo()
        {
            _connection = new AdomdConnection(dbkey);

        }

        public static AdomdConnection GetAdomdConnection()
        {
            return _connection;
        }



        /// <summary>
        /// 传化DataTable
        /// </summary>
        /// <param name="cs"></param>
        /// <returns></returns>
        private static DataTable ToDataTable(CellSet cs)
        {
            bool haveColumn = cs.Axes.Count > 1 && cs.Axes[1].Positions.Count > 0;
            int membersCount = 0;
            DataTable dt = new DataTable();

            //生成数据列对象
            if (haveColumn)
            {
                membersCount = cs.Axes[1].Positions[0].Members.Count;
                for (int i = 0; i < membersCount; i++)
                {
                    DataColumn dataColumn = new DataColumn();
                    dataColumn.Caption = cs.Axes[1].Positions[0].Members[i].LevelName.Split(']').Any()
                                             ? cs.Axes[1].Positions[0].Members[i].LevelName.Split(']')[0]
                                             : dataColumn.ColumnName;
                    dt.Columns.Add(dataColumn);
                    dataColumn = new DataColumn("Key" + dataColumn.ColumnName);
                    dataColumn.Caption = cs.Axes[1].Positions[0].Members[i].LevelName.Split(']').Any()
                                             ? cs.Axes[1].Positions[0].Members[i].LevelName.Split(']')[0] + "代码"
                                             : dataColumn.ColumnName;
                    dt.Columns.Add(dataColumn);
                }
            }
            foreach (Position p in cs.Axes[0].Positions)
            {
                DataColumn dc = new DataColumn();
                string name = p.Members.Cast<Member>().Aggregate("", (current, m) => current + m.Caption);

                dc.ColumnName = name;
                dt.Columns.Add(dc);
            }

            //添加行数据
            int pos = 0;

            if (cs.Axes.Count > 1)
            {

                foreach (Position py in cs.Axes[1].Positions)
                {
                    DataRow dr = dt.NewRow();
                    int j = -1;//用来进行Column的计数
                    for (int i = 0; i < membersCount; i++)
                    {
                        dr[++j] = py.Members[i].Caption;
                        dr[++j] = py.Members[i].UniqueName;
                    }

                    //数据列
                    for (int x = membersCount * 2; x < cs.Axes[0].Positions.Count + membersCount * 2; x++)
                    {
                        dr[x] = cs[pos++].FormattedValue.Replace(",", "");//解决数值中存在','的问题
                    }
                    dt.Rows.Add(dr);
                }
            }
            else
            {
                DataRow dr = dt.NewRow();

                //数据列
                for (int x = 0; x < cs.Axes[0].Positions.Count; x++)
                {
                    dr[x] = cs[pos++].FormattedValue.Replace(",", "");//解决数值中存在','的问题;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }


        private static DataTable ToDataTable1(CellSet cs)
        {
            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn();
            DataRow dr = null;
            //第一列：必有为维度描述（行头）
            dt.Columns.Add(new DataColumn("Description"));
            //生成数据列对象
            string name;
            foreach (Position p in cs.Axes[0].Positions)
            {
                dc = new DataColumn();
                name = "";
                foreach (Member m in p.Members)
                {
                    name = name + m.Caption + " ";
                }
                dc.ColumnName = name;
                dt.Columns.Add(dc);
            }
            //添加行数据
            int pos = 0;
            foreach (Position py in cs.Axes[1].Positions)
            {
                dr = dt.NewRow();
                //维度描述列数据（行头）
                name = "";
                foreach (Member m in py.Members)
                {
                    name = name + m.Caption + "\r\n";
                }
                dr[0] = name;
                //数据列
                for (int x = 1; x <= cs.Axes[0].Positions.Count; x++)
                {

                    dr[x] = cs[pos++].FormattedValue;

                }
                dt.Rows.Add(dr);

            }
            return dt;

        }
        public static DataTable GetDataByCon(string strMdx)
        {
            DataTable dt = new DataTable();
            try
            {
               
                    if (_connection != null)
                    {
                        if (_connection.State == ConnectionState.Closed)
                            _connection.Open();

                        Stopwatch sw = new Stopwatch();
                        //开始计时  
                        sw.Start();
                        AdomdCommand cmd = _connection.CreateCommand();
                        cmd.CommandText = strMdx;

                        var executexml = cmd.ExecuteXmlReader();
                        CellSet cellset = CellSet.LoadXml(executexml);

                        //  CellSet cellset = cmd.ExecuteCellSet();
                        sw.Stop();
                        //获取运行时间[毫秒]  
                        long times = sw.ElapsedMilliseconds;
                        Console.WriteLine($"耗时0——{times}");
                        _connection.Close();
                        dt = ToDataTable(cellset);
                    }
              
                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static DataTable GetDataByCon1(string strMdx)
        {
            DataTable dt = new DataTable();
            try
            {
              
                    if (_connection != null)
                    {
                        if (_connection.State == ConnectionState.Closed)
                            _connection.Open();

                        Stopwatch sw = new Stopwatch();
                        //开始计时  
                        sw.Start();
                        AdomdCommand cmd = _connection.CreateCommand();
                        cmd.CommandText = strMdx;

                        //var executexml = cmd.ExecuteXmlReader();
                        //CellSet cellset = CellSet.LoadXml(executexml);

                        CellSet cellset = cmd.ExecuteCellSet();
                        sw.Stop();
                        //获取运行时间[毫秒]  
                        long times = sw.ElapsedMilliseconds;
                        Console.WriteLine($"耗时1——{times}");
                        _connection.Close();
                        dt = ToDataTable(cellset);
                    }
             

                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static List<AATest> GetDataByCon2(string strMdx)
        {
            List<AATest> list=new List<AATest>();  
            try
            {
              
                    if (_connection != null)
                    {
                        if (_connection.State == ConnectionState.Closed)
                            _connection.Open();

                        Stopwatch sw = new Stopwatch();
                        //开始计时  
                        sw.Start();
                        AdomdCommand cmd = _connection.CreateCommand();
                        cmd.CommandText = strMdx;
                        AdomdDataReader reader = cmd.ExecuteReader(); //Execute query
                   
                        while (reader.Read())   // read
                        {
                            AATest dt = new AATest();  // custom class
                            dt.A = reader[0] == null ? "" : reader[0].ToString();
                            dt.B = reader[1] == null ? "" : reader[1].ToString();
                            dt.C = reader[2] == null ? "" : reader[2].ToString();
                            dt.D = reader[3]==null?"":reader[3].ToString();
                            dt.E = reader[4] == null ? "":reader[4].ToString();
                            dt.F = reader[5] == null ? "" : reader[5].ToString();
                            dt.G = reader[6] == null ? "" : reader[6].ToString();
                            //dt.H = reader[7].ToString();
                            //dt.I = reader[8].ToString();
                            //dt.J = reader[9].ToString();
                            //dt.K = reader[10].ToString();
                            list.Add(dt);
                        }



                        sw.Stop();
                        //获取运行时间[毫秒]  
                        long times = sw.ElapsedMilliseconds;
                        Console.WriteLine($"耗时2——{times}");
                        _connection.Close();

                    }
              
                return list;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        public class AATest
        { 
            public string DimName { set; get; }

            public string A { set; get; }

            public string B{ set; get; }

            public string C { set; get; }

            public string D { set; get; }

            public string E { set; get; }
            public string F { set; get; }

            public string G { set; get; }

            public string H { set; get; }
            public string I { set; get; }

            public string J { set; get; }
            public string K { set; get; }
            public string L { set; get; }
        }

    }
}
