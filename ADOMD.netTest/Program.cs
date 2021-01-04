using adomddemo;
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
    class Program
    {

        static void Main(string[] args)
        {

            //string connectionString = "Data Source=http://192.192.192.210/olap/msmdpump.dll;Initial Catalog=Cube";
            //AdomdConnection conn = new AdomdConnection();

            //AdomdHelper helper = new AdomdHelper();
            //DataTable dt = helper.GetSchemaDataSet_Catalogs(ref conn, connectionString);



            //foreach (DataRow row in dt.Rows)
            //{
            //    Console.WriteLine($"name:{row["CATALOG_NAME"]}---description:{row["DESCRIPTION"]}---roles:{row["ROLES"]}---time:{row["DATE_MODIFIED"]}");              
            //}



            //string[] cubes = helper.GetSchemaDataSet_Cubes(ref conn, connectionString);

            //foreach (var item in cubes)
            //{
            //    Console.WriteLine($"cube数据库：{item}");

            //    string[] dimList = helper.GetSchemaDataSet_Dimensions(ref conn, connectionString, item);
            //    foreach (var item1 in dimList)
            //    {
            //        Console.WriteLine($"{item}数据库的维度：{item1}");
            //    }
            //}

            {
                Stopwatch sw = new Stopwatch();
                //开始计时  
                sw.Start();
                string cubeStr = "  SELECT NON EMPTY { [Measures].[Internet Sales Count], [Measures].[Internet Sales-Freight], [Measures].[Internet Sales-Unit Price] } ON COLUMNS, NON EMPTY { ([Customer].[City].[City].ALLMEMBERS * [Customer].[Phone].[Phone].ALLMEMBERS * [Customer].[Email Address].[Email Address].ALLMEMBERS * [Date].[Date].[Date].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Analysis Services Tutorial] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                var list = TestDemo.GetDataByCon2(cubeStr);

                // 结束计时
                sw.Stop();
                //获取运行时间[毫秒]  
                long times = sw.ElapsedMilliseconds;
                Console.WriteLine($"耗时2--{times}");
            }

            {
                Stopwatch sw = new Stopwatch();
                //开始计时  
                sw.Start();
                string cubeStr = "  SELECT NON EMPTY { [Measures].[Internet Sales Count], [Measures].[Internet Sales-Freight], [Measures].[Internet Sales-Unit Price] } ON COLUMNS, NON EMPTY { ([Customer].[City].[City].ALLMEMBERS * [Customer].[Phone].[Phone].ALLMEMBERS * [Customer].[Email Address].[Email Address].ALLMEMBERS * [Date].[Date].[Date].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Analysis Services Tutorial] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";


                var list = TestDemo.GetDataByCon1(cubeStr);

                // 结束计时
                sw.Stop();
                //获取运行时间[毫秒]  
                long times = sw.ElapsedMilliseconds;
                Console.WriteLine($"耗时1--{times}");
            }

         
            {
                Stopwatch sw = new Stopwatch();
                //开始计时  
                sw.Start();
                string cubeStr = "  SELECT NON EMPTY { [Measures].[Internet Sales Count], [Measures].[Internet Sales-Freight], [Measures].[Internet Sales-Unit Price] } ON COLUMNS, NON EMPTY { ([Customer].[City].[City].ALLMEMBERS * [Customer].[Phone].[Phone].ALLMEMBERS * [Customer].[Email Address].[Email Address].ALLMEMBERS * [Date].[Date].[Date].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Analysis Services Tutorial] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                StringBuilder builder = new StringBuilder();
                builder.Append("with");
                builder.Append(" MEMBER Measures.[本月销量] AS");
                builder.Append(" sum(");
                builder.Append(" ([Date].[Date].&[20120501]:[Date].[Date].&[20120531]),");
                builder.Append(" [Measures].[Internet Sales-Order Quantity]");
                builder.Append(" )");
                builder.Append(" MEMBER Measures.[上月销量] AS");
                builder.Append(" sum(");
                builder.Append(" ([Date].[Date].&[20120401]:[Date].[Date].&[20120531]),");
                builder.Append(" [Measures].[Internet Sales-Order Quantity]");
                builder.Append(" )");
                builder.Append(" MEMBER Measures.[去年本月销量] AS");
                builder.Append(" sum(");
                builder.Append(" ([Date].[Date].&[20110501]:[Date].[Date].&[20110531]),");
                builder.Append(" [Measures].[Internet Sales-Order Quantity]");
                builder.Append(" ) ");

                builder.Append(" member [Measures].[上年销量同比] as");
                builder.Append(" (([Measures].[本月销量] - Measures.[去年本月销量]) / Measures.[去年本月销量]) ");

                builder.Append(" member [Measures].[上月销量环比] as ");
                builder.Append(" (([Measures].[本月销量]-Measures.[上月销量])/Measures.[上月销量]) ");


                builder.Append(" select");
                builder.Append(" {Measures.[本月销量],Measures.[上月销量],Measures.[去年本月销量], [Measures].[上年销量同比], [Measures].[上月销量环比]}");
                builder.Append("on 0,");
                builder.Append(" non empty{");
                builder.Append("  topcount(");
                builder.Append(" order(");
                builder.Append(" FILTER(");
                builder.Append(" ( [Customer].[City].[City]),");
                builder.Append(" ([Measures].[本月销量] < 5 and[Measures].[本月销量] > 1))");
                builder.Append(" ,[Measures].[本月销量]");
                builder.Append(" , desc) * [Date].[Date].[Date] ");
                builder.Append(" ,100,[Measures].[本月销量]");
                builder.Append(" )");
                builder.Append(" }");
                builder.Append(" on 1");
                builder.Append(" from [Analysis Services Tutorial]");
                var list = TestDemo.GetDataByCon2(builder.ToString());

                // 结束计时
                sw.Stop();
                //获取运行时间[毫秒]  
                long times = sw.ElapsedMilliseconds;
                Console.WriteLine($"耗时2--{times}");
            }
            {
                Stopwatch sw = new Stopwatch();
                //开始计时  
                sw.Start();

                StringBuilder builder = new StringBuilder();
                builder.Append("with");
                builder.Append(" MEMBER Measures.[本月销量] AS");
                builder.Append(" sum(");
                builder.Append(" ([Date].[Date].&[20120501]:[Date].[Date].&[20120531]),");
                builder.Append(" [Measures].[Internet Sales-Order Quantity]");
                builder.Append(" )");
                builder.Append(" MEMBER Measures.[上月销量] AS");
                builder.Append(" sum(");
                builder.Append(" ([Date].[Date].&[20120401]:[Date].[Date].&[20120531]),");
                builder.Append(" [Measures].[Internet Sales-Order Quantity]");
                builder.Append(" )");
                builder.Append(" MEMBER Measures.[去年本月销量] AS");
                builder.Append(" sum(");
                builder.Append(" ([Date].[Date].&[20110501]:[Date].[Date].&[20110531]),");
                builder.Append(" [Measures].[Internet Sales-Order Quantity]");
                builder.Append(" )");

                builder.Append(" member [Measures].[上年销量同比] as");
                builder.Append(" (([Measures].[本月销量] - Measures.[去年本月销量]) / Measures.[去年本月销量]) ");

                builder.Append(" member [Measures].[上月销量环比] as ");
                builder.Append(" (([Measures].[本月销量]-Measures.[上月销量])/Measures.[上月销量]) ");

                builder.Append(" select");
                builder.Append(" {Measures.[本月销量],Measures.[上月销量],Measures.[去年本月销量], [Measures].[上年销量同比], [Measures].[上月销量环比]}");

                builder.Append("on 0,");
                builder.Append(" non empty{");
                builder.Append("  topcount(");
                builder.Append(" order(");
                builder.Append(" FILTER(");
                builder.Append(" ( [Customer].[City].[City]),");
                builder.Append(" ([Measures].[本月销量] < 5 and[Measures].[本月销量] > 1))");
                builder.Append(" ,[Measures].[本月销量]");
                builder.Append(" , desc) * [Date].[Date].[Date] ");
                builder.Append(" ,100,[Measures].[本月销量]");
                builder.Append(" )");
                builder.Append(" }");
                builder.Append(" on 1");
                builder.Append(" from [Analysis Services Tutorial]");
              
                var list = TestDemo.GetDataByCon(builder.ToString());

                // 结束计时
                sw.Stop();
                //获取运行时间[毫秒]  
                long times = sw.ElapsedMilliseconds;
                Console.WriteLine($"耗时0--{times}");
            }





            Console.ReadKey();
            Console.ReadLine();



        }
    }
}
