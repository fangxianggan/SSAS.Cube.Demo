using adomddemo;
using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADOMD.netTest
{
    class Program
    {

        static void Main(string[] args)
        {

            string connectionString = "Data Source=http://192.192.192.210/olap/msmdpump.dll;Initial Catalog=Cube";
            AdomdConnection conn = new AdomdConnection();

            AdomdHelper helper = new AdomdHelper();
            DataTable dt = helper.GetSchemaDataSet_Catalogs(ref conn, connectionString);

        

            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine($"name:{row["CATALOG_NAME"]}---description:{row["DESCRIPTION"]}---roles:{row["ROLES"]}---time:{row["DATE_MODIFIED"]}");              
            }



            string[] cubes = helper.GetSchemaDataSet_Cubes(ref conn, connectionString);

            foreach (var item in cubes)
            {
                Console.WriteLine($"cube数据库：{item}");

                string[] dimList = helper.GetSchemaDataSet_Dimensions(ref conn, connectionString, item);
                foreach (var item1 in dimList)
                {
                    Console.WriteLine($"{item}数据库的维度：{item1}");
                }
            }


        
            Console.ReadKey();
            Console.ReadLine();



        }
    }
}
