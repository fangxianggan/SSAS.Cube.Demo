using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApplication1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();


        }
        public DataTable ToDataTable(CellSet cs)
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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

            string connectionString = "Data Source=localhost;Catalog=MDX Step-by-Step;ConnectTo=11.0;Integrated Security=SSPI";
            AdomdConnection _connection = new AdomdConnection(connectionString);
            if (_connection != null)
                if (_connection.State == ConnectionState.Closed)
                    _connection.Open();
            AdomdCommand command = _connection.CreateCommand();
            StringBuilder sb = new StringBuilder();
            sb.Append("WITH");
            sb.Append("  MEMBER [Product].[Category].[All Products].[X] AS 1+1");
            sb.Append("SELECT{ ([Date].[Calendar].[CY 2002]),([Date].[Calendar].[CY 2003])}*{([Measures].[Reseller Sales Amount]) } ON COLUMNS,");
            sb.Append("{ ([Product].[Category].[Accessories]),([Product].[Category].[Bikes]),([Product].[Category].[Clothing]),");
            sb.Append("([Product].[Category].[Components]),([Product].[Category].[X])} ON ROWS");
            sb.Append("  FROM [Step-by-Step]");
            command.CommandText = sb.ToString();


            var xmlreader = command.ExecuteXmlReader();
            CellSet cellSet = CellSet.LoadXml(xmlreader);

            _connection.Close();
            var dt = ToDataTable(cellSet);
           var v = dt.Rows.Count;
        }
    }
}
