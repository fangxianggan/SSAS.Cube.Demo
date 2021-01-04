using System;
using System.Data;
using System.Linq;
using Microsoft.AnalysisServices.AdomdClient;
namespace adomddemo
{
	/// <summary>
	/// Summary description for AdomdHelper.
	/// </summary>
	public class AdomdHelper
	{
		#region "== Enum ============================================================"
		public enum Versions
		{
			Server,
			Provider,
			Client
		}
		#endregion

		#region "== Methods ============================================================"
		//�ж�����AdomdConnection������State�Ƿ���Open״̬��
		public bool IsConnected(ref AdomdConnection connection)
		{
			
			return (!(connection==null))&&(connection.State!=ConnectionState.Broken)&&(connection.State!=ConnectionState.Closed);
		}
	
		/// <summary>
		/// �Ͽ�����
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="destroyConnection">�Ƿ���������</param>
		public void Disconnect(ref AdomdConnection connection,bool destroyConnection)
		{
			try
			{
				if(!(connection == null))
				{
					if(connection.State!=ConnectionState.Closed)
					{
						connection.Close();
					}
					if(destroyConnection == true)
					{
						connection.Dispose();
						connection = null;
					}
				}
			}
			catch(Exception ex)
			{
				throw ex;
			}
		}
		/// <summary>
		/// ��������
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="connectionString">�����ַ���</param>
		public void Connect(ref AdomdConnection connection,string connectionString)
		{
			if(connectionString == "")
				throw new ArgumentNullException("connectionString", "The connection string is not valid.");
			//	Ensure an AdomdConnection object exists and that its ConnectionString property is set.
			if(connection == null)
				connection = new AdomdConnection(connectionString);
			else
			{
				Disconnect(ref connection,false);
				connection.ConnectionString = connectionString;
			}
			try
			{
				connection.Open();
			}
			catch(Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// ��ȡOLAP���ݿ⡣
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="connectionString">�����ַ���</param>
		/// <returns></returns>
		public DataTable GetSchemaDataSet_Catalogs(ref AdomdConnection connection,string connectionString)
		{
			
			bool connected  = true;    //�ж�connection�ڵ��ô˺���ʱ���Ƿ��Ѿ���������״̬
			DataTable objTable = new DataTable();
			try
			{
				// Check if a valid connection was provided.
				if(IsConnected(ref connection) == false)
				{
					//������Ӳ����ڣ���������
					Connect(ref connection,connectionString);
					connected  = false;       //����connectionΪδ����״̬��		

				}
				objTable = connection.GetSchemaDataSet(AdomdSchemaGuid.Catalogs,null).Tables[0];
				if(connected == false)
				{
					//�ر�����
					Disconnect(ref connection,false);
				}
			}
			catch(Exception err)
			{
				throw err;
			}

			return objTable;
		}
		
		/// <summary>
		/// ͨ��SchemaDataSet�ķ�ʽ��ȡ������
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="connectionString">�����ַ���</param>
		/// <returns></returns>
		public string[] GetSchemaDataSet_Cubes(ref AdomdConnection connection,string connectionString)
		{
			string[] strCubes = null;
			bool connected  = true;   //�ж�connection�Ƿ��������ݿ�����
			DataTable objTable = new DataTable();
			if (IsConnected(ref connection) == false)
			{
				try
				{
					Connect(ref connection,connectionString);
					connected = false;
				}
				catch(Exception err)
				{
					throw err;
				}		
			}			
			string[] strRestriction = new string[]{null,null,null};
			objTable = connection.GetSchemaDataSet(AdomdSchemaGuid.Cubes,strRestriction).Tables[0];	
			if(connected == false)
			{
				Disconnect(ref connection,false);
			}
			strCubes = new string[objTable.Rows.Count];
			int rowcount = 0;
			foreach(DataRow tempRow in objTable.Rows)
			{
				strCubes[rowcount] = tempRow["CUBE_NAME"].ToString();
				rowcount++;
			}
			return strCubes;
		}
		
		/// <summary>
		/// ͨ��SchemaDataSet�ķ�ʽ��ȡ�ƶ��������ά��
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="connectionString">�����ַ���</param>
		/// <returns></returns>
		public string[] GetSchemaDataSet_Dimensions(ref AdomdConnection connection,string connectionString,string cubeName)
		{
			string[] strDimensions = null;
			bool connected  = true;   //�ж�connection�Ƿ��������ݿ�����
			DataTable objTable = new DataTable();
			if (IsConnected(ref connection) == false)
			{
				try
				{
					Connect(ref connection,connectionString);
					connected = false;
				}
				catch(Exception err)
				{
					throw err;
				}		
			}			
			string[] strRestriction = new string[]{null,null,cubeName,null,null};
			objTable = connection.GetSchemaDataSet(AdomdSchemaGuid.Dimensions,strRestriction).Tables[0];
			if(connected == false)
			{
				Disconnect(ref connection,false);
			}
			strDimensions = new string[objTable.Rows.Count];
			int rowcount = 0;
			foreach(DataRow tempRow in objTable.Rows)
			{
				strDimensions[rowcount] = tempRow["DIMENSION_NAME"].ToString();
				rowcount++;
			}
			return strDimensions;
		}
		
		/// <summary>
		/// ��connection�ķ�ʽ��ȡ������
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="connectionString">�����ַ���</param>
		/// <returns></returns>
		public string[] GetCubes(ref AdomdConnection connection,string connectionString)
		{
			string[] strCubesName = null;
			bool connected  = true;   //�ж�connection�Ƿ��������ݿ�����
			if (IsConnected(ref connection) == false)
			{
				try
				{
					Connect(ref connection,connection.ConnectionString);
					connected = false;
				}
				catch(Exception err)
				{
					throw err;
				}		
			}	
			
			int rowcount = connection.Cubes.Count;
			strCubesName = new string[rowcount];
			for(int i=0;i<rowcount;i++)
			{
				strCubesName[i] = connection.Cubes[i].Caption;
			}
			
			if(connected == false)
			{
				Disconnect(ref connection,false);
			}
			return strCubesName;				
		}


		/// <summary>
		/// ��ȡ�������ά��
		/// </summary>
		/// <param name="connection">AdomdConnection�����ʵ��</param>
		/// <param name="connectionString">�����ַ���</param>
		/// <param name="CubeName">����������</param>
		/// <returns></returns>
		public string[] GetDimensions(ref AdomdConnection connection,string connectionString,string CubeName)
		{
			string[] strDimensions = null;
			bool connected = true;
			if (IsConnected(ref connection) == false)
			{
				try
				{
					Connect(ref connection,connection.ConnectionString);
					connected = false;
				}
				catch(Exception err)
				{
					throw err;
				}		
			}	
			
			int rowcount = connection.Cubes[CubeName].Dimensions.Count;
			strDimensions = new string[rowcount];
			for(int i=0;i<rowcount;i++)
			{
				strDimensions[i] = connection.Cubes[CubeName].Dimensions[i].Caption.ToString();
			}
			if(connected == false)
			{
				Disconnect(ref connection,false);
			}
			return strDimensions;

		}



        /// <summary>
        /// ����DataTable
        /// </summary>
        /// <param name="cs"></param>
        /// <returns></returns>
        private DataTable ToDataTable(CellSet cs)
        {
            bool haveColumn = cs.Axes.Count > 1 && cs.Axes[1].Positions.Count > 0;
            int membersCount = 0;
            DataTable dt = new DataTable();

            //���������ж���
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
                                             ? cs.Axes[1].Positions[0].Members[i].LevelName.Split(']')[0] + "����"
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

            //���������
            int pos = 0;

            if (cs.Axes.Count > 1)
            {

                foreach (Position py in cs.Axes[1].Positions)
                {
                    DataRow dr = dt.NewRow();
                    int j = -1;//��������Column�ļ���
                    for (int i = 0; i < membersCount; i++)
                    {
                        dr[++j] = py.Members[i].Caption;
                        dr[++j] = py.Members[i].UniqueName;
                    }

                    //������
                    for (int x = membersCount * 2; x < cs.Axes[0].Positions.Count + membersCount * 2; x++)
                    {
                        dr[x] = cs[pos++].FormattedValue.Replace(",", "");//�����ֵ�д���','������
                    }
                    dt.Rows.Add(dr);
                }
            }
            else
            {
                DataRow dr = dt.NewRow();

                //������
                for (int x = 0; x < cs.Axes[0].Positions.Count; x++)
                {
                    dr[x] = cs[pos++].FormattedValue.Replace(",", "");//�����ֵ�д���','������;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// ����mdx��䡢�����ַ���
        /// </summary>
        /// <param name="strMdx"></param>
        /// <param name="ConnectionStringName"></param>
        /// <returns></returns>
        public DataTable GetDataByCon(string strMdx, string ConStr)
        {
            DataTable dt = new DataTable();
            try
            {
                using (AdomdConnection _connection = new AdomdConnection(ConStr))
                {
                    if (_connection != null)
                    {
                        if (_connection.State == ConnectionState.Closed)
                            _connection.Open();

                        AdomdCommand cmd = _connection.CreateCommand();
                        cmd.CommandText = strMdx;

                        var executexml = cmd.ExecuteXmlReader();
                        CellSet cellset = CellSet.LoadXml(executexml);
                        //_connection.GetSchemaDataSet
                        _connection.Close();
                        dt = ToDataTable(cellset);
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        
		#endregion
    }
}
