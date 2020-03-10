using Microsoft.Win32;
using System.Collections;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Windows;

namespace DBFtoDBFConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OdbcConnection obdcconn = new System.Data.Odbc.OdbcConnection();
                string fileName = "DUTY.DBF";
                //OpenFileDialog openFileDialog = new OpenFileDialog();
                //if (openFileDialog.ShowDialog() == true)
                //{
                    //lblStatus.Content = "File conversion is in progress.. Please wait..!!";
                    obdcconn.ConnectionString = GetConnectionString(System.Configuration.ConfigurationSettings.AppSettings["InputPath"]);
                    obdcconn.Open();
                    OdbcCommand oCmd = obdcconn.CreateCommand();
                    oCmd.CommandText = "SELECT * FROM " + System.Configuration.ConfigurationSettings.AppSettings["InputPath"];

                    /*Load data to table*/

                    DataTable dt1 = new DataTable();
                    dt1.Load(oCmd.ExecuteReader());

                    string currentPath = Path.GetFullPath(System.Configuration.ConfigurationSettings.AppSettings["InputPath"]);
                    currentPath = Directory.GetParent(currentPath).FullName + "\\";
                    obdcconn.Close();

                    WriteDataToTemplate(currentPath, fileName, dt1);
                    //DataSetIntoDBF(fileName, currentPath, dt1);
                    //lblStatus.Content = "File conversion completed, file is in " + dir + "\\" + fileName;
                //}
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void WriteDataToTemplate(string filePath, string fileName, DataTable table)
        {
            //CopyTemplateToWorkingDir(filePath, fileName);
            try
            {
                OdbcConnection obdcconn = new System.Data.Odbc.OdbcConnection(GetConnectionString(dir + fileName));
                OdbcCommand oCmd = null;
                oCmd = obdcconn.CreateCommand();
                obdcconn.Open();
                oCmd.CommandText = "INSERT INTO " + fileName + " VALUES ('0000',' 0000','0000','0000','0','0000',0,0,0,'')";
                oCmd.ExecuteNonQuery();

                oCmd = obdcconn.CreateCommand();
                oCmd.CommandText = "DELETE FROM " + fileName + " WHERE BOARD <> '0000';";
                var affected = oCmd.ExecuteNonQuery();

                try
                {
                    var count = 0;
                    foreach (DataRow row in table.Rows)
                    {
                        count++;
                        oCmd = obdcconn.CreateCommand();
                        string insertSql = "INSERT INTO " + fileName + " VALUES(";
                        string values = "";
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            string value = ReplaceEscape(row[i].ToString());
                            switch (i)
                            {
                                case 1:
                                    value = " " + value;
                                    break;
                                case 2:
                                    value = value.PadRight(6, ' ');
                                    break;
                                case 6:
                                    value = CalculateDay(value);
                                    values = values + value + ",";
                                    values = values + "0,";
                                    values = values + "0,";
                                    break;
                                case 7:
                                    values = "'" + value + "'," + values;
                                    break;
                                default:
                                    break;
                            }
                            if (i != 6 && i != 0 && i != 7)
                            {
                                values = values + "'" + value + "',";
                            }
                        }

                        values = values.Substring(0, values.Length - 1) + ")";
                        insertSql = insertSql + values;
                        oCmd.CommandText = insertSql;
                        oCmd.ExecuteNonQuery();
                        oCmd.Dispose();
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                oCmd = obdcconn.CreateCommand();
                oCmd.CommandText = "DELETE FROM " + fileName + " WHERE BOARD='0000'";
                oCmd.ExecuteNonQuery();
                obdcconn.Close();

                CopyTemplateToWorkingDir(System.Configuration.ConfigurationSettings.AppSettings["OutputPath"], fileName);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private static void CopyTemplateToWorkingDir(string filePath, string fileName)
        {
            string dir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string fileFullPath = filePath + fileName;

            //StreamReader reader = new StreamReader(dir);
            //string fileContent = reader.ReadToEnd();

            //StreamWriter writer = new StreamWriter(fileFullPath);
            //writer.Write(fileContent);

            File.Copy(dir + "\\" + fileName, fileFullPath, true);
        }

        public static void DataSetIntoDBF(string fileName, string filePath, DataTable table)
        {
            ArrayList list = new ArrayList();

            if (File.Exists(filePath + fileName))
            {
                MessageBox.Show("A file named DUTY.DBF already exist in " + filePath + " will be deleted!!!");
                File.Delete(filePath + fileName);
            }

            string createSql = "create table " + fileName + " (";

            foreach (DataColumn dc in table.Columns)
            {
                string fieldName = dc.ColumnName;

                string type = dc.DataType.ToString();

                switch (fieldName)
                {
                    case "BOARD":
                    case "JOURNEY":
                    case "START":
                        type = "CHAR(4)";
                        break;
                    case "DUTY":
                        type = "CHAR(5)";
                        break;
                    case "ROUTE":
                        type = "CHAR(6)";
                        break;
                    case "DIR":
                        type = "CHAR(1)";
                        break;
                    case "DAY":
                        type = "DECIMAL(4,0)";
                        break;
                    case "DESTNAME":
                        type = "CHAR(16)";
                        break;
                }

                list.Add(fieldName);

                if (fieldName == "DAY")
                {
                    fieldName = "DAYS";
                    createSql = createSql + "[" + fieldName + "]" + " " + type + ",";
                    createSql = createSql + "[BEGIN] DECIMAL(4,0),";
                    createSql = createSql + "[END] DECIMAL(4,0),";
                }
                else if (fieldName == "DESTNAME")
                { fieldName = "STAGENAME"; }

                if (fieldName != "DESTID" && fieldName != "DAYS")
                { createSql = createSql + "[" + fieldName + "]" + " " + type + ","; }

            }

            createSql = createSql.Substring(0, createSql.Length - 1) + ")";

            OleDbConnection con = new OleDbConnection(GetConnection(filePath));
            OleDbCommand cmd = new OleDbCommand
            {
                Connection = con
            };

            con.Open();
            cmd.CommandText = createSql;
            cmd.ExecuteNonQuery();

            foreach (DataRow row in table.Rows)
            {
                string insertSql = "insert into " + fileName + " values(";
                string values = "";
                for (int i = 0; i < list.Count; i++)
                {
                    string value = ReplaceEscape(row[list[i].ToString()].ToString());
                    switch (i)
                    {
                        case 1:
                            value = " " + value;
                            break;
                        case 2:
                            value = value.PadRight(6, ' ');
                            break;
                        case 6:
                            value = CalculateDay(value);
                            values = values + value + ",";
                            values = values + "0,";
                            values = values + "0,";
                            break;
                        case 7:
                            values = "'" + value + "'," + values;
                            break;
                        default:
                            break;
                    }
                    if (i != 6 && i != 0 && i != 7)
                    {
                        values = values + "'" + value + "',";
                    }
                }

                values = values.Substring(0, values.Length - 1) + ")";
                insertSql = insertSql + values;
                cmd.CommandText = insertSql;
                cmd.ExecuteNonQuery();
            }
            con.Close();

            MessageBox.Show("A file named DUTY.DBF created in " + filePath + " !!!");
            return;
        }

        private static string CalculateDay(string value)
        {
            char[] dayChar = value.ToCharArray();
            int count = 0;
            for (int item = 0; item < dayChar.Length; item++)
            {
                switch (dayChar[item])
                {
                    case 'M':
                        count += 1;
                        break;
                    case 'T':
                        if (item == 1)
                        {
                            count += 2;
                        }
                        else
                        {
                            count += 8;
                        }
                        break;
                    case 'W':
                        count += 4;
                        break;
                    case 'F':
                        count += 16;
                        break;
                    case 'S':
                        if (item == 5)
                        {
                            count += 32;
                        }
                        else
                        {
                            count += 64;
                        }

                        break;
                    default:
                        break;
                }
            }
            return count.ToString();
        }

        private static string GetConnection(string path)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=dBASE III;";
            //return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=dBASE IV;";

        }

        public static string ReplaceEscape(string str)
        {
            str = str.Replace("'", "''");
            return str;
        }

        public static string GetConnectionString(string filePath, string fileName)
        {
            //return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + fileName + ";;Database=" + filePath + ";Extended Properties=dBASE III;";
            return "Driver={Microsoft dBase Driver (*.dbf)};SourceType=DBF;SourceDB=" + filePath + fileName + "Exclusive=No; NULL=NO;DELETED=NO;BACKGROUNDFETCH=NO;";
        }

        public static string GetConnectionString(string filePath)
        {
            return "Driver={Microsoft dBase Driver (*.dbf)};SourceType=DBF;SourceDB=" + filePath + ";Exclusive=No; NULL=NO;DELETED=NO;BACKGROUNDFETCH=NO;";
        }
    }
}
