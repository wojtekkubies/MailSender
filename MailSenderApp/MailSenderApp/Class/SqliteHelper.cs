using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;

namespace MailSenderApp.Class
{
    internal class LiteHelper
    {
        SQLiteConnection sqlLiteConnector;

        public LiteHelper()
        {
            sqlLiteConnector = new SQLiteConnection();
        }

        public string LiteDbName
        {
            get
            {
                return "data.db";
            }
        }

        public string AbsolutePath
        {
            get
            {
                return Path.Combine(CurrentPath, LiteDbName).ToString();
            }
        }

        public string CurrentPath
        {
            get
            {
                return (new FileInfo(System.Reflection.Assembly.GetEntryAssembly().Location)).Directory.ToString();
            }
        }

        static string createTableQuery = @"CREATE TABLE IF NOT EXISTS [Szablony] (
                          [ID_Szablonu] INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                          [NazwaSzablonu] NVARCHAR(200)  NULL,
                          [TrescSzablonu] VARCHAR(2048)  NULL
                          )";
        public DataSet ReutrnDataSet()
        {
            DataSet ds4 = new DataSet();
            try
            {
                string query = "SELECT ID_Szablonu, NazwaSzablonu, TrescSzablonu FROM Szablony";
                using (SQLiteConnection con = new SQLiteConnection("data source=" + AbsolutePath))
                {
                    if (con.State != System.Data.ConnectionState.Open)
                    {
                        con.Open();
                    }
                    using (SQLiteCommand com = new SQLiteCommand(con))
                    {
                        SQLiteDataAdapter connect4 = new SQLiteDataAdapter(query, con);

                        connect4.Fill(ds4);
                        con.Close();
                    }
                }
                return ds4;
            }
            catch
            {
                return ds4;
            }
        }

        public bool Delete(int id)
        {
            try
            {
                string query = "DELETE FROM Szablony WHERE ID_Szablonu =" + id.ToString();
                using (SQLiteConnection con = new SQLiteConnection("data source=" + AbsolutePath))
                {
                    if (con.State != System.Data.ConnectionState.Open)
                    {
                        con.Open();
                    }
                    using (SQLiteCommand com = new SQLiteCommand(con))
                    {
                        com.CommandText = query;
                        com.ExecuteNonQuery();
                    }
                    con.Close();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreateDbAndTableIfNotExist()
        {
            try
            {
                string pat = Path.Combine(CurrentPath, LiteDbName);
                if (File.Exists(Path.Combine(CurrentPath, LiteDbName)))
                {
                    return true;
                }
                else
                {
                    SQLiteConnection.CreateFile(LiteDbName);
                    using (SQLiteConnection con = new SQLiteConnection("data source=" + AbsolutePath))
                    {
                        using (SQLiteCommand com = new SQLiteCommand(con))
                        {
                            con.Open();
                            com.CommandText = createTableQuery;
                            com.ExecuteNonQuery();
                            con.Close();
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool AddNewTemplate(string name, string text)
        {
            try
            {
                string query = "insert into Szablony (NazwaSzablonu, TrescSzablonu) values ('" + name + "','" + text + "')";
                using (SQLiteConnection con = new SQLiteConnection("data source=" + AbsolutePath))
                {
                    if (con.State != System.Data.ConnectionState.Open)
                    {
                        con.Open();
                    }
                    using (SQLiteCommand com = new SQLiteCommand(con))
                    {
                        com.CommandText = query;
                        com.ExecuteNonQuery();
                    }
                    con.Close();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public List<Template> GetAllTemplates()
        {
            List<Template> result = new List<Template>();
            try
            {
                string query = "SELECT ID_Szablonu, NazwaSzablonu, TrescSzablonu FROM Szablony";
                using (SQLiteConnection con = new SQLiteConnection("data source=" + AbsolutePath))
                {
                    if (con.State != System.Data.ConnectionState.Open)
                    {
                        con.Open();
                    }
                    using (SQLiteCommand com = new SQLiteCommand(con))
                    {
                        com.CommandText = query;
                        SQLiteDataReader reader = com.ExecuteReader();
                        while (reader.Read())
                        {
                            result.Add(new Template
                            {
                                TemplateId = int.Parse(reader["ID_Szablonu"].ToString()),
                                TemplateName = reader["NazwaSzablonu"].ToString(),
                                TemplateText = reader["TrescSzablonu"].ToString()
                            });
                        }
                        con.Close();
                    }
                }
                return result;
            }
            catch
            {
                return result;
            }
        }
    }
}
