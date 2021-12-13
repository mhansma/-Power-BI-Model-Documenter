using System;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.AnalysisServices.Tabular;
using Microsoft.Data.SqlClient;
using System.Security.Cryptography;
using System.Management;

namespace Power_BI_Model_Documenter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string PowerBiServer = "";
        private string PowerBiDatabase = "";
        private string SqlConnetionString = "";
        private string SqlAuthenticateMode = "";
        private SqlConnection TargetSqlConnection;
        private Database TargetPowerBiConnection;

        public MainWindow()
        {
            InitializeComponent();

            //Database.Text = "1f3caa8c-78e3-4daf-b3e8-595a059879e4";
            //Server.Text = "localhost:52444";

            if (App.GetSetting("TargetDatabaseName")?.Length > 0)
            {
                DatabaseName.Text = App.GetSetting("TargetDatabaseName");
            }

            if (App.GetSetting("TargetDatabaseServer")?.Length > 0)
            {
                DatabaseServer.Text = App.GetSetting("TargetDatabaseServer");
            }

            if (App.GetSetting("TargetDatabaseUsername")?.Length > 0)
            {
                DatabaseUsername.Text = App.GetSetting("TargetDatabaseUsername");
            }

            if (App.GetSetting("TargetDatabasePassword")?.Length > 0)
            {
                DatabasePassword.Password = Unprotect(App.GetSetting("TargetDatabasePassword"));
            }

            if (App.GetSetting("TargetDatabaseAuthenticateMethode")?.Length > 0)
            {
                SqlAuthenticate.Text = App.GetSetting("TargetDatabaseAuthenticateMethode");
                Debug.WriteLine(App.GetSetting("TargetDatabaseAuthenticateMethode"));

            }
        }

        private async void PushModel_Click(object sender, RoutedEventArgs e)
        {
            string SqlQuery;
            SqlCommand SqlCommand;

            Log.Text = "Pushing Model" + Environment.NewLine;
            await Task.Delay(1);

            using (SqlConnection SqlConnection = GetSqlConnection())
            {
                if (SqlConnection == null || SqlConnection.State != System.Data.ConnectionState.Open)
                {
                    return;
                }

                SqlCommand = new SqlCommand("", SqlConnection);

                using (Database PowerBiDatabase = GetPowerBiConnection())
                {
                    if (PowerBiDatabase == null || !PowerBiDatabase.Server.Connected)
                    {
                        return;
                    }

                    if (ModelName.Text.ToLower() is "model" or "modelname")
                    {
                        _ = MessageBox.Show("Please fill in a model name. This is the name the Power BI model is stored in the SQL Database.");
                        return;
                    }

                    PowerBiDatabase.Model.Name = ModelName.Text;
                    Log.AppendText("Model: " + PowerBiDatabase.Model.Name + Environment.NewLine);
                    Log.ScrollToEnd();

                    SqlQuery = "[documentation].[AddDataset] @DatasetSchema = '" + ModelName.Text + "', " +
                            "                             @DatasetName = '" + ModelName.Text + "'";
                    SqlCommand.CommandText = SqlQuery;
                    _ = await SqlCommand.ExecuteNonQueryAsync();

                    //Deactivate all dataset attributes.
                    SqlQuery = "[documentation].[DeactivateDataset] @DatasetSchema = '" + ModelName.Text + "'; ";
                    SqlCommand.CommandText = SqlQuery;
                    _ = await SqlCommand.ExecuteNonQueryAsync();

                    foreach (Microsoft.AnalysisServices.Tabular.Table Table in PowerBiDatabase.Model.Tables)
                    {
                        if (!Table.IsPrivate && !Table.Name.StartsWith("LocalDateTable_"))
                        {
                            Log.AppendText("Table: " + Table.Name + Environment.NewLine);
                            Log.ScrollToEnd();

                            SqlQuery = "[documentation].[AddTable] @DatasetSchema = '" + ModelName.Text + "', " +
                                    "                           @TableName = '" + Table.Name + "', " +
                                    "                           @IsHidden = '" + Table.IsHidden.ToString() + "'";
                            SqlCommand.CommandText = SqlQuery;
                            _ = await SqlCommand.ExecuteNonQueryAsync();

                            foreach (Column Column in Table.Columns)
                            {
                                if (Column.Type != ColumnType.RowNumber)
                                {
                                    Log.AppendText("Column: " + Column.Name + Environment.NewLine);
                                    Log.ScrollToEnd();

                                    SqlQuery = "[documentation].[AddTableColumn] @DatasetSchema = '" + ModelName.Text + "', " +
                                        "                                     @TableName = '" + Table.Name + "', " +
                                        "                                     @ColumnName = '" + Column.Name + "', " +
                                        "                                     @IsHidden = '" + Table.IsHidden.ToString() + "'";
                                    SqlCommand.CommandText = SqlQuery;
                                    _ = await SqlCommand.ExecuteNonQueryAsync();
                                }
                            }

                            foreach (Measure Measure in Table.Measures)
                            {
                                Log.AppendText("Measure: " + Measure.Name + Environment.NewLine);
                                Log.ScrollToEnd();

                                SqlQuery = "[documentation].[AddTableMeasure] @DatasetSchema = '" + ModelName.Text + "', " +
                                        "                                  @TableName = '" + Table.Name + "', " +
                                        "                                  @MeasureName = '" + Measure.Name + "', " +
                                        "                                  @Expression = '" + Measure.Expression.Replace("\'", "\'\'") + "', " +
                                        "                                  @IsHidden = '" + Table.IsHidden.ToString() + "'";
                                SqlCommand.CommandText = SqlQuery;
                                _ = await SqlCommand.ExecuteNonQueryAsync();
                            }
                        }
                    }
                    _ = PowerBiDatabase.Model.SaveChanges();
                    PowerBiDatabase.Server.Disconnect();
                    Log.AppendText("Save model" + Environment.NewLine);
                }
                SqlConnection.Close();
                Log.AppendText("Ready!" + Environment.NewLine);
                Log.ScrollToEnd();
            }
        }

        private async void PullModel_Click(object sender, RoutedEventArgs e)
        {
            string SqlQuery;
            SqlCommand SqlCommand;

            Log.Text = "Pulling Descriptions";
            await Task.Delay(1);

            using (SqlConnection SqlConnection = GetSqlConnection())
            {
                if (SqlConnection == null || SqlConnection.State != System.Data.ConnectionState.Open)
                {
                    return;
                }

                SqlCommand = new SqlCommand("", SqlConnection);

                using (Database PowerBiDatabase = GetPowerBiConnection())
                {
                    if (PowerBiDatabase == null || !PowerBiDatabase.Server.Connected)
                    {
                        return;
                    }

                    if (ModelName.Text.ToLower() is "model" or "modelname")
                    {
                        _ = MessageBox.Show("Please fill in a model name. This is the name the Power BI model is stored in the SQL Database.");
                        return;
                    }

                    PowerBiDatabase.Model.Name = ModelName.Text;
                    Log.AppendText("Model: " + PowerBiDatabase.Model.Name + Environment.NewLine);

                    foreach (Microsoft.AnalysisServices.Tabular.Table Table in PowerBiDatabase.Model.Tables)
                    {
                        Log.AppendText("Read Table: " + Table.Name + Environment.NewLine);
                        Log.ScrollToEnd();

                        SqlQuery = "[documentation].[GetDefinition] @DatasetSchema = '" + ModelName.Text + "', " +
                                "                                @TableName = '" + Table.Name + "'";
                        SqlCommand.CommandText = SqlQuery;

                        SqlDataReader reader = await SqlCommand.ExecuteReaderAsync();
                        while (reader.Read())
                        {
                            Table.Description = reader.GetString(0);
                            Log.AppendText("  Write: " + Table.Description + Environment.NewLine);
                            Log.ScrollToEnd();
                        }
                        reader.Close();

                        foreach (Column Column in Table.Columns)
                        {
                            if (Column.Type != ColumnType.RowNumber)
                            {
                                Log.AppendText("Read Column: " + Column.Name + Environment.NewLine);
                                Log.ScrollToEnd();

                                SqlQuery = "[documentation].[GetDefinition] @DatasetSchema = '" + ModelName.Text + "', " +
                                        "                                @TableName = '" + Table.Name + "', " +
                                        "                                @ColumnName = '" + Column.Name + "'";
                                SqlCommand.CommandText = SqlQuery;

                                reader = await SqlCommand.ExecuteReaderAsync();
                                while (reader.Read())
                                {
                                    Column.Description = reader.GetString(0);
                                    Log.AppendText("  Write: " + Column.Name + " (" + Column.Description + ")" + Environment.NewLine);
                                    Log.ScrollToEnd();
                                }
                                reader.Close();
                            }
                        }

                        foreach (Measure Measure in Table.Measures)
                        {
                            Log.AppendText("Read Measure: " + Measure.Name + Environment.NewLine);
                            Log.ScrollToEnd();

                            SqlQuery = "[documentation].[GetDefinition] @DatasetSchema = '" + ModelName.Text + "', " +
                                    "                                @TableName = '" + Table.Name + "', " +
                                    "                                @MeasureName = '" + Measure.Name + "'";
                            SqlCommand.CommandText = SqlQuery;

                            reader = await SqlCommand.ExecuteReaderAsync();
                            while (reader.Read())
                            {
                                Measure.Description = reader.GetString(0);
                                Log.AppendText("  Write: " + Measure.Name + " (" + Measure.Description + ")" + Environment.NewLine);
                                Log.ScrollToEnd();
                            }
                            reader.Close();
                        }
                    }
                    _ = PowerBiDatabase.Model.SaveChanges();
                    PowerBiDatabase.Server.Disconnect();
                    Log.AppendText("Save model" + Environment.NewLine);
                }
                SqlConnection.Close();
                Log.AppendText("Ready!" + Environment.NewLine);
                Log.ScrollToEnd();
            }
        }

        private SqlConnection GetSqlConnection()
        {
            if (null == TargetSqlConnection || TargetSqlConnection.State != System.Data.ConnectionState.Open)
            {
                Log.AppendText("Connecting SQL database" + Environment.NewLine);
                Log.AppendText(" Server = " + DatabaseServer.Text + "; " + Environment.NewLine);
                Log.AppendText(" Database = " + DatabaseName.Text + "; " + Environment.NewLine);

                SqlConnetionString = "";

                if ("SQL Server Authenticate" == SqlAuthenticateMode)
                {
                    SqlConnetionString = @"Server=" + DatabaseServer.Text + "; Database=" + DatabaseName.Text + "; User ID=" + DatabaseUsername.Text + "; Password=" + DatabasePassword.Password;
                    Log.AppendText(" User ID = " + DatabaseUsername.Text + "; " + Environment.NewLine);
                    Log.AppendText(" Password = *****" + Environment.NewLine);
                }
                else if ("Azure Active Directory - MFA" == SqlAuthenticateMode)
                {
                    SqlConnetionString = @"Server=" + DatabaseServer.Text + "; Database=" + DatabaseName.Text + "; Authentication=Active Directory Interactive;";
                    Log.AppendText(" Authentication=Active Directory Interactive;" + Environment.NewLine);
                }

                try
                {
                    TargetSqlConnection = new SqlConnection(SqlConnetionString);
                    TargetSqlConnection.Open();
                    Log.AppendText("Connected to SQL" + Environment.NewLine);
                }
                catch (Exception e)
                {
                    Log.AppendText("Can't connect to SQL" + Environment.NewLine);
                    Log.AppendText(e.Message + Environment.NewLine);

                    _ = MessageBox.Show("Can't connect to SQL" + Environment.NewLine +
                       e.Message);
                }
            }
            return TargetSqlConnection;
        }

        private Database GetPowerBiConnection()
        {
            if (null == TargetPowerBiConnection || !TargetPowerBiConnection.Server.Connected)
            {
                if (PowerBiServer == "Server" || PowerBiServer == "")
                {
                    Log.AppendText("Please fill in a Power BI Server." + Environment.NewLine);
                    return null;
                }
                if (PowerBiDatabase == "Database" || PowerBiDatabase == "")
                {
                    Log.AppendText("Please fill in a Power BI Database." + Environment.NewLine);
                    return null;
                }

                Log.AppendText("Connecting Power BI" + Environment.NewLine);
                Log.AppendText(" Server: " + PowerBiServer + Environment.NewLine);
                Log.AppendText(" Database: " + PowerBiDatabase + Environment.NewLine);

                try
                {
                    Server PowerBiServerConnection = new Server();
                    PowerBiServerConnection.Connect(PowerBiServer);
                    TargetPowerBiConnection = PowerBiServerConnection.Databases.GetByName(PowerBiDatabase);
                    Log.AppendText("Connected to Power BI" + Environment.NewLine);
                }
                catch (Exception e)
                {
                    Log.AppendText("Can't connect to Power BI" + Environment.NewLine);
                    Log.AppendText(e.Message + Environment.NewLine);

                    _ = MessageBox.Show("Can't connect to Power BI" + Environment.NewLine +
                                            e.Message);
                }
            }
            return TargetPowerBiConnection;
        }

        public void SetPowerBiModelName()
        {
            Database PowerBiDatabaseConnectie = GetPowerBiConnection();
            if (null != PowerBiDatabaseConnectie)
            {
                if (PowerBiDatabaseConnectie.Model.Name.ToLower() is not "model" and not "modelname")
                {
                    ModelName.Text = PowerBiDatabaseConnectie.Model.Name;
                }
            }
        }

        private void Server_TextChanged(object sender, TextChangedEventArgs e)
        {
            PowerBiServer = Server.Text;
            SetPowerBiModelName();
        }

        private void Database_TextChanged(object sender, TextChangedEventArgs e)
        {
            PowerBiDatabase = Database.Text;
            SetPowerBiModelName();
        }

        private void ModelName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (ModelName.Text != "Modelname")
            {
                App.SetSetting("ModelName", ModelName.Text);
            }
        }

        private void DatabaseName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (DatabaseName.Text != "Database")
            {
                App.SetSetting("TargetDatabaseName", DatabaseName.Text);
            }
        }

        private void DatabaseServer_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (DatabaseServer.Text != "Server")
            {
                App.SetSetting("TargetDatabaseServer", DatabaseServer.Text);
            }
        }
        private void DatabaseUsername_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (DatabaseUsername.Text != "User")
            {
                App.SetSetting("TargetDatabaseUsername", DatabaseUsername.Text);
            }
        }

        private void DatabasePassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DatabasePassword.Password != "aaaaa")
            {
                App.SetSetting("TargetDatabasePassword", Protect(DatabasePassword.Password));
            }
        }

        private void SqlAuthenticate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem typeItem = (ComboBoxItem)SqlAuthenticate.SelectedItem;
            SqlAuthenticateMode = typeItem.Content.ToString();

            if (DatabasePassword is not null && DatabaseUsername is not null)
            {
                App.SetSetting("TargetDatabaseAuthenticateMethode", SqlAuthenticateMode);

                if (SqlAuthenticateMode == "SQL Server Authenticate")
                {
                    DatabasePassword.Visibility = Visibility.Visible;
                    DatabaseUsername.Visibility = Visibility.Visible;
                }
                else
                {
                    DatabasePassword.Visibility = Visibility.Hidden;
                    DatabaseUsername.Visibility = Visibility.Hidden;
                }
            }
        }

        public string Protect(string str)
        {
            byte[] entropy = Encoding.ASCII.GetBytes(CID());
            byte[] data = Encoding.ASCII.GetBytes(str);
            string protectedData = Convert.ToBase64String(ProtectedData.Protect(data, entropy, DataProtectionScope.CurrentUser));
            return protectedData;
        }

        public string Unprotect(string str)
        {
            if (IsBase64String(str))
            {
                try
                {
                    byte[] protectedData = Convert.FromBase64String(str);
                    byte[] entropy = Encoding.ASCII.GetBytes(CID());
                    string data = Encoding.ASCII.GetString(ProtectedData.Unprotect(protectedData, entropy, DataProtectionScope.CurrentUser));
                    return data;
                }
                catch
                {

                }
            }
            return null;
        }

        public bool IsBase64String(string base64)
        {
            Span<byte> buffer = new Span<byte>(new byte[base64.Length]);
            return Convert.TryFromBase64String(base64, buffer, out int bytesParsed);
        }

        private string CID()
        {
            ManagementObjectSearcher search = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
            ManagementObjectCollection searchs = search.Get();
            string serial = ""; string cpuid = "";
            foreach (ManagementObject id in searchs)
            {
                serial = (string)id["SerialNumber"];
            }

            search = new ManagementObjectSearcher("Select ProcessorId From Win32_processor");
            searchs = search.Get();
            foreach (ManagementObject id in searchs)
            {
                cpuid = (string)id["ProcessorId"];
            }

            return serial + cpuid;
        }
    }
}
