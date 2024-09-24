using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static LogInfo_Data_Trasnfert.Form1;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace LogInfo_Data_Trasnfert
{
    public partial class Form1 : Form
    {
        public string LogPath = AppDomain.CurrentDomain.BaseDirectory + @"Log.txt";
        public string path = AppDomain.CurrentDomain.BaseDirectory + @"Config.ini";
        public string DownloadedFile = string.Empty;
        public string connetionString = ""; //null;
        public string connetionString1 = ""; //null;
        public string connetionString2 = ""; //null;
        public string connetionString3 = ""; //null;
        public string LogLevel = string.Empty;
        public string HostRemoteProgram = string.Empty;
        public SqlConnection cnn0;
        public SqlConnection cnn1;
        public SqlConnection cnn2;
        public SqlConnection cnn3;
        public string license_code = string.Empty;
        public string user_name = string.Empty;
        public Boolean ConnectionStatus = false;
        public string APIKEY = string.Empty;
        public string APIURL = string.Empty;
        public string APIUSER = string.Empty;
        public string SESSIONAME = "";
        public string ArtVente = "";
        public string ArtAchat = "";

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
                 string key, string def, StringBuilder retVal,
            int size, string filePath);
        public Form1()
        {

            InitializeComponent();
        }

        private void SetupDataGridView()
        {


        }
        //public Boolean //MessageBox.Show(string message)
        //{
        //    Boolean retVal = true;
        //    FileInfo Fi = new FileInfo(LogPath);
        //    string LogPathArch = AppDomain.CurrentDomain.BaseDirectory + @"Log_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".txt";

        //    if (LogLevel == "1")
        //    {
        //        System.IO.File.AppendAllLines(LogPath, new string[] { DateTime.Now + " : " + message });
        //    }
        //    if (Fi.Length > 20000000)
        //    {
        //        File.Move(LogPath, LogPathArch);
        //    }

        //    return retVal;
        //}

        private void ouvrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt1 = SelectTable("Societe", 0); // SelectTable contient select * from ...

            foreach (DataRow row in dt1.Rows)
            {
                comboBox1.Items.Add(row["Intitule_etablissement"]);
            }


            SetupDataGridView();
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Parcourir ...",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xml",
                Filter = "xml files (*.xml)|*.xml",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true


            };

            string xmlFilePath = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                xmlFilePath = openFileDialog1.FileName;







                XDocument xdoc = XDocument.Load(xmlFilePath);
                XNamespace ns = "urn:com.workday/picof";  // Efface les éléments existants
             
          

                DataTable dt = new DataTable();

                dt.Columns.Add("Employee_ID");
                dt.Columns.Add("Name");
                dt.Columns.Add("Marital_Status");
                dt.Columns.Add("Payroll_Company_Name");
                dt.Columns.Add("Pay_Group_ID");
                dt.Columns.Add("Pay_Group_Name");
                dt.Columns.Add("First_Name");
                dt.Columns.Add("Last_Name");
                dt.Columns.Add("Title");
                dt.Columns.Add("Country_for_Name");
                dt.Columns.Add("Birth_Date");
                dt.Columns.Add("First_Address_Line1_Data");
                dt.Columns.Add("First_Address_Line2_Data");
                dt.Columns.Add("Address_Line_Label");
                dt.Columns.Add("Address_Line_Type");
                dt.Columns.Add("First_Municipality");
                dt.Columns.Add("Employee_Status");
                dt.Columns.Add("Hire_Date");
                dt.Columns.Add("Seniority_Date");
                dt.Columns.Add("Operation_Position");
                dt.Columns.Add("Position_ID");
                dt.Columns.Add("Effective_Date");
                dt.Columns.Add("Business_Title");
                dt.Columns.Add("Position_Time_Type");
                dt.Columns.Add("Scheduled_Weekly_Hours");
                dt.Columns.Add("Default_Weekly_Hours");
                dt.Columns.Add("Total_Base_Pay");
                dt.Columns.Add("Base_Pay_Frequency");
                dt.Columns.Add("Organization_One");
                dt.Columns.Add("Organization_Two");
                dt.Columns.Add("Organization_Three");
                dt.Columns.Add("Job_Profile");
                dt.Columns.Add("Job_Family");
                dt.Columns.Add("Operation_Contract");
                dt.Columns.Add("Contract_Type");
                dt.Columns.Add("Start_Date_Contract");
                dt.Columns.Add("End_Date_Contract");
                dt.Columns.Add("Job_Family_Group");
                dt.Columns.Add("CostCenter_Code");
                dt.Columns.Add("Nbr_of_Dependents");
                dt.Columns.Add("Nationality");
                dt.Columns.Add("NomBanque");
                dt.Columns.Add("NomGuichet");
                dt.Columns.Add("ModeDePaiement");
                dt.Columns.Add("Local_Termination_Reason");
                dt.Columns.Add("Termination_Date");
                dt.Columns.Add("Flexible_Work_Arrangements");
                dt.Columns.Add("Payment_Type");
                dt.Columns.Add("Bank_Name");
                dt.Columns.Add("Branch_Code");
                dt.Columns.Add("IBAN");
                dt.Columns.Add("Niveau");
                dt.Columns.Add("Indice");
                dt.Columns.Add("Prenom2");
                dt.Columns.Add("NoBulletinModele");
                dt.Columns.Add("HoraireBase");
                dt.Columns.Add("Qualification");
                dt.Columns.Add("Service");
                dt.Columns.Add("DelivrePar");
                dt.Columns.Add("First_Address_Line_Data");
                dt.Columns.Add("nom_fichier");
                // Dictionnaire pour suivre les colonnes d'identifiants
                var identifierColumns = new Dictionary<string, int>();

                var employees = xdoc.Descendants(ns + "Employee");

                foreach (var emp in employees)
                {
                    DataRow row = dt.NewRow();
                    row["Niveau"] = DBNull.Value;
                    row["Indice"] = DBNull.Value;
                    row["NoBulletinModele"] = DBNull.Value;
                    row["Prenom2"] = DBNull.Value;
                    row["HoraireBase"] = DBNull.Value;
                    row["Qualification"] = DBNull.Value;
                    row["Service"] = DBNull.Value;
                    row["DelivrePar"] = DBNull.Value;

                    row["Employee_ID"] = emp.Descendants(ns + "Summary").Descendants(ns + "Employee_ID").FirstOrDefault()?.Value;
                    row["Name"] = emp.Descendants(ns + "Summary").Descendants(ns + "Name").FirstOrDefault()?.Value;
                    row["Payroll_Company_Name"] = emp.Descendants(ns + "Summary").Descendants(ns + "Payroll_Company_Name").FirstOrDefault()?.Value;
                    row["Pay_Group_ID"] = emp.Descendants(ns + "Summary").Descendants(ns + "Pay_Group_ID").FirstOrDefault()?.Value;
                    row["Pay_Group_Name"] = emp.Descendants(ns + "Summary").Descendants(ns + "Pay_Group_Name").FirstOrDefault()?.Value;

                    row["First_Name"] = emp.Descendants(ns + "Personal").Descendants(ns + "First_Name").FirstOrDefault()?.Value;
                    row["Last_Name"] = emp.Descendants(ns + "Personal").Descendants(ns + "Last_Name").FirstOrDefault()?.Value;
                    row["Marital_Status"] = emp.Descendants(ns + "Personal").Descendants(ns + "Marital_Status").FirstOrDefault()?.Value;
                    row["Title"] = emp.Descendants(ns + "Personal").Descendants(ns + "Title").FirstOrDefault()?.Value;
                    row["Country_for_Name"] = emp.Descendants(ns + "Personal").Descendants(ns + "Country_for_Name").FirstOrDefault()?.Value;
                    row["Birth_Date"] = emp.Descendants(ns + "Personal").Descendants(ns + "Birth_Date").FirstOrDefault()?.Value;


                    row["First_Address_Line_Data"] = DBNull.Value;
                    row["Address_Line_Label"] = DBNull.Value;
                    row["Address_Line_Type"] = DBNull.Value;
                    var addressLineDataList = emp.Descendants(ns + "Personal")
      .Descendants(ns + "First_Address_Line_Data")
      .ToList();

                    // Assigner les valeurs aux colonnes spécifiques
                    if (addressLineDataList.Count > 0)
                    {
                        row["First_Address_Line1_Data"] = addressLineDataList
                            .FirstOrDefault(x => x.Attribute(XName.Get("Label", ns.NamespaceName))?.Value == "Address Line 1")?.Value;

                        row["First_Address_Line2_Data"] = addressLineDataList
                            .FirstOrDefault(x => x.Attribute(XName.Get("Label", ns.NamespaceName))?.Value == "Address Line 2")?.Value;
                    }

                    row["First_Municipality"] = emp.Descendants(ns + "Personal").Descendants(ns + "First_Municipality").FirstOrDefault()?.Value;

                    row["Employee_Status"] = emp.Descendants(ns + "Status").Descendants(ns + "Employee_Status").FirstOrDefault()?.Value;
                    row["Hire_Date"] = emp.Descendants(ns + "Status").Descendants(ns + "Hire_Date").FirstOrDefault()?.Value;
                    row["Seniority_Date"] = emp.Descendants(ns + "Status").Descendants(ns + "Seniority_Date").FirstOrDefault()?.Value;

                    row["Termination_Date"] = emp.Descendants(ns + "Status").Descendants(ns + "Termination_Date").FirstOrDefault()?.Value;
                    row["Local_Termination_Reason"] = emp.Descendants(ns + "Status").Descendants(ns + "Local_Termination_Reason").FirstOrDefault()?.Value;

                    row["Operation_Position"] = emp.Descendants(ns + "Position").Descendants(ns + "Operation").FirstOrDefault()?.Value;
                    row["Position_ID"] = emp.Descendants(ns + "Position").Descendants(ns + "Position_ID").FirstOrDefault()?.Value;

                    row["Effective_Date"] = emp.Descendants(ns + "Position").Descendants(ns + "Effective_Date").FirstOrDefault()?.Value;
                    row["Business_Title"] = emp.Descendants(ns + "Position").Descendants(ns + "Business_Title").FirstOrDefault()?.Value;
                    row["Position_Time_Type"] = emp.Descendants(ns + "Position").Descendants(ns + "Position_Time_Type").FirstOrDefault()?.Value;
                    row["Scheduled_Weekly_Hours"] = emp.Descendants(ns + "Position").Descendants(ns + "Scheduled_Weekly_Hours").FirstOrDefault()?.Value;
                    row["Default_Weekly_Hours"] = emp.Descendants(ns + "Position").Descendants(ns + "Default_Weekly_Hours").FirstOrDefault()?.Value;
                    row["Total_Base_Pay"] = emp.Descendants(ns + "Position").Descendants(ns + "Total_Base_Pay").FirstOrDefault()?.Value;
                    row["Base_Pay_Frequency"] = emp.Descendants(ns + "Position").Descendants(ns + "Base_Pay_Frequency").FirstOrDefault()?.Value;
                    row["Organization_One"] = emp.Descendants(ns + "Position").Descendants(ns + "Organization_One").FirstOrDefault()?.Value;
                    row["Organization_Two"] = emp.Descendants(ns + "Position").Descendants(ns + "Organization_Two").FirstOrDefault()?.Value;
                    row["Organization_Three"] = emp.Descendants(ns + "Position").Descendants(ns + "Organization_Three").FirstOrDefault()?.Value;
                    row["Job_Profile"] = emp.Descendants(ns + "Position").Descendants(ns + "Job_Profile").FirstOrDefault()?.Value;
                    row["Job_Family"] = emp.Descendants(ns + "Position").Descendants(ns + "Job_Family").FirstOrDefault()?.Value;

                    row["Operation_Contract"] = emp.Descendants(ns + "Contract").Descendants(ns + "Operation").FirstOrDefault()?.Value;
                    row["Contract_Type"] = emp.Descendants(ns + "Contract").Descendants(ns + "Contract_Type").FirstOrDefault()?.Value;
                    row["Start_Date_Contract"] = emp.Descendants(ns + "Contract").Descendants(ns + "Start_Date").FirstOrDefault()?.Value;
                    row["End_Date_Contract"] = emp.Descendants(ns + "Contract").Descendants(ns + "End_Date").FirstOrDefault()?.Value;
                 
                    // Identifier
                    var identifiers = emp.Descendants(ns + "Identifier");
                    foreach (var id in identifiers)
                    {
                        string type = id.Descendants(ns + "Identifier_Type").FirstOrDefault()?.Value;
                        string value = id.Descendants(ns + "Identifier_Value").FirstOrDefault()?.Value;

                        // Générer un nom unique pour la colonne
                        string columnName = $"Identifier_Type_{type}";

                        // Trouver un nom de colonne unique si une colonne avec le nom existe déjà
                        int suffix = 1;
                        while (dt.Columns.Contains($"{columnName}_{suffix}"))
                        {
                            suffix++;
                        }
                        if (!dt.Columns.Contains(columnName))
                        {
                            dt.Columns.Add(columnName);
                        }



                        row[columnName] = value;
                    }

                    row["Job_Family_Group"] = emp.Descendants(ns + "Additional_Information").Descendants(ns + "Job_Family_Group").FirstOrDefault()?.Value;
                    row["Flexible_Work_Arrangements"] = emp.Descendants(ns + "Additional_Information").Descendants(ns + "Flexible_Work_Arrangements").FirstOrDefault()?.Value;
                    row["CostCenter_Code"] = emp.Descendants(ns + "Additional_Information").Descendants(ns + "CostCenter_Code").FirstOrDefault()?.Value;
                    row["Nbr_of_Dependents"] = emp.Descendants(ns + "Additional_Information").Descendants(ns + "Nbr_of_Dependents").FirstOrDefault()?.Value;
                    row["Payment_Type"] = emp.Descendants(ns + "Payment_Election").Descendants(ns + "Payment_Type").FirstOrDefault()?.Value;
                    row["Bank_Name"] = emp.Descendants(ns + "Payment_Election").Descendants(ns + "Bank_Name").FirstOrDefault()?.Value;
                    row["Branch_Code"] = emp.Descendants(ns + "Payment_Election").Descendants(ns + "Branch_Code").FirstOrDefault()?.Value;

                    row["IBAN"] = emp.Descendants(ns + "Payment_Election").Descendants(ns + "IBAN").FirstOrDefault()?.Value;
                    row["nom_fichier"] = Path.GetFileName(openFileDialog1.FileName).ToString();
                    dt.Rows.Add(row);
                }

                dataGridView1.DataSource = dt;
                dataGridView1.Visible = true;
                button1.Visible = true;
                dataGridView1.AllowUserToAddRows = true;
                dataGridView1.ReadOnly = false;
                dataGridView1.AllowUserToDeleteRows = true;

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;


            //DB
            connetionString = "Data Source=" + IniReadValue("PARAMS", "SERVEUR") + ";Initial Catalog=" + IniReadValue("PARAMS", "DATABASE") + ";User ID=" + IniReadValue("PARAMS", "USER") + ";Password=" + IniReadValue("PARAMS", "PWD") + "; MultipleActiveResultSets=true";
            cnn0 = new SqlConnection(connetionString);




            //bdd1
            connetionString1 = "Data Source=" + IniReadValue("PARAMS", "SERVEUR1") + ";Initial Catalog=" + IniReadValue("PARAMS", "DATABASE1") + ";User ID=" + IniReadValue("PARAMS", "USER1") + ";Password=" + IniReadValue("PARAMS", "PWD1") + "; MultipleActiveResultSets=true";
            cnn1 = new SqlConnection(connetionString1);

            //bdd2
            connetionString2 = "Data Source=" + IniReadValue("PARAMS", "SERVEUR2") + ";Initial Catalog=" + IniReadValue("PARAMS", "DATABASE2") + ";User ID=" + IniReadValue("PARAMS", "USER2") + ";Password=" + IniReadValue("PARAMS", "PWD2") + "; MultipleActiveResultSets=true";
            cnn2 = new SqlConnection(connetionString2);

            //bdd3
            connetionString3 = "Data Source=" + IniReadValue("PARAMS", "SERVEUR3") + ";Initial Catalog=" + IniReadValue("PARAMS", "DATABASE3") + ";User ID=" + IniReadValue("PARAMS", "USER3") + ";Password=" + IniReadValue("PARAMS", "PWD3") + "; MultipleActiveResultSets=true";
            cnn3 = new SqlConnection(connetionString3);

            Dictionary<string,string> tables = new Dictionary<string, string>();
              tables.Add("Config", "Config"); 
          tables.Add("Categorie", "Categorie");
            tables.Add("Civilite", "Civilite");
            tables.Add("Disability", "Disability");
            tables.Add("Duree_Contractuelle", "Durée Contrat actuelle");
            tables.Add("Fonction", "Fonction");
            tables.Add("Modalite_De_Travail", "Modalite De Travail");
            tables.Add("Mode_De_Paiement", "Mode De Paiement");
            tables.Add("Motif_De_Depart", "Motif De Depart");
            tables.Add("Natutre_De_Contrat", "Nature De Contrat");
            tables.Add("Pays", "Pays");
            tables.Add("Service", "Service");
            tables.Add("Situation_Familiale", "Situation Familiale");
            tables.Add("Societe", "Societe");
            tables.Add("Type_De_Contrat", "Type De Contrat");
            tables.Add("Type_De_Depart", "Type De Depart");
     

            // Remplir le ComboBox avec les valeurs du dictionnaire
            foreach (var kvp in tables)
            {
                // Ajoutez l'élément au comboBox avec la valeur affichée et la clé en tant qu'objet associé
                comboBox2.Items.Add(new ComboBoxItem { Key = kvp.Key, Value = kvp.Value });
            }

            // Vous devez définir l'affichage de la valeur dans le ComboBox
            comboBox2.DisplayMember = "Value";  // Indique quelle propriété afficher dans l'interface utilisateur
            comboBox2.ValueMember = "Key";      // Indique quelle propriété correspond à la clé pour mémoire








        }
        public class ComboBoxItem
        {
            public string Key { get; set; }  // Stocker la clé (pour la mémoire)
            public string Value { get; set; }  // Stocker la valeur (pour l'affichage)

            // Optionnel : Override ToString() pour afficher la valeur dans ComboBox directement
            public override string ToString()
            {
                return Value;
            }
        }


        public int InsertTable_(string MyTable, Dictionary<string, object> MyData, int nbr_bdd)
        {
            CleanData(MyData);
            int retValue = -1;
            string queryString = "INSERT INTO " + MyTable + "(";
            var keys = new List<string>(MyData.Keys);
            int nbr = 0;

            foreach (string key in keys)
            {
                queryString += (nbr == keys.Count - 1) ? key + ")" : key + ",";
                nbr++;
            }

            queryString += " VALUES(";
            var vals = new List<object>(MyData.Values);
            int cnt = 0;

            foreach (var val in vals)
            {
                string cleanedVal = (val is string strVal) ? strVal.Replace("''", "'").Replace(@"\s+", " ")
                                : (val is byte[] byteArray) ? BitConverter.ToString(byteArray)
                                : val.ToString();

                queryString += (cnt == vals.Count - 1) ? cleanedVal + ")" : cleanedVal + ",";
                cnt++;
            }
        
            Clipboard.SetText(queryString);
            SqlCommand command = null;
            SqlConnection cnn = null;

            if (nbr_bdd == 1)
                cnn = cnn1;
            else if (nbr_bdd == 2)
                cnn = cnn2;
            else if (nbr_bdd == 3)
                cnn = cnn3;
            else if (nbr_bdd == 0)
                cnn = cnn0;
          try
            {
                // Ouvrir la connexion
                if (cnn.State == ConnectionState.Closed)
                {
                    cnn.Open();
                }

                command = new SqlCommand(queryString, cnn0);
                int result = command.ExecuteNonQuery();
                MessageBox.Show("!!!");
                retValue = 1;  // Success
            }
            catch (Exception ex)
            {
                retValue = -1; // Error
                MessageBox.Show(ex.Message);
                Clipboard.SetText(ex.Message);
            }
            finally
            {
                // Fermer la connexion
                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }

                command.Dispose();
            }

            return retValue;
        }

        public int InsertTable(string MyTable, Dictionary<string, object> MyData, int nbr_bdd)
        {
            CleanData(MyData);

            int retValue = 2;
            //string queryString = "INSERT INTO " + MyTable + "(";
            //var keys = new List<string>(MyData.Keys);
            //int nbr = 0;
            //foreach (string key in keys)
            //{
            //    if (nbr == keys.Count - 1)
            //        queryString = queryString + key + ")";
            //    else
            //        queryString = queryString + key + ",";
            //    nbr++;
            //}

            //queryString = queryString + " VALUES(";

            //var vals = new List<object>(MyData.Values);
            //int cnt = 0;

            //foreach (var val in vals)
            //{
            //    string cleanedVal;

            //    if (val is string strVal)
            //    {
            //        // Nettoyage spécifique pour les chaînes
            //        cleanedVal = strVal.Replace("''", "'"); // Remplacer les guillemets doubles par un guillemet simple
            //        cleanedVal = Regex.Replace(cleanedVal, @"\s+", " "); // Remplacer les espaces multiples par un espace simple
            //    }
            //    else if (val is byte[] byteArray)
            //    {
            //        // Traiter les tableaux de bytes différemment
            //        cleanedVal = BitConverter.ToString(byteArray); // Convertir les bytes en une chaîne
            //    }
            //    else
            //    {
            //        // Convertir les autres types en chaîne
            //        cleanedVal = val.ToString();
            //    }

            //    // Append the cleaned value to the query string
            //    if (cnt == vals.Count - 1)
            //    {
            //        queryString += cleanedVal + ")"; // End of the VALUES list
            //    }
            //    else
            //    {
            //        queryString += cleanedVal + ","; // Add comma for next value
            //    }
            //    cnt++;
            //}
            ////MessageBox.Show(queryString);
            //Clipboard.SetText(queryString);


            //SqlCommand command = null;

            //// ////MessageBox.Show(queryString);
            //if (nbr_bdd == 1)
            //{
            //    command = new SqlCommand(queryString, cnn1);
            //}
            //else if (nbr_bdd == 2)
            //{
            //    command = new SqlCommand(queryString, cnn2);
            //}
            //else if (nbr_bdd == 3)
            //{
            //    command = new SqlCommand(queryString, cnn3);
            //}

            //else if (nbr_bdd == 0)
            //{
            //    command = new SqlCommand(queryString, cnn0);
            //}




            //try
            //{
            //    // OpenConnection(nbr_bdd);
            //    int result = command.ExecuteNonQuery();
            //    retValue = 1;

            //}
            //catch (Exception ex)
            //{
            //    retValue = -1;
            //    //MessageBox.Show(ex.Message);

            //}

           

            //finally
            //{
            //    command.Dispose();
            //    // CloseConnection(nbr_bdd);
            //}


            return retValue;
        }
        public int UpdateTable(string MyTable, Dictionary<string, object> MyData, int nbr_bdd, string identifierColumn, string identifierValue)
        {
            CleanData(MyData);
            int retVal = -1;
            if (MyTable != "T_SAL")
               
            {
                //Check if it exists or not in this table before modification
                //MessageBox.Show("4666");

                string queryString_ = "SELECT COUNT(*) AS PIECE  FROM  " + MyTable + " WHERE LTRIM(RTRIM(NumSalarie)) = '" + identifierValue + "' ";
                string message_ = $"Requête SQL : {queryString_}\nParamètre : NumSalarie = {identifierValue}";
                //MessageBox.Show(message_, "Détails de la Requête");

                Clipboard.SetText(message_);
                SqlCommand command_ = null;
                if (nbr_bdd == 1)
                {
                    command_ = new SqlCommand(queryString_, cnn1);

                }
                else if (nbr_bdd == 2)
                {
                    command_ = new SqlCommand(queryString_, cnn2);

                }
                else if (nbr_bdd == 3)
                {
                    command_ = new SqlCommand(queryString_, cnn3);

                }
                else if (nbr_bdd == 0)
                {
                    command_ = new SqlCommand(queryString_, cnn0);
                }


                //  OpenConnection(nbr_bdd);
                try
                {
                    var firstColumn = command_.ExecuteScalar();

                    if (firstColumn != null)
                    {
                        retVal = (int)firstColumn;
                        ////MessageBox.Show("Nombre de correspondance : "+retVal.ToString());
                    }
                }
                catch (SqlException sqlEx)
                {
                    // Gérer les exceptions SQL spécifiques
                    //MessageBox.Show($"Erreur SQL : {sqlEx.Message}");
                    
                }
                catch (Exception ex)
                {
                    // Gérer les autres exceptions
                    //MessageBox.Show($"Erreur : {ex.Message}");
                }

                command_.Dispose();
               // If it does not exist in the associated table.
                if (retVal == -1)
                {
                    //creat new 
                    int result = InsertTable(MyTable, MyData, nbr_bdd);
                    //MessageBox.Show(result != -1 ? "Les données pour " + MyTable + " ont été bien insérées." : "Erreur lors de l'insertion dans " + MyTable + ".");

                }
                // If it exist in the associated table.
                else
                {
                    if (MyData.ContainsKey("CodeFonctionEntreprise"))
                    {
                        int codefontionentreprise = CountRows("T_FONCTIONENTREPRISE", "Intitule", "'" + MyData["CodeFonctionEntreprise"].ToString() + "'", nbr_bdd);

                        if (codefontionentreprise == -1 || codefontionentreprise == 0)
                        {
                            //MessageBox.Show("code fonction d'entreprise n'existe pas en tabel ,t_fonctionentreprise");
                            //si code fonction d'entreprise n'existe pas en tabel ,t_fonctionentreprise
                            var data = new Dictionary<string, object>();
                            data.Add("Intitule", MyData["CodeFonctionEntreprise"]);
                            //incrementer le dernniere code
                            data.Add("Code", IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd));
                            int resultat = InsertTable("T_FONCTIONENTREPRISE", data, nbr_bdd);
                            //MessageBox.Show(resultat != -1 ? "Les données pour T_FONCTIONENTREPRIS ont été bien insérées le code est" + IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd) + "." : "Erreur lors de l'insertion dans T_FONCTIONENTREPRIS.");
                            MyData["CodeFonctionEntreprise"] = IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd).ToString();
                        }
                        else
                        {
                            //recuperre le code 
                            //MessageBox.Show("code fonction d'entreprise existe deja en tabel t_fonctionentreprise");
                            MyData["CodeFonctionEntreprise"] = GetIdTable("T_FONCTIONENTREPRISE", "Code", "Intitule", "'" + MyData["CodeFonctionEntreprise"].ToString() + "'", nbr_bdd);
                        }
                    }
                    var setClauses = MyData.Select(kvp =>
                    {
                        string key = kvp.Key;
                        object value = kvp.Value;

                        // Nettoyage et échappement des valeurs pour éviter les problèmes de syntaxe SQL
                        string valueString = null;

                        // Vérifier si la valeur est null ou équivalente à "NULL"
                        if (value == null || value.ToString().Trim().ToUpper() == "NULL")
                        {
                            // Ne rien faire si la valeur est nulle ou "NULL"
                            return null;
                        }

                        if (value is string stringValue)
                        {
                            // Échapper les apostrophes et remplacer les espaces multiples
                            valueString = $"{stringValue.Replace("''", "'")}"; // Échappement des apostrophes avec guillemets simples
                            valueString = Regex.Replace(valueString, @"\s+", " "); // Remplacer les espaces multiples par un espace simple
                        }
                        else if (value is byte[] byteArray)
                        {
                            // Convertir les tableaux de bytes en une chaîne
                            valueString = $"{BitConverter.ToString(byteArray).Replace("-", "")}"; // Formatage des bytes
                        }
                        else if (value is DateTime dateTimeValue)
                        {
                            // Convertir les DateTime en format SQL standard sans guillemets
                            valueString = $"{dateTimeValue:yyyy-MM-dd}";
                        }

                        // Si valueString est null, ne pas inclure cette clause
                        if (valueString == null)
                        {
                            return null;
                        }

                        return $"{key} = {valueString}";
                    }).Where(clause => !string.IsNullOrWhiteSpace(clause));

                    string queryString = $"UPDATE {MyTable} SET {string.Join(", ", setClauses)} WHERE {identifierColumn} = '{identifierValue}';";



                    //// Prepare parameter details for display
                    //var parameterDetails = new List<string>();
                    //foreach (string key in keys)
                    //{
                    //    if (MyData.TryGetValue(key, out var value))
                    //    {
                    //        string formattedValue = value == null ? "NULL" : $"{value}";
                    //        parameterDetails.Add($"{key}: {formattedValue}");
                    //    }
                    //}

                    //// Add the identifier parameter
                    //parameterDetails.Add($"{identifierColumn}: {identifierValue}'");

                    //// Combine parameter details into a single string
                    //string parameterString = string.Join(", ", parameterDetails);

                    //// Construct the final message
                    //string message = $"SQL Query: {queryString}\nParameters:\n{parameterString}";

                    // Display the message box and copy to clipboard
                    //MessageBox.Show(queryString + "Query Details");
                    //Clipboard.SetText(queryString);



                    SqlCommand command = new SqlCommand();

                    // ////MessageBox.Show(queryString);
                    if (nbr_bdd == 1)
                    {
                        command = new SqlCommand(queryString, cnn1);
                    }
                    else if (nbr_bdd == 2)
                    {
                        command = new SqlCommand(queryString, cnn2);
                    }
                    else if (nbr_bdd == 3)
                    {
                        command = new SqlCommand(queryString, cnn3);
                    }

                    else if (nbr_bdd == 0)
                    {
                        command = new SqlCommand(queryString, cnn0);
                    }


                    foreach (var kvp in MyData)
                    {
                        object value = kvp.Value;
                        if (value == null || value.ToString().Trim().ToLower() == "null")
                        {
                            // Utiliser DBNull.Value pour les valeurs null
                            command.Parameters.Add(new SqlParameter("@" + kvp.Key, DBNull.Value));
                        }
                        else if (value is string stringValue)
                        {
                            // Nettoyage des valeurs de chaîne
                            stringValue = stringValue.Replace("''", "'");
                            stringValue = Regex.Replace(stringValue, @"\s+", " ");
                            command.Parameters.Add(new SqlParameter("@" + kvp.Key, stringValue));
                        }
                        else
                        {
                            // Ajouter les autres types de données tels que int, DateTime, etc.
                            command.Parameters.Add(new SqlParameter("@" + kvp.Key, value));
                        }
                    }

                    if (nbr_bdd == 1)
                    {
                        command.Connection = cnn1;
                    }
                    else if (nbr_bdd == 2)
                    {
                        command.Connection = cnn2;
                    }
                    else if (nbr_bdd == 3)
                    {
                        command.Connection = cnn3;
                    }
                    else if (nbr_bdd == 0)
                    {
                        command.Connection = cnn0 ;
                    }



                    try
                    {
                        // OpenConnection(nbr_bdd);
                        int affectedRows = command.ExecuteNonQuery(); // Utilisez ExecuteNonQuery pour les requêtes de mise à jour
                        return affectedRows;
                    }

                    catch (SqlException sqlEx)
                    {
                        // Afficher les détails spécifiques à l'erreur SQL
                        //MessageBox.Show($"Erreur SQL : {sqlEx.Message}");
                        return -1; // Vous pouvez retourner une valeur spécifique en cas d'erreur
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show($"Erreur SQL : {ex.Message}");
                        return -1; // Vous pouvez retourner une valeur spécifique en cas d'erreur
                    }
                    finally
                    {
                        command.Dispose();
                        // CloseConnection(nbr_bdd);
                    }
                    ////MessageBox.Show(queryString);
                    //Clipboard.SetText(queryString);


                }




            }
            else
            {
                if (MyData.ContainsKey("CodeFonctionEntreprise"))
                {
                    int codefontionentreprise = CountRows("T_FONCTIONENTREPRISE", "Intitule", "'" + MyData["CodeFonctionEntreprise"].ToString() + "'", nbr_bdd);

                    if (codefontionentreprise == -1 || codefontionentreprise == 0)
                    {
                        //MessageBox.Show("code fonction d'entreprise n'existe pas en tabel ,t_fonctionentreprise");
                        //si code fonction d'entreprise n'existe pas en tabel ,t_fonctionentreprise
                        var data = new Dictionary<string, object>();
                        data.Add("Intitule", MyData["CodeFonctionEntreprise"]);
                        //incrementer le dernniere code
                        data.Add("Code", IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd));
                        int resultat = InsertTable("T_FONCTIONENTREPRISE", data, nbr_bdd);
                        //MessageBox.Show(resultat != -1 ? "Les données pour T_FONCTIONENTREPRIS ont été bien insérées le code est" + IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd) + "." : "Erreur lors de l'insertion dans T_FONCTIONENTREPRIS.");
                        MyData["CodeFonctionEntreprise"] = IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd).ToString();
                    }
                    else
                    {
                        //recuperre le code 
                        //MessageBox.Show("code fonction d'entreprise existe deja en tabel t_fonctionentreprise+++++++++++++++" + GetIdTable("T_FONCTIONENTREPRISE", "Code", "Intitule", MyData["CodeFonctionEntreprise"].ToString(), nbr_bdd));
                        MyData["CodeFonctionEntreprise"] = GetIdTable("T_FONCTIONENTREPRISE", "Code", "Intitule", "'" + MyData["CodeFonctionEntreprise"].ToString() + "'", nbr_bdd);

                    }
                }
                var setClauses = MyData.Select(kvp =>
                {
                    string key = kvp.Key;
                    object value = kvp.Value;

                    // Nettoyage et échappement des valeurs pour éviter les problèmes de syntaxe SQL
                    string valueString = null;

                    // Vérifier si la valeur est null ou équivalente à "NULL"
                    if (value == null || value.ToString().Trim().ToUpper() == "NULL")
                    {
                        // Ne rien faire si la valeur est nulle ou "NULL"
                        return null;
                    }

                    if (value is string stringValue)
                    {
                        // Échapper les apostrophes et remplacer les espaces multiples
                        valueString = $"{stringValue.Replace("''", "'")}"; // Échappement des apostrophes avec guillemets simples
                        valueString = Regex.Replace(valueString, @"\s+", " "); // Remplacer les espaces multiples par un espace simple
                    }
                    else if (value is byte[] byteArray)
                    {
                        // Convertir les tableaux de bytes en une chaîne
                        valueString = $"{BitConverter.ToString(byteArray).Replace("-", "")}"; // Formatage des bytes
                    }
                    else if (value is DateTime dateTimeValue)
                    {
                        // Convertir les DateTime en format SQL standard sans guillemets
                        valueString = $"{dateTimeValue:yyyy-MM-dd}";
                    }

                    // Si valueString est null, ne pas inclure cette clause
                    if (valueString == null)
                    {
                        return null;
                    }

                    return $"{key} = {valueString}";
                }).Where(clause => !string.IsNullOrWhiteSpace(clause));
                string queryString = $"UPDATE {MyTable} SET {string.Join(", ", setClauses)} WHERE {identifierColumn} = '{identifierValue}';";

                //// Prepare parameter details for display
                //var parameterDetails = new List<string>();
                //foreach (string key in keys)
                //{
                //    if (MyData.TryGetValue(key, out var value))
                //    {
                //        string formattedValue = value == null ? "NULL" : $"{value}";
                //        parameterDetails.Add($"{key}: {formattedValue}");
                //    }
                //}

                //// Add the identifier parameter
                //parameterDetails.Add($"{identifierColumn}: '{identifierValue}'");

                //// Combine parameter details into a single string
                //string parameterString = string.Join(", ", parameterDetails);

                //// Construct the final message
                //string message = $"SQL Query: {queryString}\nParameters:\n{parameterString}";

                //// Display the message box and copy to clipboard
                // //MessageBox.Show(queryString, "Query Details");
                //Clipboard.SetText(queryString);



                SqlCommand command = new SqlCommand();

                // ////MessageBox.Show(queryString);
                if (nbr_bdd == 1)
                {
                    command = new SqlCommand(queryString, cnn1);
                }
                else if (nbr_bdd == 2)
                {
                    command = new SqlCommand(queryString, cnn2);
                }
                else if (nbr_bdd == 3)
                {
                    command = new SqlCommand(queryString, cnn3);
                }
                else if (nbr_bdd == 0)
                {
                    command = new SqlCommand(queryString, cnn0);
                }


                //foreach (var kvp in MyData)
                //{
                //    object value = kvp.Value;
                //    if (value == null || value.ToString().Trim().ToLower() == "null")
                //    {
                //        // Utiliser DBNull.Value pour les valeurs null
                //        command.Parameters.Add(new SqlParameter("@" + kvp.Key, DBNull.Value));
                //    }
                //    else if (value is string stringValue)
                //    {
                //        // Nettoyage des valeurs de chaîne
                //        stringValue = stringValue.Replace("''", "'");
                //        stringValue = Regex.Replace(stringValue, @"\s+", " ");
                //        command.Parameters.Add(new SqlParameter("@" + kvp.Key, stringValue));
                //    }
                //    else
                //    {
                //        // Ajouter les autres types de données tels que int, DateTime, etc.
                //        command.Parameters.Add(new SqlParameter("@" + kvp.Key, value));
                //    }
                //}

                if (nbr_bdd == 1)
                {
                    command.Connection = cnn1;
                }
                else if (nbr_bdd == 2)
                {
                    command.Connection = cnn2;
                }
                else if (nbr_bdd == 3)
                {
                    command.Connection = cnn3;
                }
                else if (nbr_bdd == 0)
                {
                    command.Connection = cnn0;
                }

                try
                {
                    // OpenConnection(nbr_bdd);
                    int affectedRows = command.ExecuteNonQuery(); // Utilisez ExecuteNonQuery pour les requêtes de mise à jour
                    return affectedRows;
                }
                catch (SqlException sqlEx)
                {
                    // Affichage détaillé de l'erreur SQL
                    //MessageBox.Show($"Erreur SQL : {sqlEx.Message}");
                    return -1; // Code d'erreur spécifique pour SQL
                }
                catch (Exception ex)
                {
                    // Affichage détaillé des autres erreurs
                    //MessageBox.Show($"Erreur : {ex.Message}");
                    return -1; // Code d'erreur général pour d'autres exceptions
                }
                finally
                {
                    command.Dispose();
                    // CloseConnection(nbr_bdd);
                }
                ////MessageBox.Show(queryString);
                //Clipboard.SetText(queryString);
            }






            return -1;


        }
        public string CodeAdd(string valeur)
        {

            char[] separators = new char[] { '_', '-', ' ' };
            string[] words = valeur.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            // Initialiser une nouvelle chaîne pour stocker les premiers caractères
            string newCode = "";

            // Parcourir les mots et ajouter le premier caractère de chaque mot à la chaîne newCode
            foreach (string word in words)
            {
                if (word.Length > 0)
                {
                    newCode += word[0]; // Ajouter le premier caractère
                }
            }

            // Retourner la nouvelle chaîne
            return newCode;
        }
        public string IncrementCode(string table, string column, int nbrBdd)
        {

            string retVal="";

            string queryString = $"SELECT MAX(CAST({column} AS INT)) FROM {table}";

            string message = $"Requête SQL : {queryString}";
            ////MessageBox.Show(message, "Détails de la Requête");

            Clipboard.SetText(message);

            SqlCommand command = null;
            if (nbrBdd == 1)
            {
                command = new SqlCommand(queryString, cnn1);

            }
            else if (nbrBdd == 2)
            {
                command = new SqlCommand(queryString, cnn2);

            }
            else if (nbrBdd == 3)
            {
                command = new SqlCommand(queryString, cnn3);

            }
            else if (nbrBdd == 0)
            {
                command = new SqlCommand(queryString, cnn0);

            }


            // Ajout du paramètre

            //  OpenConnection(nbr_bdd);
            try
            {
                object result = command.ExecuteScalar();
                int lastCode = 0;
                // Check if result is null (i.e., table is empty) and set lastCode accordingly
                if (result != DBNull.Value)
                {
                    lastCode = Convert.ToInt32(result);
                }
               

                // Increment the last code
                retVal = (lastCode + 1).ToString();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            command.Dispose();
        
            return retVal;
        }

        public int CountRows(string tableName, string columnName, string value, int nbrBdd)
        {
            int firstColumn;
            string queryString = $"SELECT COUNT(*)  as count_ FROM {tableName} WHERE {columnName} = {value.Replace("''", "'")}";

            // Affichage de la requête dans une MessageBox
            string message = $"Requête SQL : {queryString}\nParamètre : {columnName} = {value}";
            //MessageBox.Show(message, "Détails de la Requête");

            //// Copie de la requête dans le presse-papiers
            Clipboard.SetText(message);
            SqlCommand command = null;
            if (nbrBdd == 1)
            {
                command = new SqlCommand(queryString, cnn1);

            }
            else if (nbrBdd == 2)
            {
                command = new SqlCommand(queryString, cnn2);

            }
            else if (nbrBdd == 3)
            {
                command = new SqlCommand(queryString, cnn3);

            }

            else if (nbrBdd == 0)
            {
                command = new SqlCommand(queryString, cnn0);

            }
            // OpenConnection(nbr_bdd);
            try
            {
               firstColumn =(int)command.ExecuteScalar();

              
            }
            catch (Exception ex)
            {
                //MessageBox.Show("error countrows");
                firstColumn = -1;

            }

            command.Dispose();
            //  CloseConnection(nbr_bdd);
            return firstColumn;



        
        }


        public int InsertTable_sal(string MyTable, Dictionary<string, object> MyData, int nbr_bdd)
        {
            CleanData(MyData);
            ////MessageBox.Show("insertTable_sal" + MyData["CodeFonctionEntreprise"].ToString());
            //int codefontionentreprise = CountRows("T_FONCTIONENTREPRISE", "Intitule", "'" + MyData["CodeFonctionEntreprise"].ToString() + "'", nbr_bdd);
            ////MessageBox.Show(codefontionentreprise.ToString()+"22");

            //if (codefontionentreprise == -1 || codefontionentreprise == 0)
            //{
            //    //MessageBox.Show("code fonction d'entreprise n'existe pas en tabel ,t_fonctionentreprise");
            //    //si code fonction d'entreprise n'existe pas en tabel ,t_fonctionentreprise
            //    var data = new Dictionary<string, object>();
            //    data.Add("Intitule", MyData["CodeFonctionEntreprise"]);
            //    //incrementer le dernniere code
            //    data.Add("Code", IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd));
            //    int resultat = InsertTable("T_FONCTIONENTREPRISE", data, nbr_bdd);
            //    //MessageBox.Show(resultat != -1 ? "Les données pour T_FONCTIONENTREPRIS ont été bien insérées le code est" + IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd) + "." : "Erreur lors de l'insertion dans T_FONCTIONENTREPRIS."); ; ;
            //    MyData["CodeFonctionEntreprise"] = IncrementCode("T_FONCTIONENTREPRISE", "Code", nbr_bdd);
            //}
            //else
            //{
            //    //recuperre le code 
            //    //MessageBox.Show("code fonction d'entreprise existe deja en tabel t_fonctionentreprise");
            //    MyData["CodeFonctionEntreprise"] = GetIdTable("T_FONCTIONENTREPRISE", "Code", "Intitule", "'" + MyData["CodeFonctionEntreprise"].ToString() + "'", nbr_bdd);
            //}
            //Boolean retValue = true;

            //string queryString = "INSERT INTO " + MyTable + "(";
            //var keys = new List<string>(MyData.Keys);
            //int nbr = 0;
            //foreach (string key in keys)
            //{
            //    if (nbr == keys.Count - 1)
            //        queryString = queryString + key + ")";
            //    else
            //        queryString = queryString + key + ",";
            //    nbr++;
            //}

            //queryString = queryString + " VALUES(";

            ////MessageBox.Show(queryString);
            //Clipboard.SetText(queryString);


            //var vals = new List<object>(MyData.Values);
            //int cnt = 0;

            //foreach (var val in vals)
            //{
            //    string cleanedVal;

            //    if (val is string strVal)
            //    {
            //        // Nettoyage spécifique pour les chaînes
            //        cleanedVal = strVal.Replace("''", "'"); // Remplacer les guillemets doubles par un guillemet simple
            //        cleanedVal = Regex.Replace(cleanedVal, @"\s+", " "); // Remplacer les espaces multiples par un espace simple
            //    }
            //    else if (val is byte[] byteArray)
            //    {
            //        // Traiter les tableaux de bytes différemment
            //        cleanedVal = BitConverter.ToString(byteArray); // Convertir les bytes en une chaîne
            //    }
            //    else
            //    {
            //        // Convertir les autres types en chaîne
            //        cleanedVal = val.ToString();
            //    }

            //    // Append the cleaned value to the query string
            //    if (cnt == vals.Count - 1)
            //    {
            //        queryString += cleanedVal + ")"; // End of the VALUES list
            //    }
            //    else
            //    {
            //        queryString += cleanedVal + ","; // Add comma for next value
            //    }
            //    cnt++;
            //}



            //SqlCommand command = null;

            ////MessageBox.Show(queryString);
            //Clipboard.SetText(queryString);

            //if (nbr_bdd == 1)
            //{
            //    command = new SqlCommand(queryString, cnn1);
            //}
            //else if (nbr_bdd == 2)
            //{
            //    command = new SqlCommand(queryString, cnn2);
            //}
            //else if (nbr_bdd == 3)
            //{
            //    command = new SqlCommand(queryString, cnn3);
            //}
            //else if (nbr_bdd == 0)
            //{
            //    command = new SqlCommand(queryString, cnn0);
            //}




            //try
            //{
            //    object result = command.ExecuteScalar();
            //    if (result != null)
            //    {
            //        int numSalarie = Convert.ToInt32(result); // Convert to integer safely
            //        return numSalarie;                                           // Use numSalarie as needed
            //    }

            //}
            //catch (Exception ex)
            //{
            //    retValue = false;
            //    //MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    command.Dispose();
            //    // CloseConnection(nbr_bdd);
            //}


            return 2;
            //MessageBox.Show("fin insertTable_sal");
        }

        //public string PrepareInsertTable(string MyTable, Dictionary<string, object> MyData)
        //{
        //    string queryString = "INSERT INTO " + MyTable + "(";
        //    var keys = new List<string>(MyData.Keys);
        //    int nbr = 0;
        //    foreach (string key in keys)
        //    {
        //        if (nbr == keys.Count - 1)
        //            queryString = queryString + key + ")";
        //        else
        //            queryString = queryString + key + ",";
        //        nbr++;

        //    }

        //    queryString = queryString + " VALUES(";

        //    var vals = new List<string>(MyData.Values);
        //    int cnt = 0;
        //    foreach (string val in vals)
        //    {
        //        if (cnt == vals.Count - 1)
        //            queryString = queryString + "'" + val.Replace("'", "''") + "'" + ")";
        //        else
        //            queryString = queryString + "'" + val.Replace("'", "''") + "'" + ",";
        //        cnt++;
        //    }

        //    ////MessageBox.Show(queryString);
        //    return queryString;
        //}

        public Boolean CloseConnection(int nbr_bdd)
        {
            if (cnn0.State == ConnectionState.Open)
            {
                try
                {
                    if (nbr_bdd == 1)
                    {
                        cnn1.Close();
                    }
                    else if (nbr_bdd == 2)
                    {
                        cnn2.Close();
                    }
                    else if (nbr_bdd == 3)

                    {
                        cnn3.Close();

                    }
                    else if (nbr_bdd == 0)

                    {
                        cnn0.Close();

                    }
                    ConnectionStatus = false;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Impossible de fermer la connexion.");
                    //MessageBox.Show(ex.Message);
                }
            }

            return ConnectionStatus;
        }

        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp,
                                            255, this.path);
            return temp.ToString();

        }

        //public Boolean //MessageBox.Show(string message)
        //{
        //    Boolean retVal = true;
        //    FileInfo Fi = new FileInfo(LogPath);
        //    string LogPathArch = AppDomain.CurrentDomain.BaseDirectory + @"Log_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".txt";

        //    if (LogLevel == "1")
        //    {
        //        System.IO.File.AppendAllLines(LogPath, new string[] { DateTime.Now + " : " + message });
        //    }
        //    if (Fi.Length > 20000000)
        //    {
        //        File.Move(LogPath, LogPathArch);
        //    }

        //    return retVal;
        //}

        public Boolean OpenConnection(int nbr_bdd)
        {
            ConnectionStatus = false;


            try
            {
                if (nbr_bdd == 1)
                {
                    cnn1.Open();
                }
                else if (nbr_bdd == 2)
                {
                    cnn2.Open();
                }
                else if (nbr_bdd == 3)
                {
                    cnn3.Open();
                }
                else if (nbr_bdd == 0)
                {
                    cnn0.Open();
                }

                ConnectionStatus = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Impossible de se connecter à la base veuiller vérifier les paramètres de connexion." + ex.Message);
                ConnectionStatus = false;
            }

            return ConnectionStatus;
        }
        public bool EmployeeExistsByMatricule(string Matricule, int nbr_bdd)
        {
            //MessageBox.Show("****");
            int retVal = 0;

            string queryString = "SELECT COUNT(*) AS PIECE  FROM  T_SAL WHERE LTRIM(RTRIM(MatriculeSalarie)) = '" + Matricule.Trim() + "' ";

            SqlCommand command = null;
            if (nbr_bdd == 1)
            {
                command = new SqlCommand(queryString, cnn1);

            }
            else if (nbr_bdd == 2)
            {
                command = new SqlCommand(queryString, cnn2);

            }
            else if (nbr_bdd == 3)
            {
                command = new SqlCommand(queryString, cnn3);

            }
            else if (nbr_bdd == 0)
            {
                command = new SqlCommand(queryString, cnn0);

            }

            //  OpenConnection(nbr_bdd);
            try
            {
                var firstColumn = command.ExecuteScalar();

                if (firstColumn != null)
                {
                    retVal = (int)firstColumn;
                    ////MessageBox.Show("Nombre de correspondance : "+retVal.ToString());
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            command.Dispose();
            //CloseConnection(nbr_bdd);
            if (retVal != 0)
                return true;
            else
                return false;
        }

        public enum Mode
        {
            AlphaNumeric = 1,
            Alpha = 2,
            Numeric = 3
        }

        public static string Increment(string text, Mode mode)
        {
            int numeric;
            if (text.Length > 0 && int.TryParse(text.Substring(text.Length - 1, 1), out numeric))
            {
                text = text + "0";
            }
            var textArr = text.ToCharArray();


            // Add legal characters
            var characters = new List<char>();

            if (mode == Mode.AlphaNumeric || mode == Mode.Numeric)
                for (char c = '0'; c <= '9'; c++)
                    characters.Add(c);

            if (mode == Mode.AlphaNumeric || mode == Mode.Alpha)
                for (char c = 'a'; c <= 'z'; c++)
                    characters.Add(c);

            // Loop from end to beginning
            for (int i = textArr.Length - 1; i >= 0; i--)
            {
                if (textArr[i] == characters.Last())
                {
                    textArr[i] = characters.First();
                }
                else
                {
                    textArr[i] = characters[characters.IndexOf(textArr[i]) + 1];
                    break;
                }
            }

            return new string(textArr);
        }

        public string GetLastMatricule(int nbr_bdd)
        {
            string retVal = "";
            string queryString = "SELECT TOP 1 MatriculeSalarie AS PIECE  FROM  T_SAL ORDER BY SA_CompteurNumero DESC";
            SqlCommand command = null;
            if (nbr_bdd == 1)
            {
                command = new SqlCommand(queryString, cnn1);

            }
            else if (nbr_bdd == 2)
            {
                command = new SqlCommand(queryString, cnn2);

            }
            else if (nbr_bdd == 3)
            {
                command = new SqlCommand(queryString, cnn3);

            }
            
            else if (nbr_bdd == 0)
            {
                command = new SqlCommand(queryString, cnn0);

            }
            // OpenConnection(nbr_bdd);
            try
            {
                var firstColumn = command.ExecuteScalar();

                if (firstColumn != null)
                {
                    retVal = firstColumn.ToString();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            command.Dispose();
            //  CloseConnection(nbr_bdd);
            return retVal;
        }
        //public Dictionary<string, string> InitiateEmployee()
        //{
        //    var MydataEmployee = new Dictionary<string, string>();

        //    MydataEmployee.Add("MatriculeSalarie", "");
        //    MydataEmployee.Add("Civilite", "");
        //    MydataEmployee.Add("Nom", "");
        //    MydataEmployee.Add("NomJeuneFille", "");
        //    MydataEmployee.Add("Prenom", "");
        //    MydataEmployee.Add("Prenom2", "");
        //    MydataEmployee.Add("Confidentialite", "");
        //    MydataEmployee.Add("Rue1", "");
        //    MydataEmployee.Add("Rue2", "");
        //    MydataEmployee.Add("Commune", "");
        //    MydataEmployee.Add("BureauDistributeur", "");
        //    MydataEmployee.Add("CodePostal", "");
        //    MydataEmployee.Add("CodePays", "");
        //    MydataEmployee.Add("Telephone", "");
        //    MydataEmployee.Add("Rue12", "");
        //    MydataEmployee.Add("Rue22", "");
        //    MydataEmployee.Add("Commune2", "");
        //    MydataEmployee.Add("BureauDistributeur2", "");
        //    MydataEmployee.Add("CodePostal2", "");
        //    MydataEmployee.Add("CodePays2", "");
        //    MydataEmployee.Add("Telephone2", "");
        //    MydataEmployee.Add("ChoixSurAdresse", "");
        //    MydataEmployee.Add("CommuneNaissance", "");
        //    MydataEmployee.Add("CodeCommuneNaissance", "");
        //    MydataEmployee.Add("NoBulletinModele", "");
        //    MydataEmployee.Add("ModeDePaiement", "");
        //    MydataEmployee.Add("EtatPaie", "");
        //    MydataEmployee.Add("Cloture", "");
        //    MydataEmployee.Add("DateDePaie", "");
        //    MydataEmployee.Add("DateDeCloture", "");
        //    MydataEmployee.Add("CompteAuxiliaire", "");
        //    MydataEmployee.Add("CumulsReposCompensateur", "");
        //    MydataEmployee.Add("ResteAPrendre", "");
        //    MydataEmployee.Add("AcquisPrecedent", "");
        //    MydataEmployee.Add("AcquisEnCours", "");
        //    MydataEmployee.Add("CumulsBrutPrecedent", "");
        //    MydataEmployee.Add("BrutCPPrecedentBis", "");
        //    MydataEmployee.Add("MoisAncienneteSociete", "");
        //    MydataEmployee.Add("MoisAncienneteEtablissement", "");
        //    MydataEmployee.Add("MoisAnciennetePoste", "");
        //    MydataEmployee.Add("DateDebutConges1", "");
        //    MydataEmployee.Add("DateFinConges1", "");
        //    MydataEmployee.Add("DateDebutConges2", "");
        //    MydataEmployee.Add("DateFinConges2", "");
        //    MydataEmployee.Add("DateDebutConges3", "");
        //    MydataEmployee.Add("DateFinConges3", "");
        //    MydataEmployee.Add("DroitSupplementaire", "");
        //    MydataEmployee.Add("NbSamedis", "");
        //    MydataEmployee.Add("MoisDeClotureDesConges", "");
        //    MydataEmployee.Add("TypeDeVentilation", "");
        //    MydataEmployee.Add("NbFichesENF", "");
        //    MydataEmployee.Add("NbFichesMUL", "");
        //    MydataEmployee.Add("ChoixPaiement", "");
        //    MydataEmployee.Add("BanqueActive", "");
        //    MydataEmployee.Add("GHRTypeConjoint", "");
        //    MydataEmployee.Add("GHRTypeDAstreinte", "");
        //    MydataEmployee.Add("GHRTypeDeSalarie", "");
        //    MydataEmployee.Add("TravailleurHandicape", "");
        //    MydataEmployee.Add("SuiteAccident", "");
        //    MydataEmployee.Add("TauxInvalidite", "");
        //    MydataEmployee.Add("CotorepDate", "");
        //    MydataEmployee.Add("SituationMilitaire", "");
        //    MydataEmployee.Add("DateDebutServiceMil", "");
        //    MydataEmployee.Add("DateFinServiceMil", "");
        //    MydataEmployee.Add("Commentaire1", "");
        //    MydataEmployee.Add("Commentaire2", "");
        //    MydataEmployee.Add("AdresseOrganismeETA", "");
        //    MydataEmployee.Add("TypeDuBulletin", "");
        //    MydataEmployee.Add("InfoDateDeNaissance", "");
        //    MydataEmployee.Add("AccordSalarie_AdCourriel", "");
        //    MydataEmployee.Add("EditiqueBulletin", "");
        //    MydataEmployee.Add("CotorepDateFin", "");
        //    MydataEmployee.Add("CotorepCategorie", "");
        //    MydataEmployee.Add("CodeInseeCommune", "");
        //    MydataEmployee.Add("RompuBilletagePrecedent", "");
        //    MydataEmployee.Add("RompuBilletageCourant", "");
        //    MydataEmployee.Add("GHRCumulAnnuelHS", "");
        //    MydataEmployee.Add("FormNiveauEtudes", "");
        //    MydataEmployee.Add("GHRSoldeRCReel", "");
        //    MydataEmployee.Add("GHRSoldeRCCalcul", "");
        //    MydataEmployee.Add("CodeFonctionEntreprise", "");
        //    MydataEmployee.Add("NumeroDeBadge", "");
        //    MydataEmployee.Add("GHRDateDernierEnregistrement", "");
        //    MydataEmployee.Add("GHRDateDernierEnregistrementHS", "");
        //    MydataEmployee.Add("GHRDateDernierEnregistrementRC", "");
        //    MydataEmployee.Add("GHRCodePAM", "");
        //    MydataEmployee.Add("GHRDernierContingentHS", "");
        //    MydataEmployee.Add("GHRContingentHSEnCours", "");
        //    MydataEmployee.Add("NumeroSQL", "");
        //    MydataEmployee.Add("CodeEmploiINSEE", "");
        //    MydataEmployee.Add("PaiementBanque", "");
        //    MydataEmployee.Add("PaiementBanqueEnTaux", "");
        //    MydataEmployee.Add("ChoixBanque1Ou2", "");
        //    MydataEmployee.Add("IdentifiantEpargne", "");
        //    MydataEmployee.Add("CleIdentifiantEpargne", "");
        //    MydataEmployee.Add("ResidentFiscalEpargne", "");
        //    MydataEmployee.Add("SoumisCSGEpargne", "");
        //    MydataEmployee.Add("DIFDernierEntProf", "");
        //    MydataEmployee.Add("DIFResteAPrendre", "");
        //    MydataEmployee.Add("EMail", "");
        //    MydataEmployee.Add("BulletinDematerialise", "");
        //    MydataEmployee.Add("RangDeNaissance", "");
        //    MydataEmployee.Add("Commentaire3", "");
        //    MydataEmployee.Add("NumeroDePortable", "");
        //    MydataEmployee.Add("TelephoneProfessionnel", "");
        //    MydataEmployee.Add("TelephPortableProfessionnel", "");
        //    MydataEmployee.Add("FaxProfessionnel", "");
        //    MydataEmployee.Add("AdresseMelProfessionnelle", "");
        //    //MydataEmployee.Add("APrevenirPersonne01","" );
        //    //MydataEmployee.Add("APrevenirTelephone01","" );
        //    //MydataEmployee.Add("APrevenirPersonne02","" );
        //    //MydataEmployee.Add("APrevenirTelephone02","" );
        //    MydataEmployee.Add("AEDTransmise", "0");
        //    //MydataEmployee.Add("ModiFiche_date","" );
        //    MydataEmployee.Add("ModiFiche_heure", "0");
        //    //MydataEmployee.Add("ModiFiche_utilisateur","" );
        //    MydataEmployee.Add("MiseAZeroCongesLorsDuDepart", "0");
        //    MydataEmployee.Add("DSNFCTTransmise", "0");
        //    //MydataEmployee.Add("DebutPeriodeRattachement","" );
        //    //MydataEmployee.Add("FinPeriodeRattachement","" );
        //    //MydataEmployee.Add("CodeDistribuALEtranger","" );
        //    MydataEmployee.Add("DebutDePriseDesConges", "0");
        //    MydataEmployee.Add("DureeDePriseDesConges", "0");
        //    MydataEmployee.Add("CPAcquisPrecedentBis", "0");
        //    MydataEmployee.Add("CPAcquisFracEnCours", "0");
        //    MydataEmployee.Add("CPAcquisFracPrecedent", "0");
        //    MydataEmployee.Add("CPAcquisFracPrecedentBis", "0");
        //    MydataEmployee.Add("CPAcquisAncEnCours", "0");
        //    MydataEmployee.Add("CPAcquisAncPrecedent", "0");
        //    MydataEmployee.Add("CPAcquisAncPrecedentBis", "0");
        //    MydataEmployee.Add("CPAcquisSupEnCours", "0");
        //    MydataEmployee.Add("CPAcquisSupPrecedent", "0");
        //    MydataEmployee.Add("CPAcquisSupPrecedentBis", "0");
        //    MydataEmployee.Add("ResteAPrendreSuivant", "0");
        //    MydataEmployee.Add("ResteAPrendrePrecedent", "0");
        //    MydataEmployee.Add("ResteAPrendreSuivantFrac", "0");
        //    MydataEmployee.Add("ResteAPrendreFrac", "0");
        //    MydataEmployee.Add("ResteAPrendrePrecedentFrac", "0");
        //    MydataEmployee.Add("ResteAPrendreSuivantAnc", "0");
        //    MydataEmployee.Add("ResteAPrendreAnc", "0");
        //    MydataEmployee.Add("ResteAPrendrePrecedentAnc", "0");
        //    MydataEmployee.Add("ResteAPrendreSuivantSup", "0");
        //    MydataEmployee.Add("ResteAPrendreSup", "0");
        //    MydataEmployee.Add("ResteAPrendrePrecedentSup", "0");
        //    MydataEmployee.Add("CongesPrisAnneePrecedente", "0");
        //    MydataEmployee.Add("ResteAPrendreALaCloture", "0");
        //    MydataEmployee.Add("ResteAPrendreFracALaCloture", "0");
        //    MydataEmployee.Add("ResteAPrendreAncALaCloture", "0");
        //    MydataEmployee.Add("ResteAPrendreSupALaCloture", "0");
        //    MydataEmployee.Add("PrisCongesInitialEnCours", "0");
        //    MydataEmployee.Add("PrisCongesInitialPrecedent", "0");
        //    MydataEmployee.Add("PrisCongesInitialPrecedentBis", "0");
        //    MydataEmployee.Add("CPFResteAPrendre", "0");
        //    MydataEmployee.Add("SalarieDesactive", "0");
        //    //MydataEmployee.Add("ChoixEtablissement","" );
        //    MydataEmployee.Add("FlagInfosGenerales", "0");
        //    //MydataEmployee.Add("Memo","" );
        //    //MydataEmployee.Add("DateNaissance","" );
        //    //MydataEmployee.Add("DeptNaissance","" );

        //    return MydataEmployee;
        //}
        // Fonction pour nettoyer les valeurs nulles ou vides
        void CleanData(Dictionary<string, object> data)
        {
            var keysToRemove = new List<string>();

            foreach (var key in data.Keys.ToList())
            {
                var value = data[key];

                // Vérifier si la valeur est null ou équivalente à "NULL"
                if (value == null ||
                  (value is string strValue && string.IsNullOrWhiteSpace(strValue)) || // Chaîne vide après nettoyage
                  (value is string strValue2 && strValue2.Trim().ToUpper() == "NULL")) // Équivalent à "NULL"
                {
                    keysToRemove.Add(key);
                    continue;
                }

                if (value is string stringValue)
                {
                    // Échapper les apostrophes et encapsuler les chaînes dans des apostrophes pour SQL
                    data[key] = $"'{stringValue}'";
                }
                else
                {
                    // Conserver les autres types de données
                    data[key] = value;
                }
            }

            // Supprimer les clés dont les valeurs ont été définies comme nulles ou spéciales
            foreach (var key in keysToRemove)
            {
                data.Remove(key);
            }
        }

        string GetIdTable(string table, string columnselect, string clmnid, string valueid, int nbrBdd)
        {
            string retVal = null; // Valeur par défaut en cas d'erreur
            string queryString = "SELECT " + columnselect + " FROM " + table + " WHERE " + clmnid + " =" + valueid.Replace("''", "'") + ";";
            //MessageBox.Show(queryString);
            Clipboard.SetText(queryString);

            SqlCommand command = null;
            if (nbrBdd == 1)
            {
                command = new SqlCommand(queryString, cnn1);

            }
            else if (nbrBdd == 2)
            {
                command = new SqlCommand(queryString, cnn2);

            }
            else if (nbrBdd == 3)
            {
                command = new SqlCommand(queryString, cnn3);

            }

            else if (nbrBdd == 0)
            {
                command = new SqlCommand(queryString, cnn0);

            }
            // OpenConnection(nbr_bdd);
            try
            {
                var firstColumn = command.ExecuteScalar();

                if (firstColumn != null)
                {
                    retVal = firstColumn.ToString();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            command.Dispose();
            //  CloseConnection(nbr_bdd);
            return retVal;
        }
        private byte ConvertToByte(string value)
        {
            try
            {
                return byte.Parse(value);
            }
            catch (FormatException ex)
            {
                // Gérer le cas où la chaîne n'est pas un nombre valide
                Console.WriteLine($"FormatException: {ex.Message}");
                throw;
            }
            catch (OverflowException ex)
            {
                // Gérer le cas où la chaîne représente un nombre hors de la plage de byte
                Console.WriteLine($"OverflowException: {ex.Message}");
                throw;
            }
        }
        public static string LimiteString(string DataIn, int Long)
        {
            if (DataIn.Length > Long)
            {
                DataIn = DataIn.Substring(0, Long);
            }

            return DataIn;
        }
//l'inserartion a la table DATA
public void display_data(Dictionary<string, object> data_param)
        {
            //MessageBox.Show("for data");
             // var data_param = new ();
            if (dataGridView1.Rows.Count > 0)
            {

                
                        var insertDataResult = InsertTable_("DATA", data_param, 0);
                        //MessageBox.Show(insertDataResult != -1 ? "Les données pour  Data ont été bien insérées." : "Erreur lors de l'insertion dans Data .");

               

                    }
        }



        void alldata(Dictionary<string, object> data_param, Dictionary<string, object> MydataEmployee_hst_famille, Dictionary<string, object> MydataEmployee_hst_secu, Dictionary<string, object> MydataEmployee_hst_nationalite, Dictionary<string, object>
       MydataEmployee_hst_infobanque, Dictionary<string, object> MydataEmployee_hst_contrat, Dictionary<string, object> MydataEmployee_hst_etablissement, Dictionary<string, object>
       MydataEmployee_hst_infossociete, Dictionary<string, object> MydataEmployee_hst_affectation, Dictionary<string, object> MydataEmployee_T_ZONESLIBRES, Dictionary<string, object>
       MydataEmployee_hst_salaire,DataGridViewRow row,int result_T_SAL, int nbr)
        {


           // MydataEmployee_hst_famille.Add("NumSalarie", result_T_SAL.ToString());

            var maritalStatus = row.Cells["Marital_Status"].Value?.ToString();
            int statusCode = -1;

            switch (maritalStatus)
            {
                case "Single":
                    statusCode = 0;
                    break;
                case "Married":
                    statusCode = 1;
                    break;
                case "Divorced":
                    statusCode = 2;
                    break;
                case "Widowed":
                    statusCode = 3;
                    break;
                default:
                    // Handle unexpected values or assign a default code
                    statusCode = -1;
                    break;
            }

            if (statusCode != -1)
            {
                MydataEmployee_hst_famille.Add("SituationFamille", statusCode.ToString());
            }


            MydataEmployee_hst_famille.Add("CodePaysNaissance", row.Cells["Nationality"].Value?.ToString());
            MydataEmployee_hst_secu.Add("NumSalarie", result_T_SAL.ToString());
            MydataEmployee_hst_secu.Add("NoSecu", row.Cells["Identifier_Type_CIN"].Value?.ToString());

            MydataEmployee_hst_nationalite.Add("NumSalarie", result_T_SAL.ToString());
            MydataEmployee_hst_nationalite.Add("NoCarte", row.Cells["Identifier_Type_CNSS"].Value?.ToString());
            MydataEmployee_hst_nationalite.Add("DelivrePar", row.Cells["DelivrePar"].Value?.ToString());

            MydataEmployee_hst_infobanque.Add("NumSalarie", result_T_SAL.ToString());
            MydataEmployee_hst_infobanque.Add("NomBanque", row.Cells["Bank_Name"].Value?.ToString());
            MydataEmployee_hst_infobanque.Add("NomGuichet", row.Cells["Branch_Code"].Value?.ToString());

            MydataEmployee_hst_contrat.Add("NumSalarie", result_T_SAL.ToString());
            //t_contrat

            DateTime dateStar_Contract;

            var cellValue_datedebutContrat = row.Cells["Start_Date_Contract"].Value?.ToString();
            if (DateTime.TryParse(cellValue_datedebutContrat, out dateStar_Contract))
            {
                dateStar_Contract = dateStar_Contract.Date;

                // Stocker la date dans le dictionnaire
                MydataEmployee_hst_contrat.Add("DateDebutContrat", dateStar_Contract);
                data_param.Add("DateDebutContrat", dateStar_Contract);

            }
            else
            {
                MydataEmployee_hst_contrat.Add("DateDebutContrat", "NULL");
                data_param.Add("DateDebutContrat", "NULL");
            }



            DateTime dateEnd_Contract;

            var cellValue_datefinContrat = row.Cells["End_Date_Contract"].Value?.ToString();

            if (DateTime.TryParse(cellValue_datefinContrat, out dateEnd_Contract))
            {
                dateEnd_Contract = dateEnd_Contract.Date;
                // Stocker la date dans le dictionnaire
                MydataEmployee_hst_contrat.Add("DateFinContrat", dateEnd_Contract);
                data_param.Add("DateFinContrat", dateEnd_Contract);

            }
            else
            {
                MydataEmployee_hst_contrat.Add("DateFinContrat", "NULL");
                data_param.Add("DateFinContrat", "NULL");
            }
            //fk
            if (row.Cells["Contract_Type"].Value?.ToString() != "")
            {
                string contract_Type = "'" + row.Cells["Contract_Type"].Value?.ToString() + "'";
                string resultatnaturedecontrat = GetIdTable("T_NATUREDECONTRAT", "Code", "Intitule", contract_Type, nbr);

                if (resultatnaturedecontrat == null)
                {

                    var data = new Dictionary<string, object>();
                    data.Add("Intitule", contract_Type);
                    //incrementer le dernniere code
                    data.Add("Code", CodeAdd(contract_Type));
                    int resultat = InsertTable("T_NATUREDECONTRAT", data, nbr);
                    //MessageBox.Show(resultat != -1 ? "Les données pour T_NATUREDECONTRAT ont été bien insérées le code est" + CodeAdd(contract_Type) + "." : "Erreur lors de l'insertion dans T_NATUREDECONTRAT.");
                    MydataEmployee_hst_contrat.Add("CodeNatureDeContrat", CodeAdd(contract_Type).ToString());

                }
                else
                {
                    //MessageBox.Show("CodeNatureDeContrat deja exist");
                    MydataEmployee_hst_contrat.Add("CodeNatureDeContrat", GetIdTable("T_NATUREDECONTRAT", "Code", "Intitule", contract_Type, nbr));


                }
                data_param.Add("CodeNatureDeContrat", contract_Type);
            }



            // Pour T_HST_ETABLISSEMENT
            MydataEmployee_hst_etablissement.Add("NumSalarie", result_T_SAL.ToString());

            DateTime dateEntree;
            DateTime dateSortie;

            var cellValue_dateentree = row.Cells["Hire_Date"].Value?.ToString();

            if (DateTime.TryParse(cellValue_dateentree, out dateEntree))
            {
                dateEntree = dateEntree.Date;
                // Stocker la date dans le dictionnaire
                MydataEmployee_hst_etablissement.Add("DateEntree", dateEntree);
                MydataEmployee_hst_infossociete.Add("DateEmbauche", dateEntree);
                data_param.Add("Hire_Date", row.Cells["Hire_Date"].Value?.ToString());

            }
            else
            {
                MydataEmployee_hst_etablissement.Add("DateEntree", "NULL");
                MydataEmployee_hst_infossociete.Add("DateEmbauche", "NULL");
                data_param.Add("Hire_Date", "NULL");
            }



            var cellValue_datesortie = row.Cells["Termination_Date"].Value?.ToString();

            if (DateTime.TryParse(cellValue_datesortie, out dateSortie))
            {

                dateSortie = dateSortie.Date;

                // Stocker la date dans le dictionnaire
                MydataEmployee_hst_etablissement.Add("DateSortie", dateSortie);
                MydataEmployee_hst_infossociete.Add("DateDepart", dateSortie);
                data_param.Add("Termination_Date", dateSortie);


            }
            else
            {
                MydataEmployee_hst_etablissement.Add("DateSortie", "NULL");
                MydataEmployee_hst_infossociete.Add("DateDepart", "NULL");
                data_param.Add("Termination_Date", "NULL");
            }


            MydataEmployee_hst_infossociete.Add("NumSalarie", result_T_SAL.ToString());


            DateTime dateanciennete;

            var cellValue_dateanciennete = row.Cells["Seniority_Date"].Value?.ToString();

            if (DateTime.TryParse(cellValue_dateanciennete, out dateanciennete))
            {
                dateanciennete = dateanciennete.Date;
                // Stocker la date dans le dictionnaire
                MydataEmployee_hst_infossociete.Add("DateAnciennete", dateanciennete);
                data_param.Add("Seniority_Date", row.Cells["Seniority_Date"].Value?.ToString());



            }
            else
            {
                MydataEmployee_hst_infossociete.Add("DateAnciennete", "NULL");
                data_param.Add("Seniority_Date", "NULL");
            }



            string code_etab = "'" + row.Cells["Local_Termination_Reason"].Value?.ToString() + "'";
            if (row.Cells["Local_Termination_Reason"].Value?.ToString() != "")
            {
                string resultatcode_etab = GetIdTable("T_MOTIFDEPART", "Code", "Intitule", code_etab, nbr);
                if (resultatcode_etab == null)
                {

                    var data = new Dictionary<string, object>();
                    data.Add("Intitule", code_etab);
                    //incrementer le dernniere code
                    data.Add("Code", CodeAdd(code_etab));
                    int resultat = InsertTable("T_MOTIFDEPART", data, nbr);
                    //MessageBox.Show(resultat != -1 ? "Les données pour T_MOTIFDEPART ont été bien insérées ." : "Erreur lors de l'insertion dans T_MOTIFDEPART.");
                    MydataEmployee_hst_infossociete.Add("CodeMotifDepart", CodeAdd(code_etab).ToString());
                }
                else
                {
                    //MessageBox.Show("CodeMotifDepart deja exist");
                    MydataEmployee_hst_infossociete.Add("CodeMotifDepart", GetIdTable("T_MOTIFDEPART", "Code", "Intitule", code_etab, nbr));

                }
            }


            MydataEmployee_hst_affectation.Add("NumSalarie", result_T_SAL.ToString());
            MydataEmployee_hst_affectation.Add("Niveau", row.Cells["Niveau"].Value?.ToString());
            MydataEmployee_hst_affectation.Add("Indice", row.Cells["Indice"].Value?.ToString());
            MydataEmployee_hst_affectation.Add("Qualification", row.Cells["Qualification"].Value?.ToString());
            MydataEmployee_hst_affectation.Add("Coefficient", row.Cells["Position_Time_Type"].Value?.ToString());
            //for T_HST_AFFECTATION
       if (row.Cells["Service"].Value?.ToString() != "")
                {
                    string code_service = "'" + row.Cells["Service"].Value?.ToString() + "'";

                    string resultatcode_service = GetIdTable("T_SERVICE", "Code", "Intitule", code_service, nbr);
                    if (resultatcode_service == null)
                    {

                        var data = new Dictionary<string, object>();
                        data.Add("Intitule", code_service);
                        //incrementer le dernniere code
                        data.Add("Code", CodeAdd(code_service));
                        int resultat = InsertTable("T_SERVICE", data, nbr);
                        //MessageBox.Show(resultat != -1 ? "Les données pour T_SERVICE ont été bien insérées ." : "Erreur lors de l'insertion dans T_SERVICE.");
                        MydataEmployee_hst_affectation.Add("Service", CodeAdd(code_etab).ToString());
                    }
                    else
                    {
                        //MessageBox.Show("CodeService deja exist");
                        MydataEmployee_hst_affectation.Add("Service", GetIdTable("T_SERVICE", "Code", "Intitule", code_etab, nbr));

                    }
                }

            if (row.Cells["Organization_Three"].Value?.ToString() != "")
            {

                string code_categorie = "'" + row.Cells["Organization_Three"].Value?.ToString() + "'";

                string resultatcode_categorie = GetIdTable("T_CATEGORIE", "Code", "Intitule", code_categorie, nbr);
                if (resultatcode_categorie == null)
                {

                    var data = new Dictionary<string, object>();
                    data.Add("Intitule", code_categorie);
                    //incrementer le dernniere code
                    data.Add("Code", CodeAdd(code_categorie));
                    int resultat = InsertTable("T_CATEGORIE", data, nbr);
                    //MessageBox.Show(resultat != -1 ? "Les données pour T_CATEGORIE ont été bien insérées ." : "Erreur lors de l'insertion dans CATEGORIE.");
                    MydataEmployee_hst_affectation.Add("Categorie", CodeAdd(code_categorie).ToString());
                }
                else
                {
                    //MessageBox.Show("CodeCATEGORIE deja exist");
                    MydataEmployee_hst_affectation.Add("Categorie", GetIdTable("T_CATEGORIE", "Code", "Intitule", code_categorie, nbr));

                }

            }


            MydataEmployee_T_ZONESLIBRES.Add("NumSalarie", result_T_SAL.ToString());
            MydataEmployee_T_ZONESLIBRES.Add("St20_2", row.Cells["Job_Family_Group"].Value?.ToString());
            MydataEmployee_T_ZONESLIBRES.Add("St20_3", row.Cells["Job_Family"].Value?.ToString());



            MydataEmployee_hst_salaire.Add("NumSalarie", result_T_SAL.ToString());
            var HoraireBase = row.Cells["HoraireBase"].Value?.ToString();

            if (byte.TryParse(HoraireBase, out byte result_HoraireBase))
            {
                MydataEmployee_hst_salaire.Add("HoraireBase", result_HoraireBase);

            }

            var SalaireBase = row.Cells["Total_Base_Pay"].Value?.ToString();
            if (byte.TryParse(SalaireBase, out byte result_SalaireBase))

            {
                MydataEmployee_hst_salaire.Add("SalaireBase", result_SalaireBase);



            }


            data_param.Add("CodePaysNaissance", row.Cells["Nationality"].Value?.ToString());
            data_param.Add("NoCarte", row.Cells["Identifier_Type_CNSS"].Value?.ToString());
            data_param.Add("DelivrePar", row.Cells["DelivrePar"].Value?.ToString());
            data_param.Add("SituationFamille", maritalStatus.ToString());
            data_param.Add("NomBanque", row.Cells["Bank_Name"].Value?.ToString());
            data_param.Add("NomGuichet", row.Cells["Branch_Code"].Value?.ToString());
            data_param.Add("Local_Termination_Reason", row.Cells["Local_Termination_Reason"].Value?.ToString());
            data_param.Add("Service", row.Cells["Service"].Value?.ToString());
            data_param.Add("Niveau", row.Cells["Niveau"].Value?.ToString());
            data_param.Add("Qualification", row.Cells["Qualification"].Value?.ToString());
            data_param.Add("Position_Time_Type", row.Cells["Position_Time_Type"].Value?.ToString());
            data_param.Add("Organization_one", row.Cells["Organization_one"].Value?.ToString());
            data_param.Add("Organization_Two", row.Cells["Organization_Two"].Value?.ToString());
            data_param.Add("Organization_Three", row.Cells["Organization_Three"].Value?.ToString());
            data_param.Add("Job_Family_Group", row.Cells["Job_Family_Group"].Value?.ToString());
            data_param.Add("Job_Family", row.Cells["Job_Family"].Value?.ToString());
            data_param.Add("HoraireBase", row.Cells["HoraireBase"].Value?.ToString());
            data_param.Add("Total_Base_Pay", row.Cells["Total_Base_Pay"].Value?.ToString());
            data_param.Add("name", row.Cells["name"].Value?.ToString());
            data_param.Add("first_address_line_data", row.Cells["first_address_line_data"].Value?.ToString());
            data_param.Add("first_municipality", row.Cells["first_municipality"].Value?.ToString());
            data_param.Add("operation_contract", row.Cells["Operation_contract"].Value?.ToString());
            data_param.Add("contract_type", row.Cells["Contract_type"].Value?.ToString());
            data_param.Add("Start_Date_Contract", row.Cells["Start_Date_Contract"].Value?.ToString());
            data_param.Add("End_Date_Contract", row.Cells["End_Date_Contract"].Value?.ToString());









        }



        private void button1_Click(object sender, EventArgs e)
        {
            

       

            // Récupérer la valeur sélectionnée
            string selectedValue = comboBox1.SelectedItem.ToString();
            string LastMatricule = "";

            // Liste pour stocker les noms et prénoms
            List<string> employeeNames = new List<string>();
            //insert for each company
            void insertcompany(int nbr)
            {
                //MessageBox.Show("first");
                OpenConnection(nbr);
                //MessageBox.Show("1111");

                //LastMatricule = GetLastMatricule(nbr);
                ////MessageBox.Show(LastMatricule.ToString());

                //LastMatricule = Increment(LastMatricule, Mode.AlphaNumeric);
                ////MessageBox.Show(LastMatricule.ToString());

                // Assurez-vous que le DataGridView a des lignes
                if (dataGridView1.Rows.Count > 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        // Ignorer les lignes de nouvelle ligne (si elles existe
                        // nt)
                        if (!row.IsNewRow)
                        {




                            //l'inserartion a la table DATA
                            var data_param = new Dictionary<string, object>();

                            //for T_SAL
                            var MydataEmployee = new Dictionary<string, object>();

                            //for T_HST_FAMILLE
                            var MydataEmployee_hst_famille = new Dictionary<string, object>();


                            //for T_HST_SECU
                            var MydataEmployee_hst_secu = new Dictionary<string, object>();

                            //for T_HST_NATIONALITE
                            var MydataEmployee_hst_nationalite = new Dictionary<string, object>();

                            //for T_INFOBANQUE
                            var MydataEmployee_hst_infobanque = new Dictionary<string, object>();

                            //for T_Contrat
                            var MydataEmployee_hst_contrat = new Dictionary<string, object>();

                            //for T_HST_ETABLISSEMENT
                            var MydataEmployee_hst_etablissement = new Dictionary<string, object>();

                            //for T_HST_INFOSSOCIETE
                            var MydataEmployee_hst_infossociete = new Dictionary<string, object>();


                            //for T_HST_AFFECTATION
                            var MydataEmployee_hst_affectation = new Dictionary<string, object>();


                            //for T_ZONESLIBRES
                            var MydataEmployee_T_ZONESLIBRES = new Dictionary<string, object>();




                            //for T_HST_SALAIRE 
                            var MydataEmployee_hst_salaire = new Dictionary<string, object>();









                            //for T_SAL
                            //  bool matricule_employee = EmployeeExistsByMatricule(row.Cells["Identifier_Type_CIN  "].Value?.ToString(), nbr);
                            //*
                            bool matricule_employee = EmployeeExistsByMatricule(row.Cells["Identifier_Type_CIN"].Value?.ToString(), nbr);
                            //MessageBox.Show("555" + matricule_employee.ToString());
                            //les information sal 

                            //MydataEmployee["MatriculeSalarie"] = row.Cells["Employee_ID"].Value?.ToString() ?? string.Empty;
                            if (row.Cells["Title"].Value?.ToString() == "Monsieur")
                            {
                                MydataEmployee["Civilite"] = Convert.ToByte("0");
                            }
                            else if (row.Cells["Title"].Value?.ToString() == "Mademoiselle")
                            {
                                MydataEmployee["Civilite"] = Convert.ToByte("1");
                            }
                            else if (row.Cells["Title"].Value?.ToString() == "Madame")
                            {
                                MydataEmployee["Civilite"] = Convert.ToByte("2");
                            }
                            MydataEmployee["Nom"] = row.Cells["Last_Name"].Value?.ToString() ?? string.Empty;
                            MydataEmployee["Prenom"] = LimiteString(row.Cells["First_Name"].Value?.ToString(), 20) ?? string.Empty;
                            MydataEmployee["Prenom2"] = LimiteString(row.Cells["Prenom2"].Value?.ToString(), 20) ?? string.Empty;
                            // Assurez-vous que la cellule n'est pas null et que la valeur est convertible en byte
                            var cellValue = row.Cells["Nbr_of_Dependents"].Value?.ToString();
                            if (byte.TryParse(cellValue, out byte result))
                            {
                                MydataEmployee["NbFichesENF"] = result;  // Stocker la valeur convertie en byte
                            }
                            else
                            {
                                MydataEmployee["NbFichesENF"] = (byte)0;  // Valeur par défaut si la conversion échoue
                            }

                            var cellValue_daten = row.Cells["Birth_Date"].Value?.ToString();

                            if (cellValue_daten != "")
                            {
                                // Stocker la date dans le dictionnaire
                                MydataEmployee["DateNaissance"] = cellValue_daten;
                            }
                            else
                            {
                                MydataEmployee["DateNaissance"] = DBNull.Value;
                            }

                            if (row.Cells["NoBulletinModele"].Value?.ToString() != "")
                            {

                                string nobulletinmodele = "'" + row.Cells["NoBulletinModele"].Value?.ToString() + "'";
                                string resultatnobulletinmodele = GetIdTable("T_BMOD", "CodeBulletinModele", "Intitule", nobulletinmodele, nbr);
                                if (resultatnobulletinmodele == null)
                                {

                                    var data = new Dictionary<string, object>();
                                    data.Add("Intitule", nobulletinmodele);
                                    //incrementer le dernniere code
                                    data.Add("Code", IncrementCode("T_BMOD", "CodeBulletinModele", nbr));
                                    int resultat = InsertTable("T_BMOD", data, nbr);
                                    //MessageBox.Show(resultat != -1 ? "Les données pour T_BMOD ont été bien insérées le code est" + CodeAdd(nobulletinmodele) + "." : "Erreur lors de l'insertion dans T_BMOD.");
                                    MydataEmployee["NoBulletinModele"] = IncrementCode("T_BMOD", "CodeBulletinModele", nbr);
                                }
                                else
                                {
                                    //MessageBox.Show("CodeNatureDeContrat deja exist");
                                    MydataEmployee["NoBulletinModele"] = GetIdTable("T_BMOD", "CodeBulletinModele", "Intitule", nobulletinmodele, nbr);


                                }

                            }

                            MydataEmployee["Rue1"] = LimiteString(row.Cells["First_Address_Line1_Data"].Value?.ToString(), 40) ?? string.Empty;
                            MydataEmployee["Rue2"] = LimiteString(row.Cells["First_Address_Line2_Data"].Value?.ToString(), 40) ?? string.Empty;
                            MydataEmployee["Commune"] = LimiteString(row.Cells["First_Municipality"].Value?.ToString(), 30) ?? string.Empty;
                            MydataEmployee["ModeDePaiement"] = row.Cells["Payment_Type"].Value?.ToString() ?? string.Empty;
                            MydataEmployee["EMail"] = LimiteString(row.Cells["CostCenter_Code"].Value?.ToString(), 128) ?? string.Empty;

                            //AJOUT EN DATA
                            data_param.Add("Title", row.Cells["Title"].Value?.ToString());
                            data_param.Add("Last_Name", row.Cells["Last_Name"].Value?.ToString());
                            data_param.Add("First_Name", row.Cells["First_Name"].Value?.ToString());
                            data_param.Add("Prenom2", row.Cells["Prenom2"].Value?.ToString());
                            data_param.Add("Nbr_of_Dependents", row.Cells["Nbr_of_Dependents"].Value?.ToString());
                            data_param.Add("Birth_Date", row.Cells["Birth_Date"].Value?.ToString());
                            data_param.Add("First_Address_Line1", row.Cells["First_Address_Line1_Data"].Value?.ToString());
                            data_param.Add("First_Address_Line2", row.Cells["First_Address_Line2_Data"].Value?.ToString());
                            data_param.Add("First_Municipality", row.Cells["First_Municipality"].Value?.ToString());
                            data_param.Add("Payment_Type", row.Cells["Payment_Type"].Value?.ToString());
                            data_param.Add("CostCenter_Code", row.Cells["CostCenter_Code"].Value?.ToString());
                            data_param.Add("Job_Profile", row.Cells["Job_Profile"].Value?.ToString());
                            data_param.Add("NoBulletinModele", row.Cells["NoBulletinModele"].Value?.ToString());
                            data_param.Add("data_import", DateTime.Now.ToString());
                            ////*
                            data_param.Add("CIN", row.Cells["Identifier_Type_CIN"].Value?.ToString());
                            data_param.Add("CNSS", row.Cells["Identifier_Type_CNSS"].Value?.ToString());
                            data_param.Add("SAGE_ID", row.Cells["Identifier_Type_SAGE_ID"].Value?.ToString());

                            //*
                            MydataEmployee["MatriculeSalarie"] = row.Cells["Identifier_Type_CIN"].Value?.ToString() ?? string.Empty;
                            //*
                          data_param.Add("MatriculeSalarie", row.Cells["Identifier_Type_CIN"].Value?.ToString());
                          data_param.Add("nom_fichier", row.Cells["nom_fichier"].Value?.ToString());
                          data_param.Add("payroll_company_name", row.Cells["payroll_company_name"].Value?.ToString());
                          data_param.Add("pay_group_id", row.Cells["pay_group_id"].Value?.ToString());
                          data_param.Add("pay_group_name", row.Cells["pay_group_name"].Value?.ToString());
                          data_param.Add("Country_for_Name", row.Cells["Country_for_Name"].Value?.ToString());
                          data_param.Add("employee_status", row.Cells["employee_status"].Value?.ToString());
                          data_param.Add("operation_Contract", row.Cells["Operation_Contract"].Value?.ToString());
                          data_param.Add("operation_position_", row.Cells["Operation_Position"].Value?.ToString());
                          data_param.Add("position_ID", row.Cells["Position_ID"].Value?.ToString());
                          data_param.Add("effective_date", row.Cells["Effective_date"].Value?.ToString());
                          data_param.Add("business_title", row.Cells["Business_title"].Value?.ToString());
                          data_param.Add("scheduled_weekly_hours", row.Cells["Scheduled_weekly_hours"].Value?.ToString());
                          data_param.Add("default_weekly_hours", row.Cells["Default_weekly_hours"].Value?.ToString());
                          data_param.Add("base_pay_frequency", row.Cells["Base_pay_frequency"].Value?.ToString());





                            MydataEmployee["CodeFonctionEntreprise"] = row.Cells["Job_Profile"].Value?.ToString() ?? string.Empty;
        

                            //si employee n'exist pas 
                            if (!matricule_employee)
                            {
                                //MessageBox.Show("employee n'exist pas  ");

                                //MessageBox.Show("333");
                                //MessageBox.Show(MydataEmployee["CodeFonctionEntreprise"].ToString());

                                int result_T_SAL = InsertTable_sal("T_SAL", MydataEmployee, nbr);
                                //MessageBox.Show("41");
                                if (result_T_SAL != -1)
                                    {   
                                        
                                        
                                        
                                        alldata(data_param,MydataEmployee_hst_famille,MydataEmployee_hst_secu, MydataEmployee_hst_nationalite,
                                           MydataEmployee_hst_infobanque,MydataEmployee_hst_contrat,MydataEmployee_hst_etablissement,
                                           MydataEmployee_hst_infossociete,MydataEmployee_hst_affectation,MydataEmployee_T_ZONESLIBRES,
                                           MydataEmployee_hst_salaire, row, result_T_SAL,  nbr);

                                        //for T_HST_FAMILLE
                                        int result_hst_famille = InsertTable("T_HST_FAMILLE", MydataEmployee_hst_famille, nbr);
                                        //MessageBox.Show(result_hst_famille != -1 ? "Les données pour T_HST_FAMILLE ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_FAMILLE.");


                                        // Pour T_HST_SECU
                                        int result_hst_secu = InsertTable("T_HST_SECU", MydataEmployee_hst_secu, nbr);
                                        //MessageBox.Show(result_hst_secu != -1 ? "Les données pour T_HST_SECU ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_SECU.");


                                        // Pour T_HST_NATIONALITE
                                        int result_hst_nationalite = InsertTable("T_HST_NATIONALITE", MydataEmployee_hst_nationalite, nbr);
                                        //MessageBox.Show(result_hst_nationalite != -1 ? "Les données pour T_HST_NATIONALITE ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_NATIONALITE.");



                                  
                         
                                        // Pour T_INFOBANQUE
                                        int result_infobanque = InsertTable("T_INFOBANQUE", MydataEmployee_hst_infobanque, nbr);
                                        //MessageBox.Show(result_infobanque != -1 ? "Les données pour T_INFOBANQUE ont été bien insérées." : "Erreur lors de l'insertion dans T_INFOBANQUE.");


                                        // Pour T_Contrat
                                        int result_contrat = InsertTable("T_HST_Contrat", MydataEmployee_hst_contrat, nbr);
                                        //MessageBox.Show(result_contrat != -1 ? "Les données pour T_Contrat ont été bien insérées." : "Erreur lors de l'insertion dans T_Contrat.");

                                        // Pour T_HST_ETABLISSEMENT




                                        int result_etablissement = InsertTable("T_HST_ETABLISSEMENT", MydataEmployee_hst_etablissement, nbr);
                                        //MessageBox.Show(result_etablissement != -1 ? "Les données pour T_HST_ETABLISSEMENT ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_ETABLISSEMENT.");


                                        // Pour T_HST_INFOSSOCIETE
                                      
                                            int result_infossociete = InsertTable("T_HST_INFOSSOCIETE", MydataEmployee_hst_infossociete, nbr);
                                            //MessageBox.Show(result_infossociete != -1 ? "Les données pour T_HST_INFOSSOCIETE ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_INFOSSOCIETE.");
                                        
                                        //for T_HST_AFFECTATION
                                       int result_affectation = InsertTable("T_HST_AFFECTATION", MydataEmployee_hst_affectation, nbr);
                                        //MessageBox.Show(result_affectation != -1 ? "Les données pour T_HST_AFFECTATION ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_AFFECTATION.");

                                        //for T_ZONESLIBRES
                                        int result_zoneslibres = InsertTable("T_ZONESLIBRES", MydataEmployee_T_ZONESLIBRES, nbr);
                                        //MessageBox.Show(result_zoneslibres != -1 ? "Les données pour T_ZONESLIBRES ont été bien insérées." : "Erreur lors de l'insertion dans T_ZONESLIBRES.");

                                        //for T_HST_SALAIRE 
                                       int result_salaire = InsertTable("T_HST_SALAIRE", MydataEmployee_hst_salaire, nbr);
                                        //MessageBox.Show(result_salaire != -1 ? "Les données pour T_HST_SALAIRE ont été bien insérées." : "Erreur lors de l'insertion dans T_HST_SALAIRE.");
                                    }



                                }
                            //si employee deja exist 
                            //else if (matricule_employee && (row.Cells["Operation_Contract"].Value?.ToString() == "REMOVE" || row.Cells["Operation_Contract"].Value?.ToString() == "MODIFY"))
                            else if (matricule_employee )
                            {
                                //MessageBox.Show("444");
                                //*
                                    //MessageBox.Show("employee deja exist" + row.Cells["Identifier_Type_CIN"].Value?.ToString());
                                //*


                                    int result_T_SAL_ = UpdateTable("T_SAL", MydataEmployee, nbr, "MatriculeSalarie", row.Cells["Identifier_Type_CIN"].Value?.ToString());

                                    if (result_T_SAL_ != -1)
                                    {

                                    //MessageBox.Show("4545");
                                        alldata(data_param, MydataEmployee_hst_famille, MydataEmployee_hst_secu, MydataEmployee_hst_nationalite,   MydataEmployee_hst_infobanque, MydataEmployee_hst_contrat, MydataEmployee_hst_etablissement,
                                        MydataEmployee_hst_infossociete, MydataEmployee_hst_affectation, MydataEmployee_T_ZONESLIBRES,
                                        MydataEmployee_hst_salaire, row, result_T_SAL_, nbr);
                                    //*
                                         string cellValue_ = "'" + row.Cells["Identifier_Type_CIN"].Value?.ToString() + "'";
                                        string result_T_SAL = GetIdTable("T_SAL", "SA_CompteurNumero", "MatriculeSalarie", cellValue_, nbr);
;
                                        //for T_HST_FAMILLE
                                        int result_hst_famille = UpdateTable("T_HST_FAMILLE", MydataEmployee_hst_famille, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_hst_famille != -1 ? "Les données pour T_HST_FAMILLE ont été bien modifiées." : "Erreur lors de l'insertion dans T_HST_FAMILLE.");


                                        // Pour T_HST_SECU
                                        int result_hst_secu = UpdateTable("T_HST_SECU", MydataEmployee_hst_secu, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_hst_secu != -1 ? "Les données pour T_HST_SECU ont été bien  modifiées." : "Erreur lors de l'insertion dans T_HST_SECU.");



                                        // Pour T_HST_NATIONALITE
                                       int result_hst_nationalite = UpdateTable("T_HST_NATIONALITE", MydataEmployee_hst_nationalite, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_hst_nationalite != -1 ? "Les données pour T_HST_NATIONALITE ont été bien  modifiées" : "Erreur lors de l'insertion dans  T_HST_NATIONALITE.");



                                        // Pour T_INFOBANQUE
                                        int result_infobanque = UpdateTable("T_INFOBANQUE", MydataEmployee_hst_infobanque, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_infobanque != -1 ? "Les données pour T_INFOBANQUE ont été bien  modifiées." : "Erreur lors de l'insertion dans T_INFOBANQUE.");



                                        // Pour T_Contrat
                                        int result_contrat = UpdateTable("T_HST_Contrat", MydataEmployee_hst_contrat, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_contrat != -1 ? "Les données pour T_Contrat ont été bien modifiées." : "Erreur lors de l'insertion dans T_Contrat.");


                                        // Pour T_HST_ETABLISSEMENT
                                       int result_etablissement = UpdateTable("T_HST_ETABLISSEMENT", MydataEmployee_hst_etablissement, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_etablissement != -1 ? "Les données pour T_HST_ETABLISSEMENT ont été bien modifie ." : "Erreur lors de l'insertion dans T_HST_ETABLISSEMENT.");


                                        // Pour T_HST_INFOSSOCIETE
                                       
                                                int result_infossociete = UpdateTable("T_HST_INFOSSOCIETE", MydataEmployee_hst_infossociete, nbr, "NumSalarie", result_T_SAL.ToString());
                                                //MessageBox.Show(result_infossociete != -1 ? "Les données pour T_HST_INFOSSOCIETE ont été bien modifiées ." : "Erreur lors de l'insertion dans T_HST_INFOSSOCIETE.");
                                       
                                        //for T_HST_AFFECTATION
                                 
                                        int result_affectation = UpdateTable("T_HST_AFFECTATION", MydataEmployee_hst_affectation, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_affectation != -1 ? "Les données pour T_HST_AFFECTATION ont été bien modifiées." : "Erreur lors de l'insertion dans T_HST_AFFECTATION.");



                                        //for T_ZONESLIBRES
                                      
                                        int result_zoneslibres = UpdateTable("T_ZONESLIBRES", MydataEmployee_T_ZONESLIBRES, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_zoneslibres != -1 ? "Les données pour T_ZONESLIBRES ont été bien modifiées." : "Erreur lors de l'insertion dans T_ZONESLIBRES.");



                                        //for T_HST_SALAIRE 
                                    
                                        int result_salaire = UpdateTable("T_HST_SALAIRE", MydataEmployee_hst_salaire, nbr, "NumSalarie", result_T_SAL.ToString());
                                        //MessageBox.Show(result_salaire != -1 ? "Les données pour T_HST_SALAIRE ont été bien modifiées." : "Erreur lors de l'insertion dans T_HST_SALAIRE.");


                                    }




                                }
                          
  display_data(data_param);





                        }

                    }
                }

                CloseConnection(nbr);
            }











            if (selectedValue == "Business Support Services Maroc")
            {
                insertcompany(1);
            }
            else if (selectedValue == "Business Casablanca 2S")
            {

                insertcompany(2);
            }
            else 
            {
                insertcompany(3);
            }



        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

           



            // Assurez-vous que le DataGridView a des données
            if (dataGridView1.DataSource is DataTable dt)
            {
                string selectedValue = comboBox1.SelectedItem?.ToString();

                if (!string.IsNullOrEmpty(selectedValue))
                {
                    // Filtrer les lignes du DataTable
                    DataView dv = dt.DefaultView;
                    dv.RowFilter = $"Organization_One = '{selectedValue}'";
                }
                else
                {
                    // Réinitialiser le filtre pour afficher toutes les lignes
                    DataView dv = dt.DefaultView;
                    dv.RowFilter = string.Empty;
                }
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void paramToolStripMenuItem_Click(object sender, EventArgs e)
        {

            panel2.Visible = true;
            panel1.Visible = false;
            panel3.Visible = false;
            DataTable dt = SelectTable("config", 0);
            dataGridView_config.DataSource = dt;
        }

        private void ajoutDuNouveauParamétrageDeLaBaseDeDonnéesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel2.Visible = false;
            panel1.Visible = false;
          
            DataTable dt = SelectTable("DATA", 0);
            dataGridView_data.DataSource = dt;


        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        DataTable SelectTable(string table, int nbr)
        {
            // Prépare la requête SQL de base
            string queryString = $"SELECT * FROM {table}"; // Modifiez cette requête si vous avez des filtres ou des paramètres supplémentaires

            // Déclaration de la connexion SQL
            SqlCommand command = null;

            // Sélection de la connexion en fonction de nbr
            switch (nbr)
            {
                case 1:
                    command = new SqlCommand(queryString, cnn1);
                    break;
                case 2:
                    command = new SqlCommand(queryString, cnn2);
                    break;
                case 3:
                    command = new SqlCommand(queryString, cnn3);
                    break;
                case 0:
                    command = new SqlCommand(queryString, cnn0);
                    break;
            }

            // Création d'un DataTable pour stocker les résultats
            DataTable dataTable = new DataTable();

            // Utilisation de SqlCommand, SqlConnection, et SqlDataAdapter

                try
                {

                    // Utilise SqlDataAdapter pour remplir le DataTable
                    using (SqlDataAdapter dataAdapter = new SqlDataAdapter(command))
                    {
                        dataAdapter.Fill(dataTable);
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Une erreur s'est produite : {ex.Message}");
                    // Retourne un DataTable vide en cas d'erreur
                    return new DataTable();
                }
            

            return dataTable;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenConnection(0);

       
            var data = new Dictionary<string, object>();
            //data.Add("ip_serveur", textBox_ipserver.Text);
            //data.Add("nom_bdd", textBox_company.Text);
            //data.Add("nom_societe", textBox_bdd.Text);

            //var insertDataResult = InsertTable("Config", data, 0);
            ////MessageBox.Show(insertDataResult != -1 ? "Les données pour  Config ont été bien insérées." : "Erreur lors de l'insertion dans Config .");

            
            //textBox_ipserver.Text = "";
            //textBox_company.Text = "";
            //textBox_bdd.Text = "";
            CloseConnection(0);

        }



        private void dd_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void fichierToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBoxItem selectedItem = (ComboBoxItem)comboBox2.SelectedItem;
            if (selectedItem != null)
            {
                string selectedKey = selectedItem.Key; 

                DataTable dt = SelectTable(selectedKey, 0);
                dataGridView_config.DataSource = dt;
            }


        }
        private void dataGridView_config_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
           
        }

        private void dataGridView_config_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

      
    }
}
