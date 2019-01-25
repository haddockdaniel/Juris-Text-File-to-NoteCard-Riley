using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            UpdateStatus("Parsing text file.", 1, 5);
            string file = Properties.Resources.tabdelim;
            string[] allLines = file.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            NoteCard nc = null;
            List<NoteCard> ncList = new List<NoteCard>();
            //store client , matter numbers and party type/names 
            for (int i = 0; i < allLines.Count(); i++)
            {
                try
                {
                    string[] items = allLines[i].Split(new[] { '\t' }, StringSplitOptions.None);
                    nc = new NoteCard();
                    nc.Client = items[0].Trim().Replace(" ", "");
                    nc.matter = items[1].Trim().Replace(" ", "").Replace(" (Cont.)", "");
                    nc.partyName = items[2].Trim().Replace("'", "");
                    nc.partyType = items[3].Trim();
                    if (items.Count() > 4)
                        nc.synopsis = items[4].Trim().Replace("'", "");
                    nc.clientID = "-1";
                    nc.matterID = "-1";
                    ncList.Add(nc);
                }
                catch (Exception ex4)
                {
                    MessageBox.Show(allLines[i]);
                }
            }

            UpdateStatus("Converting Client/Matter numbers to IDs.", 2,5);


            //convert client/matter numbers to ids

            string SQL = "";
            string exceptions = "";
            foreach (NoteCard n in ncList)
            {
                SQL = "select clisysnbr from client where clicode like '%" + n.Client + "'"; ;
                DataSet matSet = _jurisUtility.RecordsetFromSQL(SQL);
                if (matSet.Tables[0].Rows.Count > 0)
                    n.clientID = matSet.Tables[0].Rows[0][0].ToString();
                SQL = "select matsysnbr from matter where matcode like '%" + n.matter + "' and matclinbr = " + n.clientID;
                DataSet cliSet = _jurisUtility.RecordsetFromSQL(SQL);
                if (cliSet.Tables[0].Rows.Count > 0)
                    n.matterID = cliSet.Tables[0].Rows[0][0].ToString();
                if (n.matterID.Equals("-1") || n.clientID.Equals("-1"))
                    exceptions = exceptions + "\r\n" + n.Client + "\t" + n.matter + "\t" + n.clientName;
            }

            UpdateStatus("Adding Note Cards.", 3,5);

            //add notecards
            foreach (NoteCard n in ncList)
            {
                if (!n.matterID.Equals("-1"))
                {

                    SQL = "SELECT * FROM MatterNote where MNMatter = " + n.matterID + " and MNNoteIndex = '" + n.partyType + "'";
                    DataSet matSet = _jurisUtility.RecordsetFromSQL(SQL);
                    if (matSet.Tables[0].Rows.Count > 0) //if it already exists
                        SQL = "update matternote set mnnotetext = cast(mnnotetext as nvarchar(max)) + cast( char(13) + char(10) as nvarchar(max)) + cast(replace('" + n.partyName + "', '|', char(13) + char(10))as nvarchar(max)) where mnmatter = " + n.matterID + " and mnnoteindex = '" + n.partyType + "'";
                    else
                        SQL = "insert into matternote (mnmatter, mnnoteindex, mnobject, mnnotetext, mnnoteobject) Values (" + n.matterID + ", '" + n.partyType + "', '', replace('" + n.partyName + "', '|', char(13) + char(10)), null)";

                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                }
            }

            UpdateStatus("Adding Matter Remarks.", 4,5);

            //add matter remarks
            foreach (NoteCard n in ncList)
            {
                if (!string.IsNullOrEmpty(n.synopsis))
                {
                    SQL = "update matter set MatRemarks = '" + n.synopsis + "' where matsysnbr = " + n.matterID;
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                }
            }

            UpdateStatus("Utility Complete.", 5,5);

            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }




    }
}
