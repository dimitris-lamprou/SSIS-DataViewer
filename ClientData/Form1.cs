using Ionic.Zip;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;

namespace ClientData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Unzip();
            RunPackageSSIS();
            LoadClients();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            
        }

        private void Unzip()
        {
            string zipPath = Path.Combine(Application.StartupPath, "Data/Files.zip");
            string extractPath = Path.Combine(Application.StartupPath, "ExtractedFiles");

            try
            {
                if (!Directory.Exists(extractPath))
                    Directory.CreateDirectory(extractPath);

                using (ZipFile zip = ZipFile.Read(zipPath))
                {
                    zip.Password = "FirstCall13";
                    zip.ExtractAll(extractPath, ExtractExistingFileAction.OverwriteSilently);
                }

                MessageBox.Show("Το unzip ολοκληρώθηκε επιτυχώς!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά το unzip: " + ex.Message);
            }
        }

        private void RunPackageSSIS()
        {
            string ssisPackagePath = Path.Combine(Application.StartupPath, @"SSIS Package\Package.dtsx");
            string[] possibleDTexecPaths = {
                                            @"C:\Program Files\Microsoft SQL Server\150\DTS\Binn\dtexec.exe",
                                            @"C:\Program Files\Microsoft SQL Server\160\DTS\Binn\dtexec.exe",
                                            @"C:\Program Files (x86)\Microsoft SQL Server\150\DTS\Binn\dtexec.exe",
                                            @"C:\Program Files\Microsoft SQL Server\140\DTS\Binn\dtexec.exe"
            };

            string dtexecPath = possibleDTexecPaths.FirstOrDefault(File.Exists);
            if (dtexecPath == null) throw new FileNotFoundException("Cannot find dtexec.exe");

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = dtexecPath;
            psi.Arguments = $@"/FILE ""{ssisPackagePath}"" /CHECKPOINTING OFF";

            psi.CreateNoWindow = false;
            psi.UseShellExecute = false;

            Process proc = new Process();
            proc.StartInfo = psi;
            proc.Start();

            proc.WaitForExit();

            if (proc.ExitCode == 0)
            {
                // Εδώ η εκτέλεση ήταν επιτυχής
                MessageBox.Show("Το SSIS package ολοκληρώθηκε επιτυχώς!");
            }
            else
            {
                // Κάποιο σφάλμα συνέβη
                MessageBox.Show("Σφάλμα κατά την εκτέλεση του SSIS package");
            }
        }
        private void LoadClients()
        {
            string connectionString = @"Data Source=LAPTOP-K0MBAPHV\SQLSERVERDEVED;Initial Catalog=FirstCall Project;Integrated Security=True;";
            string query = "SELECT DISTINCT lastName FROM Client ORDER BY lastName";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        comboBox1.Items.Clear();
                        while (reader.Read())
                        {
                            comboBox1.Items.Add(reader["lastName"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά τη φόρτωση πελατών: " + ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        { 
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            string selectedLastName = comboBox1.SelectedItem.ToString();

            string connectionString = @"Data Source=LAPTOP-K0MBAPHV\SQLSERVERDEVED;Initial Catalog=FirstCall Project;Integrated Security=True;";

            string query = @"
        SELECT 
            ISNULL(STRING_AGG(ph.phone, ', '), '') AS Phones,
            p.paymentAmount,
            cs.id AS CaseID,
            cs.cardOrLoan,
            cs.caseCode,
            cs.acc,
            cs.productCode,
            cs.limitOrLoanAmount,
            cs.balance,
            cs.interestRate,
            cs.overdueAmount,
            cs.statementBalance,
            cs.interest,
            cs.assignmentDate,
            cs.caseStatus,
            cs.expiryDate,
            cs.storeCode,
            cs.assignmentCode,
            cs.totalBalance,
            cs.assignmentTotalBalance,
            cs.assignmentEndlessAmount,
            cs.assignmentOverdueAmount,
            cs.availableBalance,
            cs.revocationCode,
            cs.caseBucket,
            cs.daysLate,
            cs.loanPurpose            
        FROM Client c
        LEFT JOIN Phone ph ON ph.clientID = c.id
        LEFT JOIN [Relation] r ON r.clientID = c.id
        LEFT JOIN [Case] cs ON cs.id = r.caseID
        LEFT JOIN Payment p ON p.caseID = cs.id
        WHERE c.lastName = @lastName
        GROUP BY 
            p.paymentAmount,
            cs.id,
            cs.cardOrLoan,
            cs.caseCode,
            cs.acc,
            cs.productCode,
            cs.limitOrLoanAmount,
            cs.balance,
            cs.interestRate,
            cs.overdueAmount,
            cs.statementBalance,
            cs.interest,
            cs.assignmentDate,
            cs.caseStatus,
            cs.expiryDate,
            cs.storeCode,
            cs.assignmentCode,
            cs.totalBalance,
            cs.assignmentTotalBalance,
            cs.assignmentEndlessAmount,
            cs.assignmentOverdueAmount,
            cs.availableBalance,
            cs.revocationCode,
            cs.caseBucket,
            cs.daysLate,
            cs.loanPurpose";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@lastName", selectedLastName);

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά τη φόρτωση δεδομένων: " + ex.Message);
            }
        }
    }
}
