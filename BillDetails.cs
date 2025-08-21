using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Geenova
{
    public partial class BillDetails : Form
    {
        private string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\mydb.mdf;Integrated Security=True;Connect Timeout=30";

        private DataTable originalBillItemsData;
        private string currentUserRole;
        private string currentUsername;


        public BillDetails(string userRole, string username)
        {
            InitializeComponent();

            currentUserRole = userRole;
            currentUsername = username;

            rolelbl.Text = currentUserRole;
            rolenamelbl.Text = currentUsername;
        }

        public BillDetails()
        {
        }

        private void OrderDetails_Load(object sender, EventArgs e)
        {
            LoadAndDisplayBillData();
            LoadAndDisplayBillItemsData();
            ConfigureDataGridViews();

            // Store the original BillItems data
            originalBillItemsData = (DataTable)BillItemsTable.DataSource;
        }

        private void LoadAndDisplayBillData()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string billQuery = "SELECT BillNumber, BillDate, TotalAmount, Discount, NetTotal, Payment, Balance, SellerRole, SellerName FROM BillTbl ORDER BY BillDate DESC";
                    SqlDataAdapter billAdapter = new SqlDataAdapter(billQuery, con);
                    DataTable billTable = new DataTable();
                    billAdapter.Fill(billTable);
                    BillTable.DataSource = billTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading bill data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAndDisplayBillItemsData()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string itemsQuery = "SELECT BillNumber, ProductID, ProductName, Quantity, Price, Total FROM BillItemsTbl";
                    SqlDataAdapter itemsAdapter = new SqlDataAdapter(itemsQuery, con);
                    DataTable itemsTable = new DataTable();
                    itemsAdapter.Fill(itemsTable);
                    BillItemsTable.DataSource = itemsTable;

                    // Store the original data if not already stored
                    if (originalBillItemsData == null)
                    {
                        originalBillItemsData = itemsTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading bill items: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BillTable_SelectionChanged(object sender, EventArgs e)
        {
            // Only filter if there's no active search in searchtxt2
            if (BillTable.SelectedRows.Count > 0 && string.IsNullOrEmpty(searchtxt2.Text))
            {
                string billNumber = BillTable.SelectedRows[0].Cells["BillNumber"].Value.ToString();
                FilterBillItems(billNumber);
            }
        }

        private void FilterBillItems(string billNumber)
        {
            if (BillItemsTable.DataSource is DataTable itemsTable)
            {
                itemsTable.DefaultView.RowFilter = $"BillNumber = '{billNumber}'";
            }
        }

        private void ConfigureDataGridViews()
        {
            // Configure BillTable
            BillTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            BillTable.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            BillTable.MultiSelect = false;
            BillTable.ReadOnly = true;
            BillTable.AllowUserToAddRows = false;
            BillTable.RowHeadersVisible = false;

            // Format columns
            if (BillTable.Columns.Contains("BillDate"))
            {
                BillTable.Columns["BillDate"].HeaderText = "Date";
                BillTable.Columns["BillDate"].DefaultCellStyle.Format = "g";
            }
            if (BillTable.Columns.Contains("TotalAmount"))
            {
                BillTable.Columns["TotalAmount"].HeaderText = "Total";
                BillTable.Columns["TotalAmount"].DefaultCellStyle.Format = "N2";
            }
            if (BillTable.Columns.Contains("NetTotal"))
            {
                BillTable.Columns["NetTotal"].HeaderText = "Net Total";
                BillTable.Columns["NetTotal"].DefaultCellStyle.Format = "N2";
            }
            if (BillTable.Columns.Contains("Payment"))
            {
                BillTable.Columns["Payment"].HeaderText = "Payment";
                BillTable.Columns["Payment"].DefaultCellStyle.Format = "N2";
            }
            if (BillTable.Columns.Contains("Balance"))
            {
                BillTable.Columns["Balance"].HeaderText = "Balance";
                BillTable.Columns["Balance"].DefaultCellStyle.Format = "N2";
            }

            // Configure BillItemsTable
            BillItemsTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            BillItemsTable.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            BillItemsTable.MultiSelect = false;
            BillItemsTable.ReadOnly = true;
            BillItemsTable.AllowUserToAddRows = false;
            BillItemsTable.RowHeadersVisible = false;

            // Format columns
            if (BillItemsTable.Columns.Contains("Price"))
            {
                BillItemsTable.Columns["Price"].DefaultCellStyle.Format = "N2";
            }
            if (BillItemsTable.Columns.Contains("Total"))
            {
                BillItemsTable.Columns["Total"].DefaultCellStyle.Format = "N2";
            }
        }

        private void Searchbtn1_Click(object sender, EventArgs e)
        {
            string billNumber = searchtxt1.Text.Trim();

            if (string.IsNullOrEmpty(billNumber))
            {
                MessageBox.Show("Please enter a Bill Number to search.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT BillNumber, ProductID, ProductName, Quantity, Price, Total FROM BillItemsTbl WHERE BillNumber = @BillNumber";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@BillNumber", billNumber);
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            BillItemsTable.DataSource = dt;
                        }
                        else
                        {
                            MessageBox.Show("No items found for the given Bill Number.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error searching bill items: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Searchbtn2_Click(object sender, EventArgs e)
        {
            string billNumber = searchtxt2.Text.Trim();

            if (string.IsNullOrEmpty(billNumber))
            {
                MessageBox.Show("Please enter a Bill Number to search.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT BillNumber, BillDate, TotalAmount, Discount, NetTotal, Payment, Balance, SellerRole, SellerName FROM BillTbl WHERE BillNumber = @BillNumber";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@BillNumber", billNumber);
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            BillTable.DataSource = dt;
                        }
                        else
                        {
                            MessageBox.Show("No bill found with the given Bill Number.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error searching bill: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void showAllBtn_Click(object sender, EventArgs e)
        {
            // Restore original BillItems data
            BillItemsTable.DataSource = originalBillItemsData;
            searchtxt1.Text = "";
        }

        private void showAllbtn2_Click(object sender, EventArgs e)
        {
            LoadAndDisplayBillData();
            searchtxt2.Text = "";

            // Clear any filtering on BillItemsTable
            if (BillItemsTable.DataSource is DataTable dt)
            {
                dt.DefaultView.RowFilter = "";
            }
        }

        private void procatbtn_Click(object sender, EventArgs e)
        {   
            this.Hide();
            productcat productForm = new productcat(currentUserRole, currentUsername);
            productForm.Show();
        }

        private void cashierbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Cashier cashierForm = new Cashier(currentUserRole, currentUsername);
            cashierForm.Show();
        }

        private void empbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Employee empForm = new Employee(currentUserRole, currentUsername);
            empForm.Show();
        }

        private void logoutbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login loginForm = new Login();
            loginForm.Show();
        }

        
    }
}