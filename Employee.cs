using System;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace Geenova
{
    public partial class Employee : Form
    {
        string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\mydb.mdf;Integrated Security=True;Connect Timeout=30";

        int selectedEmpID = 0;
        private string currentUserRole;
        private string currentUsername;

        private string currentPasswordHash = ""; // Stores current hashed password

        public Employee(string userRole, string username)
        {
            InitializeComponent();
            empnametxt.MaxLength = 100;
            empmobiletxt.MaxLength = 20;

            currentUserRole = userRole;
            currentUsername = username;

            rolelbl.Text = currentUserRole;
            rolenamelbl.Text = currentUsername;
        }

        public Employee() { }

        private void Employee_Load(object sender, EventArgs e)
        {
            roleComboBox.Items.Clear();
            roleComboBox.Items.Add("Staff");
            roleComboBox.Items.Add("Cashier");
            roleComboBox.Items.Add("Manager");
            roleComboBox.SelectedIndex = -1;

            LoadEmployees();
        }

        private void LoadEmployees()
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = "SELECT EmpID, Name, Password, Role, MobileNumber FROM Employee";
                SqlDataAdapter adapter = new SqlDataAdapter(query, con);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                Employeetbl.DataSource = dt;
            }
        }

        private void ClearFields()
        {
            empnametxt.Clear();
            emppwtxt.Clear();
            empmobiletxt.Clear();
            searchtxt.Clear();
            selectedEmpID = 0;
            roleComboBox.SelectedIndex = -1;
            currentPasswordHash = "";
        }

        private void Addbtn_Click_1(object sender, EventArgs e)
        {
            string name = empnametxt.Text.Trim();
            string mobile = empmobiletxt.Text.Trim();
            string password = emppwtxt.Text.Trim();
            string role = roleComboBox.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(mobile) || string.IsNullOrEmpty(role))
            {
                MessageBox.Show("Please fill in all required fields: Name, Mobile, and Role.");
                return;
            }

            if ((role == "Manager" || role == "Cashier" || role == "Admin") && string.IsNullOrEmpty(password))
            {
                MessageBox.Show($"{role} must have a password.");
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                string countQuery = "SELECT COUNT(*) FROM Employee";
                using (SqlCommand countCmd = new SqlCommand(countQuery, con))
                {
                    int count = (int)countCmd.ExecuteScalar();
                    if (count == 0)
                    {
                        using (SqlCommand resetCmd = new SqlCommand("DBCC CHECKIDENT ('Employee', RESEED, 0)", con))
                        {
                            resetCmd.ExecuteNonQuery();
                        }
                    }
                }

                object passwordParam = string.IsNullOrEmpty(password) ? (object)DBNull.Value : ComputeSha256Hash(password);

                string insertQuery = @"INSERT INTO Employee (Name, Password, MobileNumber, Role)
                                       VALUES (@Name, @Password, @MobileNumber, @Role)";
                using (SqlCommand insertCmd = new SqlCommand(insertQuery, con))
                {
                    insertCmd.Parameters.AddWithValue("@Name", name);
                    insertCmd.Parameters.AddWithValue("@Password", passwordParam);
                    insertCmd.Parameters.AddWithValue("@MobileNumber", mobile);
                    insertCmd.Parameters.AddWithValue("@Role", role);

                    insertCmd.ExecuteNonQuery();
                }

                MessageBox.Show("Employee added successfully.");
                LoadEmployees();
                ClearFields();
            }
        }

        private string ComputeSha256Hash(string rawData)
        {
            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

        private void Updatebtn_Click(object sender, EventArgs e)
        {
            if (selectedEmpID == 0)
            {
                MessageBox.Show("Please select an employee to update.");
                return;
            }

            string role = roleComboBox.SelectedItem?.ToString().Trim();
            string newPassword = emppwtxt.Text.Trim();

            if (string.IsNullOrWhiteSpace(empnametxt.Text) || string.IsNullOrWhiteSpace(empmobiletxt.Text) || string.IsNullOrEmpty(role))
            {
                MessageBox.Show("Please fill all required fields.");
                return;
            }

            // ---------- Enforce password requirement for Cashier or Manager ----------
            if ((role == "Manager" || role == "Cashier") && string.IsNullOrEmpty(newPassword))
            {
                MessageBox.Show($"{role} must have a password.");
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;

                // ---------- Only update password if user entered a new one ----------
                if (!string.IsNullOrEmpty(newPassword) && newPassword != currentPasswordHash)
                {
                    cmd.Parameters.AddWithValue("@Password", ComputeSha256Hash(newPassword));
                    cmd.CommandText = "UPDATE Employee SET Name=@Name, Password=@Password, Role=@Role, MobileNumber=@Mobile WHERE EmpID=@EmpID";
                }
                else
                {
                    // Do not change password
                    cmd.CommandText = "UPDATE Employee SET Name=@Name, Role=@Role, MobileNumber=@Mobile WHERE EmpID=@EmpID";
                }

                cmd.Parameters.AddWithValue("@Name", empnametxt.Text.Trim());
                cmd.Parameters.AddWithValue("@Role", role);
                cmd.Parameters.AddWithValue("@Mobile", empmobiletxt.Text.Trim());
                cmd.Parameters.AddWithValue("@EmpID", selectedEmpID);

                cmd.ExecuteNonQuery();
            }

            MessageBox.Show("Employee updated successfully.");
            LoadEmployees();
            ClearFields();
        }

        private void Deletebtn_Click(object sender, EventArgs e)
        {
            if (selectedEmpID == 0)
            {
                MessageBox.Show("Please select an employee to delete.");
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                string deleteQuery = "DELETE FROM Employee WHERE EmpID = @EmpID";
                using (SqlCommand deleteCmd = new SqlCommand(deleteQuery, con))
                {
                    deleteCmd.Parameters.AddWithValue("@EmpID", selectedEmpID);
                    deleteCmd.ExecuteNonQuery();
                }

                string countQuery = "SELECT COUNT(*) FROM Employee";
                using (SqlCommand countCmd = new SqlCommand(countQuery, con))
                {
                    int count = (int)countCmd.ExecuteScalar();
                    if (count == 0)
                    {
                        using (SqlCommand resetCmd = new SqlCommand("DBCC CHECKIDENT ('Employee', RESEED, 0)", con))
                        {
                            resetCmd.ExecuteNonQuery();
                        }
                    }
                }

                MessageBox.Show("Employee deleted successfully.");
                LoadEmployees();
                ClearFields();
            }
        }

        private void Clearbtn_Click(object sender, EventArgs e) => ClearFields();

        private void Searchbtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(searchtxt.Text) || !int.TryParse(searchtxt.Text, out int empId))
            {
                MessageBox.Show("Please enter a valid EmpID to search.");
                return;
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query = "SELECT EmpID, Name, Password, Role, MobileNumber FROM Employee WHERE EmpID = @EmpID";
                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@EmpID", empId);
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        Employeetbl.DataSource = dt;
                        selectedEmpID = Convert.ToInt32(dt.Rows[0]["EmpID"]);
                        empnametxt.Text = dt.Rows[0]["Name"].ToString();
                        empmobiletxt.Text = dt.Rows[0]["MobileNumber"].ToString();
                        roleComboBox.SelectedItem = dt.Rows[0]["Role"].ToString();

                        // ---------- Show hashed password in textbox ----------
                        currentPasswordHash = dt.Rows[0]["Password"] == DBNull.Value ? "" : dt.Rows[0]["Password"].ToString();
                        emppwtxt.Text = currentPasswordHash;
                    }
                    else
                    {
                        MessageBox.Show("Employee not found.");
                        ClearFields();
                    }
                }
            }
        }

        private void Employeetbl_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = Employeetbl.Rows[e.RowIndex];
                selectedEmpID = Convert.ToInt32(row.Cells["EmpID"].Value);
                empnametxt.Text = row.Cells["Name"].Value?.ToString();
                empmobiletxt.Text = row.Cells["MobileNumber"].Value?.ToString();
                roleComboBox.SelectedItem = row.Cells["Role"].Value?.ToString();

                // ---------- Show hashed password in textbox ----------
                currentPasswordHash = row.Cells["Password"].Value == DBNull.Value ? "" : row.Cells["Password"].Value?.ToString();
                emppwtxt.Text = currentPasswordHash;
            }
        }

        private void SearchEmployees(string searchText)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                string query;
                SqlDataAdapter adapter;

                if (int.TryParse(searchText, out int empId))
                {
                    query = @"SELECT EmpID, Name, Password, Role, MobileNumber FROM Employee WHERE EmpID = @empId";
                    adapter = new SqlDataAdapter(query, con);
                    adapter.SelectCommand.Parameters.AddWithValue("@empId", empId);
                }
                else
                {
                    query = @"SELECT EmpID, Name, Password, Role, MobileNumber FROM Employee 
                              WHERE Name LIKE @search OR MobileNumber LIKE @search";
                    adapter = new SqlDataAdapter(query, con);
                    adapter.SelectCommand.Parameters.AddWithValue("@search", "%" + searchText + "%");
                }

                DataTable dt = new DataTable();
                adapter.Fill(dt);
                Employeetbl.DataSource = dt;
            }
        }

        private void searchbtn_Click_1(object sender, EventArgs e)
        {
            string searchText = searchtxt.Text.Trim();
            SearchEmployees(searchText);
        }

        private void showAllBtn_Click(object sender, EventArgs e)
        {
            LoadEmployees();
            searchtxt.Text = "";
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
            Cashier cashier = new Cashier(currentUserRole, currentUsername);
            cashier.Show();
        }

        private void billbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            BillDetails billDetails = new BillDetails(currentUserRole, currentUsername);
            billDetails.Show();
        }

        private void logoutbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login login = new Login();
            login.Show();
        }

        private void empnametxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                emppwtxt.Focus();
            }
        }
    }
}
