using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;

namespace Geenova
{
    public partial class Login : Form
    {
        private  string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\mydb.mdf;Integrated Security=True;Connect Timeout=30";



        public Login()
        {
            InitializeComponent();

            roleselectioncombobox.Items.AddRange(new string[] { "Cashier", "Manager" });
            roleselectioncombobox.SelectedIndex = -1;
        }

        
        

        private bool ValidateUser(string username, string password, string role)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                string query = @"SELECT EmpID, Name, Password, Role 
                 FROM Employee 
                 WHERE Name = @Name AND Role = @Role";


                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@Name", username);
                    cmd.Parameters.AddWithValue("@Role", role);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (!reader.Read())
                            return false;

                        string storedHash = reader["Password"].ToString();

                        if (storedHash == ComputeSha256Hash(password))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private string ComputeSha256Hash(string rawData)
        {
            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                StringBuilder builder = new StringBuilder();
                foreach (var b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

       



       

        private void loginbtn_Click_1(object sender, EventArgs e)
        {
            string username = usernametxt.Text.Trim();
            string password = passwordtxt.Text.Trim();
            string selectedRole = roleselectioncombobox.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(username) ||
                string.IsNullOrEmpty(password) ||
                string.IsNullOrEmpty(selectedRole))
            {
                MessageBox.Show("Please enter all required fields.");
                return;
            }

            bool isValidLogin = ValidateUser(username, password, selectedRole);
            if (!isValidLogin)
            {
                MessageBox.Show("Invalid username, password, or role.");
                return;
            }

            this.Hide();

            if (selectedRole == "Cashier")
            {
                Cashier cash = new Cashier(selectedRole, username);
                cash.FormClosed += (s, args) => this.Close();
                cash.Show();
            }
            else if (selectedRole == "Manager")
            {
                productcat productForm = new productcat(selectedRole, username);
                productForm.FormClosed += (s, args) => this.Close();
                productForm.Show();
            }
        }

        private void usernametxt_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                passwordtxt.Focus();
            }
        }

        private void clearbtn_Click_1(object sender, EventArgs e)
        {
            usernametxt.Clear();
            passwordtxt.Clear();
            roleselectioncombobox.SelectedIndex = -1;
        }

        private void Closebtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
