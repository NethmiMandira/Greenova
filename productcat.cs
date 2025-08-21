using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Geenova
{
    public partial class productcat : Form
    {
        string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\mydb.mdf;Integrated Security=True;Connect Timeout=30";

        int selectedProductId = -1;
        private string currentUserRole;
        private string currentUsername;

        public productcat(string userRole, string username)
        {
            InitializeComponent();

            currentUserRole = userRole;
            currentUsername = username;

            // Display both username and role in the label
            rolelbl.Text =  currentUserRole;
            rolenamelbl.Text =   currentUsername;

            ProductCatTable.DataBindingComplete += ProductCatTable_DataBindingComplete;
            ProductCatTable.CellClick += ProductCatTable_CellClick;

            proidtxt.KeyPress += proidtxt_KeyPress;
        }

        public productcat() { }

        private void productcat_Load(object sender, EventArgs e)
        {
            LoadProductData();
        }

        private void LoadProductData()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT * FROM producttbl ORDER BY InsertedAt ASC";
                    SqlDataAdapter sda = new SqlDataAdapter(query, con);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    ProductCatTable.Columns.Clear();
                    ProductCatTable.AutoGenerateColumns = true;
                    ProductCatTable.DataSource = dt;

                    ProductCatTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    ProductCatTable.ColumnHeadersVisible = true;

                    foreach (DataGridViewColumn col in ProductCatTable.Columns)
                        col.SortMode = DataGridViewColumnSortMode.NotSortable;

                    if (ProductCatTable.Columns.Contains("InsertedAt"))
                        ProductCatTable.Columns["InsertedAt"].Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load product data: " + ex.Message);
            }
        }

        private void ProductCatTable_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (ProductCatTable.Columns.Contains("ProID"))
                ProductCatTable.Columns["ProID"].HeaderText = "Product ID";
            if (ProductCatTable.Columns.Contains("ProName"))
                ProductCatTable.Columns["ProName"].HeaderText = "Product Name";
            if (ProductCatTable.Columns.Contains("Description"))
                ProductCatTable.Columns["Description"].HeaderText = "Description";
            if (ProductCatTable.Columns.Contains("ProQTY"))
                ProductCatTable.Columns["ProQTY"].HeaderText = "Quantity";

            if (ProductCatTable.Columns.Contains("Price"))
            {
                ProductCatTable.Columns["Price"].HeaderText = "Price";
                ProductCatTable.Columns["Price"].DefaultCellStyle.Format = "N2";
            }

            if (ProductCatTable.Columns.Contains("AddedDate"))
                ProductCatTable.Columns["AddedDate"].HeaderText = "Date Added";

            foreach (DataGridViewRow row in ProductCatTable.Rows)
            {
                if (row.Cells["ProQTY"].Value != null &&
                    int.TryParse(row.Cells["ProQTY"].Value.ToString(), out int qty) &&
                    qty == 0)
                {
                    row.DefaultCellStyle.ForeColor = Color.Red;
                }
            }
        }

        private void ProductCatTable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = ProductCatTable.Rows[e.RowIndex];

                proidtxt.Text = row.Cells["ProID"].Value?.ToString();
                pronametxt.Text = row.Cells["ProName"].Value?.ToString();
                desctxt.Text = row.Cells["Description"].Value?.ToString();
                proqtytxt.Text = row.Cells["ProQTY"].Value?.ToString();
                pricetxt.Text = row.Cells["Price"].Value?.ToString();

                if (int.TryParse(row.Cells["ProID"].Value?.ToString(), out int id))
                    selectedProductId = id;

                if (DateTime.TryParse(row.Cells["AddedDate"].Value?.ToString(), out DateTime dateValue))
                    DateTimePicker.Value = dateValue;
                else
                    DateTimePicker.Value = DateTime.Now;
            }
        }

        private void Addbtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(proidtxt.Text) ||
                string.IsNullOrWhiteSpace(pronametxt.Text) ||
                string.IsNullOrWhiteSpace(proqtytxt.Text) ||
                string.IsNullOrWhiteSpace(pricetxt.Text))
            {
                MessageBox.Show("Please enter Product ID, Name, Quantity, and Price.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = @"INSERT INTO producttbl (ProID, ProName, Description, ProQTY, Price, AddedDate)
                                     VALUES (@ProID, @ProName, @Description, @ProQTY, @Price, @AddedDate)";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ProID", int.Parse(proidtxt.Text.Trim()));
                        cmd.Parameters.AddWithValue("@ProName", pronametxt.Text.Trim());
                        cmd.Parameters.AddWithValue("@Description", desctxt.Text.Trim());
                        cmd.Parameters.AddWithValue("@ProQTY", int.Parse(proqtytxt.Text.Trim()));
                        cmd.Parameters.AddWithValue("@Price", double.Parse(pricetxt.Text.Trim()));
                        cmd.Parameters.AddWithValue("@AddedDate", DateTimePicker.Value.Date);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            MessageBox.Show("Product added successfully!");
                            AddNewRowToDataGridView();
                            ClearInputFields();
                        }
                        else
                        {
                            MessageBox.Show("Failed to add product.");
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Number == 2627)
            {
                MessageBox.Show("A product with this ID already exists.", "Duplicate ID", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddNewRowToDataGridView()
        {
            if (ProductCatTable.DataSource is DataTable dt)
            {
                DataRow newRow = dt.NewRow();
                newRow["ProID"] = int.Parse(proidtxt.Text.Trim());
                newRow["ProName"] = pronametxt.Text.Trim();
                newRow["Description"] = desctxt.Text.Trim();
                newRow["ProQTY"] = int.Parse(proqtytxt.Text.Trim());
                newRow["Price"] = double.Parse(pricetxt.Text.Trim());
                newRow["AddedDate"] = DateTimePicker.Value.Date;

                dt.Rows.Add(newRow);
                dt.AcceptChanges();

                int lastRowIndex = ProductCatTable.Rows.Count - 1;
                if (lastRowIndex >= 0)
                    ProductCatTable.FirstDisplayedScrollingRowIndex = lastRowIndex;
            }
        }

        private void Updatebtn_Click(object sender, EventArgs e)
        {
            if (selectedProductId == -1)
            {
                MessageBox.Show("Please select a product to update.");
                return;
            }

            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = @"UPDATE producttbl
                                     SET ProName = @ProName,
                                         Description = @Description,
                                         ProQTY = @ProQTY,
                                         Price = @Price,
                                         AddedDate = @AddedDate
                                     WHERE ProID = @ProID";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ProName", pronametxt.Text.Trim());
                        cmd.Parameters.AddWithValue("@Description", desctxt.Text.Trim());
                        cmd.Parameters.AddWithValue("@ProQTY", int.Parse(proqtytxt.Text.Trim()));
                        cmd.Parameters.AddWithValue("@Price", double.Parse(pricetxt.Text.Trim()));
                        cmd.Parameters.AddWithValue("@AddedDate", DateTimePicker.Value.Date);
                        cmd.Parameters.AddWithValue("@ProID", selectedProductId);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Product updated successfully.");
                            LoadProductData();
                            ClearInputFields();
                        }
                        else
                        {
                            MessageBox.Show("Failed to update product.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating product: " + ex.Message);
            }
        }

        private void Remove_Click(object sender, EventArgs e)
        {
            if (selectedProductId == -1)
            {
                MessageBox.Show("Please select a row to delete.");
                return;
            }

            DialogResult confirm = MessageBox.Show("Are you sure you want to delete this product?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        con.Open();
                        string deleteQuery = "DELETE FROM producttbl WHERE ProID = @ProID";
                        using (SqlCommand cmd = new SqlCommand(deleteQuery, con))
                        {
                            cmd.Parameters.AddWithValue("@ProID", selectedProductId);
                            cmd.ExecuteNonQuery();
                        }

                        MessageBox.Show("Product deleted successfully.");
                        LoadProductData();
                        ClearInputFields();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting product: " + ex.Message);
                }
            }
        }

        private void ClearInputFields()
        {
            proidtxt.Text = "";
            pronametxt.Text = "";
            desctxt.Text = "";
            proqtytxt.Text = "";
            pricetxt.Text = "";
            DateTimePicker.Value = DateTime.Now;
            selectedProductId = -1;
            ProductCatTable.ClearSelection();
            searchtxt.Text = "";
        }

        private void Clearbtn_Click(object sender, EventArgs e)
        {
            ClearInputFields();
        }

        private void Searchtn_Click(object sender, EventArgs e)
        {
            string searchValue = searchtxt.Text.Trim();

            if (string.IsNullOrWhiteSpace(searchValue))
            {
                MessageBox.Show("Please enter a Product ID to search.");
                return;
            }

            if (!int.TryParse(searchValue, out int proID))
            {
                MessageBox.Show("Please enter a valid numeric Product ID.");
                return;
            }

            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT * FROM producttbl WHERE ProID = @ProID";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ProID", proID);
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            ProductCatTable.DataSource = dt;
                            ProductCatTable_DataBindingComplete(null, null);
                        }
                        else
                        {
                            MessageBox.Show("No product found with the given ID.");
                            ProductCatTable.DataSource = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error searching product: " + ex.Message);
            }
        }

        private void proidtxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
                e.Handled = true;
        }

        private void showAllBtn_Click_1(object sender, EventArgs e)
        {
            LoadProductData();
            searchtxt.Text = "";
        }

        private void procatbtn_Click(object sender, EventArgs e)
        {
          
        }

        private void cashierbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Cashier c = new Cashier(currentUserRole, currentUsername);
            c.Show();
        }

        private void empbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Employee emp=new Employee(currentUserRole, currentUsername);
            emp.Show();

        }

        private void billbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            BillDetails b=new BillDetails(currentUserRole, currentUsername);
            b.Show();
        }

        private void logoutbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login loginForm = new Login();
            loginForm.Show();
        }

        private void proidtxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                pronametxt.Focus();
            }
        }

        private void pronametxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                desctxt.Focus();
            }
        }

        private void desctxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                proqtytxt.Focus();
            }
        }

        private void proqtytxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                pricetxt.Focus();
            }
        }
    }
}
