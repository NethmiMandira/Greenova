using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Geenova
{
    public partial class Cashier : Form
    {
        string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\mydb.mdf;Integrated Security=True;Connect Timeout=30";

        private string selectedProductId = null;
        private string selectedProductName = null;
        private decimal selectedProductPrice = 0;
        private int selectedBillRowIndex = -1;

        // Variables for printing
        private DataTable billDataForPrinting;
        private string billNumberForPrinting;
        private string billDateForPrinting;
        private string totalAmountForPrinting;
        private string discountForPrinting;
        private string netTotalForPrinting;
        private string paymentForPrinting;
        private string balanceForPrinting;
        private string sellerRoleForPrinting;
        private string sellerNameForPrinting;
        private string currentUserRole;
        private string currentUsername;


        public Cashier(string userRole, string username)
        {
            InitializeComponent();
            Load += Cashier_Load;
            ProductCatTable.CellClick += ProductCatTable_CellClick;
            BillTbl.CellClick += BillTbl_CellClick;
            distxt.KeyPress += NumericTextBox_KeyPress;
            Paymenttxt.KeyPress += NumericTextBox_KeyPress;
            proqtytxt.KeyPress += NumericTextBox_KeyPress;

            // Setup print document
            printDocument1.PrintPage += PrintDocument1_PrintPage;

            currentUserRole = userRole;
            currentUsername = username;

            // Display both username and role in the label
            rolelbl.Text = currentUserRole;
            rolenamelbl.Text = currentUsername;
        }

        public Cashier()
        {
        }

        private string GenerateUniqueBillNumber()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT MAX(BillNumber) FROM BillTbl";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        object result = cmd.ExecuteScalar();
                        int lastNumber = 0;

                        if (result != DBNull.Value)
                            int.TryParse(result.ToString(), out lastNumber);

                        lastNumber++;
                        if (lastNumber > 9999) lastNumber = 1; // wrap around if > 9999

                        return lastNumber.ToString("D4"); // always 4 digits
                    }
                }
            }
            catch
            {
                return "0001"; // fallback in case of error
            }
        }


        private void Cashier_Load(object sender, EventArgs e)
        {
            LoadProductCatalog();
            DateTimePicker.Value = DateTime.Now;

            procatbtn.Visible = currentUserRole != "Cashier";
            billbtn.Visible = currentUserRole != "Cashier";
            empbtn.Visible = currentUserRole != "Cashier";
            cashierbtn.Visible = currentUserRole != "Cashier";
        }

        private void LoadProductCatalog()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    // Still select ProQTY for internal use but don't display it
                    string query = "SELECT ProID, ProName, Price, ProQTY FROM producttbl WHERE ProQTY > 0";
                    SqlDataAdapter adapter = new SqlDataAdapter(query, con);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    ProductCatTable.DataSource = dt;

                    // Hide the ProQTY column after binding
                    if (ProductCatTable.Columns.Contains("ProQTY"))
                    {
                        ProductCatTable.Columns["ProQTY"].Visible = false;
                    }

                    ProductCatTable_DataBindingComplete(null, null);

                    // Highlight low stock items (quantity <= 5)
                    foreach (DataGridViewRow row in ProductCatTable.Rows)
                    {
                        if (row.Cells["ProQTY"].Value != null &&
                            int.TryParse(row.Cells["ProQTY"].Value.ToString(), out int qty))
                        {
                            if (qty <= 5)
                            {
                                row.DefaultCellStyle.ForeColor = Color.MediumVioletRed;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading product data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductCatTable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && ProductCatTable.Rows[e.RowIndex].Cells["ProID"].Value != null)
            {
                DataGridViewRow row = ProductCatTable.Rows[e.RowIndex];
                selectedProductId = row.Cells["ProID"].Value.ToString();
                selectedProductName = row.Cells["ProName"].Value.ToString();
                selectedProductPrice = SafeParseDecimal(row.Cells["Price"].Value.ToString());

                proidtxt.Text = selectedProductId;
                pronametxt.Text = selectedProductName;
                pricetxt.Text = selectedProductPrice.ToString("N2");
                proqtytxt.Focus();
            }
        }

        private void AddToBillBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedProductId))
            {
                MessageBox.Show("Please select a product first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!int.TryParse(proqtytxt.Text.Trim(), out int quantity) || quantity <= 0)
            {
                MessageBox.Show("Please enter a valid quantity.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Check stock availability
            int availableQty = GetProductQuantity(selectedProductId);
            if (availableQty < quantity)
            {
                MessageBox.Show($"Only {availableQty} items available in stock.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal total = quantity * selectedProductPrice;

            if (BillTbl.DataSource == null)
            {
                DataTable billTable = new DataTable();
                billTable.Columns.Add("Product ID");
                billTable.Columns.Add("Product Name");
                billTable.Columns.Add("Quantity");
                billTable.Columns.Add("Price");
                billTable.Columns.Add("Total");

                BillTbl.DataSource = billTable;
            }

            DataTable currentTable = (DataTable)BillTbl.DataSource;
            DataRow newRow = currentTable.NewRow();
            newRow["Product ID"] = selectedProductId;
            newRow["Product Name"] = selectedProductName;
            newRow["Quantity"] = quantity;
            newRow["Price"] = selectedProductPrice.ToString("N2");
            newRow["Total"] = total.ToString("N2");

            currentTable.Rows.Add(newRow);

            // Update stock quantity in database
            UpdateProductQuantity(selectedProductId, quantity, false);

            ClearInputs();
            LoadProductCatalog();
            UpdateTotalLabel();
        }

        private int GetProductQuantity(string productId)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT ProQTY FROM producttbl WHERE ProID = @ProID";
                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@ProID", productId);
                        object result = cmd.ExecuteScalar();
                        return result != null ? Convert.ToInt32(result) : 0;
                    }
                }
            }
            catch
            {
                return 0;
            }
        }

        private void UpdateProductQuantity(string productId, int quantity, bool isReturning)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = isReturning
                        ? "UPDATE producttbl SET ProQTY = ProQTY + @Quantity WHERE ProID = @ProID"
                        : "UPDATE producttbl SET ProQTY = ProQTY - @Quantity WHERE ProID = @ProID";

                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@Quantity", quantity);
                        cmd.Parameters.AddWithValue("@ProID", productId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating product quantity: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BillTbl_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = BillTbl.Rows[e.RowIndex];
                proidtxt.Text = row.Cells["Product ID"].Value.ToString();
                pronametxt.Text = row.Cells["Product Name"].Value.ToString();
                pricetxt.Text = row.Cells["Price"].Value.ToString();
                proqtytxt.Text = row.Cells["Quantity"].Value.ToString();
                selectedBillRowIndex = e.RowIndex;
            }
        }

        private void Changeqtybtn_Click(object sender, EventArgs e)
        {
            if (selectedBillRowIndex < 0)
            {
                MessageBox.Show("Please select a product from the bill.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!int.TryParse(proqtytxt.Text.Trim(), out int newQty) || newQty <= 0)
            {
                MessageBox.Show("Please enter a valid quantity.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataTable billTable = (DataTable)BillTbl.DataSource;
            DataRow row = billTable.Rows[selectedBillRowIndex];

            string productId = row["Product ID"].ToString();
            int oldQty = Convert.ToInt32(row["Quantity"]);
            decimal price = SafeParseDecimal(row["Price"].ToString());

            // Check if we have enough stock (considering we're returning the old quantity first)
            int availableQty = GetProductQuantity(productId) + oldQty;
            if (availableQty < newQty)
            {
                MessageBox.Show($"Only {availableQty} items available in stock.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Update stock - first return old quantity, then subtract new quantity
            UpdateProductQuantity(productId, oldQty, true);
            UpdateProductQuantity(productId, newQty, false);

            decimal newTotal = newQty * price;
            row["Quantity"] = newQty;
            row["Total"] = newTotal.ToString("N2");

            selectedBillRowIndex = -1;
            ClearInputs();
            UpdateTotalLabel();
            LoadProductCatalog(); // Refresh product catalog to show updated quantities
        }

        private void Removebtn_Click(object sender, EventArgs e)
        {
            if (selectedBillRowIndex < 0)
            {
                MessageBox.Show("Please select a product to remove.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataTable billTable = (DataTable)BillTbl.DataSource;
            DataRow row = billTable.Rows[selectedBillRowIndex];

            // Return the quantity to stock
            string productId = row["Product ID"].ToString();
            int qty = Convert.ToInt32(row["Quantity"]);
            UpdateProductQuantity(productId, qty, true);

            billTable.Rows.RemoveAt(selectedBillRowIndex);

            selectedBillRowIndex = -1;
            ClearInputs();
            UpdateTotalLabel();
            LoadProductCatalog(); // Refresh product catalog to show updated quantities
        }

        private void Searchtn_Click(object sender, EventArgs e)
        {
            string searchValue = searchtxt.Text.Trim();

            if (!int.TryParse(searchValue, out int proID))
            {
                MessageBox.Show("Please enter a valid numeric Product ID.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT ProID, ProName, Price, ProQTY FROM producttbl WHERE ProID = @ProID AND ProQTY > 0";
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
                            MessageBox.Show("No product found with the given ID or product is out of stock.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ProductCatTable.DataSource = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error searching product: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductCatTable_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (ProductCatTable.Columns.Contains("ProID"))
                ProductCatTable.Columns["ProID"].HeaderText = "Product ID";

            if (ProductCatTable.Columns.Contains("ProName"))
                ProductCatTable.Columns["ProName"].HeaderText = "Product Name";

            if (ProductCatTable.Columns.Contains("Price"))
            {
                ProductCatTable.Columns["Price"].HeaderText = "Price";
                ProductCatTable.Columns["Price"].DefaultCellStyle.Format = "N2";
                ProductCatTable.Columns["Price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            // Ensure ProQTY column is hidden
            if (ProductCatTable.Columns.Contains("ProQTY"))
            {
                ProductCatTable.Columns["ProQTY"].Visible = false;
            }
        }

        private void ClearInputs()
        {
            proidtxt.Clear();
            pronametxt.Clear();
            pricetxt.Clear();
            proqtytxt.Clear();
            searchtxt.Clear();
            selectedProductId = null;
            selectedProductName = null;
            selectedProductPrice = 0;
        }

        private void UpdateTotalLabel()
        {
            decimal grandTotal = 0;
            DataTable billTable = (DataTable)BillTbl.DataSource;

            if (billTable != null)
            {
                foreach (DataRow row in billTable.Rows)
                {
                    if (decimal.TryParse(row["Total"].ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal total))
                    {
                        grandTotal += total;
                    }
                }
            }

            showtotallbl.Text = grandTotal.ToString("N2");
        }

        private void Calbtn_Click(object sender, EventArgs e)
        {
            UpdateTotalLabel();

            if (!decimal.TryParse(showtotallbl.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal total))
            {
                MessageBox.Show("Invalid total amount. Please check your bill.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal discount = 0;
            if (!string.IsNullOrWhiteSpace(distxt.Text))
            {
                if (!decimal.TryParse(distxt.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out discount) || discount < 0 || discount > 100)
                {
                    MessageBox.Show("Please enter a valid discount percentage (0-100).", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            decimal netTotal = total - (total * discount / 100);
            shownettotallbl.Text = netTotal.ToString("N2");
        }

        private void generatebillbtn_Click(object sender, EventArgs e)
        {
            if (BillTbl.Rows.Count == 0)
            {
                MessageBox.Show("Please add items to the bill first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(shownettotallbl.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal netTotal))
            {
                MessageBox.Show("Please calculate the net total first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(Paymenttxt.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal payment))
            {
                MessageBox.Show("Please enter a valid payment amount.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (payment < netTotal)
            {
                MessageBox.Show($"Payment amount must be at least {netTotal.ToString("N2")}.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal balance = payment - netTotal;
            showbalancelbl.Text = balance.ToString("N2");

            // Generate 4-digit bill number
            billNumberForPrinting = GenerateUniqueBillNumber();

            // Store data for printing and database
            billDataForPrinting = (DataTable)BillTbl.DataSource;
            billDateForPrinting = DateTimePicker.Value.ToString("dd/MM/yyyy hh:mm tt");
            totalAmountForPrinting = showtotallbl.Text;
            discountForPrinting = string.IsNullOrWhiteSpace(distxt.Text) ? "0" : distxt.Text;
            netTotalForPrinting = shownettotallbl.Text;
            paymentForPrinting = Paymenttxt.Text;
            balanceForPrinting = showbalancelbl.Text;
            sellerRoleForPrinting = rolelbl.Text;
            sellerNameForPrinting = rolenamelbl.Text;

            // Save to database
            SaveBillToDatabase(netTotal, payment, balance);

            // Show print preview
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();

            // Clear the bill after generation
            ClearBill();
        }

        private void PrintDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            // Fonts
            Font headerFont = new Font("Arial", 16, FontStyle.Bold);
            Font subHeaderFont = new Font("Arial", 12, FontStyle.Bold);
            Font bodyFont = new Font("Arial", 10);
            Font footerFont = new Font("Arial", 10, FontStyle.Bold);

            // Brushes and Pens
            Brush brush = Brushes.Black;
            Pen thinLinePen = new Pen(Color.Black, 1);
            Pen thickLinePen = new Pen(Color.Black, 2);

            // Margins and positions
            float leftMargin = e.MarginBounds.Left;
            float rightMargin = e.MarginBounds.Right;
            float yPos = 50;
            float centerX = leftMargin + (e.MarginBounds.Width / 2);
            float lineHeight = bodyFont.GetHeight(e.Graphics) + 5;

            // Column positions
            float col1 = leftMargin;               // Product ID
            float col2 = leftMargin + 60;          // Product Name
            float col3 = leftMargin + 200;         // Qty
            float col4 = col3 + 200;                // Price (gap increased)
            float col5 = rightMargin - 10;         // Total (right aligned to margin)

            // Header
            string header = "GREENOVA SUPERMARKET";
            float headerTextWidth = e.Graphics.MeasureString(header, headerFont).Width;
            e.Graphics.DrawString(header, headerFont, brush, centerX - (headerTextWidth / 2), yPos);
            yPos += lineHeight * 1.5f;

            // Bill info
            e.Graphics.DrawString($"Bill No: {billNumberForPrinting}", bodyFont, brush, leftMargin, yPos);
            string billDateStr = $"Date: {billDateForPrinting}";
            float billDateWidth = e.Graphics.MeasureString(billDateStr, bodyFont).Width;
            e.Graphics.DrawString(billDateStr, bodyFont, brush, rightMargin - billDateWidth, yPos);
            yPos += lineHeight;

            e.Graphics.DrawString($"Cashier: {sellerNameForPrinting}", bodyFont, brush, leftMargin, yPos);
            string roleStr = $"Role: {sellerRoleForPrinting}";
            float roleWidth = e.Graphics.MeasureString(roleStr, bodyFont).Width;
            e.Graphics.DrawString(roleStr, bodyFont, brush, rightMargin - roleWidth, yPos);
            yPos += lineHeight * 1.5f;

            // Draw line separator
            e.Graphics.DrawLine(thinLinePen, leftMargin, yPos, rightMargin, yPos);
            yPos += lineHeight;

            // Column headers
            e.Graphics.DrawString("ID", subHeaderFont, brush, col1, yPos);
            e.Graphics.DrawString("Item", subHeaderFont, brush, col2, yPos);
            e.Graphics.DrawString("Qty", subHeaderFont, brush, col3 + 40, yPos, new StringFormat() { Alignment = StringAlignment.Far });
            e.Graphics.DrawString("Price(Rs.)", subHeaderFont, brush, col4 + 40, yPos, new StringFormat() { Alignment = StringAlignment.Far });
            e.Graphics.DrawString("Total", subHeaderFont, brush, col5, yPos, new StringFormat() { Alignment = StringAlignment.Far });
            yPos += lineHeight;

            // Draw line separator
            e.Graphics.DrawLine(thinLinePen, leftMargin, yPos, rightMargin, yPos);
            yPos += lineHeight;

            // Right-align format
            StringFormat rightAlign = new StringFormat() { Alignment = StringAlignment.Far };

            // Items
            foreach (DataRow row in billDataForPrinting.Rows)
            {
                string productId = row["Product ID"].ToString();
                string productName = row["Product Name"].ToString();
                string qty = row["Quantity"].ToString();
                string price = row["Price"].ToString();
                string total = row["Total"].ToString();

                if (productName.Length > 20)
                {
                    productName = productName.Substring(0, 17) + "...";
                }

                e.Graphics.DrawString(productId, bodyFont, brush, col1, yPos);
                e.Graphics.DrawString(productName, bodyFont, brush, col2, yPos);
                e.Graphics.DrawString(qty, bodyFont, brush, col3 + 40, yPos, rightAlign);
                e.Graphics.DrawString(price, bodyFont, brush, col4 + 40, yPos, rightAlign);
                e.Graphics.DrawString(total, bodyFont, brush, col5, yPos, rightAlign);

                yPos += lineHeight;
            }

            // Draw line separator after items
            e.Graphics.DrawLine(thinLinePen, leftMargin, yPos, rightMargin, yPos);
            yPos += lineHeight;

            // Payment summary
            float summaryLeft = leftMargin;
            float summaryRight = rightMargin;

            // Sub Total
            e.Graphics.DrawString("Sub Total:", footerFont, brush, summaryLeft, yPos);
            float totalAmountWidth = e.Graphics.MeasureString(totalAmountForPrinting, footerFont).Width;
            e.Graphics.DrawString(totalAmountForPrinting, footerFont, brush, summaryRight - totalAmountWidth, yPos);
            yPos += lineHeight;

            e.Graphics.DrawLine(thinLinePen, summaryLeft, yPos, summaryRight, yPos);
            yPos += lineHeight;

            // Discount
            decimal discount = SafeParseDecimal(discountForPrinting);
            string discountStr = $"Discount ({discount}%):";
            e.Graphics.DrawString(discountStr, footerFont, brush, summaryLeft, yPos);
            decimal totalAmount = SafeParseDecimal(totalAmountForPrinting);
            decimal discountAmount = totalAmount * discount / 100;
            string discountAmountStr = discountAmount.ToString("N2");
            float discountAmountWidth = e.Graphics.MeasureString(discountAmountStr, footerFont).Width;
            e.Graphics.DrawString(discountAmountStr, footerFont, brush, summaryRight - discountAmountWidth, yPos);
            yPos += lineHeight;

            e.Graphics.DrawLine(thinLinePen, summaryLeft, yPos, summaryRight, yPos);
            yPos += lineHeight;

            // Net Total
            e.Graphics.DrawString("Net Total:", footerFont, brush, summaryLeft, yPos);
            float netTotalWidth = e.Graphics.MeasureString(netTotalForPrinting, footerFont).Width;
            e.Graphics.DrawString(netTotalForPrinting, footerFont, brush, summaryRight - netTotalWidth, yPos);
            yPos += lineHeight;

            // Thick line before payment
            e.Graphics.DrawLine(thickLinePen, summaryLeft, yPos, summaryRight, yPos);
            yPos += lineHeight;

            // Payment
            e.Graphics.DrawString("Payment:", footerFont, brush, summaryLeft, yPos);
            string paymentFormatted = SafeParseDecimal(paymentForPrinting).ToString("N2");
            float paymentWidth = e.Graphics.MeasureString(paymentFormatted, footerFont).Width;
            e.Graphics.DrawString(paymentFormatted, footerFont, brush, summaryRight - paymentWidth, yPos);
            yPos += lineHeight;

            e.Graphics.DrawLine(thinLinePen, summaryLeft, yPos, summaryRight, yPos);
            yPos += lineHeight;

            // Balance
            e.Graphics.DrawString("Balance:", footerFont, brush, summaryLeft, yPos);
            float balanceWidth = e.Graphics.MeasureString(balanceForPrinting, footerFont).Width;
            e.Graphics.DrawString(balanceForPrinting, footerFont, brush, summaryRight - balanceWidth, yPos);
            yPos += lineHeight * 1.5f;

            // Final thick line
            e.Graphics.DrawLine(thickLinePen, leftMargin, yPos, rightMargin, yPos);
            yPos += lineHeight * 1.5f;

            // Footer
            string thankYou = "Thank you for shopping with us!";
            float thankYouWidth = e.Graphics.MeasureString(thankYou, bodyFont).Width;
            e.Graphics.DrawString(thankYou, bodyFont, brush, centerX - (thankYouWidth / 2), yPos);
            yPos += lineHeight;

            string visitAgain = "Please visit again";
            float visitAgainWidth = e.Graphics.MeasureString(visitAgain, bodyFont).Width;
            e.Graphics.DrawString(visitAgain, bodyFont, brush, centerX - (visitAgainWidth / 2), yPos);
        }



        private void SaveBillToDatabase(decimal netTotal, decimal payment, decimal balance)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    SqlTransaction transaction = con.BeginTransaction();

                    try
                    {
                        // Parse all values with proper decimal handling
                        decimal totalAmount = SafeParseDecimal(totalAmountForPrinting);
                        decimal discount = SafeParseDecimal(discountForPrinting);
                        decimal netTotalDec = SafeParseDecimal(netTotalForPrinting);
                        decimal paymentDec = SafeParseDecimal(paymentForPrinting);
                        decimal balanceDec = SafeParseDecimal(balanceForPrinting);

                        // Insert bill header
                        string billQuery = @"
                            INSERT INTO BillTbl (
                                BillNumber,
                                BillDate,
                                TotalAmount,
                                Discount,
                                NetTotal,
                                Payment,
                                Balance,
                                SellerRole,
                                SellerName
                            ) VALUES (
                                @BillNumber,
                                @BillDate,
                                @TotalAmount,
                                @Discount,
                                @NetTotal,
                                @Payment,
                                @Balance,
                                @SellerRole,
                                @SellerName
                            )";

                        using (SqlCommand billCmd = new SqlCommand(billQuery, con, transaction))
                        {
                            billCmd.Parameters.Add("@BillNumber", SqlDbType.VarChar, 4).Value = billNumberForPrinting;
                            billCmd.Parameters.Add("@BillDate", SqlDbType.DateTime).Value = DateTimePicker.Value;
                            billCmd.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = totalAmount;
                            billCmd.Parameters.Add("@Discount", SqlDbType.Decimal).Value = discount;
                            billCmd.Parameters.Add("@NetTotal", SqlDbType.Decimal).Value = netTotalDec;
                            billCmd.Parameters.Add("@Payment", SqlDbType.Decimal).Value = paymentDec;
                            billCmd.Parameters.Add("@Balance", SqlDbType.Decimal).Value = balanceDec;
                            billCmd.Parameters.Add("@SellerRole", SqlDbType.NVarChar).Value = sellerRoleForPrinting;
                            billCmd.Parameters.Add("@SellerName", SqlDbType.NVarChar).Value = sellerNameForPrinting;

                            billCmd.ExecuteNonQuery();
                        }

                        // Insert bill items
                        string itemQuery = @"
                            INSERT INTO BillItemsTbl (
                                BillNumber,
                                ProductID,
                                ProductName,
                                Quantity,
                                Price,
                                Total
                            ) VALUES (
                                @BillNumber,
                                @ProductID,
                                @ProductName,
                                @Quantity,
                                @Price,
                                @Total
                            )";

                        foreach (DataRow row in billDataForPrinting.Rows)
                        {
                            using (SqlCommand itemCmd = new SqlCommand(itemQuery, con, transaction))
                            {
                                itemCmd.Parameters.Add("@BillNumber", SqlDbType.VarChar, 4).Value = billNumberForPrinting;
                                itemCmd.Parameters.Add("@ProductID", SqlDbType.NVarChar).Value = row["Product ID"].ToString();
                                itemCmd.Parameters.Add("@ProductName", SqlDbType.NVarChar).Value = row["Product Name"].ToString();
                                itemCmd.Parameters.Add("@Quantity", SqlDbType.Int).Value = SafeParseInt(row["Quantity"].ToString());
                                itemCmd.Parameters.Add("@Price", SqlDbType.Decimal).Value = SafeParseDecimal(row["Price"].ToString());
                                itemCmd.Parameters.Add("@Total", SqlDbType.Decimal).Value = SafeParseDecimal(row["Total"].ToString());

                                itemCmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                        MessageBox.Show("Bill saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show($"Error saving bill: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database connection error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private decimal SafeParseDecimal(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 0m;

            value = value.Replace("$", "").Replace(",", "").Trim();

            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                return result;
            }

            MessageBox.Show($"Invalid decimal value: {value} - Using 0 instead", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return 0m;
        }

        private int SafeParseInt(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 0;

            if (int.TryParse(value, out int result))
            {
                return result;
            }

            MessageBox.Show($"Invalid integer value: {value} - Using 0 instead", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return 0;
        }

        private void ClearBill()
        {
            BillTbl.DataSource = null;
            showtotallbl.Text = "0.00";
            shownettotallbl.Text = "0.00";
            showbalancelbl.Text = "0.00";
            distxt.Clear();
            Paymenttxt.Clear();
        }

        private void NumericTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
            {
                e.Handled = true;
            }
        }

        private void Resetbtn_Click(object sender, EventArgs e)
        {
            // Return all items to stock before clearing
            if (BillTbl.DataSource != null)
            {
                DataTable billTable = (DataTable)BillTbl.DataSource;
                foreach (DataRow row in billTable.Rows)
                {
                    string productId = row["Product ID"].ToString();
                    int qty = Convert.ToInt32(row["Quantity"]);
                    UpdateProductQuantity(productId, qty, true);
                }
            }

            ClearInputs();
            ClearBill();
            LoadProductCatalog();
        }

        private void procatbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            productcat productForm = new productcat(currentUserRole, currentUsername);
            productForm.Show();
        }

        private void empbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            Employee employee = new Employee(currentUserRole, currentUsername);
            employee.Show();
        }

        private void billbtn_Click(object sender, EventArgs e)
        {
            this.Hide();
            BillDetails billDetails= new BillDetails(currentUserRole, currentUsername);
            billDetails.Show();
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

        private void pricetxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                distxt.Focus();
            }
        }

        

        private void distxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Prevents beep sound
                Paymenttxt.Focus();
            }
        }

      
    }
}