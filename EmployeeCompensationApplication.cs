using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmployeeCompensationApplication
{
    public partial class Form1 : Form
    {
        string connectionString = "Server=DESKTOP-NIKS9CH;Database=employee;Trusted_Connection=True;";

        ComboBox comboBoxRole, comboBoxLocation;
        TextBox txtIncrement;
        Button btnLoadData, btnApplyIncrement, btnExport, btnGroupExperience, btnSave, btnExportCsv;
        CheckBox chkIncludeInactive;
        Label lblSummary, lblMin, lblMax, lblTotal;
        DataGridView dataGrid;
        Chart chart;

        public Form1()
        {
            InitializeComponent();
            BuildUI();
            this.Load += Form1_Load;
        }

        private void BuildUI()
        {
            this.Text = "Employee Compensation App";
            this.Size = new Size(1000, 800);

            comboBoxRole = new ComboBox() { Location = new Point(20, 20), Width = 150 };
            comboBoxLocation = new ComboBox() { Location = new Point(180, 20), Width = 150 };
            chkIncludeInactive = new CheckBox() { Text = "Include Inactive", Location = new Point(340, 20), Width = 130 };

            btnLoadData = new Button() { Text = "Load Data", Location = new Point(480, 20) };
            btnLoadData.Click += BtnLoadData_Click;

            txtIncrement = new TextBox() { Location = new Point(20, 60), Width = 100 };
            btnApplyIncrement = new Button() { Text = "Apply %", Location = new Point(130, 60) };
            btnApplyIncrement.Click += BtnApplyIncrement_Click;

            btnExport = new Button() { Text = "Export to Excel", Location = new Point(250, 60) };
            btnExport.Click += BtnExport_Click;

            btnGroupExperience = new Button() { Text = "Group by", Location = new Point(370, 60) };
            btnGroupExperience.Click += BtnGroupExperience_Click;

            btnSave = new Button() { Text = "Save to DB", Location = new Point(480, 60) };
            btnSave.Click += BtnSave_Click;

            btnExportCsv = new Button() { Text = "Export to CSV", Location = new Point(600, 60) };
            btnExportCsv.Click += BtnExportCsv_Click;

            lblSummary = new Label() { Location = new Point(720, 60), AutoSize = true };
            lblMin = new Label() { Location = new Point(20, 380), AutoSize = true };
            lblMax = new Label() { Location = new Point(300, 380), AutoSize = true };
            lblTotal = new Label() { Location = new Point(600, 380), AutoSize = true };

            dataGrid = new DataGridView()
            {
                Location = new Point(20, 100),
                Size = new Size(940, 270),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            chart = new Chart()
            {
                Location = new Point(20, 420),
                Size = new Size(940, 300)
            };
            ChartArea area = new ChartArea();
            chart.ChartAreas.Add(area);

            this.Controls.AddRange(new Control[] {
                comboBoxRole, comboBoxLocation, chkIncludeInactive,
                btnLoadData, txtIncrement, btnApplyIncrement,
                btnExport, btnGroupExperience, btnSave, btnExportCsv,
                lblSummary, lblMin, lblMax, lblTotal,
                dataGrid, chart
            });
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                SqlDataAdapter roleAdapter = new SqlDataAdapter("SELECT Role_Name FROM Role", con);
                DataTable roleTable = new DataTable();
                roleAdapter.Fill(roleTable);
                comboBoxRole.DataSource = roleTable;
                comboBoxRole.DisplayMember = "Role_Name";

                SqlDataAdapter locAdapter = new SqlDataAdapter("SELECT Location_Name FROM Location", con);
                DataTable locTable = new DataTable();
                locAdapter.Fill(locTable);
                comboBoxLocation.DataSource = locTable;
                comboBoxLocation.DisplayMember = "Location_Name";
            }
        }

        private void BtnLoadData_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                string query = @"SELECT e.Employee_ID, e.Name, r.Role_Name, l.Location_Name,
                                e.Years_of_Experience, ea.Compensation AS Current_Compensation, ea.Compensation AS Updated_Compensation
                                FROM Employee e
                                JOIN Employee_Assignment ea ON e.Employee_ID = ea.Employee_ID
                                JOIN Role r ON ea.Role_ID = r.Role_ID
                                JOIN Location l ON ea.Location_ID = l.Location_ID
                                WHERE r.Role_Name = @Role AND l.Location_Name = @Location";

                if (!chkIncludeInactive.Checked)
                    query += " AND ea.Active = 'Y'";

                SqlCommand cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@Role", comboBoxRole.Text);
                cmd.Parameters.AddWithValue("@Location", comboBoxLocation.Text);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dataGrid.DataSource = dt;

                if (dt.Rows.Count > 0)
                {
                    lblSummary.Text = $"Employees: {dt.Rows.Count}, Avg Salary: ₹{dt.AsEnumerable().Average(r => Convert.ToDouble(r["Current_Compensation"])):N0}";
                    lblMin.Text = $"Min: ₹{dt.AsEnumerable().Min(r => Convert.ToDouble(r["Current_Compensation"])):N0}";
                    lblMax.Text = $"Max: ₹{dt.AsEnumerable().Max(r => Convert.ToDouble(r["Current_Compensation"])):N0}";
                    lblTotal.Text = $"Total: ₹{dt.AsEnumerable().Sum(r => Convert.ToDouble(r["Current_Compensation"])):N0}";
                    DrawChart(dt);
                }
                else
                {
                    lblSummary.Text = "No records found.";
                    chart.Series.Clear();
                }
            }
        }

        private void BtnApplyIncrement_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(txtIncrement.Text, out double percent))
            {
                MessageBox.Show("Please enter a valid percentage.");
                return;
            }

            if (dataGrid.DataSource is DataTable dt)
            {
                foreach (DataRow row in dt.Rows)
                {
                    double curr = Convert.ToDouble(row["Current_Compensation"]);
                    row["Updated_Compensation"] = Math.Round(curr * (1 + percent / 100), 2);
                }

                lblSummary.Text = $"Updated Avg: ₹{dt.AsEnumerable().Average(r => Convert.ToDouble(r["Updated_Compensation"])):N0}";
                DrawChart(dt);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (dataGrid.Rows.Count > 0)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel._Worksheet worksheet = workbook.Sheets[1];
                worksheet.Name = "Export";

                for (int i = 0; i < dataGrid.Columns.Count; i++)
                    worksheet.Cells[1, i + 1] = dataGrid.Columns[i].HeaderText;

                for (int i = 0; i < dataGrid.Rows.Count; i++)
                    for (int j = 0; j < dataGrid.Columns.Count; j++)
                        worksheet.Cells[i + 2, j + 1] = dataGrid.Rows[i].Cells[j].Value?.ToString();

                excelApp.Visible = true;
            }
            else
            {
                MessageBox.Show("No data to export.");
            }
        }

        private void BtnExportCsv_Click(object sender, EventArgs e)
        {
            if (dataGrid.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV file (*.csv)|*.csv";
                sfd.Title = "Save CSV File";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sfd.FileName))
                    {
                        for (int i = 0; i < dataGrid.Columns.Count; i++)
                        {
                            sw.Write(dataGrid.Columns[i].HeaderText);
                            if (i < dataGrid.Columns.Count - 1)
                                sw.Write(",");
                        }
                        sw.WriteLine();

                        foreach (DataGridViewRow row in dataGrid.Rows)
                        {
                            for (int i = 0; i < dataGrid.Columns.Count; i++)
                            {
                                sw.Write(row.Cells[i].Value?.ToString());
                                if (i < dataGrid.Columns.Count - 1)
                                    sw.Write(",");
                            }
                            sw.WriteLine();
                        }
                    }

                    MessageBox.Show("Data exported to CSV successfully.");
                }
            }
        }

        private void BtnGroupExperience_Click(object sender, EventArgs e)
        {
            if (dataGrid.DataSource is DataTable dt && dt.Rows.Count > 0)
            {
                chart.Series.Clear();
                Series series = new Series("Experience Groups") { ChartType = SeriesChartType.Column };

                var groups = new[]
                {
                    new { Label = "0–1",   Min = 0.0, Max = 1.0 },
                    new { Label = "1–2",   Min = 1.0, Max = 2.0 },
                    new { Label = "2–5",   Min = 2.0, Max = 5.0 },
                    new { Label = "5+",    Min = 5.0, Max = double.MaxValue }
                };

                foreach (var group in groups)
                {
                    int count = dt.AsEnumerable()
                        .Count(row =>
                        {
                            double exp = Convert.ToDouble(row["Years_of_Experience"]);
                            return exp > group.Min && exp <= group.Max;
                        });

                    series.Points.AddXY(group.Label, count);
                }

                chart.Series.Add(series);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (dataGrid.DataSource is DataTable dt && dt.Rows.Count > 0)
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    foreach (DataRow row in dt.Rows)
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE Employee_Assignment SET Compensation = @Comp WHERE Employee_ID = @EmpID", con);
                        cmd.Parameters.AddWithValue("@Comp", row["Updated_Compensation"]);
                        cmd.Parameters.AddWithValue("@EmpID", row["Employee_ID"]);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Updated compensation saved to database successfully.");
            }
        }

        private void DrawChart(DataTable dt)
        {
            chart.Series.Clear();
            Series series = new Series("Updated Compensation") { ChartType = SeriesChartType.Column };

            var grouped = dt.AsEnumerable()
                .GroupBy(r => r["Location_Name"].ToString())
                .Select(g => new
                {
                    Location = g.Key,
                    Avg = g.Average(r => Convert.ToDouble(r["Updated_Compensation"]))
                });

            foreach (var item in grouped)
                series.Points.AddXY(item.Location, item.Avg);

            chart.Series.Add(series);
        }
    }
}
