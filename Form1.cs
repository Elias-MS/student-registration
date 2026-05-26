using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Student_and_Teacher_manegement
{
    public partial class Form1 : Form
    {
        private string connectionString;
        private int selectedMarkID = -1;

        public Form1()
        {
            InitializeComponent();
        }

        // =========================================
        // FORM LOAD
        // =========================================
        private void Form1_Load(object sender, EventArgs e)
        {
            // 1. Auto-discover SQL instance and initialize database/tables
            InitializeDatabase();

            // 2. Populate dropdown values (merging database values with defaults)
            RefreshDropdownData();

            // 3. Load existing marks into teacher tab grid
            LoadMarks();

            // 4. Hook up real-time total mark calculations
            txtAssignment.TextChanged += (s, ev) => CalculateTotalMarks();
            txtMid.TextChanged += (s, ev) => CalculateTotalMarks();
            txtFinal.TextChanged += (s, ev) => CalculateTotalMarks();

            // 5. Add direct CellClick listener for improved grid responsiveness
            dgvMarks.CellClick += (s, ev) => PopulateTeacherFieldsFromGrid();

            // 6. Style all components programmatically for a premium theme
            StyleControls();
        }

        // =========================================
        // DATABASE BOOTSTRAPPING
        // =========================================
        private void InitializeDatabase()
        {
            // List of common local SQL Server connection strings
            string[] masterConnStrings = new string[] {
                @"Data Source=.\SQLEXPRESS;Initial Catalog=master;Integrated Security=True;TrustServerCertificate=True",
                @"Data Source=.;Initial Catalog=master;Integrated Security=True;TrustServerCertificate=True",
                @"Data Source=.\SQLEXPRESS01;Initial Catalog=master;Integrated Security=True;TrustServerCertificate=True",
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=master;Integrated Security=True"
            };

            SqlConnection conn = null;
            string workingMaster = null;

            // Find a working SQL instance
            foreach (var str in masterConnStrings)
            {
                try
                {
                    conn = new SqlConnection(str);
                    conn.Open();
                    workingMaster = str;
                    break;
                }
                catch
                {
                    if (conn != null) conn.Dispose();
                }
            }

            if (workingMaster == null)
            {
                // Default fallback if SQL server isn't running or accessible
                connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=StudentDB;Integrated Security=True;TrustServerCertificate=True";
                return;
            }

            try
            {
                // Create database if not exists
                string dbCreateQuery = "IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'StudentDB') CREATE DATABASE StudentDB;";
                using (SqlCommand cmd = new SqlCommand(dbCreateQuery, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error checking/creating database: " + ex.Message, "Database Setup Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            // Derive connection string targeting our database
            connectionString = workingMaster.Replace("Initial Catalog=master", "Initial Catalog=StudentDB");

            // Build schema
            try
            {
                using (SqlConnection dbConn = new SqlConnection(connectionString))
                {
                    dbConn.Open();

                    string createStudents = @"
                        IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Students]') AND type in (N'U'))
                        BEGIN
                            CREATE TABLE [dbo].[Students] (
                                [StudID] VARCHAR(50) NOT NULL PRIMARY KEY,
                                [FirstName] NVARCHAR(100) NOT NULL,
                                [LastName] NVARCHAR(100) NOT NULL,
                                [Year] NVARCHAR(50) NULL,
                                [Faculty] NVARCHAR(100) NULL,
                                [Department] NVARCHAR(100) NULL
                            );
                        END";

                    string createCourses = @"
                        IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Courses]') AND type in (N'U'))
                        BEGIN
                            CREATE TABLE [dbo].[Courses] (
                                [CourseID] VARCHAR(50) NOT NULL PRIMARY KEY,
                                [CourseName] NVARCHAR(100) NOT NULL,
                                [Faculty] NVARCHAR(100) NOT NULL,
                                [Department] NVARCHAR(100) NOT NULL,
                                [Year] NVARCHAR(50) NULL
                            );
                        END";

                    string createMarks = @"
                        IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Marks]') AND type in (N'U'))
                        BEGIN
                            CREATE TABLE [dbo].[Marks] (
                                [MarkID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
                                [StudID] VARCHAR(50) NOT NULL,
                                [Year] NVARCHAR(50) NULL,
                                [Faculty] NVARCHAR(100) NULL,
                                [Department] NVARCHAR(100) NULL,
                                [CourseName] NVARCHAR(100) NOT NULL,
                                [Assignments] FLOAT NOT NULL DEFAULT 0,
                                [MidExam] FLOAT NOT NULL DEFAULT 0,
                                [FinalExam] FLOAT NOT NULL DEFAULT 0,
                                [Total] FLOAT NOT NULL DEFAULT 0
                            );
                        END";

                    using (SqlCommand cmd = new SqlCommand(createStudents, dbConn)) cmd.ExecuteNonQuery();
                    using (SqlCommand cmd = new SqlCommand(createCourses, dbConn)) cmd.ExecuteNonQuery();
                    using (SqlCommand cmd = new SqlCommand(createMarks, dbConn)) cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing tables: " + ex.Message, "Schema Setup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================
        // DROPDOWN DATA MANAGEMENT
        // =========================================
        private void RefreshDropdownData()
        {
            var fallbackYears = new string[] { "1st year", "2nd year", "3rd year", "4th year" };
            var fallbackFaculties = new string[] { "Computing & IT", "Engineering", "Business & Economics", "Health Sciences" };
            var fallbackDepts = new string[] { "Computer Science", "Information Systems", "Software Engineering", "Electrical Engineering", "Management", "Finance" };
            var fallbackCourses = new string[] { "Advanced Programming", "Database Systems", "Software Engineering II", "Web Development", "Computer Networks" };

            cmbYear.Items.Clear();
            cmbMarkYear.Items.Clear();
            cmbTab2Year.Items.Clear();

            cmbFaculty.Items.Clear();
            cmbMarkFaculty.Items.Clear();

            cmbDept.Items.Clear();
            cmbMarkDept.Items.Clear();

            cmbMarkCourse.Items.Clear();

            cmbYear.Items.AddRange(fallbackYears);
            cmbMarkYear.Items.AddRange(fallbackYears);
            cmbTab2Year.Items.AddRange(fallbackYears);
            
            cmbYear.SelectedIndex = 0;
            cmbMarkYear.SelectedIndex = 0;
            cmbTab2Year.SelectedIndex = 0;

            List<string> faculties = new List<string>();
            List<string> departments = new List<string>();
            List<string> courses = new List<string>();

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand("SELECT DISTINCT Faculty FROM Courses UNION SELECT DISTINCT Faculty FROM Students", conn))
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            string f = dr[0].ToString();
                            if (!string.IsNullOrWhiteSpace(f)) faculties.Add(f);
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT DISTINCT Department FROM Courses UNION SELECT DISTINCT Department FROM Students", conn))
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            string d = dr[0].ToString();
                            if (!string.IsNullOrWhiteSpace(d)) departments.Add(d);
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT DISTINCT CourseName FROM Courses", conn))
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            string c = dr[0].ToString();
                            if (!string.IsNullOrWhiteSpace(c)) courses.Add(c);
                        }
                    }
                }
            }
            catch
            {
                // Fallback will cover it
            }

            foreach (var f in fallbackFaculties) if (!faculties.Contains(f)) faculties.Add(f);
            foreach (var d in fallbackDepts) if (!departments.Contains(d)) departments.Add(d);
            foreach (var c in fallbackCourses) if (!courses.Contains(c)) courses.Add(c);

            cmbFaculty.Items.AddRange(faculties.ToArray());
            cmbMarkFaculty.Items.AddRange(faculties.ToArray());

            cmbDept.Items.AddRange(departments.ToArray());
            cmbMarkDept.Items.AddRange(departments.ToArray());

            cmbMarkCourse.Items.AddRange(courses.ToArray());

            if (cmbFaculty.Items.Count > 0) cmbFaculty.SelectedIndex = 0;
            if (cmbMarkFaculty.Items.Count > 0) cmbMarkFaculty.SelectedIndex = 0;
            if (cmbDept.Items.Count > 0) cmbDept.SelectedIndex = 0;
            if (cmbMarkDept.Items.Count > 0) cmbMarkDept.SelectedIndex = 0;
            if (cmbMarkCourse.Items.Count > 0) cmbMarkCourse.SelectedIndex = 0;
        }

        // =========================================
        // STUDENT REGISTER
        // =========================================
        private void btnsaveStudent_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtID.Text))
            {
                MessageBox.Show("Student ID is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtID.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(txtFirst.Text) || string.IsNullOrWhiteSpace(txtLast.Text))
            {
                MessageBox.Show("First Name and Last Name are required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        MERGE INTO Students AS target
                        USING (SELECT @StudID AS StudID) AS source
                        ON (target.StudID = source.StudID)
                        WHEN MATCHED THEN
                            UPDATE SET FirstName = @FirstName, LastName = @LastName, Year = @Year, Faculty = @Faculty, Department = @Department
                        WHEN NOT MATCHED THEN
                            INSERT (StudID, FirstName, LastName, Year, Faculty, Department)
                            VALUES (@StudID, @FirstName, @LastName, @Year, @Faculty, @Department);";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@StudID", txtID.Text.Trim());
                        cmd.Parameters.AddWithValue("@FirstName", txtFirst.Text.Trim());
                        cmd.Parameters.AddWithValue("@LastName", txtLast.Text.Trim());
                        cmd.Parameters.AddWithValue("@Year", cmbYear.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Faculty", cmbFaculty.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Department", cmbDept.SelectedItem?.ToString() ?? "");

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Student registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                txtID.Clear();
                txtFirst.Clear();
                txtLast.Clear();

                RefreshDropdownData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving student: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================
        // FACULTY REGISTRATION (COURSE CREATION)
        // =========================================
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTab2CourseID.Text))
            {
                MessageBox.Show("Course ID is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTab2CourseID.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(txtTab2CourseName.Text))
            {
                MessageBox.Show("Course Name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(txtTab2Faculty.Text) || string.IsNullOrWhiteSpace(txtTab2Dept.Text))
            {
                MessageBox.Show("Faculty and Department names are required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        MERGE INTO Courses AS target
                        USING (SELECT @CourseID AS CourseID) AS source
                        ON (target.CourseID = source.CourseID)
                        WHEN MATCHED THEN
                            UPDATE SET CourseName = @CourseName, Faculty = @Faculty, Department = @Department, Year = @Year
                        WHEN NOT MATCHED THEN
                            INSERT (CourseID, CourseName, Faculty, Department, Year)
                            VALUES (@CourseID, @CourseName, @Faculty, @Department, @Year);";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@CourseID", txtTab2CourseID.Text.Trim());
                        cmd.Parameters.AddWithValue("@CourseName", txtTab2CourseName.Text.Trim());
                        cmd.Parameters.AddWithValue("@Faculty", txtTab2Faculty.Text.Trim());
                        cmd.Parameters.AddWithValue("@Department", txtTab2Dept.Text.Trim());
                        cmd.Parameters.AddWithValue("@Year", cmbTab2Year.SelectedItem?.ToString() ?? "");

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Course/Faculty registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtTab2CourseID.Clear();
                txtTab2CourseName.Clear();
                txtTab2Faculty.Clear();
                txtTab2Dept.Clear();

                RefreshDropdownData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving course/faculty: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================
        // VIEW RESULT
        // =========================================
        private void btnViewResult_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtResultStudentID.Text))
            {
                MessageBox.Show("Please enter a Student ID.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtResultStudentID.Focus();
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        SELECT CourseName AS [Course Name], 
                               Assignments AS [Assignments], 
                               MidExam AS [Mid Exam], 
                               FinalExam AS [Final Exam], 
                               Total AS [Total Mark],
                               CASE 
                                   WHEN Total >= 90 THEN 'A+'
                                   WHEN Total >= 85 THEN 'A'
                                   WHEN Total >= 80 THEN 'A-'
                                   WHEN Total >= 75 THEN 'B+'
                                   WHEN Total >= 70 THEN 'B'
                                   WHEN Total >= 65 THEN 'B-'
                                   WHEN Total >= 60 THEN 'C+'
                                   WHEN Total >= 50 THEN 'C'
                                   ELSE 'F'
                               END AS [Letter Grade]
                        FROM Marks 
                        WHERE StudID = @StudentID 
                        ORDER BY CourseName";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@StudentID", txtResultStudentID.Text.Trim());

                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        
                        dgvResults.DataSource = dt;

                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show("No results found for student ID: " + txtResultStudentID.Text, "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading student results: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================
        // TEACHER REGISTRATION (MARKS CRUD)
        // =========================================
        private void LoadMarks(string filterQuery = "", List<SqlParameter> parameters = null)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "SELECT MarkID, StudID AS [Student ID], Year, Faculty, Department, CourseName AS [Course Name], Assignments, MidExam AS [Mid Exam], FinalExam AS [Final Exam], Total FROM Marks";
                    if (!string.IsNullOrEmpty(filterQuery))
                    {
                        query += " WHERE " + filterQuery;
                    }
                    query += " ORDER BY MarkID DESC";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters.ToArray());
                        }

                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dgvMarks.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading marks: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CalculateTotalMarks()
        {
            double assignment = 0;
            double mid = 0;
            double final = 0;

            double.TryParse(txtAssignment.Text, out assignment);
            double.TryParse(txtMid.Text, out mid);
            double.TryParse(txtFinal.Text, out final);

            txtTotal.Text = (assignment + mid + final).ToString("0.##");
        }

        private void ClearTeacherInputs()
        {
            txtMarkStudentID.Clear();
            txtAssignment.Clear();
            txtMid.Clear();
            txtFinal.Clear();
            txtTotal.Clear();
            txtSearch.Clear();
            txtSearchDep.Clear();
            
            if (cmbMarkYear.Items.Count > 0) cmbMarkYear.SelectedIndex = 0;
            if (cmbMarkFaculty.Items.Count > 0) cmbMarkFaculty.SelectedIndex = 0;
            if (cmbMarkDept.Items.Count > 0) cmbMarkDept.SelectedIndex = 0;
            if (cmbMarkCourse.Items.Count > 0) cmbMarkCourse.SelectedIndex = 0;
        }

        // Auto-fills faculty & department for convenience if student ID is registered
        private void txtMarkStudentID_TextChanged(object sender, EventArgs e)
        {
            string studentId = txtMarkStudentID.Text.Trim();
            if (studentId.Length >= 2)
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string query = "SELECT Year, Faculty, Department FROM Students WHERE StudID = @StudID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@StudID", studentId);
                            using (SqlDataReader dr = cmd.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    string year = dr["Year"].ToString();
                                    string faculty = dr["Faculty"].ToString();
                                    string dept = dr["Department"].ToString();

                                    if (cmbMarkYear.Items.Contains(year)) cmbMarkYear.SelectedItem = year;
                                    if (cmbMarkFaculty.Items.Contains(faculty)) cmbMarkFaculty.SelectedItem = faculty;
                                    if (cmbMarkDept.Items.Contains(dept)) cmbMarkDept.SelectedItem = dept;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Fail silently
                }
            }
        }

        // Add Button
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMarkStudentID.Text))
            {
                MessageBox.Show("Student ID is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMarkStudentID.Focus();
                return;
            }

            CalculateTotalMarks();

            double assignment = 0, mid = 0, final = 0, total = 0;
            double.TryParse(txtAssignment.Text, out assignment);
            double.TryParse(txtMid.Text, out mid);
            double.TryParse(txtFinal.Text, out final);
            double.TryParse(txtTotal.Text, out total);

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        INSERT INTO Marks (StudID, Year, Faculty, Department, CourseName, Assignments, MidExam, FinalExam, Total)
                        VALUES (@StudID, @Year, @Faculty, @Department, @CourseName, @Assignments, @MidExam, @FinalExam, @Total)";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@StudID", txtMarkStudentID.Text.Trim());
                        cmd.Parameters.AddWithValue("@Year", cmbMarkYear.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Faculty", cmbMarkFaculty.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Department", cmbMarkDept.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@CourseName", cmbMarkCourse.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Assignments", assignment);
                        cmd.Parameters.AddWithValue("@MidExam", mid);
                        cmd.Parameters.AddWithValue("@FinalExam", final);
                        cmd.Parameters.AddWithValue("@Total", total);

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Marks added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearTeacherInputs();
                LoadMarks();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding marks: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Update Button
        private void btnUpdateTeacher_Click(object sender, EventArgs e)
        {
            if (selectedMarkID == -1)
            {
                MessageBox.Show("Please select a mark record from the grid first to update.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CalculateTotalMarks();

            double assignment = 0, mid = 0, final = 0, total = 0;
            double.TryParse(txtAssignment.Text, out assignment);
            double.TryParse(txtMid.Text, out mid);
            double.TryParse(txtFinal.Text, out final);
            double.TryParse(txtTotal.Text, out total);

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        UPDATE Marks 
                        SET StudID = @StudID, Year = @Year, Faculty = @Faculty, Department = @Department, 
                            CourseName = @CourseName, Assignments = @Assignments, MidExam = @MidExam, 
                            FinalExam = @FinalExam, Total = @Total
                        WHERE MarkID = @MarkID";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@StudID", txtMarkStudentID.Text.Trim());
                        cmd.Parameters.AddWithValue("@Year", cmbMarkYear.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Faculty", cmbMarkFaculty.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Department", cmbMarkDept.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@CourseName", cmbMarkCourse.SelectedItem?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@Assignments", assignment);
                        cmd.Parameters.AddWithValue("@MidExam", mid);
                        cmd.Parameters.AddWithValue("@FinalExam", final);
                        cmd.Parameters.AddWithValue("@Total", total);
                        cmd.Parameters.AddWithValue("@MarkID", selectedMarkID);

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Marks updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearTeacherInputs();
                selectedMarkID = -1;
                LoadMarks();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating marks: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Delete Button
        private void btnDeleteTeacher_Click(object sender, EventArgs e)
        {
            if (selectedMarkID == -1)
            {
                MessageBox.Show("Please select a mark record from the grid first to delete.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show("Are you sure you want to delete the selected marks record?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes) return;

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = "DELETE FROM Marks WHERE MarkID = @MarkID";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@MarkID", selectedMarkID);
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Marks deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearTeacherInputs();
                selectedMarkID = -1;
                LoadMarks();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting marks: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Refresh Button
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            btnRefresh_Click_Logic();
        }

        private void btnRefresh_Click_Logic()
        {
            ClearTeacherInputs();
            selectedMarkID = -1;
            LoadMarks();
        }

        // Search Student ID Button
        private void btnSearchTeacher_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                MessageBox.Show("Please enter a Student ID to search.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSearch.Focus();
                return;
            }

            var parameters = new List<SqlParameter> {
                new SqlParameter("@StudID", "%" + txtSearch.Text.Trim() + "%")
            };
            LoadMarks("StudID LIKE @StudID", parameters);
        }

        // Search Department Button
        private void btnSearchbyDepartment_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSearchDep.Text))
            {
                MessageBox.Show("Please enter a Department name to search.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSearchDep.Focus();
                return;
            }

            var parameters = new List<SqlParameter> {
                new SqlParameter("@Dept", "%" + txtSearchDep.Text.Trim() + "%")
            };
            LoadMarks("Department LIKE @Dept", parameters);
        }

        private void dgvMarks_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            PopulateTeacherFieldsFromGrid();
        }

        private void PopulateTeacherFieldsFromGrid()
        {
            if (dgvMarks.CurrentRow != null)
            {
                try
                {
                    var row = dgvMarks.CurrentRow;
                    
                    // Verify if columns are loaded
                    if (row.Cells["MarkID"].Value == DBNull.Value || row.Cells["MarkID"].Value == null)
                        return;

                    selectedMarkID = Convert.ToInt32(row.Cells["MarkID"].Value);
                    
                    txtMarkStudentID.Text = row.Cells["Student ID"].Value.ToString();
                    
                    string year = row.Cells["Year"].Value.ToString();
                    if (cmbMarkYear.Items.Contains(year)) cmbMarkYear.SelectedItem = year;
                    
                    string faculty = row.Cells["Faculty"].Value.ToString();
                    if (cmbMarkFaculty.Items.Contains(faculty)) cmbMarkFaculty.SelectedItem = faculty;
                    
                    string department = row.Cells["Department"].Value.ToString();
                    if (cmbMarkDept.Items.Contains(department)) cmbMarkDept.SelectedItem = department;
                    
                    string courseName = row.Cells["Course Name"].Value.ToString();
                    if (cmbMarkCourse.Items.Contains(courseName)) cmbMarkCourse.SelectedItem = courseName;
                    
                    txtAssignment.Text = row.Cells["Assignments"].Value.ToString();
                    txtMid.Text = row.Cells["Mid Exam"].Value.ToString();
                    txtFinal.Text = row.Cells["Final Exam"].Value.ToString();
                    txtTotal.Text = row.Cells["Total"].Value.ToString();
                }
                catch
                {
                    // Fail-safe selection populate
                }
            }
        }

        // Empty event handler leftovers to maintain designer compatibility
        private void button1_Click_Legacy(object sender, EventArgs e) {}

        // =========================================
        // PREMIUM FLAT THEME UI STYLING
        // =========================================
        private void StyleControls()
        {
            // Main Form background - Obsidian Black
            this.BackColor = Color.FromArgb(11, 12, 16);
            this.ForeColor = Color.FromArgb(226, 232, 240);
            this.Font = new Font("Segoe UI", 10F, FontStyle.Regular);

            tabControl1.BackColor = Color.FromArgb(11, 12, 16);
            
            // TabPages background - Sleek Charcoal Card
            Color tabColor = Color.FromArgb(21, 23, 30);
            foreach (TabPage tp in tabControl1.TabPages)
            {
                tp.BackColor = tabColor;
                tp.ForeColor = Color.FromArgb(226, 232, 240);
                StyleControlsRecursive(tp.Controls);
            }

            StyleDataGridView(dgvMarks);
            StyleDataGridView(dgvResults);
        }

        private void StyleControlsRecursive(Control.ControlCollection controls)
        {
            foreach (Control ctrl in controls)
            {
                if (ctrl is Label)
                {
                    // Clean Ice-Blue Labels
                    ctrl.ForeColor = Color.FromArgb(197, 198, 199);
                    ctrl.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
                }
                else if (ctrl is TextBox || ctrl is ComboBox)
                {
                    // Dark Obsidian inputs with white text
                    ctrl.BackColor = Color.FromArgb(11, 12, 16);
                    ctrl.ForeColor = Color.FromArgb(243, 244, 246);
                    ctrl.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
                    
                    if (ctrl is TextBox txt)
                    {
                        txt.BorderStyle = BorderStyle.FixedSingle;
                    }
                    if (ctrl is ComboBox cmb)
                    {
                        cmb.FlatStyle = FlatStyle.Flat;
                    }
                }
                else if (ctrl is Button btn)
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderSize = 0;
                    btn.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                    btn.ForeColor = Color.White;
                    btn.Height = 32;

                    string btnText = btn.Text.ToLower();
                    // Premium accent colors
                    if (btnText.Contains("save") || btnText.Contains("add") || btnText.Contains("update") || btnText.Contains("viewresult"))
                    {
                        btn.BackColor = Color.FromArgb(99, 102, 241); // Royal Indigo
                    }
                    else if (btnText.Contains("delete") || btnText.Contains("refresh"))
                    {
                        btn.BackColor = Color.FromArgb(244, 63, 94); // Neon Coral/Pink
                    }
                    else if (btnText.Contains("search"))
                    {
                        btn.BackColor = Color.FromArgb(6, 182, 212); // Cyber Cyan/Turquoise
                    }
                    else
                    {
                        btn.BackColor = Color.FromArgb(71, 85, 105); // Cool Gray
                    }
                }

                if (ctrl.HasChildren)
                {
                    StyleControlsRecursive(ctrl.Controls);
                }
            }
        }

        private void StyleDataGridView(DataGridView dgv)
        {
            if (dgv == null) return;

            dgv.BackgroundColor = Color.FromArgb(21, 23, 30);
            dgv.ForeColor = Color.FromArgb(226, 232, 240);
            dgv.GridColor = Color.FromArgb(45, 55, 72);
            dgv.BorderStyle = BorderStyle.None;
            dgv.RowHeadersVisible = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.AllowUserToAddRows = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(11, 12, 16);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(6, 182, 212); // Cyan headers
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.ColumnHeadersHeight = 35;

            dgv.DefaultCellStyle.BackColor = Color.FromArgb(28, 30, 38);
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(226, 232, 240);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(99, 102, 241); // Indigo selection
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);

            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(35, 38, 48);
        }
    }
}