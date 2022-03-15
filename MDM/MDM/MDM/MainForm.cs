using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Deployment.Application;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.ComponentModel;
using System.Threading;
using System.IO.Ports;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace MDM
{
    public partial class MainForm : Form
    {
        //Variables for drag panel
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        //Form size and location
        public int WW, WH;
        //For SQL
        private string connectionString;
        private SqlDataAdapter dataAdapter = new SqlDataAdapter();
        private SqlDataAdapter InfoAdapter = new SqlDataAdapter();
        private SqlDataAdapter ApprAdapter = new SqlDataAdapter();
        private SqlDataAdapter DimAdapter = new SqlDataAdapter();
        private SqlDataAdapter RevAdapter = new SqlDataAdapter();
        private SqlDataAdapter UserAdapter = new SqlDataAdapter();
        private SqlDataAdapter ToolAdapter = new SqlDataAdapter();
        private SqlDataAdapter ResultAdapter = new SqlDataAdapter();
        public int CurrentPartID;
        public static SqlConnection cnn;
        //Other variables
        private bool changeUnsaved;
        private int UserLevel;
        private string CurrentTool = "";
        private DateTime StartMeasureTime;
        private int StartMeasureRowIndex = 0;
        private int PreviousRow = 0, PreviousCol = 0;
        //Excel communication variables
        DataExporter Exporter = new DataExporter();

        public MainForm()
        {
            InitializeComponent();
            UsernameStripLbl.Text = "Not connected.";
            WW = Properties.Settings.Default.ScreenResW;
            WH = Properties.Settings.Default.ScreenResH;
            //Prepare resouce files.
        }
        //Basic form elements--------------------------------------------------------------------
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g;
            g = e.Graphics;
            Pen myPen = new Pen(Color.DarkBlue)
            {
                Width = 2
            };
            g.DrawRectangle(myPen, 0, 1, WW, WH - 1);
        }
        private void DragPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void ExitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ResetAllPanel()
        {
            //Hide all panel
            loginTablePnl.Visible = false;
            LogPanel.Visible = false;
            PartSelectPanel.Visible = false;
            PartInfoEdit.Visible = false;
            ManagePnl.Visible = false;
            ChkSheetPanel.Visible = false;
            MeasurePanel.Visible = false;
            //Reset menu color to gray
            PartDataMenu.BackColor = Color.LightGray;
            ChecksheetMenu.BackColor = Color.LightGray;
            MeasureMenu.BackColor = Color.LightGray;
            ManageMenu.BackColor = Color.LightGray;
            //Enable all menu items
            PartDataMenu.Enabled = true;
            ChecksheetMenu.Enabled = true;
            MeasureMenu.Enabled = true;
            //Hide buttons and checkboxes in part select panel
            NewPartBt.Visible = false;
            DeletePartBt.Visible = false;
            PartSelectOpenBt.Visible = false;
            AllDimChkBt.Visible = false;
            QADimChkBt.Visible = false;
            IPQCDimChkBt.Visible = false;
            //Disable items in measure interface
            FilterGrb.Enabled = false;
            QuickJudgeGrb.Enabled = false;
            ToolGrb.Enabled = false;
        }
        //Form load, close events-----------------------------------------------------------------
        private void MainForm_Load(object sender, EventArgs e)
        {
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F); // for design in 96 DPI
            //Do some sizing and positioning
            Size = new Size(WW, WH);
            CenterToScreen();
            MainMenuStrip.Location = new Point(120, 2);
            signOutToolStripMenuItem.Enabled = false;
            statusStrip.Location = new Point(2, WH - 23);
            statusStrip.Width = WW - 4;
            ExitButton.Location = new Point(WW - 54, 1);
            TitleBarPnl.Width = WW - 2;
            ResizeRedraw = true;
            loginTablePnl.Location = new Point(WW / 2 - loginTablePnl.Width / 2, WH / 2 - 40);
            PartSelectPanel.Size = new Size(WW / 2, WH - 100);
            PartSelectPanel.Location = new Point(15, 50);
            PartSelectGrV.Size = new Size(650, PartSelectPanel.Height - 50);
            //Screen Resolution checkboxes
            if (WW == 800)
            {
                x900ToolStripMenuItem.Checked = false;
                x1080ToolStripMenuItem.Checked = false;
                maximizeToolStripMenuItem.Checked = false;
            }
            else if (WW == 1600)
            {
                x900ToolStripMenuItem.Checked = true;
                x1080ToolStripMenuItem.Checked = false;
                maximizeToolStripMenuItem.Checked = false;
            }
            else if (WW == 1920)
            {
                x900ToolStripMenuItem.Checked = false;
                x1080ToolStripMenuItem.Checked = true;
                maximizeToolStripMenuItem.Checked = false;
            }
            else if (WW == Screen.PrimaryScreen.WorkingArea.Width)
            {
                x900ToolStripMenuItem.Checked = false;
                x1080ToolStripMenuItem.Checked = false;
                maximizeToolStripMenuItem.Checked = true;
            }
            //Factory checkboxes
            if(Properties.Settings.Default.Factory == "SVN2")
            {
                sVN2ToolStripMenuItem.Checked = true;
                SVN2ChkBx.Checked = true;
                SVN3ChkBx.Checked = false;
            }
            else if(Properties.Settings.Default.Factory == "SVN3")
            {
                sVN3ToolStripMenuItem.Checked = true;
                SVN2ChkBx.Checked = false;
                SVN3ChkBx.Checked = true;
            }
            else
            {
                allToolStripMenuItem.Checked = true;
                SVN2ChkBx.Checked = true;
                SVN3ChkBx.Checked = true;
            }
            //Initial values
            ExportPathTxb.Text = Properties.Settings.Default.DefaultExportPath;
            UsernameTxb.Text = Properties.Settings.Default.UserName;
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cnn != null)
            {
                SignOutToolStripMenuItem_Click(null, EventArgs.Empty);
                cnn.Close();
                cnn.Dispose();
            }
            Properties.Settings.Default.Save();
        }

        //For settings menu-----------------------------------------------------------------------
        private void X1080ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ScreenResW = 1920;
            Properties.Settings.Default.ScreenResH = 1080;
            MessageBox.Show("Change take effect after application restart.");
        }
        private void X900ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ScreenResW = 1600;
            Properties.Settings.Default.ScreenResH = 800;
            MessageBox.Show("Change take effect after application restart.");
        }
        private void MaximizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ScreenResW = Screen.PrimaryScreen.WorkingArea.Width;
            Properties.Settings.Default.ScreenResH = Screen.PrimaryScreen.WorkingArea.Height;
            MessageBox.Show("Change take effect after application restart.");
        }
        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                MessageBox.Show("Measurement Data Management Software" + Environment.NewLine
                + "Version: " + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4) + Environment.NewLine
                + "2020 Santomas VietNam Jsc." + Environment.NewLine
                + "All rights reserved." + Environment.NewLine, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void allToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Factory = "All";
            allToolStripMenuItem.Checked = true;
            sVN2ToolStripMenuItem.Checked = false;
            sVN3ToolStripMenuItem.Checked = false;
        }
        private void sVN2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Factory = "SVN2";
            allToolStripMenuItem.Checked = false;
            sVN2ToolStripMenuItem.Checked = true;
            sVN3ToolStripMenuItem.Checked = false;
        }
        private void sVN3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Factory = "SVN3";
            allToolStripMenuItem.Checked = false;
            sVN2ToolStripMenuItem.Checked = false;
            sVN3ToolStripMenuItem.Checked = true;
        }


        //Sign in, sign out code-------------------------------------------------------------------
        private void ConnectBt_Click(object sender, EventArgs e)
        {
            if (UsernameTxb.Text == "" || PasswordTxb.Text == "")
            {
                MessageBox.Show("Please input username and password!");
                return;
            }
            //connectionString = "Data Source=192.168.0.243;Database=AMS;User ID=auto; Password=12qw!@QW;";
            connectionString = Properties.Settings.Default.SqlConnectionString;
            cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
            }
            catch (SqlException r)
            {
                MessageBox.Show("Login failed." + Environment.NewLine + "Kết nối server thất bại." + Environment.NewLine + r.ToString());
                return;
            }
            //Check username, password and then get user level.
            using (SqlDataAdapter UserAdapter = new SqlDataAdapter("Select * from UserAccount where (Username = '" + UsernameTxb.Text + "')", cnn))
            {
                _ = new SqlCommandBuilder(UserAdapter);
                DataTable table = new DataTable { };
                UserAdapter.Fill(table);
                if (table.Rows.Count < 1)
                {
                    MessageBox.Show("Wrong username.");
                    cnn.Close();
                    cnn.Dispose();
                    return;
                }
                else if (table.Rows.Count >= 1)
                {
                    if (table.Rows[0]["Password"].ToString().Trim() != PasswordTxb.Text)
                    {
                        MessageBox.Show("Wrong password.");
                        cnn.Close();
                        cnn.Dispose();
                        return;
                    }
                    //else if (table.Rows[0]["Active"].ToString() == "True")
                    //{
                    //    MessageBox.Show("User already signed in on another device");
                    //    cnn.Close();
                    //    cnn.Dispose();
                    //    return;
                    //}
                    else
                    {
                        UserLevel = int.Parse(table.Rows[0]["UserLevel"].ToString());
                        table.Rows[0]["Active"] = "True";
                        if (table.Rows[0]["Role"].ToString().Trim() == "QA")
                        {
                            QADimChkBt.Checked = true;
                        }
                        else if (table.Rows[0]["Role"].ToString().Trim() == "IPQC")
                        {
                            IPQCDimChkBt.Checked = true;
                        }
                        else
                        {
                            AllDimChkBt.Checked = true;
                        }
                        UserAdapter.Update(table);
                    }
                }
            }
            UsernameStripLbl.Text = UsernameTxb.Text + " logged in sucessfully.";
            UserLvlLbl.Text += UserLevel;
            ResetAllPanel();
            signOutToolStripMenuItem.Enabled = true;
            statisticToolStripMenuItem.Enabled = true;
            logToolStripMenuItem.Enabled = true;
            uWaveCommunicationToolStripMenuItem.Enabled = true;
            //reset to level 1
            QuickJudgeGrb.Enabled = false;
            SavingGrb.Enabled = false;
            DimCheckGrV.ReadOnly = true;
            ApprCheckGrV.ReadOnly = true;
            DimAddBt.Enabled = false;
            DimRemoveBt.Enabled = false;
            RevisionGrV.ReadOnly = true;
            RevAddBt.Enabled = false;
            RevRemoveBt.Enabled = false;
            SaveInfoBt.Enabled = false;
            ManageMenu.Enabled = false;
            MeasureMenu.Enabled = false;
            DeleteBatchBt.Enabled = false;
            ApprAddBt.Enabled = false;
            ApprRemoveBt.Enabled = false;
            ToolGrb.Enabled = false;
            MeasureResultGrV.Columns["Value"].ReadOnly = true;
            MeasureResultGrV.Columns["Judge"].ReadOnly = true;
            ClearStripMenu.Enabled = false;
            //LVL? Enable
            if (UserLevel > 1)
            {
                //LV2: only measure menu enable
                QuickJudgeGrb.Enabled = true;
                SavingGrb.Enabled = true;
                ToolGrb.Enabled = true;
                MeasureResultGrV.Columns["Value"].ReadOnly = false;
                MeasureResultGrV.Columns["Judge"].ReadOnly = false;
                MeasureMenu.Enabled = true;
            }
            if (UserLevel > 2)
            {
                //LV3: LV2 + Edit part data
                DimCheckGrV.ReadOnly = false;
                ApprCheckGrV.ReadOnly = false;
                PartInfoGrV.ReadOnly = false;
                DimAddBt.Enabled = true;
                DimRemoveBt.Enabled = true;
                RevisionGrV.ReadOnly = false;
                RevAddBt.Enabled = true;
                RevRemoveBt.Enabled = true;
                SaveInfoBt.Enabled = true;
                DeleteBatchBt.Enabled = true;
                ApprAddBt.Enabled = true;
                ApprRemoveBt.Enabled = true;
            }
            if (UserLevel > 3)
            {
                //LV4: LV3 + Edit User, Tool
                ManageMenu.Enabled = true;
                ClearStripMenu.Enabled = true;
            }
            //Store current username
            Properties.Settings.Default.UserName = UsernameTxb.Text;
        }
        private void SignOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Visibility and colors
            ResetAllPanel();
            signOutToolStripMenuItem.Enabled = false;
            ChecksheetMenu.Enabled = false;
            ManageMenu.Enabled = false;
            MeasureMenu.Enabled = false;
            PartDataMenu.Enabled = false;
            statisticToolStripMenuItem.Enabled = false;
            logToolStripMenuItem.Enabled = false;
            loginTablePnl.Visible = true;
            uWaveCommunicationToolStripMenuItem.Enabled = false;
            //Settings and others
            UsernameStripLbl.Text = "Disconnected.";
            UserLvlLbl.Text = "User Level: ";
            PasswordTxb.Clear();
            //reset user role, release tool, reset active status
            UserLevel = 1;
            ReleaseAllTool();
            using (SqlCommand ResetActive = new SqlCommand("Update UserAccount set Active = 'False' where UserName = '" + UsernameTxb.Text + "'", cnn))
            {
                if (cnn.State == ConnectionState.Open)
                {
                    ResetActive.ExecuteNonQuery();
                }
            }
        }
        private void PasswordTxb_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ConnectBt_Click(null, EventArgs.Empty);
            }
        }
        //Statistic function-----------------------------------------------------------------------
        private void StatisticToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Statistic = "Statistic:" + Environment.NewLine;
            using (SqlCommand GetMaxPartID = new SqlCommand("Select count(ID) from PartInfo", cnn))
            {
                Statistic += "Part count:  " + GetMaxPartID.ExecuteScalar().ToString() + Environment.NewLine;
            }
            using (SqlCommand GetMaxApprID = new SqlCommand("Select count(ApprID) from AppearanceCheckPoint", cnn))
            {
                Statistic += "Appearance checkpoint count: " + GetMaxApprID.ExecuteScalar().ToString() + Environment.NewLine;
            }
            using (SqlCommand GetMaxDimensionID = new SqlCommand("Select count(DimID) from DimensionCheckPoint", cnn))
            {
                Statistic += "Dimension checkpoint count: " + GetMaxDimensionID.ExecuteScalar().ToString() + Environment.NewLine;
            }
            using (SqlCommand GetMaxMeasurementResult = new SqlCommand("Select count(MeasID) from MeasurementResult", cnn))
            {
                Statistic += "Measurement results count: " + GetMaxMeasurementResult.ExecuteScalar().ToString() + Environment.NewLine;
            }
            using (SqlCommand GetMaxDataID = new SqlCommand("Select max(DataID) from UWaveData", cnn))
            {
                Statistic += "Received data from U-Wave Tools count: " + GetMaxDataID.ExecuteScalar().ToString() + Environment.NewLine;
            }
            MessageBox.Show(Statistic, "Statistic", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //Record to Activity log function-----------------------------------------------------------
        private void RecordLog(string content, string ID)
        {
            SqlCommand Record = new SqlCommand("Insert into ActivityLog (ID, Time, LogContent) values (@ID, @DateTime ,@Content)", cnn);
            Record.Parameters.AddWithValue("DateTime", DateTime.Now);
            Record.Parameters.AddWithValue("ID", ID);
            Record.Parameters.AddWithValue("Content", content);
            Record.ExecuteNonQuery();
        }
        private void LogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogPanel.Size = new Size(WW - 25, WH - 60);
            LogPanel.Location = new Point(15, 45);
            LogGrV.Size = new Size(WW - 100, WH - 100);
            ResetAllPanel();
            LogPanel.Visible = true;
            DataTable table = new DataTable { };
            SqlDataAdapter LogAdapter = new SqlDataAdapter("Select * from ActivityLog order by LogID DESC", connectionString);
            _ = new SqlCommandBuilder(LogAdapter);
            LogAdapter.Fill(table);
            LogGrV.DataSource = table;
        }

        //Part data interface functions------------------------------------------------------------
        private void PartDataMenu_Click(object sender, EventArgs e)
        {
            //Allow editting depend on user
            if (UserLevel == 1)
            {
                PartInfoGrV.ReadOnly = true;
                ApprCheckGrV.ReadOnly = true;
                DimCheckGrV.ReadOnly = true;
                SaveInfoBt.Enabled = false;
            }
            else if (UserLevel == 3)
            {
                PartInfoGrV.ReadOnly = false;
                ApprCheckGrV.ReadOnly = false;
                DimCheckGrV.ReadOnly = false;
                SaveInfoBt.Enabled = true;
            }
            //Sizing and positioning
            PartInfoEdit.Size = new Size(WW - 25, WH - 70);
            PartInfoEdit.Location = new Point(15, 45);
            PartInfoTabs.Size = new Size(PartInfoEdit.Width, PartInfoEdit.Height - 100);
            PartInfoTabs.Location = new Point(0, 0);
            PartInfoGrV.Size = new Size(PartInfoTabs.Width - 15, PartInfoTabs.Height - 70);
            DimCheckGrV.Size = new Size(PartInfoTabs.Width - 15, PartInfoTabs.Height - 100);
            ApprCheckGrV.Size = new Size(PartInfoTabs.Width - 15, PartInfoTabs.Height - 100);
            RevisionGrV.Size = new Size(PartInfoTabs.Width - 15, PartInfoTabs.Height - 100);
            //Visibility and coloring
            ResetAllPanel();
            SVN2ChkBx_CheckedChanged(null, EventArgs.Empty);
            PartSelectPanel.Visible = true;
            NewPartBt.Visible = true;
            DeletePartBt.Visible = true;
            PartDataMenu.BackColor = Color.LightBlue;
            PartDataMenu.Enabled = false;
            PartSelectOpenBt.Visible = true;
        }
        private void NewPartBt_Click(object sender, EventArgs e)
        {
            if (MoldNoTxb.Text == "")
            {
                MessageBox.Show("Please input Mold no." + Environment.NewLine + "Nhập Mold No để tạo part mới.");
                return;
            }
            using (SqlCommand CheckExisting = new SqlCommand("Select top 1 1 from PartInfo where (MoldNo = '" + MoldNoTxb.Text + "')", cnn))
            {
                if (CheckExisting.ExecuteScalar() != null)
                {
                    var confirmResult = MessageBox.Show("Mold no already exist, do you still want to add new part data?" + Environment.NewLine +
                        "Mold No này đã có trong database. Bạn vẫn muốn tạo part data mới?", "Confirm", MessageBoxButtons.YesNo);
                    if (confirmResult == DialogResult.No)
                    {
                        return;
                    }
                }
            }

            using (SqlCommand AddNewPart = new SqlCommand("Insert into PartInfo (MoldNo, PartNo, NoOfCav, DateAdded) values ('" + MoldNoTxb.Text + "','RCXXXXX','4',@DateTimeNow)", cnn))
            {
                AddNewPart.Parameters.AddWithValue("DateTimeNow", DateTime.Now);
                AddNewPart.ExecuteNonQuery();
            }
            RecordLog(UsernameTxb.Text + ": New part added" + MoldNoTxb.Text, "NA");
            GetInfo(MoldNoTxb.Text, 1);
            dataAdapter.Update((DataTable)InfoBindingSource.DataSource);
        }
        private void DeletePartBt_Click(object sender, EventArgs e)
        {
            string PartID = PartSelectGrV.CurrentRow.Cells["InfoID"].Value.ToString();
            string MoldToDelete = PartSelectGrV.CurrentRow.Cells["InfoMoldNo"].Value.ToString();
            var confirm = MessageBox.Show("Are you sure want to delete " + MoldToDelete + "?" + Environment.NewLine +
                "Bạn có chắc chắn muốn xóa " + MoldToDelete + " ? "
                , "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm == DialogResult.No)
            {
                return;
            }
            else if (confirm == DialogResult.Yes)
            {
                using (SqlCommand DeleteInfo = new SqlCommand("Delete from PartInfo where ID ='" + PartID + "'", cnn))
                {
                    DeleteInfo.ExecuteNonQuery();
                }
                using (SqlCommand DeleteAppr = new SqlCommand("Delete from AppearanceCheckPoint where ID ='" + PartID + "'", cnn))
                {
                    DeleteAppr.ExecuteNonQuery();
                }
                using (SqlCommand DeleteDim = new SqlCommand("Delete from DimensionCheckPoint where ID ='" + PartID + "'", cnn))
                {
                    DeleteDim.ExecuteNonQuery();
                }
                GetInfo("", 1);
                RecordLog(UsernameTxb.Text + ": Part deleted.", CurrentPartID.ToString());
            }
        }
        private void ApprAdd_Click(object sender, EventArgs e)
        {
            DataTable ApprTable = (DataTable)ApprBindingSource.DataSource;
            DataRow newAppr = ApprTable.NewRow();
            newAppr["ID"] = CurrentPartID;
            ApprTable.Rows.Add(newAppr);
            ApprCheckGrV.Rows[ApprCheckGrV.RowCount - 1].Cells["ApprItemNo"].Style.BackColor = Color.LightSalmon;
            changeUnsaved = true;
            RecordLog(UsernameTxb.Text + ": Appearance checkpoint added.", CurrentPartID.ToString());
        }
        private void ApprRemove_Click(object sender, EventArgs e)
        {
            var Confirm = MessageBox.Show("Are you sure want to delete selected checkpoint?" + Environment.NewLine +
                "Bạn có chắc chắn muốn xóa checkpoint đã chọn?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (Confirm == DialogResult.No)
            {
                return;
            }
            DataTable ApprTable = (DataTable)ApprBindingSource.DataSource;
            ApprTable.Rows[ApprCheckGrV.CurrentRow.Index].Delete();
            RecordLog(UsernameTxb.Text + ": Appearance checkpoint removed.", CurrentPartID.ToString());
        }
        private void DimAdd_Click(object sender, EventArgs e)
        {
            DataTable DimTable = (DataTable)DimBindingSource.DataSource;
            DataRow newDim = DimTable.NewRow();
            DimTable.Rows.Add(newDim);
            newDim["ID"] = CurrentPartID;
            newDim["DimTool"] = "NA";
            DimCheckGrV.Rows[DimCheckGrV.RowCount - 1].Cells["DimNo"].Style.BackColor = Color.LightSalmon;
            changeUnsaved = true;
            RecordLog(UsernameTxb.Text + ": Dimension checkpoint added.", CurrentPartID.ToString());
        }
        private void DimRemove_Click(object sender, EventArgs e)
        {
            var Confirm = MessageBox.Show("Are you sure want to delete selected checkpoint?" + Environment.NewLine +
                "Bạn có chắc chắn muốn xóa checkpoint đã chọn?", "Confirm",
                 MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (Confirm == DialogResult.No)
            {
                return;
            }
            DataTable DimTable = (DataTable)DimBindingSource.DataSource;
            DimTable.Rows[DimCheckGrV.CurrentRow.Index].Delete();
            RecordLog(UsernameTxb.Text + ": Dimension checkpoint removed.", CurrentPartID.ToString());
        }
        private void RevAddBt_Click(object sender, EventArgs e)
        {
            DataTable RevTable = (DataTable)RevBindingSource.DataSource;
            DataRow newRev = RevTable.NewRow();
            newRev["ID"] = CurrentPartID;
            RevTable.Rows.Add(newRev);
            changeUnsaved = true;
        }
        private void RevRemoveBt_Click(object sender, EventArgs e)
        {
            DataTable RevTable = (DataTable)RevBindingSource.DataSource;
            RevTable.Rows[RevisionGrV.CurrentRow.Index].Delete();
        }
        private void SaveBt_Click(object sender, EventArgs e)
        {
            //Check for empty cells or invalid data on required columns
            if (Convert.ToInt32(PartInfoGrV.Rows[0].Cells["Cav"].Value) < 1)
            {
                MessageBox.Show("Cavity number can not be empty or zero." + Environment.NewLine +
                    "Số Cavity không thể để trống hoặc bằng không.");
                return;
            }
            dataAdapter.Update((DataTable)InfoBindingSource.DataSource);
            ApprAdapter.Update((DataTable)ApprBindingSource.DataSource);
            DimAdapter.Update((DataTable)DimBindingSource.DataSource);
            RevAdapter.Update((DataTable)RevBindingSource.DataSource);
            changeUnsaved = false;
            RecordLog(UsernameTxb.Text + ": Part data editted and saved.", CurrentPartID.ToString());
        }
        private void CloseInfoBt_Click(object sender, EventArgs e)
        {
            if (changeUnsaved)
            {
                var confirm = MessageBox.Show("Do you want to save the changes you made?" + Environment.NewLine +
                    "Bạn có muốn lưu dữ liệu", "Unsaved change", MessageBoxButtons.YesNoCancel);
                if (confirm == DialogResult.Yes)
                {
                    SaveBt_Click(null, EventArgs.Empty);
                }
                else if (confirm == DialogResult.Cancel)
                {
                    return;
                }
                else if (confirm == DialogResult.No)
                {
                    RecordLog(UsernameTxb.Text + ": Part data editted but not saved.", CurrentPartID.ToString());
                    changeUnsaved = false;
                }
            }

            PartDataMenu_Click(null, EventArgs.Empty);
        }
        private void PartInfoGrV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(UserLevel < 3)
            {
                return;
            }
            if (e.ColumnIndex == PartInfoGrV.Columns["Illustration"].Index)
            {
                //Image file for checksheets
                using (OpenFileDialog GetImage = new OpenFileDialog())
                {
                    GetImage.Filter = "Image Files(*.BMP;*.JPG;*.JPEG;*.GIF;*.PNG) | *.BMP; *.JPG; *.JPEG; *.GIF; *.PNG";
                    GetImage.Title = "Select image to include in QA and IPQC inspection sheet";
                    var result = GetImage.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        string SourceImage = GetImage.FileName.ToString();
                        string StoreImage = Properties.Settings.Default.DefaultIllustrationPath + GetImage.FileName.ToString() + CurrentPartID.ToString() + Path.GetExtension(SourceImage);
                        File.Copy(SourceImage, StoreImage, true);
                        PartInfoGrV.CurrentRow.Cells["Illustration"].Value = StoreImage;
                    }
                }
            }
            if (e.ColumnIndex == PartInfoGrV.Columns["WI1"].Index)
            {
                //Work instruction file
                using (OpenFileDialog GetWIFile = new OpenFileDialog())
                {
                    GetWIFile.Title = "Select work instruction file";
                    var result = GetWIFile.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        string SourceFile = GetWIFile.FileName.ToString();
                        string StoreFile = Properties.Settings.Default.DefaultIllustrationPath + GetWIFile.FileName.ToString() + CurrentPartID.ToString() + Path.GetExtension(SourceFile);
                        File.Copy(SourceFile, StoreFile, true);
                        PartInfoGrV.CurrentRow.Cells["WI1"].Value = StoreFile;
                    }
                }
            }
            if (e.ColumnIndex == PartInfoGrV.Columns["WI2"].Index)
            {
                //Work instruction file
                using (OpenFileDialog GetWIFile = new OpenFileDialog())
                {
                    GetWIFile.Title = "Select work instruction file";
                    var result = GetWIFile.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        string SourceFile = GetWIFile.FileName.ToString();
                        string StoreFile = Properties.Settings.Default.DefaultIllustrationPath + GetWIFile.FileName.ToString() + CurrentPartID.ToString() + Path.GetExtension(SourceFile);
                        File.Copy(SourceFile, StoreFile, true);
                        PartInfoGrV.CurrentRow.Cells["WI2"].Value = StoreFile;
                    }
                }
            }
            if (e.ColumnIndex == PartInfoGrV.Columns["DrawingFile"].Index)
            {
                //Drawing file
                using (OpenFileDialog GetFile = new OpenFileDialog())
                {
                    GetFile.Title = "Select drawing file";
                    var result = GetFile.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        string SourceFile = GetFile.FileName.ToString();
                        string StoreFile = Properties.Settings.Default.DefaultIllustrationPath + GetFile.FileName.ToString() + CurrentPartID.ToString() + Path.GetExtension(SourceFile);
                        File.Copy(SourceFile, StoreFile, true);
                        PartInfoGrV.CurrentRow.Cells["DrawingFile"].Value = StoreFile;
                    }
                }
            }
            if (e.ColumnIndex == PartInfoGrV.Columns["HistoryFile"].Index)
            {
                //History file
                using (OpenFileDialog GetFile = new OpenFileDialog())
                {
                    GetFile.Title = "Select history file";
                    var result = GetFile.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        string SourceFile = GetFile.FileName.ToString();
                        string StoreFile = Properties.Settings.Default.DefaultIllustrationPath + GetFile.FileName.ToString() + CurrentPartID.ToString() + Path.GetExtension(SourceFile);
                        File.Copy(SourceFile, StoreFile, true);
                        PartInfoGrV.CurrentRow.Cells["HistoryFile"].Value = StoreFile;
                    }
                }
            }
        }
        private void GrV_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (PartInfoTabs.SelectedTab == InfoTab)
            {
                PartInfoGrV.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightSalmon;
            }
            else if (PartInfoTabs.SelectedTab == ApprTab)
            {
                ApprCheckGrV.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightSalmon;
            }
            else if (PartInfoTabs.SelectedTab == DimTab)
            {
                DimCheckGrV.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightSalmon;
            }
            else if (PartInfoTabs.SelectedTab == RevTab)
            {
                RevisionGrV.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightSalmon;
            }
            changeUnsaved = true;
        }
        private void GrV_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Wrong data type." + Environment.NewLine +
                "Sai kiểu dữ liệu.", PartInfoGrV.Columns[e.ColumnIndex].HeaderText, MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.Cancel = true;
        }
        private void GrV_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
            {
                if (PartInfoTabs.SelectedTab == InfoTab)
                {
                    PartInfoGrV.CurrentCell = PartInfoGrV.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    DiameterStripMenu.Visible = false;
                    sameAsAboveToolStripMenuItem.Visible = false;
                    if (e.ColumnIndex > 16)
                    {
                        openFileToolStripMenuItem.Visible = true;
                    }
                    else
                    {
                        openFileToolStripMenuItem.Visible = false;
                    }

                }
                else if (PartInfoTabs.SelectedTab == ApprTab)
                {
                    ApprCheckGrV.CurrentCell = ApprCheckGrV.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    DiameterStripMenu.Visible = false;
                    sameAsAboveToolStripMenuItem.Visible = true;
                    openFileToolStripMenuItem.Visible = false;
                }
                else if (PartInfoTabs.SelectedTab == DimTab)
                {
                    DimCheckGrV.CurrentCell = DimCheckGrV.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    DiameterStripMenu.Visible = true;
                    sameAsAboveToolStripMenuItem.Visible = true;
                    openFileToolStripMenuItem.Visible = false;
                }
                else if (PartInfoTabs.SelectedTab == RevTab)
                {
                    return;
                }
                this.ContextMenuStrip = UltilityMenuStrip;
                UltilityMenuStrip.Show(this, MousePosition.X - this.Location.X, MousePosition.Y - this.Location.Y);
            }

        }

        //Gridview ultility right-click menu event----------------------------------------------------
        private void ClearStripMenu_Click(object sender, EventArgs e)
        {
            if (PartInfoTabs.SelectedTab == InfoTab)
            {
                PartInfoGrV.CurrentCell.Value = DBNull.Value;
                PartInfoGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                changeUnsaved = true;
            }
            else if (PartInfoTabs.SelectedTab == ApprTab)
            {
                if (ApprCheckGrV.CurrentCell.ColumnIndex > 1)
                {
                    ApprCheckGrV.CurrentCell.Value = DBNull.Value;
                    ApprCheckGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                    changeUnsaved = true;
                }
            }
            else if (PartInfoTabs.SelectedTab == DimTab)
            {
                if (DimCheckGrV.CurrentCell.ColumnIndex > 1)
                {
                    DimCheckGrV.CurrentCell.Value = DBNull.Value;
                    DimCheckGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                    changeUnsaved = true;
                }
            }
            else if (PartInfoTabs.SelectedTab == RevTab)
            {
                RevisionGrV.CurrentCell.Value = DBNull.Value;
                RevisionGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                changeUnsaved = true;
            }
        }
        private void DiameterStripMenu_Click(object sender, EventArgs e)
        {
            if (PartInfoTabs.SelectedTab == DimTab)
            {
                if (DimCheckGrV.CurrentCell.ColumnIndex > 1)
                {
                    DimCheckGrV.CurrentCell.Value += "Ø";
                    DimCheckGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                    changeUnsaved = true;
                }
            }
        }
        private void SameAsAboveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (PartInfoTabs.SelectedTab == ApprTab)
            {
                if (ApprCheckGrV.CurrentCell.ColumnIndex > 1 && ApprCheckGrV.CurrentCell.RowIndex > 0)
                {
                    ApprCheckGrV.CurrentCell.Value = ApprCheckGrV.Rows[ApprCheckGrV.CurrentCell.RowIndex - 1].Cells[ApprCheckGrV.CurrentCell.ColumnIndex].Value;
                    ApprCheckGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                    changeUnsaved = true;
                }
            }
            else if (PartInfoTabs.SelectedTab == DimTab)
            {
                if (DimCheckGrV.CurrentCell.ColumnIndex > 1 && DimCheckGrV.CurrentCell.RowIndex > 0)
                {
                    DimCheckGrV.CurrentCell.Value = DimCheckGrV.Rows[DimCheckGrV.CurrentCell.RowIndex - 1].Cells[DimCheckGrV.CurrentCell.ColumnIndex].Value;
                    DimCheckGrV.CurrentCell.Style.BackColor = Color.LightSalmon;
                    changeUnsaved = true;
                }
            }
        }
        private void OpenFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Path = PartInfoGrV.CurrentCell.Value.ToString();
            FileInfo FileToOpen = new FileInfo(Path);
            try
            {
                Process.Start(Path);
            }
            catch (Exception)
            {
                MessageBox.Show("File not found or not aceesible! Please add new file." + Environment.NewLine +
                    "Không tìm thấy file. Cần thêm file mới.");
            }
        }

        //Gridview for selecting part-----------------------------------------------------------------
        private void PartSelectGrV_CellMouseDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!PartDataMenu.Enabled)
            {
                if (e.RowIndex >= 0)
                {
                    CurrentPartID = (int)PartSelectGrV.CurrentRow.Cells["InfoID"].Value;
                    InfoBindingSource.DataSource = GetInfoByID(PartSelectGrV.CurrentRow.Cells["InfoID"].Value.ToString());
                    ApprBindingSource.DataSource = GetApprCheck(PartSelectGrV.CurrentRow.Cells["InfoID"].Value.ToString());
                    DimBindingSource.DataSource = GetDimCheck(PartSelectGrV.CurrentRow.Cells["InfoID"].Value.ToString());
                    RevBindingSource.DataSource = GetRevision(PartSelectGrV.CurrentRow.Cells["InfoID"].Value.ToString());
                    ResetAllPanel();
                    MeasureMenu.Enabled = false;
                    ChecksheetMenu.Enabled = false;
                    PartDataMenu.BackColor = Color.LightBlue;
                    PartInfoEdit.Visible = true;
                    NewPartBt.Enabled = false;
                    DeletePartBt.Enabled = false;
                    PartInfoGrV.Columns["ID"].Visible = false;
                    ApprCheckGrV.Columns["UniqueID"].Visible = false;
                    DimCheckGrV.Columns["UniqueDimID"].Visible = false;
                    RevisionGrV.Columns["RevPartID"].Visible = false;
                }
            }
            else if (!ChecksheetMenu.Enabled)
            {
            }
            else if (!MeasureMenu.Enabled)
            {
                if (e.RowIndex >= 0)
                {
                    CurrentPartID = (int)PartSelectGrV.CurrentRow.Cells["InfoID"].Value;
                    MoldNoLbl.Text = PartSelectGrV.CurrentRow.Cells["InfoMoldNo"].Value.ToString();
                    PartCodeLbl.Text = PartSelectGrV.CurrentRow.Cells["InfoPartNo"].Value.ToString();
                    ResetAllPanel();
                    //Disable other tabs
                    ManageMenu.Enabled = false;
                    PartDataMenu.Enabled = false;
                    ChecksheetMenu.Enabled = false;
                    MeasureMenu.BackColor = Color.LightBlue;
                    MeasurePanel.Visible = true;
                    InjectionDatePicker_ValueChanged(null, EventArgs.Empty);
                    string Location = PartSelectGrV.CurrentRow.Cells["InfoLocation"].Value.ToString();
                    int h = DateTime.Now.Hour;
                    if (Location.Contains("SVN3"))
                    {
                        if (h >= 8 && h <= 17)
                        {
                            ShiftTxb.Text = "MS";
                        }
                        else
                        {
                            ShiftTxb.Text = "NS";
                        }
                    }
                    else
                    {
                        if (h >= 6 && h < 14)
                        {
                            ShiftTxb.Text = "MS";
                        }
                        else if (h >= 14 && h < 22)
                        {
                            ShiftTxb.Text = "AS";
                        }
                        else
                        {
                            ShiftTxb.Text = "NS";
                        }
                    }
                }
            }
        }
        private void PartSelectGrV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!ChecksheetMenu.Enabled)
            {
                ResultInjectionDatePicker_ValueChanged(null, EventArgs.Empty);
            }

        }
        private void PartSelectOpenBt_Click(object sender, EventArgs e)
        {
            PartSelectGrV_CellMouseDoubleClick(this.PartSelectGrV, new DataGridViewCellEventArgs(this.PartSelectGrV.CurrentCell.ColumnIndex, this.PartSelectGrV.CurrentRow.Index));
        }
        private void MoldNoTxb_KeyUp(object sender, KeyEventArgs e)
        {
            GetInfo(MoldNoTxb.Text, 1);
            dataAdapter.Update((DataTable)InfoBindingSource.DataSource);
            PartCodeTxb.Clear();
        }
        private void PartCodeTxb_KeyUp(object sender, KeyEventArgs e)
        {
            GetInfo(PartCodeTxb.Text, 2);
            dataAdapter.Update((DataTable)InfoBindingSource.DataSource);
            MoldNoTxb.Clear();
        }
        private void SVN2ChkBx_CheckedChanged(object sender, EventArgs e)
        {
            if (SVN2ChkBx.Checked)
            {
                if (!SVN3ChkBx.Checked)
                {
                    GetInfo("SVN2", 3);
                }
                else
                {
                    GetInfo("", 1);
                }
            }
            else
            {
                if (!SVN3ChkBx.Checked)
                {
                    SVN3ChkBx.Checked = true;
                }
                SVN3ChkBx_CheckedChanged(null, EventArgs.Empty);
            }
            dataAdapter.Update((DataTable)InfoBindingSource.DataSource);
        }
        private void SVN3ChkBx_CheckedChanged(object sender, EventArgs e)
        {
            if (SVN3ChkBx.Checked)
            {
                if (!SVN2ChkBx.Checked)
                {
                    GetInfo("SVN3", 3);
                }
                else
                {
                    GetInfo("", 1);
                }
            }
            else
            {
                if (!SVN2ChkBx.Checked)
                {
                    SVN2ChkBx.Checked = true;
                }
                SVN2ChkBx_CheckedChanged(null, EventArgs.Empty);
            }
            dataAdapter.Update((DataTable)InfoBindingSource.DataSource);

        }
        //Function for retrieving data from server to datagridview-----------------------------------
        private void GetInfo(string Clue, int type)
        {
            string Adapterstring;
            if (type == 1)
            {
                Adapterstring = "Select * from PartInfo where MoldNo like '%" + Clue + "%' order by ID DESC";
            }
            else if (type == 2)
            {
                Adapterstring = "Select * from PartInfo where PartNo like '%" + Clue + "%' order by ID DESC";
            }
            else
            {
                Adapterstring = "Select * from PartInfo where Location like '%" + Clue + "%' order by ID DESC";
            }
            InfoAdapter = new SqlDataAdapter(Adapterstring, connectionString);
            _ = new SqlCommandBuilder(InfoAdapter);
            DataTable table = new DataTable { };
            InfoAdapter.Fill(table);
            InfoBindingSource.DataSource = table;
        }
        public DataTable GetInfoByID(string ID)
        {
            DataTable table = new DataTable { };
            dataAdapter = new SqlDataAdapter("Select * from PartInfo where ID = '" + ID + "'", connectionString);
            _ = new SqlCommandBuilder(dataAdapter);
            dataAdapter.Fill(table);
            return table;
        }
        public DataTable GetApprCheck(string ID)
        {
            DataTable table = new DataTable { };
            ApprAdapter = new SqlDataAdapter("Select * from AppearanceCheckPoint where ID = '" + ID + "'", connectionString);
            _ = new SqlCommandBuilder(ApprAdapter);
            ApprAdapter.Fill(table);
            return table;
        }
        public DataTable GetDimCheck(string ID)
        {
            DataTable table = new DataTable { };
            DimAdapter = new SqlDataAdapter("Select * from DimensionCheckPoint where ID = '" + ID + "'", connectionString);
            _ = new SqlCommandBuilder(DimAdapter);
            DimAdapter.Fill(table);
            return table;
        }
        private DataTable GetRevision(string ID)
        {
            DataTable table = new DataTable { };
            RevAdapter = new SqlDataAdapter("Select * from CheckingStanRev where ID = '" + ID + "'", connectionString);
            _ = new SqlCommandBuilder(RevAdapter);
            RevAdapter.Fill(table);
            return table;
        }
        private DataTable GetUser()
        {
            DataTable table = new DataTable { };
            UserAdapter = new SqlDataAdapter("Select * from UserAccount", connectionString);
            _ = new SqlCommandBuilder(UserAdapter);
            UserAdapter.Fill(table);
            return table;
        }
        private DataTable GetAllTool()
        {
            DataTable table = new DataTable { };
            ToolAdapter = new SqlDataAdapter("Select * from ToolList", connectionString);
            _ = new SqlCommandBuilder(ToolAdapter);
            ToolAdapter.Fill(table);
            return table;
        }
        private string[] GetBatch(DateTime Date, int ID)
        {
            string date = Date.ToString("dd/MM/yyyy");
            List<string> result = new List<string>();
            SqlCommand CollectBatch = new SqlCommand("Select distinct Batch from MeasurementResult where InjectionDate ='" + date + "' and ID ='" + ID.ToString() + "'", cnn);
            using (SqlDataReader Reader = CollectBatch.ExecuteReader())
            {
                while (Reader.Read())
                {
                    result.Add(Reader["Batch"].ToString());
                }
            }
            string[] Batch = result.ToArray();
            return Batch;
        }
        public DataTable GetAllBatch(string date, int ID)
        {
            DataTable table = new DataTable { };
            SqlDataAdapter BatchAdapter = new SqlDataAdapter("Select distinct InjectionDate, Batch, MeasureDate, MachineNo, Note from MeasurementResult where ID = '" + ID + "'", connectionString);
            _ = new SqlCommandBuilder(BatchAdapter);
            BatchAdapter.Fill(table);
            DataTable table1 = table.Clone();

            //remove duplicated rows
            if (table.Rows.Count == 0)
            {
                return table;
            }
            if (table.Rows[table.Rows.Count - 1]["InjectionDate"].ToString() == date
                || table.Rows[table.Rows.Count - 1]["InjectionDate"].ToString() == "all")
            {
                table1.ImportRow(table.Rows[table.Rows.Count - 1]);
            }
            for (int i = table.Rows.Count - 1; i > 0; i--)
            {
                if ((table.Rows[i]["InjectionDate"].ToString() != table.Rows[i - 1]["InjectionDate"].ToString()
                    || table.Rows[i]["Batch"].ToString() != table.Rows[i - 1]["Batch"].ToString())
                    && (table.Rows[i - 1]["InjectionDate"].ToString() == date || date == "all")
                    && table.Rows[i - 1]["Batch"].ToString() != "")
                {
                    table1.ImportRow(table.Rows[i - 1]);
                }
            }
            return table1;
        }
        public DataTable GetResult(string batch, string date, string ID)
        {
            DataTable table = new DataTable { };
            ResultAdapter = new SqlDataAdapter("Select * from MeasurementResult where InjectionDate = '" + date + "' and Batch = '" + batch + "' and ID = '" + ID + "'", connectionString);
            _ = new SqlCommandBuilder(ResultAdapter);
            ResultAdapter.Fill(table);
            return table;
        }
        public DataTable GetToolByType(string Type)
        {
            DataTable table = new DataTable { };
            ToolAdapter = new SqlDataAdapter("Select * from ToolList where ToolType = '" + Type + "'", connectionString);
            _ = new SqlCommandBuilder(ToolAdapter);
            ToolAdapter.Fill(table);
            return table;
        }
        public DataTable GetToolInfo(string ToolName)
        {
            DataTable table = new DataTable { };
            ToolAdapter = new SqlDataAdapter("Select * from ToolList where ToolName = '" + ToolName + "'", connectionString);
            _ = new SqlCommandBuilder(ToolAdapter);
            ToolAdapter.Fill(table);
            return table;
        }
        private void ReleaseAllTool()
        {
            using (SqlCommand ReleaseTool = new SqlCommand("Update ToolList set CurrentUser = null where CurrentUser ='" + UsernameTxb.Text + "'", cnn))
            {
                if (cnn.State == ConnectionState.Open)
                {
                    ReleaseTool.ExecuteNonQuery();
                }
            }
        }

        //Checksheet and result functions-------------------------------------------------------------
        private void ChecksheetsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Data
            SVN2ChkBx_CheckedChanged(null, EventArgs.Empty);
            BlankChkBx.Checked = true;
            if (BatchListGrV.DataSource != null)
            {
                DataTable table = (DataTable)BatchListGrV.DataSource;
                table.Clear();
            }
            //Sizing and positioning
            ChkSheetPanel.Size = new Size(Convert.ToInt32(WW * 0.5), WH - 100);
            ChkSheetPanel.Location = new Point(WW / 2 - 80, 50);
            //Visibility and coloring
            ResetAllPanel();
            ChkSheetPanel.Visible = true;
            PartSelectPanel.Visible = true;
            ChecksheetMenu.BackColor = Color.LightBlue;
            ChecksheetMenu.Enabled = false;
        }
        private void BlankChkBx_CheckedChanged(object sender, EventArgs e)
        {
            if (BlankChkBx.Checked)
            {
                InjectionDateChkBx.Checked = false;
                ResultInjectionDatePicker.Enabled = false;
                BatchListGrV.Enabled = false;
                EngChkBx.Enabled = true;
            }
            else
            {
                EngChkBx.Checked = false;
                EngChkBx.Enabled = false;
                InjectionDateChkBx.Checked = true;
            }
        }
        private void FolderBrowseBt_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog ChkSheetFolder = new FolderBrowserDialog())
            {
                DialogResult result = ChkSheetFolder.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ExportPathTxb.Text = ChkSheetFolder.SelectedPath.ToString();
                    Properties.Settings.Default.DefaultExportPath = ChkSheetFolder.SelectedPath.ToString();
                    Properties.Settings.Default.Save();
                }
            }
        }
        private void ExportBT_Click(object sender, EventArgs e)
        {
            //Check for error
            if (ExportPathTxb.Text == "")
            {
                MessageBox.Show("Please choose export path first." + Environment.NewLine +
                    "Chọn thư mục xuất checksheet.");
                FolderBrowseBt_Click(null, EventArgs.Empty);
            }
            if (!EngChkBx.Checked && !QaChkBx.Checked && !IpqcChkBx.Checked && !CanonChkBx.Checked)
            {
                MessageBox.Show("Please select at least one type of checksheet to export!" + Environment.NewLine +
                    "Chọn ít nhất một loại checksheet để xuất");
                return;
            }
            if (BatchListGrV.CurrentRow == null && !BlankChkBx.Checked)
            {
                MessageBox.Show("Please select batch." + Environment.NewLine + "Chọn batch.");
                return;
            }
            //Get part info data from current selected row on datagridview to an array.
            string IDSelected = PartSelectGrV.CurrentRow.Cells["InfoID"].Value.ToString();
            DataTable InfoTable = GetInfoByID(IDSelected);
            DataTable ApprTable = GetApprCheck(IDSelected);
            DataTable DimTable = GetDimCheck(IDSelected);
            //DataTable ResultTable = GetResult(BatchListGrV.CurrentRow.Cells["ResultBatch"].Value.ToString(), BatchListGrV.CurrentRow.Cells["BatchInjectionDate"].Value.ToString(), CurrentPartID.ToString());
            String ExportResult = "Exported checksheet to:";
            string QAFileName = Exporter.GenerateFileName(InfoTable.Rows[0]["MoldNo"].ToString(), InfoTable.Rows[0]["PartNo"].ToString(), "QA");
            string ENGFileName = Exporter.GenerateFileName(InfoTable.Rows[0]["MoldNo"].ToString(), InfoTable.Rows[0]["PartNo"].ToString(), "CheckingStan");
            string IPQCFileName = Exporter.GenerateFileName(InfoTable.Rows[0]["MoldNo"].ToString(), InfoTable.Rows[0]["PartNo"].ToString(), "IPQC");
            if (EngChkBx.Checked)
            {
                try
                {
                    File.WriteAllBytes(ExportPathTxb.Text + "\\" + ENGFileName, Properties.Resources.CheckSheetTemplate);
                }
                catch (Exception)
                {
                    MessageBox.Show("Cant access selected folder. Please choose another path!" + Environment.NewLine +
                        "Không thể lưu file vào đường dẫn đã chọn. Xin chọn thư mục khác.", "Access denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FileInfo DesFile = new FileInfo(ExportPathTxb.Text + "\\" + ENGFileName);
                ExcelPackage DestinationFile = new ExcelPackage(DesFile);
                Exporter.AddENGInfo(DestinationFile.Workbook.Worksheets["Checking Standard"], InfoTable);
                Exporter.AddEngAppr(DestinationFile.Workbook.Worksheets["Checking Standard"], ApprTable);
                Exporter.AddEngDim(DestinationFile.Workbook.Worksheets["Checking Standard"], DimTable);
                Exporter.AddEngRev(DestinationFile.Workbook.Worksheets["Checking Standard"], GetRevision(IDSelected));
                DestinationFile.Workbook.Worksheets.Delete("QA Inspection sheet");
                DestinationFile.Workbook.Worksheets.Delete("IPQC Inspection sheet");
                DestinationFile.Save();
                ExportResult += Environment.NewLine + ENGFileName;
            }
            if (QaChkBx.Checked)
            {
                if (InjectionDateChkBx.Checked)
                {
                    QAFileName = QAFileName.Replace(".xlsx", "_");
                    QAFileName += ResultInjectionDatePicker.Value.ToString("dd-MM-yyyy") + "_Batch" + BatchListGrV.CurrentRow.Cells["ResultBatch"].Value.ToString() + ".xlsx";
                }
                try
                {
                    File.WriteAllBytes(ExportPathTxb.Text + "\\" + QAFileName, Properties.Resources.CheckSheetTemplate);
                }
                catch (Exception)
                {
                    MessageBox.Show("Cant access selected folder. Please choose another path!" + Environment.NewLine +
                        "Không thể lưu file vào đường dẫn đã chọn. Xin chọn thư mục khác.", "Access denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FileInfo DesFile = new FileInfo(ExportPathTxb.Text + "\\" + QAFileName);
                ExcelPackage DestinationFile = new ExcelPackage(DesFile);
                Exporter.AddQAInfo(DestinationFile.Workbook.Worksheets["QA Inspection sheet"], InfoTable);
                Exporter.AddQAAppr(DestinationFile.Workbook.Worksheets["QA Inspection sheet"], ApprTable);
                Exporter.AddQADim(DestinationFile.Workbook.Worksheets["QA Inspection sheet"], DimTable);
                if (!BlankChkBx.Checked)
                {
                    Exporter.AddResult(DestinationFile.Workbook.Worksheets["QA Inspection sheet"],
                        GetResult(BatchListGrV.CurrentRow.Cells["ResultBatch"].Value.ToString(), BatchListGrV.CurrentRow.Cells["BatchInjectionDate"].Value.ToString(), CurrentPartID.ToString()), InfoTable);
                }
                DestinationFile.Workbook.Worksheets.Delete("Checking Standard");
                DestinationFile.Workbook.Worksheets.Delete("IPQC Inspection sheet");
                DestinationFile.Save();
                ExportResult += Environment.NewLine + QAFileName;
            }
            if (IpqcChkBx.Checked)
            {
                if (InjectionDateChkBx.Checked)
                {
                    IPQCFileName = IPQCFileName.Replace(".xlsx", "_");
                    IPQCFileName += ResultInjectionDatePicker.Value.ToString("dd-MM-yyyy") + "_Batch" + BatchListGrV.CurrentRow.Cells["ResultBatch"].Value.ToString() + ".xlsx";
                }

                try
                {
                    File.WriteAllBytes(ExportPathTxb.Text + "\\" + IPQCFileName, Properties.Resources.CheckSheetTemplate);
                }
                catch (Exception)
                {
                    MessageBox.Show("Cant access selected folder. Please choose another path!" + Environment.NewLine +
                        "Không thể lưu file vào đường dẫn đã chọn. Xin chọn thư mục khác.", "Access denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FileInfo DesFile = new FileInfo(ExportPathTxb.Text + "\\" + IPQCFileName);
                ExcelPackage DestinationFile = new ExcelPackage(DesFile);
                Exporter.AddIPQCInfo(DestinationFile.Workbook.Worksheets["IPQC Inspection sheet"], InfoTable);
                Exporter.AddIPQCAppr(DestinationFile.Workbook.Worksheets["IPQC Inspection sheet"], ApprTable);
                Exporter.AddIPQCDim(DestinationFile.Workbook.Worksheets["IPQC Inspection sheet"], DimTable);
                if (!BlankChkBx.Checked)
                {
                    Exporter.AddResult(DestinationFile.Workbook.Worksheets["IPQC Inspection sheet"],
                        GetResult(BatchListGrV.CurrentRow.Cells["ResultBatch"].Value.ToString(), BatchListGrV.CurrentRow.Cells["BatchInjectionDate"].Value.ToString(), CurrentPartID.ToString()), InfoTable);
                }
                DestinationFile.Workbook.Worksheets.Delete("Checking Standard");
                DestinationFile.Workbook.Worksheets.Delete("QA Inspection sheet");
                DestinationFile.Save();
                ExportResult += Environment.NewLine + IPQCFileName;
            }
            //RecordLog(ExportResult, IDSelected.ToString());
            var OpenResult = MessageBox.Show(ExportResult + Environment.NewLine + "Open checksheet file now?" + Environment.NewLine +
                "Mở file checksheet?", "Result", MessageBoxButtons.YesNo);
            //open file with excel
            if (OpenResult == DialogResult.Yes)
            {
                FileInfo QAFI = new FileInfo(ExportPathTxb.Text + "\\" + QAFileName);
                FileInfo ENGFI = new FileInfo(ExportPathTxb.Text + "\\" + ENGFileName);
                FileInfo IPQCFI = new FileInfo(ExportPathTxb.Text + "\\" + IPQCFileName);
                if (QaChkBx.Checked && QAFI.Exists)
                {
                    Process.Start(@ExportPathTxb.Text + "\\" + QAFileName);
                }
                if (EngChkBx.Checked && ENGFI.Exists)
                {
                    Process.Start(@ExportPathTxb.Text + "\\" + ENGFileName);
                }
                if (IpqcChkBx.Checked && IPQCFI.Exists)
                {
                    Process.Start(@ExportPathTxb.Text + "\\" + IPQCFileName);
                }
            }
        }
        private void SelectInjectionDateChkBx_CheckedChanged(object sender, EventArgs e)
        {
            if (InjectionDateChkBx.Checked)
            {
                ResultInjectionDatePicker_ValueChanged(null, EventArgs.Empty);
                BlankChkBx.Checked = false;
                ResultInjectionDatePicker.Enabled = true;
                BatchListGrV.Enabled = true;
            }
            else
            {
                BlankChkBx.Checked = true;
                AllBatchChkBx.Checked = false;
                BatchListGrV.Enabled = false;
            }

        }
        private void ResultInjectionDatePicker_ValueChanged(object sender, EventArgs e)
        {
            CurrentPartID = (int)PartSelectGrV.CurrentRow.Cells["InfoID"].Value;
            if (AllBatchChkBx.Checked)
            {
                BatchListGrV.DataSource = GetAllBatch("all", CurrentPartID);
            }
            else
            {
                BatchListGrV.DataSource = GetAllBatch(ResultInjectionDatePicker.Value.ToString("dd/MM/yyyy"), CurrentPartID);
            }

        }
        private void OpenExportBt_Click(object sender, EventArgs e)
        {

        }
        private void AllBatchChkBx_CheckedChanged(object sender, EventArgs e)
        {
            if (AllBatchChkBx.Checked)
            {
                InjectionDateChkBx.Checked = true;
            }
            SelectInjectionDateChkBx_CheckedChanged(null, EventArgs.Empty);

        }
        private void ViewResultBt_Click(object sender, EventArgs e)
        {
            if (BatchListGrV.CurrentRow != null)
            {
                ResultForm ViewResult;
                ViewResult = new ResultForm();
                ViewResult.Show(this);
                CurrentPartID = (int)PartSelectGrV.CurrentRow.Cells["InfoID"].Value;
            }
        }

        //UserAccount and Tool manage functions-------------------------------------------------------
        private void ManageMenuItem_Click(object sender, EventArgs e)
        {
            ResetAllPanel();
            ManageMenu.BackColor = Color.LightBlue;
            ManagePnl.Size = new Size(WW - 25, WH - 70);
            ManagePnl.Location = new Point(15, 45);
            ManagePnl.Visible = true;
            UserGrV.Size = new Size(400, WH - 100);
            ToolGrV.Size = new Size(500, WH - 100);
            UserGrV.DataSource = GetUser();
            ToolGrV.DataSource = GetAllTool();
        }
        private void UserGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value == null)
            {
                return;
            }
            if (UserGrV.Columns[e.ColumnIndex].Index == UserGrV.Columns["Password"].Index && e.Value != null)
            {
                UserGrV.Rows[e.RowIndex].Tag = e.Value;
                e.Value = new String('*', e.Value.ToString().Length);
            }
        }
        private void UserSaveBt_Click(object sender, EventArgs e)
        {
            try
            {
                UserAdapter.Update((DataTable)UserGrV.DataSource);
                RecordLog(UsernameTxb.Text + ": User data changed.", "00");
            }
            catch (SqlException)
            {
                MessageBox.Show("Save error! Username, Password or User Level cannot be blank." + Environment.NewLine +
                    "Lưu thất bại. Username Password và User Level không được để trống.");
            }

        }
        private void ToolSaveBt_Click(object sender, EventArgs e)
        {
            ToolAdapter.Update((DataTable)ToolGrV.DataSource);
        }

        //Measure interface functions------------------------------------------------------------------
        private void MeasureMenu_Click(object sender, EventArgs e)
        {
            //Visibility
            SVN2ChkBx_CheckedChanged(null, EventArgs.Empty);
            ResetAllPanel();
            PartSelectOpenBt.Visible = true;
            AllDimChkBt.Visible = true;
            QADimChkBt.Visible = true;
            IPQCDimChkBt.Visible = true;
            MeasureMenu.Enabled = false;
            MeasureMenu.BackColor = Color.LightBlue;
            PartSelectPanel.Visible = true;
            //Sizing and positioning
            MeasurePanel.Size = new Size(WW - 25, WH - 70);
            MeasurePanel.Location = new Point(15, 45);
        }
        private void MeaSaveBt_Click(object sender, EventArgs e)
        {
            DataTable Result = (DataTable)MeasureResultGrV.DataSource;
            //Compress data from all Cavity column into ValueList Column
            for (int i = 0; i < Result.Rows.Count; i++)
            {
                Result.Rows[i]["ValueList"] = "";
                for (int j = 25; j < Result.Columns.Count; j++)
                {
                    if (Result.Rows[i][j].ToString() == "")
                    {
                        Result.Rows[i][j] = "_";
                    }
                    Result.Rows[i]["ValueList"] += Result.Rows[i][j] + ",";
                }
                //remove last ','
                Result.Rows[i]["ValueList"] = Result.Rows[i]["ValueList"].ToString().TrimEnd(',');
            }
            bool IncompleteFlag = false;
            for (int i = 0; i < Result.Rows.Count; i++)
            {
                Result.Rows[0]["Shift"] = ShiftTxb.Text.Trim();
                Result.Rows[0]["Temp"] = TempTxb.Text.Trim();
                Result.Rows[0]["Humid"] = HumidTxb.Text.Trim();
                Result.Rows[0]["MachineNo"] = MCNoTxb.Text.Trim();
                Result.Rows[0]["Note"] = NoteTxb.Text.Trim();
                if (Result.Rows[i]["Judge"].ToString() == "False")
                {
                    IncompleteFlag = true;
                }
                if (Result.Rows[i]["ApprID"].ToString() == "" && Result.Rows[i]["ValueList"].ToString() == "")
                {
                    IncompleteFlag = true;
                }
            }
            if (IncompleteFlag)
            {
                var Confirm = MessageBox.Show("There is empty value or NG checkpoint(s). Do you still want to save?" + Environment.NewLine +
                    "Vẫn còn checkpoint chưa có kết quả hoặc có kết quả NG. Bạn vẫn muốn lưu kết quả?",
                    "Incomplete Measurement", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (Confirm == DialogResult.No)
                {
                    return;
                }
            }
            ResultAdapter = new SqlDataAdapter("Select * from MeasurementResult", connectionString);
            _ = new SqlCommandBuilder(ResultAdapter);
            ResultAdapter.Update(Result);
            RecordLog(UsernameTxb.Text + " save measurement data Inj date =" + Result.Rows[1]["InjectionDate"].ToString() +
                " batch =" + Result.Rows[1]["Batch"].ToString(), CurrentPartID.ToString());
            changeUnsaved = false;
        }
        private void MeaCloseBt_Click(object sender, EventArgs e)
        {
            if (changeUnsaved)
            {
                var confirm = MessageBox.Show("Do you want to save the changes you made?" + Environment.NewLine +
                    "Bạn có muốn lưu thay đổi?", "Unsaved change", MessageBoxButtons.YesNoCancel);
                if (confirm == DialogResult.Yes)
                {
                    MeaSaveBt_Click(null, EventArgs.Empty);
                }
                else if (confirm == DialogResult.No)
                {
                }
                else if (confirm == DialogResult.Cancel)
                {
                    return;
                }
            }
            MeasureMenu_Click(null, EventArgs.Empty);
            BatchCBx.Text = "";
            BatchLbl.Text = "Batch:";
            DataTable table = (DataTable)MeasureResultGrV.DataSource;
            if (table != null)
            {
                table.Clear();
            }
            BatchCBx.Items.Clear();
            MCNoTxb.Clear();
            ShiftTxb.Clear();
            MeaSaveBt.Enabled = false;
            changeUnsaved = false;
            //release all tool used by this user
            ReleaseAllTool();
        }
        private void InjectionDatePicker_ValueChanged(object sender, EventArgs e)
        {
            BatchCBx.Items.Clear();
            BatchCBx.Items.AddRange(GetBatch(InjectionDatePicker.Value, CurrentPartID));
            if (UserLevel > 1)
            {
                BatchCBx.Items.Add("New batch");
            }
            DataTable table = (DataTable)MeasureResultGrV.DataSource;
            if (table != null)
            {
                table.Clear();
            }
        }
        private void BatchCBx_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable Info = GetInfoByID(CurrentPartID.ToString());
            DataTable Appr = GetApprCheck(CurrentPartID.ToString());
            DataTable Dim = GetDimCheck(CurrentPartID.ToString());
            string InjectionDate = InjectionDatePicker.Value.ToString("dd/MM/yyyy");
            string MeasuredBy = UsernameTxb.Text;
            if (BatchCBx.SelectedItem.ToString() == "New batch")
            {
                int NewBatch = 1;
                if (BatchCBx.Items.Count >= 2)
                {
                    NewBatch = int.Parse(BatchCBx.Items[BatchCBx.Items.Count - 2].ToString()) + 1;
                }
                BatchLbl.Text = "Batch: " + NewBatch.ToString();
                string Select = "all";
                if (AllDimChkBt.Checked)
                {
                    Select = "all";
                }
                else if (QADimChkBt.Checked)
                {
                    Select = "QA";
                }
                else if (IPQCDimChkBt.Checked)
                {
                    Select = "IPQC";
                }
                DataTable MeasureList = Exporter.CreateBatchTable(Info, Appr, Dim, NewBatch, InjectionDate, MeasuredBy, Select);
                MeasureList.Rows[0]["Shift"] = ShiftTxb.Text;
                MeasureList.Rows[0]["Temp"] = TempTxb.Text;
                MeasureList.Rows[0]["Humid"] = HumidTxb.Text;
                MeasureList.Rows[0]["MachineNo"] = MCNoTxb.Text;
                MeasureResultGrV.DataSource = MeasureList;
                MCNoTxb.Clear();
            }
            else
            {
                int SelectedBatch = int.Parse(BatchCBx.Text);
                BatchLbl.Text = "Batch: " + BatchCBx.Text;
                DataTable BatchResults = GetResult(SelectedBatch.ToString(), InjectionDate, CurrentPartID.ToString());
                MeasureResultGrV.DataSource = Exporter.FillInfoToResult(Info, Appr, Dim, BatchResults);
                ShiftTxb.Text = BatchResults.Rows[0]["Shift"].ToString();
                TempTxb.Text = BatchResults.Rows[0]["Temp"].ToString();
                HumidTxb.Text = BatchResults.Rows[0]["Humid"].ToString();
                MCNoTxb.Text = BatchResults.Rows[0]["MachineNo"].ToString();
                NoteTxb.Text = BatchResults.Rows[0]["Note"].ToString();
            }
            MeasureGrVHighlight();
            MeasureResultGrV.Focus();
            MeaSaveBt.Enabled = true;
            FilterGrb.Enabled = true;
            if (UserLevel > 1)
            {
                QuickJudgeGrb.Enabled = true;
                ToolGrb.Enabled = true;
            }


        }
        private void ApprOKBt_Click(object sender, EventArgs e)
        {
            DataTable table = (DataTable)MeasureResultGrV.DataSource;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["DimID"].ToString() == "")
                {
                    table.Rows[i]["Judge"] = true;
                    for (int j = 25; j < table.Columns.Count; j++)
                    {
                        table.Rows[i][j] = "OK";
                    }
                }
            }
            MeasureResultGrV.DataSource = table;
            MeasureGrVHighlight();
            HideApprChkBx.Checked = false;
        }
        private void Filter_CheckedChanged(object sender, EventArgs e)
        {
            MeasureResultGrV.Columns["MeasSpecs"].Width = 300;
            for (int i = 0; i < MeasureResultGrV.Rows.Count; i++)
            {
                MeasureResultGrV.RowHeadersVisible = false;
                //Set all to visible
                MeasureResultGrV.Rows[i].Visible = true;
                MeasureResultGrV.CurrentCell = null;
                //Hide appr
                if (MeasureResultGrV.Rows[i].Cells["MeasDimID"].Value.ToString() == "" && HideApprChkBx.Checked)
                {
                    MeasureResultGrV.Rows[i].Visible = false;
                    MeasureResultGrV.Columns["MeasSpecs"].Width = 80;
                }
                //Hide dim
                if (MeasureResultGrV.Rows[i].Cells["MeasApprID"].Value.ToString() == "" && HideDimChkBx.Checked)
                {
                    MeasureResultGrV.Rows[i].Visible = false;
                }
                //Hide OK
                if (MeasureResultGrV.Rows[i].Cells["Judge"].Value.ToString() == "True" && HideAllOKChkBx.Checked)
                {
                    MeasureResultGrV.Rows[i].Visible = false;
                }
                MeasureResultGrV.RowHeadersVisible = true;
            }
        }
        private void DeleteBatchBt_Click(object sender, EventArgs e)
        {
            if (BatchCBx.Text == "New batch" || BatchCBx.Text == "")
            {
                return;
            }
            var result = MessageBox.Show("Confirm delete batch: " + BatchCBx.Text + " from date: " + InjectionDatePicker.Value.ToString("dd/MM/yyyy") + "."
                , "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                using (SqlCommand DeleteBatch = new SqlCommand("Delete from MeasurementResult where InjectionDate ='"
                    + InjectionDatePicker.Value.ToString("dd/MM/yyyy") + "' and Batch ='" + BatchCBx.Text + "'", cnn))
                {
                    DeleteBatch.ExecuteNonQuery();
                }
                RecordLog(UsernameTxb.Text + " delete measurement data Inj date =" + InjectionDatePicker.Value.ToString("dd/MM/yyyy") +
                " batch =" + BatchCBx.Text, CurrentPartID.ToString());
                DataTable table = (DataTable)MeasureResultGrV.DataSource;
                table.Clear();
                MeasureResultGrV.DataSource = table;
                BatchCBx.Text = "";
                BatchCBx.Items.Clear();
                BatchCBx.Items.AddRange(GetBatch(InjectionDatePicker.Value, CurrentPartID));
                BatchCBx.Items.Add("New batch");
                MeaSaveBt.Enabled = false;
            }
            else
            {
                return;
            }
        }
        private void MeasureGrVHighlight()
        {
            for (int i = 0; i < MeasureResultGrV.Rows.Count; i++)
            {
                if (MeasureResultGrV.Rows[i].Cells["Judge"].Value.ToString() == "True")
                {
                    MeasureResultGrV.Rows[i].Cells["Judge"].Style.BackColor = Color.LawnGreen;
                }
                else
                {
                    MeasureResultGrV.Rows[i].Cells["Judge"].Style.BackColor = Color.Red;
                }
                //Highlight OK/NG dimensions checkpoints
                if (AutoDimJudgeChkBx.Checked && MeasureResultGrV.Rows[i].Cells["MeasApprID"].Value.ToString() == "")
                {
                    //translate Range to boundary values
                    string Range = MeasureResultGrV.Rows[i].Cells["MeasRange"].Value.ToString();
                    if (float.TryParse(Range.Split('~')[0], out float Lower) &&
                        float.TryParse(Range.Split('~')[1], out float Upper))
                    {
                        bool FA1 = float.TryParse(MeasureResultGrV.Rows[i].Cells["MeasFaAcceptMax"].Value.ToString(), out float FaMax);
                        bool FA2 = float.TryParse(MeasureResultGrV.Rows[i].Cells["MeasFaAcceptMin"].Value.ToString(), out float FaMin);
                        //checking if all cavity values on the editted rows are within boundary values
                        for (int j = 26; j < MeasureResultGrV.Columns.Count; j++)
                        {
                            if (float.TryParse(MeasureResultGrV.Rows[i].Cells[j].Value.ToString(), out float Value))
                            {
                                if ((Value >= Lower || (FA2 && Value >= FaMin)) && (Value <= Upper || (FA1 && Value <= FaMax)))
                                {
                                    MeasureResultGrV.Rows[i].Cells[j].Style.BackColor = Color.LightGreen;
                                }
                                else
                                {
                                    MeasureResultGrV.Rows[i].Cells[j].Style.BackColor = Color.LightSalmon;
                                }
                            }
                        }
                    }
                }
            }

        }
        private void DimFilterBt_Click(object sender, EventArgs e)
        {
            MeasureResultGrV.Sort(MeasureResultGrV.Columns["MeasTool"], ListSortDirection.Ascending);
            MeasureGrVHighlight();
            HideApprChkBx.Checked = true;
            HideDimChkBx.Checked = false;
            MeasureResultGrV.Focus();
        }
        private void ApprFilterBt_Click(object sender, EventArgs e)
        {
            HideApprChkBx.Checked = false;
            HideDimChkBx.Checked = true;
            Filter_CheckedChanged(null, EventArgs.Empty);
            MeasureResultGrV.Focus();
        }
        private void ViewAllBt_Click(object sender, EventArgs e)
        {
            HideApprChkBx.Checked = false;
            HideDimChkBx.Checked = false;
            MeasureResultGrV.Sort(MeasureResultGrV.Columns["MeasDimID"], System.ComponentModel.ListSortDirection.Ascending);
            MeasureGrVHighlight();
            MeasureResultGrV.Focus();
        }
        private void MeasureResultGrV_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //Auto dim judge function
            if (AutoDimJudgeChkBx.Checked && MeasureResultGrV.Rows[e.RowIndex].Cells["MeasApprID"].Value.ToString() == ""
                && e.ColumnIndex > 25)
            {
                DataTable Result = (DataTable)MeasureResultGrV.DataSource;
                //translate Range to boundary values
                string Range = MeasureResultGrV.Rows[e.RowIndex].Cells["MeasRange"].Value.ToString();
                if (float.TryParse(Range.Split('~')[0], out float Lower) &&
                    float.TryParse(Range.Split('~')[1], out float Upper))
                {
                    bool FA1 = float.TryParse(MeasureResultGrV.Rows[e.RowIndex].Cells["MeasFaAcceptMax"].Value.ToString(), out float FaMax);
                    bool FA2 = float.TryParse(MeasureResultGrV.Rows[e.RowIndex].Cells["MeasFaAcceptMin"].Value.ToString(), out float FaMin);
                    //checking if all cavity values on the editted rows are within boundary values
                    //First assume that all cavities are OK
                    Result.Rows[e.RowIndex]["Judge"] = "True";
                    for (int i = 26; i < MeasureResultGrV.Columns.Count; i++)
                    {
                        if (float.TryParse(MeasureResultGrV.Rows[e.RowIndex].Cells[i].Value.ToString(), out float Value))
                        {
                            if ((Value >= Lower || (FA2 && Value >= FaMin)) && (Value <= Upper || (FA1 && Value <= FaMax)))
                            {
                                MeasureResultGrV.Rows[e.RowIndex].Cells[i].Style.BackColor = Color.LightGreen;
                            }
                            else
                            {
                                //if there is one value not within specs, mark judge NG
                                Result.Rows[e.RowIndex]["Judge"] = "False";
                            }
                        }
                        else
                        {
                            Result.Rows[e.RowIndex]["Judge"] = "False";
                        }
                    }
                }
            }
            if (e.ColumnIndex == MeasureResultGrV.Columns["Judge"].Index)
            {
                MeasureResultGrV.Rows[e.RowIndex].Cells["MeasToolUsed"].Value = CurrentToolCBx.Text;
            }
            MeasureGrVHighlight();
            changeUnsaved = true;
        }

        //File open buttons
        private void WIOpenBt_Click(object sender, EventArgs e)
        {
            DataTable table = GetInfoByID(CurrentPartID.ToString());
            string Path = table.Rows[0]["WIFile1"].ToString();
            try
            {
                Process.Start(Path);
            }
            catch (Exception)
            {
                MessageBox.Show("Open file failed. Check file path.");
            }
        }
        private void WIOpen2Bt_Click(object sender, EventArgs e)
        {
            DataTable table = GetInfoByID(CurrentPartID.ToString());
            string Path = table.Rows[0]["WIFile2"].ToString();
            try
            {
                Process.Start(Path);
            }
            catch (Exception)
            {
                MessageBox.Show("Open file failed. Check file path.");
            }

        }
        private void OpenDWGBt_Click(object sender, EventArgs e)
        {
            DataTable table = GetInfoByID(CurrentPartID.ToString());
            string Path = table.Rows[0]["DrawingFile"].ToString();
            try
            {
                Process.Start(Path);
            }
            catch (Exception)
            {
                MessageBox.Show("Open file failed. Check file path.");
            }
        }
        private void OpenHistoryBt_Click(object sender, EventArgs e)
        {
            DataTable table = GetInfoByID(CurrentPartID.ToString());
            string Path = table.Rows[0]["HistoryFile"].ToString();
            try
            {
                Process.Start(Path);
            }
            catch (Exception)
            {
                MessageBox.Show("Open file failed. Check file path.");
            }
        }

        //For auto input from U-Wave tools
        private void MeasureResultGrV_CurrentCellChanged(object sender, EventArgs e)
        {
            if (AutoInputChkBx.Checked)
            {
                if (MeasureResultGrV.CurrentCell == null || UserLevel < 2
                    || CurrentTool == MeasureResultGrV.CurrentRow.Cells["MeasTool"].Value.ToString())
                {
                    return;
                }
                //if user change from one tool to another
                if (CurrentTool != MeasureResultGrV.CurrentRow.Cells["MeasTool"].Value.ToString() && CurrentTool != "")
                {
                    CurrentToolCBx.Focus();
                }
                //record current tool
                CurrentTool = MeasureResultGrV.CurrentRow.Cells["MeasTool"].Value.ToString();
                //for appearance checkpoints
                if (MeasureResultGrV.CurrentRow.Cells["MeasDimID"].Value.ToString() == "")
                {
                    CurrentToolCBx.Items.Clear();
                    CurrentToolCBx.Items.Add(MeasureResultGrV.CurrentRow.Cells["MeasTool"].Value.ToString());
                    CurrentToolCBx.SelectedIndex = 0;
                    CurrentToolCBx.Enabled = false;
                    TAddressLbl.Text = "None";
                    RAddressLbl.Text = "None";
                }
                //for dimension checkpoints
                else if (MeasureResultGrV.CurrentRow.Cells["MeasApprID"].Value.ToString() == "")
                {
                    if (CurrentToolCBx.Text == "Vis")
                    {
                        CurrentToolCBx.Text = "";
                    }
                    CurrentToolCBx.Enabled = true;
                    CurrentToolCBx.Items.Clear();
                    //Get available tool and add them to combo box for select
                    DataTable table = GetToolByType(MeasureResultGrV.CurrentRow.Cells["MeasTool"].Value.ToString());
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        if (table.Rows[i]["CurrentUser"].ToString() == "" ||
                            table.Rows[i]["CurrentUser"].ToString() == UsernameTxb.Text || !AutoInputChkBx.Checked
                            || (DateTime.Now - (DateTime)table.Rows[i]["LeaseTime"]).TotalMinutes > 30)
                        {
                            //Only add free tool or tool used by this user. Or if not using auto input, show all tool.
                            //Also add tool that is leased for more thatn 60 minutes.
                            CurrentToolCBx.Items.Add(table.Rows[i]["ToolName"].ToString());
                        }
                    }
                    if (CurrentToolCBx.Items.Count == 1)
                    {
                        //if there is only 1 tool available, select it
                        CurrentToolCBx.SelectedIndex = 0;
                    }
                    if (table.Rows.Count == 0 || CurrentToolCBx.Items.Count == 0)
                    {
                        //if there is no auto input tool available
                        CurrentToolCBx.Items.Add(MeasureResultGrV.CurrentRow.Cells["MeasTool"].Value.ToString());
                        CurrentToolCBx.SelectedIndex = 0;
                    }
                    if (CurrentToolCBx.Items.Count > 1)
                    {
                        //Open the combo box for user to select tool.
                        CurrentToolCBx.DroppedDown = true;
                    }
                }
            }
        }
        private void MeasureResultGrV_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (MeasureResultGrV.CurrentRow.Cells["MeasDimID"].Value.ToString() == "")
            {
                return;
            }
        }
        private void CurrentToolCBx_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReleaseAllTool();
            //Take current tool
            using (SqlCommand TakeTool = new SqlCommand("Update ToolList set CurrentUser = '" + UsernameTxb.Text + "' where ToolName ='" + CurrentToolCBx.Text + "'", cnn))
            {
                TakeTool.ExecuteNonQuery();
            }
            //Update label with current channel address
            DataTable Tool = GetToolInfo(CurrentToolCBx.Text);
            if (Tool.Rows.Count == 0)
            {
                TAddressLbl.Text = "None";
                RAddressLbl.Text = "None";
                UWaveCheck.Enabled = false;
            }
            else
            {
                TAddressLbl.Text = Tool.Rows[0]["TChannel"].ToString();
                RAddressLbl.Text = Tool.Rows[0]["RChannel"].ToString();
                StartMeasureTime = DateTime.Now;
                StartMeasureRowIndex = MeasureResultGrV.CurrentCell.RowIndex;
                UWaveCheck.Enabled = true;
            }
            MeasureResultGrV.Focus();
        }
        private void AutoInputChkBx_CheckedChanged(object sender, EventArgs e)
        {
            if (AutoInputChkBx.Checked)
            {
                MeasureResultGrV_CurrentCellChanged(null, new DataGridViewCellEventArgs(MeasureResultGrV.CurrentCell.ColumnIndex, MeasureResultGrV.CurrentCell.RowIndex));
            }
            else
            {
                UWaveCheck.Enabled = false;
            }
        }






        //U-Wave communications functions----------------------------------------------------------------
        private void UWaveCommunicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ResetAllPanel();
            UWaveForm UWaveInterface;
            UWaveInterface = new UWaveForm();
            UWaveInterface.Show(this);
        }

        //Timer for checking for new U-Wave data
        private void UWaveCheck_Tick(object sender, EventArgs e)
        {
            if (AutoInputChkBx.Checked
                && CurrentToolCBx.SelectedItem != null
                && TAddressLbl.Text != "None" && RAddressLbl.Text != "None"
                && StartMeasureTime != null)
            {
                //looking into database to find any untaken measurement results
                //of current tool and time after user choose tool.
                DataTable Result = new DataTable { };
                SqlCommand FindNewData = new SqlCommand("Select * from UWaveData where Transmitter = @TAddress and Receiver = @RAddress and Time > @Time and Taken = 0", cnn);
                FindNewData.Parameters.AddWithValue("TAddress", TAddressLbl.Text);
                FindNewData.Parameters.AddWithValue("RAddress", RAddressLbl.Text);
                FindNewData.Parameters.AddWithValue("Time", StartMeasureTime);
                using (SqlDataReader Reader = FindNewData.ExecuteReader())
                {
                    Result.Load(Reader);
                }
                if (Result.Rows.Count > 0)
                {
                    //move to the first cavity column if current cell is not on any cavity columns.
                    if (MeasureResultGrV.CurrentCell.ColumnIndex < 26)
                    {
                        MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[MeasureResultGrV.CurrentCell.RowIndex].Cells[26];
                    }
                    //Insert the value to table or move back to the last inserted value
                    if (float.TryParse(Result.Rows[Result.Rows.Count - 1]["Value"].ToString(), out float Value))
                    {
                        MeasureResultGrV.CurrentCell.Value = Value;
                        //Trigger cell end edit to judge the checkpoint
                        MeasureResultGrV_CellEndEdit(null, new DataGridViewCellEventArgs(PreviousCol, PreviousRow));
                        //Move to next cell
                        PreviousRow = MeasureResultGrV.CurrentCell.RowIndex;
                        PreviousCol = MeasureResultGrV.CurrentCell.ColumnIndex;
                        if (HorizontalMoveChkBt.Checked)
                        {
                            if (PreviousCol < MeasureResultGrV.Columns.Count - 1)
                            {
                                MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[PreviousRow].Cells[PreviousCol + 1];
                            }
                            else
                            {
                                if (PreviousRow < MeasureResultGrV.Rows.Count - 1)
                                {
                                    MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[PreviousRow + 1].Cells[26];
                                }
                            }
                        }
                        else if (VerticalMoveChkBt.Checked)
                        {
                            if (PreviousRow < MeasureResultGrV.Rows.Count - 1)
                            {
                                //check if next row is the same tool or not
                                if (MeasureResultGrV.Rows[PreviousRow].Cells["MeasTool"].Value.ToString() == MeasureResultGrV.Rows[PreviousRow + 1].Cells["MeasTool"].Value.ToString())
                                {
                                    //Go to next row of the same columns
                                    MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[PreviousRow + 1].Cells[PreviousCol];
                                }
                                else if (PreviousCol < MeasureResultGrV.Columns.Count - 1)
                                {
                                    //Go to next col, at the row the user start measure with this tool
                                    MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[StartMeasureRowIndex].Cells[PreviousCol + 1];
                                }
                                else
                                {
                                    //Go to next row and first cavity
                                    MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[PreviousRow + 1].Cells[26];
                                }
                            }
                            else
                            {
                                if (PreviousCol < MeasureResultGrV.Columns.Count - 1)
                                {
                                    //Go to next col, at the row user start measure with this tool
                                    MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[StartMeasureRowIndex].Cells[PreviousCol + 1];
                                }
                            }

                        }

                    }
                    else if (Result.Rows[Result.Rows.Count - 1]["Value"].ToString() == "DataCancel")
                    {
                        if (PreviousRow > 0 && PreviousCol > 0)
                        {
                            MeasureResultGrV.CurrentCell = MeasureResultGrV.Rows[PreviousRow].Cells[PreviousCol];
                        }
                    }
                    //Mark the record as taken
                    SqlCommand MarkAsTaken = new SqlCommand("Update UWaveData set Taken = 1 where DataID = '" + Result.Rows[Result.Rows.Count - 1]["DataID"] + "'", cnn);
                    MarkAsTaken.ExecuteNonQuery();
                }
            }
        }

    }
}
