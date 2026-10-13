using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using HTKK.Properties;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

public class frmConfigUpdate : Form
{
	private string rootPath = "";

	private string rootUpdate = "";

	private static XmlDocument doc = new XmlDocument();

	private IContainer components = null;

	private TabControl tabControl1;

	private TabPage tabPage1;

	private GroupBox groupBox1;

	private Label lbVersion;

	private Label label3;

	private Label lbTimeCapNhat;

	private Label label1;

	private GroupBox groupBox3;

	private Button btnExit;

	private DataGridView gridData;

	private DataGridViewTextBoxColumn PathFile;

	private DataGridViewTextBoxColumn STT;

	private DataGridViewTextBoxColumn Version;

	private DataGridViewTextBoxColumn CreateDate;

	private Button btnUpdateNow;

	private Label label7;

	private Label lblVersionNew;

	private Panel panel1;

	private Label label8;

	public frmConfigUpdate()
	{
		InitializeComponent();
		rootPath = HTKKConfig.APPPATH.Substring(0, HTKKConfig.APPPATH.LastIndexOf("\\"));
		rootUpdate = rootPath + "\\AutoUpdate";
	}

	private void frmConfigUpdate_Load(object sender, EventArgs e)
	{
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		if (Directory.Exists(rootUpdate))
		{
			LoadInfoConfig();
			LoadHistoryUpdateVersion();
			((Control)lblVersionNew).Text = WinSettings.NewVersionOnServer;
			if (string.IsNullOrEmpty(WinSettings.NewVersionOnServer) || ((Control)lbVersion).Text == WinSettings.NewVersionOnServer)
			{
				((Control)btnUpdateNow).Enabled = false;
			}
		}
		else
		{
			MessageBox.Show("Hệ thống chưa tích hợp auto update", "Thông báo", (MessageBoxButtons)0);
		}
	}

	private void LoadHistoryUpdateVersion()
	{
		try
		{
			if (Directory.Exists(rootPath + "\\Version\\VersionInfo"))
			{
				DataTable dataTable = new DataTable();
				dataTable.Columns.Add("STT", typeof(int));
				dataTable.Columns.Add("Version", typeof(string));
				dataTable.Columns.Add("CreateDate", typeof(string));
				dataTable.Columns.Add("PathFile", typeof(string));
				dataTable.Columns.Add("Order", typeof(int));
				dataTable.Columns.Add("OrderDate", typeof(DateTime));
				DirectoryInfo directoryInfo = new DirectoryInfo(rootPath + "\\Version\\VersionInfo");
				FileInfo[] files = directoryInfo.GetFiles("*.txt");
				int num = 1;
				int num2 = files.Length;
				FileInfo[] array = files;
				foreach (FileInfo fileInfo in array)
				{
					dataTable.Rows.Add(num2--, fileInfo.Name, fileInfo.CreationTime.ToString("dd/MM/yyyy HH:mm:ss"), fileInfo.FullName, num++, fileInfo.CreationTime);
				}
				DataView defaultView = dataTable.DefaultView;
				defaultView.Sort = "OrderDate DESC";
				gridData.AutoGenerateColumns = false;
				gridData.DataSource = defaultView.ToTable();
				for (int j = 0; j < gridData.RowCount; j++)
				{
					gridData.Rows[j].Cells["Stt"].Value = j + 1;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void LoadInfoConfig()
	{
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = rootUpdate + "\\AppSchedulerCf.xml";
			if (!File.Exists(text))
			{
				MessageBox.Show("Không tìm thấy file cấu hình.Vui lòng kiểm tra lại", "Thông báo", (MessageBoxButtons)0);
				return;
			}
			doc.Load(text);
			if (((XmlNode)doc).SelectNodes("/AppSchedule/Task/IsProxy").Count <= 0 || ((XmlNode)doc).SelectNodes("/AppSchedule/Task/IsProxy")[0].InnerText == "1")
			{
			}
			if (((XmlNode)doc).SelectNodes("/AppSchedule/Task/IsUser").Count <= 0 || ((XmlNode)doc).SelectNodes("/AppSchedule/Task/IsUser")[0].InnerText == "1")
			{
			}
			((Control)lbTimeCapNhat).Text = ((((XmlNode)doc).SelectNodes("/AppSchedule/Task/TimeExcute").Count > 0) ? ((XmlNode)doc).SelectNodes("/AppSchedule/Task/TimeExcute")[0].InnerText : string.Empty);
			((Control)lbVersion).Text = ((((XmlNode)doc).SelectNodes("/AppSchedule/Task/AppVersion").Count > 0) ? ((XmlNode)doc).SelectNodes("/AppSchedule/Task/AppVersion")[0].InnerText : string.Empty);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnUpdate_Click(object sender, EventArgs e)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Expected O, but got Unknown
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = rootUpdate + "\\AppSchedulerCf.xml";
			if (!File.Exists(text))
			{
				MessageBox.Show("Không tìm thấy file cấu hình.Vui lòng kiểm tra lại", "Thông báo", (MessageBoxButtons)0);
				return;
			}
			doc.Load(text);
			XmlTextWriter val = new XmlTextWriter(text, Encoding.UTF8);
			try
			{
				val.Formatting = (Formatting)1;
				doc.Save((XmlWriter)(object)val);
				MessageBox.Show("Cập nhật thành công", "Thông báo", (MessageBoxButtons)0);
			}
			finally
			{
				((IDisposable)val)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnThoat_Click(object sender, EventArgs e)
	{
		try
		{
			((Form)this).Close();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void gridData_CellClick(object sender, DataGridViewCellEventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		DataGridView val = (DataGridView)sender;
		val.EndEdit();
	}

	private void gridData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		if (e.RowIndex != -1)
		{
			string text = gridData.Rows[e.RowIndex].Cells["PathFile"].Value.ToString();
			string version = Path.GetFileNameWithoutExtension(text).Replace("Version", string.Empty);
			frmVersionInfo frmVersionInfo2 = new frmVersionInfo(version, text);
			((Form)frmVersionInfo2).ShowDialog();
		}
	}

	private void btnUpdateNow_Click(object sender, EventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Invalid comparison between Unknown and I4
		try
		{
			DialogResult val = MessageBox.Show("Bạn muốn cập nhật phiên bản mới?", "Thông báo", (MessageBoxButtons)4);
			if ((int)val == 6)
			{
				Process.Start(rootUpdate + "\\Client_VersionUpdate.exe");
				Environment.Exit(0);
			}
		}
		catch (Exception ex)
		{
			Application.Exit();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmConfigUpdate_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Expected O, but got Unknown
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Expected O, but got Unknown
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Expected O, but got Unknown
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0085: Expected O, but got Unknown
		//IL_0086: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected O, but got Unknown
		//IL_0091: Unknown result type (might be due to invalid IL or missing references)
		//IL_009b: Expected O, but got Unknown
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a6: Expected O, but got Unknown
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Expected O, but got Unknown
		//IL_00b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Expected O, but got Unknown
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c7: Expected O, but got Unknown
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d2: Expected O, but got Unknown
		//IL_00d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dd: Expected O, but got Unknown
		//IL_00de: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Expected O, but got Unknown
		//IL_0179: Unknown result type (might be due to invalid IL or missing references)
		//IL_023e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0264: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_03a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_054e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0558: Expected O, but got Unknown
		//IL_0566: Unknown result type (might be due to invalid IL or missing references)
		//IL_0570: Expected O, but got Unknown
		//IL_077f: Unknown result type (might be due to invalid IL or missing references)
		//IL_07a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0816: Unknown result type (might be due to invalid IL or missing references)
		//IL_08bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_093c: Unknown result type (might be due to invalid IL or missing references)
		//IL_09bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a3a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b35: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b99: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ba3: Expected O, but got Unknown
		//IL_0bf2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c65: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c6f: Expected O, but got Unknown
		//IL_0c9e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d90: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d9a: Expected O, but got Unknown
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmConfigUpdate));
		tabControl1 = new TabControl();
		tabPage1 = new TabPage();
		btnExit = new Button();
		groupBox3 = new GroupBox();
		gridData = new DataGridView();
		PathFile = new DataGridViewTextBoxColumn();
		STT = new DataGridViewTextBoxColumn();
		Version = new DataGridViewTextBoxColumn();
		CreateDate = new DataGridViewTextBoxColumn();
		groupBox1 = new GroupBox();
		btnUpdateNow = new Button();
		lbVersion = new Label();
		label7 = new Label();
		label3 = new Label();
		lblVersionNew = new Label();
		lbTimeCapNhat = new Label();
		label1 = new Label();
		panel1 = new Panel();
		label8 = new Label();
		((Control)tabControl1).SuspendLayout();
		((Control)tabPage1).SuspendLayout();
		((Control)groupBox3).SuspendLayout();
		((ISupportInitialize)gridData).BeginInit();
		((Control)groupBox1).SuspendLayout();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)tabControl1).Controls.Add((Control)(object)tabPage1);
		((Control)tabControl1).Dock = (DockStyle)2;
		((Control)tabControl1).Location = new Point(0, 28);
		((Control)tabControl1).Margin = new Padding(4, 4, 4, 4);
		((Control)tabControl1).Name = "tabControl1";
		tabControl1.SelectedIndex = 0;
		((Control)tabControl1).Size = new Size(664, 394);
		((Control)tabControl1).TabIndex = 0;
		((Control)tabPage1).BackColor = SystemColors.Control;
		((Control)tabPage1).Controls.Add((Control)(object)btnExit);
		((Control)tabPage1).Controls.Add((Control)(object)groupBox3);
		((Control)tabPage1).Controls.Add((Control)(object)groupBox1);
		tabPage1.Location = new Point(4, 25);
		((Control)tabPage1).Margin = new Padding(4, 4, 4, 4);
		((Control)tabPage1).Name = "tabPage1";
		((Control)tabPage1).Padding = new Padding(4, 4, 4, 4);
		((Control)tabPage1).Size = new Size(656, 365);
		tabPage1.TabIndex = 0;
		((Control)tabPage1).Text = "Thông tin cập nhật";
		((Control)btnExit).Anchor = (AnchorStyles)14;
		((ButtonBase)btnExit).Image = (Image)(object)Resources.cancel;
		((ButtonBase)btnExit).ImageAlign = (ContentAlignment)16;
		((Control)btnExit).Location = new Point(279, 322);
		((Control)btnExit).Margin = new Padding(4, 4, 4, 4);
		((Control)btnExit).Name = "btnExit";
		((Control)btnExit).Size = new Size(107, 31);
		((Control)btnExit).TabIndex = 15;
		((Control)btnExit).Text = "&Thoát";
		((ButtonBase)btnExit).UseVisualStyleBackColor = true;
		((Control)btnExit).Click += btnThoat_Click;
		((Control)groupBox3).Controls.Add((Control)(object)gridData);
		((Control)groupBox3).Location = new Point(8, 106);
		((Control)groupBox3).Margin = new Padding(4, 4, 4, 4);
		((Control)groupBox3).Name = "groupBox3";
		((Control)groupBox3).Padding = new Padding(4, 4, 4, 4);
		((Control)groupBox3).Size = new Size(635, 203);
		((Control)groupBox3).TabIndex = 14;
		groupBox3.TabStop = false;
		((Control)groupBox3).Text = "Lịch sử cập nhật";
		gridData.AllowUserToAddRows = false;
		gridData.AllowUserToDeleteRows = false;
		gridData.AllowUserToResizeColumns = false;
		gridData.AllowUserToResizeRows = false;
		gridData.BackgroundColor = Color.White;
		gridData.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		gridData.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[4]
		{
			(DataGridViewColumn)PathFile,
			(DataGridViewColumn)STT,
			(DataGridViewColumn)Version,
			(DataGridViewColumn)CreateDate
		});
		((Control)gridData).Dock = (DockStyle)5;
		((Control)gridData).Location = new Point(4, 19);
		((Control)gridData).Margin = new Padding(4, 4, 4, 4);
		((Control)gridData).Name = "gridData";
		gridData.ReadOnly = true;
		gridData.RowHeadersVisible = false;
		gridData.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)gridData).Size = new Size(627, 180);
		((Control)gridData).TabIndex = 0;
		gridData.CellClick += new DataGridViewCellEventHandler(gridData_CellClick);
		gridData.CellDoubleClick += new DataGridViewCellEventHandler(gridData_CellDoubleClick);
		((DataGridViewColumn)PathFile).DataPropertyName = "PathFile";
		((DataGridViewColumn)PathFile).HeaderText = "Đường dẫn file";
		((DataGridViewColumn)PathFile).Name = "PathFile";
		((DataGridViewBand)PathFile).ReadOnly = true;
		((DataGridViewBand)PathFile).Visible = false;
		((DataGridViewColumn)STT).DataPropertyName = "STT";
		val.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)STT).DefaultCellStyle = val;
		((DataGridViewColumn)STT).HeaderText = "Stt";
		((DataGridViewColumn)STT).Name = "STT";
		((DataGridViewBand)STT).ReadOnly = true;
		((DataGridViewColumn)STT).Width = 30;
		((DataGridViewColumn)Version).AutoSizeMode = (DataGridViewAutoSizeColumnMode)16;
		((DataGridViewColumn)Version).DataPropertyName = "Version";
		((DataGridViewColumn)Version).HeaderText = "Phiên bản";
		((DataGridViewColumn)Version).Name = "Version";
		((DataGridViewBand)Version).ReadOnly = true;
		((DataGridViewColumn)CreateDate).DataPropertyName = "CreateDate";
		((DataGridViewColumn)CreateDate).HeaderText = "Ngày cập nhật";
		((DataGridViewColumn)CreateDate).Name = "CreateDate";
		((DataGridViewBand)CreateDate).ReadOnly = true;
		((DataGridViewColumn)CreateDate).Width = 150;
		((Control)groupBox1).Controls.Add((Control)(object)btnUpdateNow);
		((Control)groupBox1).Controls.Add((Control)(object)lbVersion);
		((Control)groupBox1).Controls.Add((Control)(object)label7);
		((Control)groupBox1).Controls.Add((Control)(object)label3);
		((Control)groupBox1).Controls.Add((Control)(object)lblVersionNew);
		((Control)groupBox1).Controls.Add((Control)(object)lbTimeCapNhat);
		((Control)groupBox1).Controls.Add((Control)(object)label1);
		((Control)groupBox1).Location = new Point(8, 7);
		((Control)groupBox1).Margin = new Padding(4, 4, 4, 4);
		((Control)groupBox1).Name = "groupBox1";
		((Control)groupBox1).Padding = new Padding(4, 4, 4, 4);
		((Control)groupBox1).Size = new Size(635, 91);
		((Control)groupBox1).TabIndex = 14;
		groupBox1.TabStop = false;
		((Control)groupBox1).Text = "Thông tin cập nhật";
		((Control)btnUpdateNow).Location = new Point(387, 52);
		((Control)btnUpdateNow).Margin = new Padding(4, 4, 4, 4);
		((Control)btnUpdateNow).Name = "btnUpdateNow";
		((Control)btnUpdateNow).Size = new Size(232, 31);
		((Control)btnUpdateNow).TabIndex = 20;
		((Control)btnUpdateNow).Text = "&Cập nhật ngay bây giờ";
		((ButtonBase)btnUpdateNow).UseVisualStyleBackColor = true;
		((Control)btnUpdateNow).Click += btnUpdateNow_Click;
		((Control)lbVersion).AutoSize = true;
		((Control)lbVersion).Location = new Point(223, 58);
		((Control)lbVersion).Margin = new Padding(4, 0, 4, 0);
		((Control)lbVersion).Name = "lbVersion";
		((Control)lbVersion).Size = new Size(46, 17);
		((Control)lbVersion).TabIndex = 19;
		((Control)lbVersion).Text = "label2";
		((Control)label7).AutoSize = true;
		((Control)label7).Location = new Point(383, 28);
		((Control)label7).Margin = new Padding(4, 0, 4, 0);
		((Control)label7).Name = "label7";
		((Control)label7).Size = new Size(134, 17);
		((Control)label7).TabIndex = 18;
		((Control)label7).Text = "Phiên bản mới nhất:";
		((Control)label3).AutoSize = true;
		((Control)label3).Location = new Point(39, 58);
		((Control)label3).Margin = new Padding(4, 0, 4, 0);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(126, 17);
		((Control)label3).TabIndex = 18;
		((Control)label3).Text = "Phiên bản hiện tại:";
		((Control)lblVersionNew).AutoSize = true;
		((Control)lblVersionNew).Location = new Point(516, 28);
		((Control)lblVersionNew).Margin = new Padding(4, 0, 4, 0);
		((Control)lblVersionNew).Name = "lblVersionNew";
		((Control)lblVersionNew).Size = new Size(102, 17);
		((Control)lblVersionNew).TabIndex = 17;
		((Control)lblVersionNew).Text = "Không tìm thấy";
		((Control)lbTimeCapNhat).AutoSize = true;
		((Control)lbTimeCapNhat).Location = new Point(223, 28);
		((Control)lbTimeCapNhat).Margin = new Padding(4, 0, 4, 0);
		((Control)lbTimeCapNhat).Name = "lbTimeCapNhat";
		((Control)lbTimeCapNhat).Size = new Size(46, 17);
		((Control)lbTimeCapNhat).TabIndex = 17;
		((Control)lbTimeCapNhat).Text = "label2";
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(39, 28);
		((Control)label1).Margin = new Padding(4, 0, 4, 0);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(155, 17);
		((Control)label1).TabIndex = 16;
		((Control)label1).Text = "Lần cập nhật gần nhất:";
		((Control)panel1).BackgroundImage = (Image)componentResourceManager.GetObject("panel1.BackgroundImage");
		((Control)panel1).BackgroundImageLayout = (ImageLayout)3;
		((Control)panel1).Controls.Add((Control)(object)label8);
		((Control)panel1).Dock = (DockStyle)1;
		((Control)panel1).Location = new Point(0, 0);
		((Control)panel1).Margin = new Padding(4, 4, 4, 4);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(664, 26);
		((Control)panel1).TabIndex = 26;
		((Control)label8).AutoSize = true;
		((Control)label8).BackColor = Color.Transparent;
		((Control)label8).Font = new Font("Tahoma", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label8).ForeColor = Color.White;
		((Control)label8).Location = new Point(5, 2);
		((Control)label8).Margin = new Padding(4, 0, 4, 0);
		((Control)label8).Name = "label8";
		((Control)label8).Size = new Size(164, 21);
		((Control)label8).TabIndex = 0;
		((Control)label8).Text = "Cấu hình cập nhật";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(664, 422);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)tabControl1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).Margin = new Padding(4, 4, 4, 4);
		((Control)this).Name = "frmConfigUpdate";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Cấu hình cập nhật phiên bản";
		((Form)this).FormClosed += new FormClosedEventHandler(frmConfigUpdate_FormClosed);
		((Form)this).Load += frmConfigUpdate_Load;
		((Control)tabControl1).ResumeLayout(false);
		((Control)tabPage1).ResumeLayout(false);
		((Control)groupBox3).ResumeLayout(false);
		((ISupportInitialize)gridData).EndInit();
		((Control)groupBox1).ResumeLayout(false);
		((Control)groupBox1).PerformLayout();
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
