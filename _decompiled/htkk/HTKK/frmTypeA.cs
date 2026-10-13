using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Xml;

namespace HTKK;

public class frmTypeA : Form
{
	private static XmlNode xmlNode;

	private static frmHTKKSystems frmHTKKSystems;

	private IContainer components = null;

	private PictureBox pictureBox1;

	private Button cmdOK;

	private Button cmdClose;

	private Label lblNgayDau;

	private TextBox txtNgayDau;

	private TextBox txtNgayCuoi;

	private Label lblNgayCuoi;

	private TextBox txtYear;

	private Label lblYear;

	private TextBox txtSolan;

	private Label lblSolan;

	private RadioButton OptChinhthuc;

	private RadioButton OptBosung;

	private GroupBox groupBox1;

	public frmTypeA(frmHTKKSystems _frmHTKKSystems, XmlNode _xmlNode)
	{
		InitializeComponent();
		frmHTKKSystems = _frmHTKKSystems;
		xmlNode = _xmlNode;
	}

	private void frmPeriod_Load(object sender, EventArgs e)
	{
	}

	private void cmdOK_Click(object sender, EventArgs e)
	{
		string text = ((Control)txtYear).Text;
		string text2 = ((Control)txtNgayDau).Text;
		string text3 = ((Control)txtNgayCuoi).Text;
		string empty = string.Empty;
		if (OptChinhthuc.Checked)
		{
		}
		if (OptBosung.Checked)
		{
			empty = ((Control)txtSolan).Text;
		}
	}

	private void cmdClose_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
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
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f2: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmTypeA));
		pictureBox1 = new PictureBox();
		cmdOK = new Button();
		cmdClose = new Button();
		lblNgayDau = new Label();
		txtNgayDau = new TextBox();
		txtNgayCuoi = new TextBox();
		lblNgayCuoi = new Label();
		txtYear = new TextBox();
		lblYear = new Label();
		txtSolan = new TextBox();
		lblSolan = new Label();
		OptChinhthuc = new RadioButton();
		OptBosung = new RadioButton();
		groupBox1 = new GroupBox();
		((ISupportInitialize)pictureBox1).BeginInit();
		((Control)groupBox1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)pictureBox1).Dock = (DockStyle)1;
		pictureBox1.Image = (Image)componentResourceManager.GetObject("pictureBox1.Image");
		((Control)pictureBox1).Location = new Point(0, 0);
		((Control)pictureBox1).Name = "pictureBox1";
		((Control)pictureBox1).Size = new Size(372, 19);
		pictureBox1.TabIndex = 0;
		pictureBox1.TabStop = false;
		((Control)cmdOK).Location = new Point(97, 168);
		((Control)cmdOK).Name = "cmdOK";
		((Control)cmdOK).Size = new Size(75, 23);
		((Control)cmdOK).TabIndex = 3;
		((Control)cmdOK).Text = "Đồn&g ý";
		((ButtonBase)cmdOK).UseVisualStyleBackColor = true;
		((Control)cmdOK).Click += cmdOK_Click;
		((Control)cmdClose).Location = new Point(195, 168);
		((Control)cmdClose).Name = "cmdClose";
		((Control)cmdClose).Size = new Size(75, 23);
		((Control)cmdClose).TabIndex = 4;
		((Control)cmdClose).Text = "Đó&ng";
		((ButtonBase)cmdClose).UseVisualStyleBackColor = true;
		((Control)cmdClose).Click += cmdClose_Click;
		((Control)lblNgayDau).AutoSize = true;
		((Control)lblNgayDau).Location = new Point(2, 53);
		((Control)lblNgayDau).Name = "lblNgayDau";
		((Control)lblNgayDau).Size = new Size(50, 13);
		((Control)lblNgayDau).TabIndex = 7;
		((Control)lblNgayDau).Text = "Từ tháng";
		((Control)txtNgayDau).Location = new Point(54, 50);
		((Control)txtNgayDau).Name = "txtNgayDau";
		((Control)txtNgayDau).Size = new Size(100, 20);
		((Control)txtNgayDau).TabIndex = 8;
		((Control)txtNgayCuoi).Location = new Point(237, 50);
		((Control)txtNgayCuoi).Name = "txtNgayCuoi";
		((Control)txtNgayCuoi).Size = new Size(100, 20);
		((Control)txtNgayCuoi).TabIndex = 10;
		((Control)lblNgayCuoi).AutoSize = true;
		((Control)lblNgayCuoi).Location = new Point(175, 53);
		((Control)lblNgayCuoi).Name = "lblNgayCuoi";
		((Control)lblNgayCuoi).Size = new Size(57, 13);
		((Control)lblNgayCuoi).TabIndex = 9;
		((Control)lblNgayCuoi).Text = "Đến tháng";
		((Control)txtYear).Location = new Point(54, 19);
		((Control)txtYear).Name = "txtYear";
		((Control)txtYear).Size = new Size(100, 20);
		((Control)txtYear).TabIndex = 14;
		((Control)lblYear).AutoSize = true;
		((Control)lblYear).Location = new Point(19, 22);
		((Control)lblYear).Name = "lblYear";
		((Control)lblYear).Size = new Size(29, 13);
		((Control)lblYear).TabIndex = 16;
		((Control)lblYear).Text = "Năm";
		((Control)txtSolan).Location = new Point(237, 96);
		((Control)txtSolan).Name = "txtSolan";
		((Control)txtSolan).Size = new Size(50, 20);
		((Control)txtSolan).TabIndex = 22;
		((Control)lblSolan).AutoSize = true;
		((Control)lblSolan).Location = new Point(197, 99);
		((Control)lblSolan).Name = "lblSolan";
		((Control)lblSolan).Size = new Size(25, 13);
		((Control)lblSolan).TabIndex = 21;
		((Control)lblSolan).Text = "Lần";
		((Control)OptChinhthuc).AutoSize = true;
		OptChinhthuc.Checked = true;
		((Control)OptChinhthuc).Location = new Point(51, 76);
		((Control)OptChinhthuc).Name = "OptChinhthuc";
		((Control)OptChinhthuc).Size = new Size(100, 17);
		((Control)OptChinhthuc).TabIndex = 41;
		OptChinhthuc.TabStop = true;
		((Control)OptChinhthuc).Text = "Tờ khai lần đầu";
		((ButtonBase)OptChinhthuc).UseVisualStyleBackColor = true;
		((Control)OptBosung).AutoSize = true;
		((Control)OptBosung).Location = new Point(51, 99);
		((Control)OptBosung).Name = "OptBosung";
		((Control)OptBosung).Size = new Size(102, 17);
		((Control)OptBosung).TabIndex = 42;
		((Control)OptBosung).Text = "Tờ khai bổ sung";
		((ButtonBase)OptBosung).UseVisualStyleBackColor = true;
		((Control)groupBox1).Controls.Add((Control)(object)lblYear);
		((Control)groupBox1).Controls.Add((Control)(object)lblNgayDau);
		((Control)groupBox1).Controls.Add((Control)(object)OptBosung);
		((Control)groupBox1).Controls.Add((Control)(object)txtYear);
		((Control)groupBox1).Controls.Add((Control)(object)OptChinhthuc);
		((Control)groupBox1).Controls.Add((Control)(object)lblSolan);
		((Control)groupBox1).Controls.Add((Control)(object)txtSolan);
		((Control)groupBox1).Controls.Add((Control)(object)lblNgayCuoi);
		((Control)groupBox1).Controls.Add((Control)(object)txtNgayCuoi);
		((Control)groupBox1).Controls.Add((Control)(object)txtNgayDau);
		((Control)groupBox1).Location = new Point(11, 29);
		((Control)groupBox1).Name = "groupBox1";
		((Control)groupBox1).Size = new Size(349, 133);
		((Control)groupBox1).TabIndex = 41;
		groupBox1.TabStop = false;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(372, 207);
		((Control)this).Controls.Add((Control)(object)cmdOK);
		((Control)this).Controls.Add((Control)(object)groupBox1);
		((Control)this).Controls.Add((Control)(object)cmdClose);
		((Control)this).Controls.Add((Control)(object)pictureBox1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmTypeA";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "frmPeriod";
		((Form)this).Load += frmPeriod_Load;
		((ISupportInitialize)pictureBox1).EndInit();
		((Control)groupBox1).ResumeLayout(false);
		((Control)groupBox1).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
