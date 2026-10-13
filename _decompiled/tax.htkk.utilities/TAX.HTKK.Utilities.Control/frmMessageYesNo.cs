using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Utilities.Control;

public class frmMessageYesNo : Form
{
	public string ID;

	public int Error;

	public bool Yes = false;

	private IContainer components = null;

	private PictureBox picInfo;

	private Label lblHeader;

	private Button btnCo;

	private Label lblThongBao;

	private Button btnKhong;

	private Panel panel1;

	public DialogResult Result { get; set; }

	public frmMessageYesNo(string id, int error)
	{
		InitializeComponent();
		ID = id;
		Error = error;
		Message();
	}

	public void Message()
	{
		try
		{
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			if (Error == 0)
			{
				((Control)lblHeader).Text = "Cảnh báo";
			}
			else if (Error == 1)
			{
				((Control)lblHeader).Text = "Thông báo";
			}
			((Control)lblThongBao).Text = xMLMessage.GetMessage(ID);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnKhong_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void btnCo_Click(object sender, EventArgs e)
	{
		Yes = true;
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
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Expected O, but got Unknown
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Expected O, but got Unknown
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Expected O, but got Unknown
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_00e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f1: Expected O, but got Unknown
		//IL_01fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0206: Expected O, but got Unknown
		picInfo = new PictureBox();
		lblHeader = new Label();
		btnCo = new Button();
		lblThongBao = new Label();
		btnKhong = new Button();
		panel1 = new Panel();
		((ISupportInitialize)picInfo).BeginInit();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)picInfo).Location = new Point(6, 28);
		((Control)picInfo).Name = "picInfo";
		((Control)picInfo).Size = new Size(43, 39);
		picInfo.SizeMode = (PictureBoxSizeMode)1;
		picInfo.TabIndex = 0;
		picInfo.TabStop = false;
		((Control)lblHeader).BackColor = SystemColors.GradientActiveCaption;
		((Control)lblHeader).Font = new Font("Times New Roman", 11.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblHeader).Location = new Point(-3, 0);
		((Control)lblHeader).Name = "lblHeader";
		((Control)lblHeader).Size = new Size(333, 23);
		((Control)lblHeader).TabIndex = 1;
		((Control)lblHeader).Text = "Thông báo";
		lblHeader.TextAlign = (ContentAlignment)16;
		((ButtonBase)btnCo).FlatStyle = (FlatStyle)1;
		((Control)btnCo).Location = new Point(155, 107);
		((Control)btnCo).Name = "btnCo";
		((Control)btnCo).Size = new Size(75, 23);
		((Control)btnCo).TabIndex = 2;
		((Control)btnCo).Text = "&Có";
		((ButtonBase)btnCo).UseVisualStyleBackColor = true;
		((Control)btnCo).Click += btnCo_Click;
		((Control)lblThongBao).Font = new Font("Times New Roman", 11.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblThongBao).Location = new Point(63, 32);
		((Control)lblThongBao).Name = "lblThongBao";
		((Control)lblThongBao).Size = new Size(259, 72);
		((Control)lblThongBao).TabIndex = 3;
		((Control)lblThongBao).Text = "Infor here";
		((ButtonBase)btnKhong).FlatStyle = (FlatStyle)1;
		((Control)btnKhong).Location = new Point(248, 108);
		((Control)btnKhong).Name = "btnKhong";
		((Control)btnKhong).Size = new Size(75, 23);
		((Control)btnKhong).TabIndex = 4;
		((Control)btnKhong).Text = "&Không";
		((ButtonBase)btnKhong).UseVisualStyleBackColor = true;
		((Control)btnKhong).Click += btnKhong_Click;
		panel1.BorderStyle = (BorderStyle)1;
		((Control)panel1).Controls.Add((Control)(object)lblHeader);
		((Control)panel1).Controls.Add((Control)(object)btnKhong);
		((Control)panel1).Controls.Add((Control)(object)lblThongBao);
		((Control)panel1).Controls.Add((Control)(object)btnCo);
		((Control)panel1).Controls.Add((Control)(object)picInfo);
		((Control)panel1).Location = new Point(3, 0);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(330, 142);
		((Control)panel1).TabIndex = 5;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(333, 144);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Control)this).Name = "frmMessageYesNo";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thông Báo";
		((Form)this).TopMost = true;
		((ISupportInitialize)picInfo).EndInit();
		((Control)panel1).ResumeLayout(false);
		((Control)this).ResumeLayout(false);
	}
}
