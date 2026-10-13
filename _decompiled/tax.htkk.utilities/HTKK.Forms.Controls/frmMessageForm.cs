using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK.Forms.Controls;

public class frmMessageForm : Form
{
	private string MsgStyle;

	private string iCon;

	private string title;

	public bool Yes = false;

	public bool Cancel = false;

	private IContainer components = null;

	private TextBox txtMessge;

	private Button btn_1;

	private Button btn_2;

	private Button btn_3;

	private Label label1;

	public frmMessageForm()
	{
		InitializeComponent();
	}

	public frmMessageForm(string title, string message)
	{
		InitializeComponent();
		((Control)this).Text = title;
		((Control)txtMessge).Text = message;
	}

	public frmMessageForm(string pMsgID, string pMsgStyle, string pTitle, string outMessage)
	{
		InitializeComponent();
		MsgStyle = pMsgStyle;
		title = pTitle;
		iCon = "1";
		XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
		((Control)label1).Text = xMLMessage.GetMessage(pMsgID);
		Message(outMessage);
	}

	private void btnClose_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	public void Message(string outMessage)
	{
		try
		{
			((Control)txtMessge).Text = outMessage;
			int num = ((Control)this).Width - 150 - 30;
			switch (MsgStyle)
			{
			case "2":
				((Control)btn_1).Visible = false;
				((Control)btn_2).Visible = true;
				((Control)btn_2).Text = "Đồn&g ý";
				((Control)btn_3).Visible = true;
				((Control)btn_3).Text = "&Hủy bỏ";
				((Control)btn_2).Click += cmd_OK;
				((Control)btn_3).Click += cmd_Cancel;
				((Control)btn_2).Location = new Point(num, ((Control)btn_2).Location.Y);
				((Control)btn_3).Location = new Point(num + 75 + 15, ((Control)btn_2).Location.Y);
				break;
			case "3":
				((Control)btn_1).Visible = false;
				((Control)btn_3).Visible = false;
				((Control)btn_2).Visible = true;
				((Control)btn_2).Text = "Đó&ng";
				((Control)btn_2).Click += cmd_OK;
				break;
			case "5":
				((Control)btn_1).Visible = false;
				((Control)btn_2).Visible = true;
				((Control)btn_2).Text = "&Có";
				((Control)btn_3).Visible = true;
				((Control)btn_3).Text = "&Không";
				((Control)btn_2).Click += cmd_OK;
				((Control)btn_3).Click += cmd_No;
				((Control)btn_2).Location = new Point(num, ((Control)btn_2).Location.Y);
				((Control)btn_3).Location = new Point(num + 75 + 15, ((Control)btn_2).Location.Y);
				break;
			default:
				((Control)btn_1).Visible = true;
				((Control)btn_1).Text = "&Có";
				((Control)btn_2).Visible = true;
				((Control)btn_2).Text = "&Không";
				((Control)btn_3).Visible = true;
				((Control)btn_3).Text = "&Hủy bỏ";
				((Control)btn_1).Click += cmd_OK;
				((Control)btn_2).Click += cmd_No;
				((Control)btn_3).Click += cmd_Cancel;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cmd_OK(object sender, EventArgs e)
	{
		Yes = true;
		Cancel = false;
		((Form)this).Close();
	}

	private void cmd_No(object sender, EventArgs e)
	{
		Yes = false;
		Cancel = false;
		((Form)this).Close();
	}

	private void cmd_Cancel(object sender, EventArgs e)
	{
		Yes = false;
		Cancel = true;
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
		//IL_00e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Expected O, but got Unknown
		//IL_017b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0185: Expected O, but got Unknown
		//IL_0212: Unknown result type (might be due to invalid IL or missing references)
		//IL_021c: Expected O, but got Unknown
		//IL_02a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b3: Expected O, but got Unknown
		//IL_03b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bd: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmMessageForm));
		txtMessge = new TextBox();
		btn_1 = new Button();
		btn_2 = new Button();
		btn_3 = new Button();
		label1 = new Label();
		((Control)this).SuspendLayout();
		((Control)txtMessge).Location = new Point(12, 29);
		((TextBoxBase)txtMessge).Multiline = true;
		((Control)txtMessge).Name = "txtMessge";
		((TextBoxBase)txtMessge).ReadOnly = true;
		txtMessge.ScrollBars = (ScrollBars)3;
		((Control)txtMessge).Size = new Size(467, 335);
		((Control)txtMessge).TabIndex = 0;
		((ButtonBase)btn_1).FlatStyle = (FlatStyle)1;
		((Control)btn_1).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)btn_1).Location = new Point(208, 370);
		((Control)btn_1).Name = "btn_1";
		((Control)btn_1).Size = new Size(80, 25);
		((Control)btn_1).TabIndex = 8;
		((Control)btn_1).Text = "&Có";
		((ButtonBase)btn_1).UseVisualStyleBackColor = true;
		((ButtonBase)btn_2).FlatStyle = (FlatStyle)1;
		((Control)btn_2).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)btn_2).Location = new Point(304, 370);
		((Control)btn_2).Name = "btn_2";
		((Control)btn_2).Size = new Size(80, 25);
		((Control)btn_2).TabIndex = 7;
		((Control)btn_2).Text = "&Không";
		((ButtonBase)btn_2).UseVisualStyleBackColor = true;
		((ButtonBase)btn_3).FlatStyle = (FlatStyle)1;
		((Control)btn_3).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)btn_3).Location = new Point(399, 370);
		((Control)btn_3).Name = "btn_3";
		((Control)btn_3).Size = new Size(80, 25);
		((Control)btn_3).TabIndex = 6;
		((Control)btn_3).Text = "&Hủy bỏ";
		((ButtonBase)btn_3).UseVisualStyleBackColor = true;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 10f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(12, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(261, 17);
		((Control)label1).TabIndex = 9;
		((Control)label1).Text = "Dữ liệu kê khai sai. Có muốn ghi không?";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(491, 405);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)btn_1);
		((Control)this).Controls.Add((Control)(object)btn_2);
		((Control)this).Controls.Add((Control)(object)btn_3);
		((Control)this).Controls.Add((Control)(object)txtMessge);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Control)this).Name = "frmMessageForm";
		((Form)this).StartPosition = (FormStartPosition)4;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
