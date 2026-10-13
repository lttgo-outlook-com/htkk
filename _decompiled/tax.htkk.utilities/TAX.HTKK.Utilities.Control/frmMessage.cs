using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Utilities.Control;

public class frmMessage : Form
{
	private string ID;

	private string MsgStyle;

	private string iCon;

	private string title;

	public bool Yes = false;

	public bool Cancel = false;

	private IContainer components = null;

	private PictureBox picInfo;

	private Label lblHeader;

	private Button btn_3;

	private Label lblThongBao;

	private Panel panel1;

	private Button btn_1;

	private Button btn_2;

	public frmMessage(string pMsgStyle, string pTitle, string outMessage)
	{
		InitializeComponent();
		MsgStyle = pMsgStyle;
		title = pTitle;
		iCon = "1";
		Message(outMessage);
	}

	public frmMessage(string pMsgID, string outMessage, string pMsgStyle, string pIcon, string pTitle)
	{
		InitializeComponent();
		ID = pMsgID;
		MsgStyle = pMsgStyle;
		iCon = pIcon;
		title = pTitle;
		Message(outMessage);
	}

	public frmMessage(string pMsgID, string pMsgStyle, string pIcon, string pTitle)
	{
		InitializeComponent();
		ID = pMsgID;
		MsgStyle = pMsgStyle;
		iCon = pIcon;
		title = pTitle;
		Message("");
	}

	public void Message(string outMessage)
	{
		try
		{
			string text = "";
			text = (iCon.Equals("2") ? "..\\Pictures\\question.ico" : (iCon.Equals("3") ? "..\\Pictures\\exclamation.ico" : ((!iCon.Equals("4")) ? "..\\Pictures\\information.ico" : "..\\Pictures\\critical.ico")));
			if (!string.IsNullOrEmpty(outMessage))
			{
				((Control)lblThongBao).Text = outMessage;
			}
			else
			{
				XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
				((Control)lblThongBao).Text = xMLMessage.GetMessage(ID);
			}
			((Control)lblHeader).Text = title;
			if (!HTKKFormater.isNullString(text))
			{
				try
				{
					FileStream fileStream = new FileStream(text, FileMode.Open);
					picInfo.Image = Image.FromStream((Stream)fileStream);
					fileStream.Close();
				}
				catch
				{
				}
			}
			int num = (((Control)this).Width - 150 - 30) / 2;
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
				((Control)btn_3).Location = new Point(((Control)this).Width - num - 75, ((Control)btn_2).Location.Y);
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
				((Control)btn_3).Location = new Point(((Control)this).Width - num - 75, ((Control)btn_2).Location.Y);
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

	public void setLabelButton()
	{
		((Control)btn_2).Text = "&Kết thúc";
		((Control)btn_3).Text = "&Tiếp tục";
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
		//IL_0103: Unknown result type (might be due to invalid IL or missing references)
		//IL_010d: Expected O, but got Unknown
		//IL_0196: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a0: Expected O, but got Unknown
		//IL_0220: Unknown result type (might be due to invalid IL or missing references)
		//IL_022a: Expected O, but got Unknown
		//IL_038a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0394: Expected O, but got Unknown
		//IL_041e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0428: Expected O, but got Unknown
		//IL_04ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f9: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmMessage));
		picInfo = new PictureBox();
		lblHeader = new Label();
		btn_3 = new Button();
		lblThongBao = new Label();
		panel1 = new Panel();
		btn_1 = new Button();
		btn_2 = new Button();
		((ISupportInitialize)picInfo).BeginInit();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)picInfo).Location = new Point(10, 28);
		((Control)picInfo).Name = "picInfo";
		((Control)picInfo).Size = new Size(43, 39);
		picInfo.SizeMode = (PictureBoxSizeMode)1;
		picInfo.TabIndex = 0;
		picInfo.TabStop = false;
		((Control)lblHeader).BackColor = SystemColors.GradientActiveCaption;
		((Control)lblHeader).Font = new Font("Tahoma", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblHeader).Location = new Point(1, 0);
		((Control)lblHeader).Name = "lblHeader";
		((Control)lblHeader).Size = new Size(362, 23);
		((Control)lblHeader).TabIndex = 1;
		((Control)lblHeader).Text = "Thông báo";
		lblHeader.TextAlign = (ContentAlignment)16;
		((ButtonBase)btn_3).FlatStyle = (FlatStyle)1;
		((Control)btn_3).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)btn_3).Location = new Point(237, 154);
		((Control)btn_3).Name = "btn_3";
		((Control)btn_3).Size = new Size(80, 25);
		((Control)btn_3).TabIndex = 2;
		((Control)btn_3).Text = "&Hủy bỏ";
		((ButtonBase)btn_3).UseVisualStyleBackColor = true;
		((Control)lblThongBao).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblThongBao).Location = new Point(58, 30);
		((Control)lblThongBao).Name = "lblThongBao";
		((Control)lblThongBao).Size = new Size(296, 121);
		((Control)lblThongBao).TabIndex = 3;
		((Control)lblThongBao).Text = "Infor here";
		panel1.BorderStyle = (BorderStyle)1;
		((Control)panel1).Controls.Add((Control)(object)btn_1);
		((Control)panel1).Controls.Add((Control)(object)btn_2);
		((Control)panel1).Controls.Add((Control)(object)lblHeader);
		((Control)panel1).Controls.Add((Control)(object)btn_3);
		((Control)panel1).Controls.Add((Control)(object)lblThongBao);
		((Control)panel1).Controls.Add((Control)(object)picInfo);
		((Control)panel1).Location = new Point(1, 0);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(364, 186);
		((Control)panel1).TabIndex = 4;
		((ButtonBase)btn_1).FlatStyle = (FlatStyle)1;
		((Control)btn_1).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)btn_1).Location = new Point(46, 154);
		((Control)btn_1).Name = "btn_1";
		((Control)btn_1).Size = new Size(80, 25);
		((Control)btn_1).TabIndex = 5;
		((Control)btn_1).Text = "&Có";
		((ButtonBase)btn_1).UseVisualStyleBackColor = true;
		((ButtonBase)btn_2).FlatStyle = (FlatStyle)1;
		((Control)btn_2).Font = new Font("Tahoma", 8.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)btn_2).Location = new Point(142, 154);
		((Control)btn_2).Name = "btn_2";
		((Control)btn_2).Size = new Size(80, 25);
		((Control)btn_2).TabIndex = 4;
		((Control)btn_2).Text = "&Không";
		((ButtonBase)btn_2).UseVisualStyleBackColor = true;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(368, 188);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Control)this).Name = "frmMessage";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thông Báo";
		((Form)this).TopMost = true;
		((ISupportInitialize)picInfo).EndInit();
		((Control)panel1).ResumeLayout(false);
		((Control)this).ResumeLayout(false);
	}
}
