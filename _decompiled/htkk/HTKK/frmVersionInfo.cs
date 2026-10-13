using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using HTKK.Properties;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

public class frmVersionInfo : Form
{
	private IContainer components = null;

	private RichTextBox txtVersionInfo;

	private Button btnDongY;

	private Panel panel1;

	public frmVersionInfo(string version, string pathFileContent)
	{
		InitializeComponent();
		((Control)this).Text = ((Control)this).Text + version;
		LoadVersionInfo(pathFileContent);
	}

	private void LoadVersionInfo(string pathFile)
	{
		try
		{
			StreamReader streamReader = new StreamReader(pathFile);
			((TextBoxBase)txtVersionInfo).AppendText(streamReader.ReadToEnd());
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnDongY_Click(object sender, EventArgs e)
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
		//IL_0084: Unknown result type (might be due to invalid IL or missing references)
		//IL_008e: Expected O, but got Unknown
		//IL_027c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0286: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmVersionInfo));
		txtVersionInfo = new RichTextBox();
		btnDongY = new Button();
		panel1 = new Panel();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)txtVersionInfo).Anchor = (AnchorStyles)15;
		((Control)txtVersionInfo).BackColor = Color.White;
		((TextBoxBase)txtVersionInfo).BorderStyle = (BorderStyle)0;
		((Control)txtVersionInfo).Font = new Font("Microsoft Sans Serif", 10f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtVersionInfo).Location = new Point(8, 3);
		((Control)txtVersionInfo).Name = "txtVersionInfo";
		((TextBoxBase)txtVersionInfo).ReadOnly = true;
		((Control)txtVersionInfo).Size = new Size(667, 397);
		((Control)txtVersionInfo).TabIndex = 0;
		((Control)txtVersionInfo).Text = "";
		((ButtonBase)btnDongY).Image = (Image)(object)Resources.cancel;
		((ButtonBase)btnDongY).ImageAlign = (ContentAlignment)16;
		((Control)btnDongY).Location = new Point(305, 413);
		((Control)btnDongY).Name = "btnDongY";
		((Control)btnDongY).Size = new Size(75, 23);
		((Control)btnDongY).TabIndex = 1;
		((Control)btnDongY).Text = "Thoát";
		((ButtonBase)btnDongY).UseVisualStyleBackColor = true;
		((Control)btnDongY).Click += btnDongY_Click;
		((Control)panel1).BackColor = Color.White;
		((Control)panel1).Controls.Add((Control)(object)txtVersionInfo);
		((Control)panel1).Location = new Point(-1, 1);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(678, 406);
		((Control)panel1).TabIndex = 2;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(677, 441);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)btnDongY);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmVersionInfo";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thông tin cập nhật phiên bản : ";
		((Control)panel1).ResumeLayout(false);
		((Control)this).ResumeLayout(false);
	}
}
