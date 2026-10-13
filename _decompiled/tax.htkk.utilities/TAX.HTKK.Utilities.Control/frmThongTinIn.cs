using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Control;

public class frmThongTinIn : Form
{
	private int options;

	private string sStart;

	private string sEnd;

	private IContainer components = null;

	private RadioButton rdCheckAll;

	private RadioButton rdNonCheckAll;

	private RadioButton rdCheckNumber;

	private Label lblFrom;

	private TextBox txtFrom;

	private Label lblTo;

	private TextBox txtTo;

	private Button btnOk;

	private Button btnClose;

	public frmThongTinIn()
	{
		InitializeComponent();
	}

	public int GetOptions()
	{
		return options;
	}

	public string GetStart()
	{
		return sStart;
	}

	public string GetEnd()
	{
		return sEnd;
	}

	private void btnOk_Click(object sender, EventArgs e)
	{
		try
		{
			if (rdCheckNumber.Checked)
			{
				options = 3;
				sStart = ((Control)txtFrom).Text.Trim();
				sEnd = ((Control)txtTo).Text.Trim();
			}
			else if (rdCheckAll.Checked)
			{
				options = 1;
			}
			else if (rdNonCheckAll.Checked)
			{
				options = 2;
			}
			((Form)this).Close();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnClose_Click(object sender, EventArgs e)
	{
		((Component)this).Dispose();
		((Form)this).Close();
	}

	private void rdCheckNumber_CheckedChanged(object sender, EventArgs e)
	{
		((Control)txtFrom).Enabled = true;
		((Control)txtTo).Enabled = true;
		((Control)txtFrom).Focus();
	}

	private void txtFrom_KeyPress(object sender, KeyPressEventArgs e)
	{
		try
		{
			e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != '\b';
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtTo_KeyPress(object sender, KeyPressEventArgs e)
	{
		try
		{
			e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != '\b';
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
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_02cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d5: Expected O, but got Unknown
		//IL_03a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ae: Expected O, but got Unknown
		rdCheckAll = new RadioButton();
		rdNonCheckAll = new RadioButton();
		rdCheckNumber = new RadioButton();
		lblFrom = new Label();
		txtFrom = new TextBox();
		lblTo = new Label();
		txtTo = new TextBox();
		btnOk = new Button();
		btnClose = new Button();
		((Control)this).SuspendLayout();
		((Control)rdCheckAll).AutoSize = true;
		((Control)rdCheckAll).Location = new Point(22, 9);
		((Control)rdCheckAll).Name = "rdCheckAll";
		((Control)rdCheckAll).Size = new Size(80, 17);
		((Control)rdCheckAll).TabIndex = 0;
		rdCheckAll.TabStop = true;
		((Control)rdCheckAll).Text = "Chọn tất cả";
		((ButtonBase)rdCheckAll).UseVisualStyleBackColor = true;
		((Control)rdNonCheckAll).AutoSize = true;
		((Control)rdNonCheckAll).Location = new Point(22, 31);
		((Control)rdNonCheckAll).Name = "rdNonCheckAll";
		((Control)rdNonCheckAll).Size = new Size(95, 17);
		((Control)rdNonCheckAll).TabIndex = 1;
		rdNonCheckAll.TabStop = true;
		((Control)rdNonCheckAll).Text = "Bỏ chọn tất cả";
		((ButtonBase)rdNonCheckAll).UseVisualStyleBackColor = true;
		((Control)rdCheckNumber).AutoSize = true;
		((Control)rdCheckNumber).Location = new Point(22, 52);
		((Control)rdCheckNumber).Name = "rdCheckNumber";
		((Control)rdCheckNumber).Size = new Size(118, 17);
		((Control)rdCheckNumber).TabIndex = 2;
		rdCheckNumber.TabStop = true;
		((Control)rdCheckNumber).Text = "Chọn theo số thứ tự";
		((ButtonBase)rdCheckNumber).UseVisualStyleBackColor = true;
		rdCheckNumber.CheckedChanged += rdCheckNumber_CheckedChanged;
		((Control)lblFrom).AutoSize = true;
		((Control)lblFrom).Location = new Point(22, 78);
		((Control)lblFrom).Name = "lblFrom";
		((Control)lblFrom).Size = new Size(20, 13);
		((Control)lblFrom).TabIndex = 3;
		((Control)lblFrom).Text = "Từ";
		((Control)txtFrom).Enabled = false;
		((Control)txtFrom).Location = new Point(48, 75);
		((Control)txtFrom).Name = "txtFrom";
		((Control)txtFrom).Size = new Size(60, 20);
		((Control)txtFrom).TabIndex = 4;
		((Control)txtFrom).KeyPress += new KeyPressEventHandler(txtFrom_KeyPress);
		((Control)lblTo).AutoSize = true;
		((Control)lblTo).Location = new Point(136, 78);
		((Control)lblTo).Name = "lblTo";
		((Control)lblTo).Size = new Size(27, 13);
		((Control)lblTo).TabIndex = 5;
		((Control)lblTo).Text = "Đến";
		((Control)txtTo).Enabled = false;
		((Control)txtTo).Location = new Point(169, 75);
		((Control)txtTo).Name = "txtTo";
		((Control)txtTo).Size = new Size(60, 20);
		((Control)txtTo).TabIndex = 6;
		((Control)txtTo).KeyPress += new KeyPressEventHandler(txtTo_KeyPress);
		((Control)btnOk).Location = new Point(49, 103);
		((Control)btnOk).Name = "btnOk";
		((Control)btnOk).Size = new Size(83, 25);
		((Control)btnOk).TabIndex = 7;
		((Control)btnOk).Text = "Đồn&g ý";
		((ButtonBase)btnOk).UseVisualStyleBackColor = true;
		((Control)btnOk).Click += btnOk_Click;
		((Control)btnClose).Location = new Point(138, 103);
		((Control)btnClose).Name = "btnClose";
		((Control)btnClose).Size = new Size(80, 25);
		((Control)btnClose).TabIndex = 8;
		((Control)btnClose).Text = "Đó&ng";
		((ButtonBase)btnClose).UseVisualStyleBackColor = true;
		((Control)btnClose).Click += btnClose_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(256, 135);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)btnClose);
		((Control)this).Controls.Add((Control)(object)btnOk);
		((Control)this).Controls.Add((Control)(object)txtTo);
		((Control)this).Controls.Add((Control)(object)lblTo);
		((Control)this).Controls.Add((Control)(object)txtFrom);
		((Control)this).Controls.Add((Control)(object)lblFrom);
		((Control)this).Controls.Add((Control)(object)rdCheckNumber);
		((Control)this).Controls.Add((Control)(object)rdNonCheckAll);
		((Control)this).Controls.Add((Control)(object)rdCheckAll);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Control)this).Name = "frmThongTinIn";
		((Control)this).Text = "Chọn";
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
