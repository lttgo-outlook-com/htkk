using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace HTKK;

public class UcNgayCD : UserControl
{
	private IContainer components = null;

	private MaskedTextBox txtBegin;

	private Label lblTuNgay;

	private Label lblEnd;

	private MaskedTextBox txtEnd;

	private MaskedTextBox txtLanXuatBan;

	private Label lblLanXuat;

	public UcNgayCD()
	{
		InitializeComponent();
		((Control)txtBegin).Text = "01" + DateTime.Now.ToString("MMyyyy");
		((Control)txtEnd).Text = DateTime.Now.ToString("ddMMyyyy");
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((ContainerControl)this).Dispose(disposing);
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
		txtBegin = new MaskedTextBox();
		lblTuNgay = new Label();
		lblEnd = new Label();
		txtEnd = new MaskedTextBox();
		txtLanXuatBan = new MaskedTextBox();
		lblLanXuat = new Label();
		((Control)this).SuspendLayout();
		((Control)txtBegin).Location = new Point(60, 6);
		txtBegin.Mask = "00/00/0000";
		((Control)txtBegin).Name = "txtBegin";
		((Control)txtBegin).Size = new Size(70, 20);
		((Control)txtBegin).TabIndex = 56;
		txtBegin.TextAlign = (HorizontalAlignment)2;
		txtBegin.ValidatingType = typeof(DateTime);
		((Control)lblTuNgay).AutoSize = true;
		((Control)lblTuNgay).Location = new Point(9, 9);
		((Control)lblTuNgay).Name = "lblTuNgay";
		((Control)lblTuNgay).Size = new Size(46, 13);
		((Control)lblTuNgay).TabIndex = 55;
		((Control)lblTuNgay).Text = "Từ ngày";
		((Control)lblEnd).AutoSize = true;
		((Control)lblEnd).Location = new Point(138, 9);
		((Control)lblEnd).Name = "lblEnd";
		((Control)lblEnd).Size = new Size(53, 13);
		((Control)lblEnd).TabIndex = 55;
		((Control)lblEnd).Text = "Đến ngày";
		((Control)txtEnd).Location = new Point(194, 6);
		txtEnd.Mask = "00/00/0000";
		((Control)txtEnd).Name = "txtEnd";
		((Control)txtEnd).Size = new Size(70, 20);
		((Control)txtEnd).TabIndex = 56;
		txtEnd.TextAlign = (HorizontalAlignment)2;
		txtEnd.ValidatingType = typeof(DateTime);
		((Control)txtLanXuatBan).Location = new Point(343, 6);
		txtLanXuatBan.Mask = "00";
		((Control)txtLanXuatBan).Name = "txtLanXuatBan";
		txtLanXuatBan.PromptChar = ' ';
		((Control)txtLanXuatBan).Size = new Size(80, 20);
		((Control)txtLanXuatBan).TabIndex = 58;
		txtLanXuatBan.TextAlign = (HorizontalAlignment)1;
		((Control)lblLanXuat).AutoSize = true;
		((Control)lblLanXuat).Location = new Point(270, 9);
		((Control)lblLanXuat).Name = "lblLanXuat";
		((Control)lblLanXuat).Size = new Size(69, 13);
		((Control)lblLanXuat).TabIndex = 57;
		((Control)lblLanXuat).Text = "Lần xuất bán";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).Controls.Add((Control)(object)txtLanXuatBan);
		((Control)this).Controls.Add((Control)(object)lblLanXuat);
		((Control)this).Controls.Add((Control)(object)txtEnd);
		((Control)this).Controls.Add((Control)(object)lblEnd);
		((Control)this).Controls.Add((Control)(object)txtBegin);
		((Control)this).Controls.Add((Control)(object)lblTuNgay);
		((Control)this).Name = "UcNgayCD";
		((Control)this).Size = new Size(432, 32);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
