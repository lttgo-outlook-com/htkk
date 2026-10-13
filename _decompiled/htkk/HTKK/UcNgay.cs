using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace HTKK;

public class UcNgay : UserControl
{
	private IContainer components = null;

	private MaskedTextBox txtBegin;

	private Label lblTuNgay;

	private Label txtEnd_1;

	private MaskedTextBox txtEnd;

	public UcNgay()
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
		txtBegin = new MaskedTextBox();
		lblTuNgay = new Label();
		txtEnd_1 = new Label();
		txtEnd = new MaskedTextBox();
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
		((Control)txtEnd_1).AutoSize = true;
		((Control)txtEnd_1).Location = new Point(145, 9);
		((Control)txtEnd_1).Name = "txtEnd_1";
		((Control)txtEnd_1).Size = new Size(53, 13);
		((Control)txtEnd_1).TabIndex = 55;
		((Control)txtEnd_1).Text = "Đến ngày";
		((Control)txtEnd).Location = new Point(202, 6);
		txtEnd.Mask = "00/00/0000";
		((Control)txtEnd).Name = "txtEnd";
		((Control)txtEnd).Size = new Size(70, 20);
		((Control)txtEnd).TabIndex = 56;
		txtEnd.TextAlign = (HorizontalAlignment)2;
		txtEnd.ValidatingType = typeof(DateTime);
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).Controls.Add((Control)(object)txtEnd);
		((Control)this).Controls.Add((Control)(object)txtEnd_1);
		((Control)this).Controls.Add((Control)(object)txtBegin);
		((Control)this).Controls.Add((Control)(object)lblTuNgay);
		((Control)this).Name = "UcNgay";
		((Control)this).Size = new Size(280, 33);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
