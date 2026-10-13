using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace HTKK;

public class UcThang : UserControl
{
	private IContainer components = null;

	private Label label2;

	private Label label1;

	private MaskedTextBox txtEnd;

	private MaskedTextBox txtBegin;

	public UcThang()
	{
		InitializeComponent();
		((Control)txtBegin).Text = DateTime.Now.Month.ToString().PadLeft(2, '0') + "/" + (DateTime.Now.Year - 1);
		((Control)txtEnd).Text = DateTime.Now.Month.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Year;
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
		label2 = new Label();
		label1 = new Label();
		txtEnd = new MaskedTextBox();
		txtBegin = new MaskedTextBox();
		((Control)this).SuspendLayout();
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(122, 9);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(57, 13);
		((Control)label2).TabIndex = 56;
		((Control)label2).Text = "Đến tháng";
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(9, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(50, 13);
		((Control)label1).TabIndex = 57;
		((Control)label1).Text = "Từ tháng";
		((Control)txtEnd).Location = new Point(185, 6);
		txtEnd.Mask = "00/0000";
		((Control)txtEnd).Name = "txtEnd";
		((Control)txtEnd).Size = new Size(50, 20);
		((Control)txtEnd).TabIndex = 62;
		txtEnd.TextAlign = (HorizontalAlignment)2;
		((Control)txtBegin).Location = new Point(65, 6);
		txtBegin.Mask = "00/0000";
		((Control)txtBegin).Name = "txtBegin";
		((Control)txtBegin).Size = new Size(50, 20);
		((Control)txtBegin).TabIndex = 63;
		txtBegin.TextAlign = (HorizontalAlignment)2;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).Controls.Add((Control)(object)txtEnd);
		((Control)this).Controls.Add((Control)(object)txtBegin);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Name = "UcThang";
		((Control)this).Size = new Size(245, 32);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
