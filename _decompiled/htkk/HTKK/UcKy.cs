using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace HTKK;

public class UcKy : UserControl
{
	private IContainer components = null;

	private MaskedTextBox txtEnd;

	private Label label2;

	private MaskedTextBox txtBegin;

	private Label label1;

	public UcKy()
	{
		InitializeComponent();
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
		txtEnd = new MaskedTextBox();
		label2 = new Label();
		txtBegin = new MaskedTextBox();
		label1 = new Label();
		((Control)this).SuspendLayout();
		((Control)txtEnd).Location = new Point(163, 6);
		txtEnd.Mask = "00/0000";
		((Control)txtEnd).Name = "txtEnd";
		((Control)txtEnd).Size = new Size(50, 20);
		((Control)txtEnd).TabIndex = 58;
		txtEnd.TextAlign = (HorizontalAlignment)2;
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(121, 9);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(41, 13);
		((Control)label2).TabIndex = 56;
		((Control)label2).Text = "Đến kỳ";
		((Control)txtBegin).Location = new Point(45, 6);
		txtBegin.Mask = "00/0000";
		((Control)txtBegin).Name = "txtBegin";
		((Control)txtBegin).Size = new Size(50, 20);
		((Control)txtBegin).TabIndex = 59;
		txtBegin.TextAlign = (HorizontalAlignment)2;
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(9, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(34, 13);
		((Control)label1).TabIndex = 57;
		((Control)label1).Text = "Từ kỳ";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).Controls.Add((Control)(object)txtEnd);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)txtBegin);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Name = "UcKy";
		((Control)this).Size = new Size(229, 33);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
