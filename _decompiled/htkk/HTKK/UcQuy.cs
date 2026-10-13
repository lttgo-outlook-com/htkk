using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

public class UcQuy : UserControl
{
	private IContainer components = null;

	private Label label2;

	private Label label1;

	private MaskedTextBox txtEnd;

	private MaskedTextBox txtBegin;

	public UcQuy()
	{
		InitializeComponent();
		Dictionary<string, string> dictionary = DictData();
		string text = dictionary[GetCurentQuy()];
		if (!string.IsNullOrEmpty(text))
		{
			((Control)txtBegin).Text = text.Split(new char[1] { '_' })[0];
			((Control)txtEnd).Text = text.Split(new char[1] { '_' })[1];
		}
	}

	private string GetCurentQuy()
	{
		try
		{
			string result = string.Empty;
			int month = DateTime.Now.Month;
			if (month > 0 && month <= 3)
			{
				result = "1";
			}
			if (month > 3 && month <= 6)
			{
				result = "2";
			}
			if (month > 6 && month <= 9)
			{
				result = "3";
			}
			if (month > 9 && month <= 12)
			{
				result = "4";
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private Dictionary<string, string> DictData()
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			dictionary.Add("1", "01" + DateTime.Now.Year + "_03" + DateTime.Now.Year);
			dictionary.Add("2", "04" + DateTime.Now.Year + "_06" + DateTime.Now.Year);
			dictionary.Add("3", "07" + DateTime.Now.Year + "_09" + DateTime.Now.Year);
			dictionary.Add("4", "10" + DateTime.Now.Year + "_12" + DateTime.Now.Year);
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
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
		((Control)label2).Location = new Point(109, 9);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(47, 13);
		((Control)label2).TabIndex = 56;
		((Control)label2).Text = "Đến quý";
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(9, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(40, 13);
		((Control)label1).TabIndex = 57;
		((Control)label1).Text = "Từ quý";
		((Control)txtEnd).Location = new Point(162, 6);
		txtEnd.Mask = "00/0000";
		((Control)txtEnd).Name = "txtEnd";
		((Control)txtEnd).Size = new Size(50, 20);
		((Control)txtEnd).TabIndex = 62;
		txtEnd.TextAlign = (HorizontalAlignment)2;
		((Control)txtBegin).Location = new Point(53, 6);
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
		((Control)this).Name = "UcQuy";
		((Control)this).Size = new Size(223, 32);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
