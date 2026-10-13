using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Business;

namespace HTKK.Class;

public class CheckedComboBoxCellType : TextCellType
{
	private ToolStripDropDown dropDown;

	private CheckedListBox clb;

	public string[] Options { get; set; }

	public CheckedComboBoxCellType()
	{
		Options = new string[0];
	}

	public override void PaintCell(Graphics g, Rectangle r, Appearance appearance, object value, bool isSelected, bool isLocked, float zoomFactor)
	{
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Expected O, but got Unknown
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0083: Expected O, but got Unknown
		//IL_00ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Expected O, but got Unknown
		Color color = (isSelected ? SystemColors.Highlight : appearance.BackColor);
		SolidBrush val = new SolidBrush(color);
		try
		{
			g.FillRectangle((Brush)(object)val, r);
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
		string text = ((value != null) ? value.ToString() : "");
		Rectangle rectangle = new Rectangle(r.Left + 2, r.Top + 2, r.Width - 20, r.Height - 4);
		StringFormat val2 = new StringFormat();
		try
		{
			val2.FormatFlags = (StringFormatFlags)0;
			val2.Trimming = (StringTrimming)0;
			val2.Alignment = (StringAlignment)0;
			val2.LineAlignment = (StringAlignment)0;
			SolidBrush val3 = new SolidBrush(isSelected ? SystemColors.HighlightText : appearance.ForeColor);
			try
			{
				g.DrawString(text, appearance.Font, (Brush)(object)val3, (RectangleF)rectangle, val2);
			}
			finally
			{
				((IDisposable)val3)?.Dispose();
			}
		}
		finally
		{
			((IDisposable)val2)?.Dispose();
		}
		int num = 16;
		Rectangle rectangle2 = new Rectangle(r.Right - num, r.Top, num, r.Height);
		ControlPaint.DrawComboButton(g, rectangle2, (ButtonState)16384);
		Point point = new Point(rectangle2.Left + rectangle2.Width / 2, rectangle2.Top + rectangle2.Height / 2);
		Point[] array = new Point[3]
		{
			new Point(point.X - 4, point.Y - 2),
			new Point(point.X + 4, point.Y - 2),
			new Point(point.X, point.Y + 3)
		};
		g.FillPolygon(Brushes.Black, array);
	}

	public void ShowMultiCheckDropdown(SheetView sheet, int row, int col, string controlType, int rowSelected, int colSelected)
	{
		//IL_0066: Unknown result type (might be due to invalid IL or missing references)
		//IL_006d: Expected O, but got Unknown
		//IL_02b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c1: Expected O, but got Unknown
		//IL_037b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0382: Expected O, but got Unknown
		//IL_03a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ab: Expected O, but got Unknown
		//IL_03ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_03fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0404: Expected O, but got Unknown
		Rectangle cellRectangle = sheet.FpSpread.GetCellRectangle(0, 0, row, col);
		CheckedListBox val = new CheckedListBox();
		val.CheckOnClick = true;
		((ListBox)val).BorderStyle = (BorderStyle)0;
		clb = val;
		if (Options != null && Options.Length > 0)
		{
			((ObjectCollection)clb.Items).AddRange((object[])Options);
		}
		string text = "";
		string text2 = sheet.Cells[rowSelected, colSelected].Text;
		if (text2 == ";")
		{
			text2 = "";
		}
		if (!string.IsNullOrEmpty(text2))
		{
			string key = controlType + "_" + col;
			string[] array2;
			if (GeneralFunctions.dicDms.ContainsKey(key))
			{
				Dictionary<string, string> source = GeneralFunctions.dicDms[key];
				string[] array = text2.Split(new string[1] { ";" }, StringSplitOptions.None);
				string text3 = "";
				array2 = array;
				foreach (string ma in array2)
				{
					if (!string.IsNullOrEmpty(ma))
					{
						text3 = text3 + source.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == ma).Value + ";";
					}
				}
				if (text3.Length > 1)
				{
					text3 = text3.Substring(0, text3.Length - 1);
				}
				text = text3;
			}
			string[] array3 = text.Split(new string[1] { ";" }, StringSplitOptions.RemoveEmptyEntries);
			array2 = array3;
			foreach (string text4 in array2)
			{
				int num = ((ObjectCollection)clb.Items).IndexOf((object)text4.Trim());
				if (num >= 0)
				{
					clb.SetItemChecked(num, true);
				}
			}
		}
		string result = text;
		clb.ItemCheck += (ItemCheckEventHandler)delegate
		{
			((Control)clb).BeginInvoke((Delegate)(System.Action)delegate
			{
				result = string.Join(";", ((IEnumerable)clb.CheckedItems).Cast<string>().ToArray());
				TextCellType textCellType = (sheet.Cells[row, col].CellType as TextCellType) ?? new TextCellType();
				textCellType.Multiline = true;
				textCellType.WordWrap = true;
				textCellType.MaxLength = 1000;
				sheet.Cells[row, col].CellType = textCellType;
				sheet.Cells[row, col].Value = result;
				AutoResizeRow(sheet, row, col);
			});
		};
		int num2 = 0;
		Graphics val2 = ((Control)clb).CreateGraphics();
		try
		{
			foreach (object item in (ObjectCollection)clb.Items)
			{
				num2 = Math.Max(num2, (int)val2.MeasureString(item.ToString(), ((Control)clb).Font).Width + 70);
			}
		}
		finally
		{
			((IDisposable)val2)?.Dispose();
		}
		((Control)clb).Width = num2;
		ToolStripControlHost val3 = new ToolStripControlHost((Control)(object)clb);
		((ToolStripItem)val3).AutoSize = false;
		((ToolStripItem)val3).Size = new Size(num2, 150);
		ToolStripControlHost val4 = val3;
		ToolStripDropDown val5 = new ToolStripDropDown();
		((Control)val5).Padding = Padding.Empty;
		val5.AutoClose = true;
		dropDown = val5;
		((ToolStrip)dropDown).Items.Clear();
		((ToolStrip)dropDown).Items.Add((ToolStripItem)(object)val4);
		dropDown.Closing += (ToolStripDropDownClosingEventHandler)delegate
		{
			TextCellType textCellType = (sheet.Cells[row, col].CellType as TextCellType) ?? new TextCellType();
			textCellType.Multiline = true;
			textCellType.WordWrap = true;
			textCellType.MaxLength = 1000;
			sheet.Cells[row, col].CellType = textCellType;
			sheet.Cells[row, col].Value = result;
			AutoResizeRow(sheet, row, col);
			sheet.FpSpread.StopCellEditing();
			string key2 = controlType + "_" + col;
			if (GeneralFunctions.dicDms.ContainsKey(key2))
			{
				Dictionary<string, string> source2 = GeneralFunctions.dicDms[key2];
				string[] array4 = result.Split(new string[1] { ";" }, StringSplitOptions.None);
				string text5 = "";
				string[] array5 = array4;
				foreach (string ten in array5)
				{
					if (!string.IsNullOrEmpty(ten))
					{
						text5 = text5 + source2.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == ten).Key + ";";
					}
				}
				if (text5.Length > 1)
				{
					text5 = text5.Substring(0, text5.Length - 1);
				}
				sheet.Cells[rowSelected, colSelected].Value = text5;
				WinDoWorkEvent.CellChange(sheet.FpSpread, sheet.FpSpread.ActiveSheetIndex, row + 1, col + 1);
			}
		};
		Point point = ((Control)sheet.FpSpread).PointToScreen(new Point(cellRectangle.Left, cellRectangle.Bottom));
		dropDown.Show(point);
	}

	private void AutoResizeRow(SheetView sheet, int row, int col)
	{
		string text = sheet.Cells[row, col].Text;
		if (string.IsNullOrEmpty(text))
		{
			return;
		}
		Font defaultFont = SystemFonts.DefaultFont;
		int num = sheet.FpSpread.GetCellRectangle(0, 0, row, col).Width - 6;
		Graphics val = ((Control)sheet.FpSpread).CreateGraphics();
		try
		{
			SizeF sizeF = val.MeasureString(text, defaultFont, num);
			float num2 = defaultFont.Height + 2;
			int num3 = (int)Math.Ceiling(sizeF.Height / num2);
			sheet.Rows[row].Height = Math.Max(17f, (float)num3 * num2);
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
	}
}
