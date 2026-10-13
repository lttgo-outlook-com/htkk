using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using FarPoint.Win.Spread.Model;
using FarPoint.Win.SuperEdit;
using HTKK.Class;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class WinSpreadProcess
{
	public static ProgressForm _form = null;

	public static void StartForm(ProgressForm form)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Invalid comparison between Unknown and I4
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Invalid comparison between Unknown and I4
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		_form = form;
		DialogResult val = ((Form)form).ShowDialog();
		if ((int)val == 2)
		{
			MessageBox.Show("Operation has been cancelled");
		}
		if ((int)val == 3)
		{
			MessageBox.Show("Exception:" + Environment.NewLine + form.Result.Error.Message);
		}
	}

	public static void WinSetDataHeader(SheetView sView, HTKKHeaderInfo headerInfo, XmlDocument xmlConfigHeader, string dataFile)
	{
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_005b: Expected O, but got Unknown
		try
		{
			XmlNodeList val = ((XmlNode)xmlConfigHeader).SelectNodes("/Sections/Section/Cells[@DataFile='" + dataFile + "']");
			if (val == null || val.Count <= 0)
			{
				return;
			}
			XmlNodeList childNodes = val[0].ChildNodes;
			foreach (XmlNode item in childNodes)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["PropertyInfo"]).Value;
				string value2 = ((XmlNode)val2.Attributes["CellID"]).Value;
				string value3 = HTKKProperty.GetPropertyValue(headerInfo, value).ToString();
				WinSetValue(sView, value2, value3);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetDataMaSoThue(SheetView sView, string maSoThue)
	{
		try
		{
			if (!string.IsNullOrEmpty(maSoThue))
			{
				char[] array = maSoThue.ToCharArray();
				string[] array2 = new string[0];
				array2 = ((array.Length != 10) ? new string[13]
				{
					"C_2", "D_2", "E_2", "F_2", "G_2", "H_2", "I_2", "J_2", "K_2", "L_2",
					"N_2", "O_2", "P_2"
				} : new string[10] { "C_2", "D_2", "E_2", "F_2", "G_2", "H_2", "I_2", "J_2", "K_2", "L_2" });
				for (int i = 0; i < array2.Length; i++)
				{
					WinSetValue(sView, array2[i], array[i].ToString());
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetValue(SheetView sView, string cellIndex, string value)
	{
		try
		{
			string[] array = cellIndex.Split(new char[1] { '_' });
			if (array.Length == 2)
			{
				int num = int.Parse(array[1]) - 1;
				int num2 = HTKKFile.ColumnLettersToIndex(array[0]) - 1;
				if (num >= 0 && num2 >= 0)
				{
					sView.Cells[num, num2].Value = value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetValueByIndex(SheetView sView, string cellIndex, string value)
	{
		try
		{
			string[] array = cellIndex.Split(new char[1] { '_' });
			if (array.Length == 2)
			{
				int num = int.Parse(array[1]) - 1;
				int num2 = int.Parse(array[0]) - 1;
				if (num >= 0 && num2 >= 0)
				{
					sView.Cells[num, num2].Value = value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FillMstHeader(SheetView sView, string value)
	{
		try
		{
			char[] array = value.ToCharArray();
			int num = array.Length;
			int row = 1;
			int num2 = 2;
			int num3 = 0;
			for (int i = 0; i < 14; i++)
			{
				sView.Cells[1, 2 + i].Value = "";
			}
			if (num == 10)
			{
				for (int i = 0; i < num; i++)
				{
					sView.Cells[row, num2].Value = array[num3];
					num2++;
					num3++;
				}
			}
			if (num != 13 && num != 12)
			{
				return;
			}
			for (int i = 0; i < num; i++)
			{
				if (i != 10)
				{
					sView.Cells[row, num2].Value = array[num3];
					num2++;
					num3++;
				}
				else
				{
					sView.Cells[row, num2 + 1].Value = array[num3];
					num2 += 2;
					num3++;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetValueToSpread(SheetView sView, HTKKSheetInfo sheetInfo, Dictionary<string, List<DictInfo>> dictData)
	{
		try
		{
			if (dictData != null && dictData.Count > 0)
			{
				if (sheetInfo.DynamicData == "0")
				{
					foreach (KeyValuePair<string, List<DictInfo>> dictDatum in dictData)
					{
						List<DictInfo> value = dictDatum.Value;
						if (value.Count <= 0)
						{
							continue;
						}
						foreach (DictInfo item in value)
						{
							if (!string.IsNullOrEmpty(item.Value))
							{
								WinSetValue(sView, item.Key, item.Value);
							}
						}
					}
				}
				else
				{
					int num = 0;
					foreach (KeyValuePair<string, List<DictInfo>> dictDatum2 in dictData)
					{
						if (num == 0)
						{
							List<DictInfo> value = dictDatum2.Value;
							if (value.Count > 0)
							{
								foreach (DictInfo item2 in value)
								{
									if (!string.IsNullOrEmpty(item2.Value))
									{
										WinSetValue(sView, item2.Key, item2.Value);
									}
								}
							}
						}
						else
						{
							sView.AddRows(int.Parse(dictDatum2.Key), 1);
							sView.CopyRange(int.Parse(dictDatum2.Key) - 1, 1, int.Parse(dictDatum2.Key), 1, 1, 100, dataOnly: false);
							List<DictInfo> value = dictDatum2.Value;
							if (value.Count > 0)
							{
								foreach (DictInfo item3 in value)
								{
									WinSetValue(sView, item3.Key, item3.Value);
								}
							}
						}
						num++;
					}
				}
			}
			sView.Rows.Count = sView.NonEmptyRowCount + 1;
			sView.Columns.Count = sView.NonEmptyColumnCount + 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetValueToSpread(SheetView sView, HTKKSheetInfo sheetInfo, SortedDictionary<int, List<DictInfo>> dictData)
	{
		try
		{
			if (dictData == null || dictData.Count <= 0)
			{
				return;
			}
			if (sheetInfo.DynamicData == "0")
			{
				foreach (KeyValuePair<int, List<DictInfo>> dictDatum in dictData)
				{
					List<DictInfo> value = dictDatum.Value;
					if (value.Count > 0)
					{
						foreach (DictInfo item in value)
						{
							if (!string.IsNullOrEmpty(item.Value))
							{
								WinSetValue(sView, item.Key, item.Value);
							}
						}
					}
				}
				return;
			}
			int num = 0;
			foreach (KeyValuePair<int, List<DictInfo>> dictDatum2 in dictData)
			{
				if (num == 0)
				{
					List<DictInfo> value = dictDatum2.Value;
					if (value.Count > 0)
					{
						foreach (DictInfo item2 in value)
						{
							if (!string.IsNullOrEmpty(item2.Value))
							{
								WinSetValue(sView, item2.Key, item2.Value);
							}
						}
					}
				}
				else
				{
					sView.AddRows(dictDatum2.Key - 1, 1);
					sView.CopyRange(dictDatum2.Key - 2, 0, dictDatum2.Key - 1, 0, 1, 100, dataOnly: false);
					List<DictInfo> value = dictDatum2.Value;
					if (value.Count > 0)
					{
						foreach (DictInfo item3 in value)
						{
							WinSetValue(sView, item3.Key, item3.Value);
						}
					}
				}
				num++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string WinGetText(SheetView sView, string cellIndex)
	{
		try
		{
			string[] array = cellIndex.Split(new char[1] { '_' });
			return sView.GetText(int.Parse(array[1]) - 1, HTKKFile.ColumnLettersToIndex(array[0]));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static void WinSetupSpread(FpSpread winFpSpread)
	{
		try
		{
			winFpSpread.SubEditorOpening += delegate(object sender, SubEditorOpeningEventArgs args)
			{
				args.Cancel = true;
			};
			winFpSpread.AllowCellOverflow = true;
			winFpSpread.AllowEditOverflow = true;
			winFpSpread.AllowColumnMove = false;
			winFpSpread.AllowRowMove = false;
			winFpSpread.AllowDragDrop = false;
			winFpSpread.AllowSheetMove = false;
			winFpSpread.EditModePermanent = true;
			winFpSpread.EditModeReplace = true;
			winFpSpread.TextTipPolicy = TextTipPolicy.Floating;
			winFpSpread.BorderStyle = (BorderStyle)0;
			winFpSpread.RowSplitBoxPolicy = SplitBoxPolicy.Never;
			winFpSpread.ColumnSplitBoxPolicy = SplitBoxPolicy.Never;
			winFpSpread.TabStrip.Editable = false;
			winFpSpread.TabStripInsertTab = false;
			winFpSpread.SetCursor(CursorType.LockedCell, Cursors.Default);
			winFpSpread.SkipMouseProcessing = true;
			winFpSpread.VerticalScrollBarSmallChange = 1;
			winFpSpread.HorizontalScrollBarSmallChange = 1;
			winFpSpread.VisualStyles = VisualStyles.Off;
			winFpSpread.InterfaceRenderer = null;
			winFpSpread.FocusRenderer = null;
			winFpSpread.TabStripPolicy = TabStripPolicy.AsNeeded;
			winFpSpread.AllowUserFormulas = false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetDataToTable(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				int num = 0;
				do
				{
					num++;
					sInfo.Tables[i].addNewDefaultRow(0);
				}
				while (sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + num, sInfo.Tables[i].iColumnSTT].Text == "aa");
				fpSpReindexingTableTemplate(sView, sInfo.Tables[i]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetupSpreadSheet(SheetView sView)
	{
		try
		{
			GridLine horizontalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
			GridLine verticalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
			sView.HorizontalGridLine = horizontalGridLine;
			sView.VerticalGridLine = verticalGridLine;
			sView.RowHeader.Visible = false;
			sView.ColumnHeader.Visible = false;
			sView.Protect = true;
			for (int i = 0; i < sView.NonEmptyColumnCount; i++)
			{
				Column column = sView.Columns[i];
				column.Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetupSpreadReportSheet(SheetView sView)
	{
		try
		{
			GridLine horizontalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
			GridLine verticalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
			sView.HorizontalGridLine = horizontalGridLine;
			sView.VerticalGridLine = verticalGridLine;
			sView.RowHeader.Visible = false;
			sView.ColumnHeader.Visible = false;
			sView.Protect = true;
			for (int i = 0; i < sView.NonEmptyColumnCount; i++)
			{
				Column column = sView.Columns[i];
				column.Locked = true;
			}
			int nonEmptyRowCount = sView.NonEmptyRowCount;
			int nonEmptyColumnCount = sView.NonEmptyColumnCount;
			DefaultSheetStyleModel defaultSheetStyleModel = (DefaultSheetStyleModel)sView.Models.Style;
			StyleInfo styleInfo = null;
			for (int i = 0; i < nonEmptyRowCount; i++)
			{
				for (int j = 0; j < nonEmptyColumnCount; j++)
				{
					styleInfo = defaultSheetStyleModel.GetCompositeInfo(i, j, -1, null);
					if (!(styleInfo.Border is ComplexBorder))
					{
						continue;
					}
					ComplexBorder complexBorder = (ComplexBorder)styleInfo.Border;
					if (complexBorder != null && (complexBorder.Bottom.Width > 0 || complexBorder.Left.Width > 0 || complexBorder.Top.Width > 0 || complexBorder.Right.Width > 0))
					{
						bool left = false;
						bool right = false;
						bool top = false;
						bool bottom = false;
						if (complexBorder.Bottom.Width > 0)
						{
							bottom = true;
						}
						if (complexBorder.Left.Width > 0)
						{
							left = true;
						}
						if (complexBorder.Top.Width > 0)
						{
							top = true;
						}
						if (complexBorder.Right.Width > 0)
						{
							right = true;
						}
						sView.Cells[i, j].Border = new LineBorder(complexBorder.Bottom.Color, 1, left, top, right, bottom);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinRemoveRow(SheetView sView, Dictionary<int, int> listCount, int rowIndex, int rowConfig, int activeSheetIndex)
	{
		try
		{
			sView.Rows.Remove(rowIndex - 1, 1);
			listCount[activeSheetIndex]--;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinSetBackColor(SheetView sView, string cellIndex)
	{
		try
		{
			string[] array = cellIndex.Split(new char[1] { '_' });
			StyleInfo styleInfo = new StyleInfo();
			styleInfo.BackColor = Color.White;
			sView.SetStyleInfo(int.Parse(array[1]), HTKKFile.ColumnLettersToIndex(array[0]), styleInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertDataRows(SheetView sView, int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount)
	{
		if (rowCount != 0)
		{
			sView.Rows.Add(frRow, rowCount);
			int num = (int)Math.Floor(Math.Log(rowCount, 2.0));
			for (int i = 0; i < num; i++)
			{
				sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, i) - 1.0), toCol, (int)Math.Pow(2.0, i), colCount + 2, dataOnly: false);
				string text = ((int)((double)toRow + Math.Pow(2.0, i) - 1.0)).ToString();
				string text2 = ((int)Math.Pow(2.0, i)).ToString();
			}
			sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, num) - 1.0), toCol, rowCount - (int)Math.Pow(2.0, num) + 1, colCount + 2, dataOnly: false);
		}
	}

	public static List<HTKKCMCRowValue> GetDataFromExcelFile(FpSpread winFpSpread, HTKKCMCRowInfo rowInfo, int DataLocationTaiBangKeRowStart, string flagStop, out int rowStartNext, out string status)
	{
		rowStartNext = DataLocationTaiBangKeRowStart;
		status = "";
		int num = HTKKFile.ColumnLettersToIndex("B");
		if (WinSettings.MenuID == "Q5")
		{
			num = HTKKFile.ColumnLettersToIndex("A");
		}
		List<HTKKCMCRowValue> list;
		try
		{
			list = new List<HTKKCMCRowValue>();
			for (int i = DataLocationTaiBangKeRowStart - 1; i <= winFpSpread.ActiveSheet.NonEmptyRowCount - 1; i++)
			{
				int num2 = 0;
				if (i == 5)
				{
					num2 = i;
				}
				string text = winFpSpread.ActiveSheet.Cells[i, num - 1].Text;
				if ((flagStop.Contains(text) && WinSettings.MenuID != "Q5") || (flagStop.Contains(text) && !string.IsNullOrEmpty(text) && WinSettings.MenuID == "Q5"))
				{
					if (WinSettings.MenuID == "P6" || WinSettings.MenuID == "P2")
					{
						rowStartNext = i + 2;
					}
					if (WinSettings.MenuID == "X7")
					{
						rowStartNext = i + 2;
					}
					for (int j = i + 1; j <= winFpSpread.ActiveSheet.NonEmptyRowCount - 1; j++)
					{
						if (WinSettings.MenuID == "03")
						{
							if (!HTKKFormater.isNullString(winFpSpread.ActiveSheet.Cells[j, num - 1].Text.Trim()))
							{
								bool flag = false;
								switch (winFpSpread.ActiveSheet.Cells[j, num - 1].Text.Trim())
								{
								case "aa":
									rowStartNext = j;
									flag = true;
									break;
								case "bb":
									rowStartNext = j;
									flag = true;
									break;
								case "cc":
									rowStartNext = j;
									flag = true;
									break;
								case "dd":
									rowStartNext = j;
									flag = true;
									break;
								case "ee":
									rowStartNext = j;
									flag = true;
									break;
								case "ff":
									rowStartNext = j;
									flag = true;
									break;
								case "gg":
									rowStartNext = j;
									flag = true;
									break;
								case "hh":
									rowStartNext = j;
									flag = true;
									break;
								}
								if (flag)
								{
									break;
								}
							}
						}
						else if (!HTKKFormater.isNullString(winFpSpread.ActiveSheet.Cells[j - 1, num - 1].Text.Trim()) && flagStop.IndexOf(winFpSpread.ActiveSheet.Cells[j - 1, num - 1].Text.Trim()) > -1)
						{
							rowStartNext = j + 2;
							if (WinSettings.MenuID == "68")
							{
								rowStartNext--;
							}
							if (WinSettings.MenuID == "95")
							{
								rowStartNext += 2;
							}
							if (WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8")
							{
								rowStartNext += 3;
							}
							if (WinSettings.MenuID == "79")
							{
								rowStartNext += 2;
							}
							if (WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5")
							{
								rowStartNext += 3;
							}
							if (WinSettings.MenuID == "L4" || WinSettings.MenuID == "N7" || WinSettings.MenuID == "L8")
							{
								rowStartNext--;
							}
							if (WinSettings.MenuID == "D1" || WinSettings.MenuID.Equals("Z5"))
							{
								rowStartNext += 3;
							}
							if (WinSettings.MenuID == "P6")
							{
								rowStartNext--;
							}
							if (WinSettings.MenuID == "P2")
							{
								rowStartNext--;
							}
							if (WinSettings.MenuID == "Q5")
							{
								rowStartNext++;
							}
							break;
						}
					}
					break;
				}
				HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
				for (int j = 0; j < rowInfo.Cells.Count; j++)
				{
					if (rowInfo.Cells[j].CellBangKe == null)
					{
						continue;
					}
					if (WinSettings.MenuID == "95" && DataLocationTaiBangKeRowStart > 40 && "#02TH;3;5;6;7;10;11;12;13;14;15;16;17;18;19".Split(new char[1] { ';' }).Contains(j.ToString()))
					{
						hTKKCMCRowValue.RowValue.Add("");
					}
					else if (rowInfo.Cells[j].CellBangKe.ColIndex >= 1)
					{
						Type type = winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].GetType();
						if (rowInfo.Cells[j].Controltype == "14")
						{
							string item = "";
							if (!HTKKFormater.isNullString(winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Text.Trim()))
							{
								if (winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Text.IndexOf("/") > -1)
								{
									item = winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Text.Trim();
								}
								else if (ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Value) > 0.0)
								{
									item = DateTime.FromOADate(Convert.ToDouble(winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Value, CultureInfo.InvariantCulture)).ToString("dd/MM/yyyy");
								}
							}
							hTKKCMCRowValue.RowValue.Add(item);
							continue;
						}
						string text2 = winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Text.Trim();
						if (rowInfo.Cells[j].Controltype == "16")
						{
							if (text2.IndexOf("E") > -1)
							{
								string text3 = winFpSpread.ActiveSheet.Cells[i, rowInfo.Cells[j].CellBangKe.ColIndex - 1].Value.ToString();
								if (text3.LastIndexOf(".0") > -1)
								{
									text3 = text3.Substring(0, text3.LastIndexOf(".0") - 1);
								}
								hTKKCMCRowValue.RowValue.Add(text3);
								if ((text3.ToString().IndexOf(".") > -1 || text3.ToString().IndexOf(",") > -1) && rowInfo.Cells[j].DefaultValue.IndexOf(".") == -1 && rowInfo.Cells[j].DefaultValue.IndexOf(",") == -1)
								{
									status = HTKKFile.ColumnIndexToLetters(rowInfo.Cells[j].CellBangKe.ColIndex) + (i + 1);
									throw new Exception("Valid format");
								}
							}
							else
							{
								hTKKCMCRowValue.RowValue.Add(text2);
								if ((text2.ToString().IndexOf(".") > -1 || text2.ToString().IndexOf(",") > -1) && rowInfo.Cells[j].DefaultValue.IndexOf(".") == -1 && rowInfo.Cells[j].DefaultValue.IndexOf(",") == -1)
								{
									status = HTKKFile.ColumnIndexToLetters(rowInfo.Cells[j].CellBangKe.ColIndex) + (i + 1);
									throw new Exception("Valid format");
								}
							}
						}
						else
						{
							hTKKCMCRowValue.RowValue.Add(text2);
						}
					}
					else
					{
						hTKKCMCRowValue.RowValue.Add("");
					}
				}
				list.Add(hTKKCMCRowValue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return list;
	}

	public static Dictionary<int, List<HTKKCMCRowValue>> GetDataFromExcelLarge(FpSpread BangKeWinFpSpread, HTKKCMCSheetInfo sInfo, List<HTKKCMCTableInfo> tables, bool themDuLieu, ref string status)
	{
		List<string> list = new List<string>();
		if (WinSettings.MenuID == "N7" && sInfo.SheetIndex == 1)
		{
			list.Add("II. Thông tin chung nhà cung cấp nước ngoài");
		}
		if (WinSettings.MenuID == "N7" && sInfo.SheetIndex == 3)
		{
			list.Add("II. Thông tin tài khoản nhà cung cấp nước ngoài");
		}
		if (WinSettings.MenuID == "Q8" && sInfo.SheetIndex == 2)
		{
			list.Add("II. MÁY MÓC, THIẾT BỊ");
		}
		if (WinSettings.MenuID == "L8" && sInfo.SheetIndex == 1)
		{
			list.Add("Tổng cộng");
			list.Add("I");
			list.Add("II");
			list.Add("III");
			list.Add("IV");
			list.Add("V");
		}
		if (WinSettings.MenuID == "Y6" && sInfo.SheetIndex == 1)
		{
			list.Add("B. CÁ NHÂN CÓ THU NHẬP TỪ ĐẦU TƯ VỐN");
			list.Add("C. CÁ NHÂN KHÔNG CƯ TRÚ CÓ DOANH THU TỪ KINH DOANH");
			list.Add("D. CÁ NHÂN CÓ THU NHẬP TỪ CÁC KHOẢN THU KHÁC");
			list.Add("II");
			list.Add("III");
			list.Add("IV");
			list.Add("V");
			list.Add("VI");
		}
		if (WinSettings.MenuID == "99" && sInfo.SheetIndex == 3)
		{
			list.Add("aa");
			list.Add("bb");
			list.Add("cc");
		}
		if (WinSettings.MenuID == "G8")
		{
			list.Add("Tổng cộng");
		}
		if (WinSettings.MenuID == "R4")
		{
			list.Add("Tổng cộng");
		}
		if (WinSettings.MenuID == "X8")
		{
			list.Add("Tổng cộng");
		}
		if (WinSettings.MenuID == "U7")
		{
			list.Add("Tổng cộng");
		}
		if (WinSettings.MenuID == "P6")
		{
			list.Add("Tiểu mục 1003 - Thuế thu nhập từ hoạt động sản xuất, kinh doanh của cá nhân");
		}
		if (WinSettings.MenuID == "X7")
		{
			list.Add("Tiểu mục 1099 - Thuế thu nhập doanh nghiệp khác");
			list.Add("II. BẢNG\u00a0KÊ CHI TIẾT SỐ TIỀN NỘP THUẾ KHẤU TRỪ CỦA HỘ, CÁ NHÂN KINH DOANH");
			list.Add("Tiểu mục 1003 - Thuế thu nhập từ hoạt động sản xuất, kinh doanh của cá nhân");
		}
		if (WinSettings.MenuID == "P2")
		{
			list.Add("Thông tin hộ, cá nhân nước ngoài");
		}
		if (WinSettings.MenuID == "W1" && sInfo.SheetIndex == 1)
		{
			list.Add("Trường hợp cá nhân thuộc diện điều chuyển (tổ chức có đánh dấu vào chỉ tiêu [04] trên Tờ khai 05/QTT-TNCN)");
			list.Add("Trường hợp cá nhân có khoản thu nhập được miễn thuế hoặc có số thuế được miễn theo quy định của pháp luật thuế thu nhập cá nhân");
		}
		Dictionary<int, List<HTKKCMCRowValue>> dictionary = new Dictionary<int, List<HTKKCMCRowValue>>();
		dictionary.Add(0, new List<HTKKCMCRowValue>());
		try
		{
			int num = 0;
			int num2 = ((sInfo.Tables[num].Location.DataLocationTaiBangKeRowStart >= 0) ? sInfo.Tables[num].Location.DataLocationTaiBangKeRowStart : 0);
			for (int i = num2 - 1; i <= BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex].NonEmptyRowCount - 1; i++)
			{
				try
				{
					while (true)
					{
						IL_03a6:
						HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
						HTKKCMCRowInfo rowsInfo = tables[num].RowsInfo;
						int num3 = 0;
						while (true)
						{
							if (num3 < rowsInfo.Cells.Count)
							{
								try
								{
									if (rowsInfo.Cells[num3].CellBangKe != null)
									{
										if (rowsInfo.Cells[num3].CellBangKe.ColIndex > 0)
										{
											Cell cell = BangKeWinFpSpread.ActiveSheet.Cells[i, rowsInfo.Cells[num3].CellBangKe.ColIndex - 1];
											if (rowsInfo.Cells[num3].Controltype == "14")
											{
												string text = "";
												if (!HTKKFormater.isNullString(cell.Text.Trim()))
												{
													if (cell.Text.IndexOf("/") > -1)
													{
														text = cell.Text.Trim();
													}
													else if (ConvertHelpers.ConvertToDouble(cell.Value) > 0.0)
													{
														text = DateTime.FromOADate(Convert.ToDouble(cell.Value, CultureInfo.InvariantCulture)).ToString("dd/MM/yyyy");
													}
												}
												hTKKCMCRowValue.RowValue.Add(text.Replace("/", ""));
											}
											else if (rowsInfo.Cells[num3].Controltype == "16")
											{
												if (cell.Value != null)
												{
													if (!GeneralFunctions.IsNumeric(cell.Value.ToString()))
													{
														status = status + HTKKFile.ColumnIndexToLetters(rowsInfo.Cells[num3].CellBangKe.ColIndex) + (i + 1) + ",";
														hTKKCMCRowValue.RowValue.Add("0");
													}
													else
													{
														hTKKCMCRowValue.RowValue.Add(cell.Value.ToString());
													}
												}
												else
												{
													hTKKCMCRowValue.RowValue.Add("0");
												}
											}
											else if (rowsInfo.Cells[num3].Controltype == "2")
											{
												hTKKCMCRowValue.RowValue.Add((!string.IsNullOrEmpty(cell.Text.Trim()) && (cell.Text.Trim().ToLower() == "x" || cell.Text.Trim() == "1")) ? "1" : "0");
											}
											else
											{
												if (list.IndexOf(cell.Text.Trim()) >= 0 && num3 < 2)
												{
													num++;
													i = ((num == 1 && WinSettings.MenuID == "L8" && sInfo.SheetIndex == 1) ? (i + 6) : ((!(WinSettings.MenuID == "G8")) ? ((!(WinSettings.MenuID == "R4")) ? ((!(WinSettings.MenuID == "X8")) ? ((!(WinSettings.MenuID == "U7")) ? ((num == 2 && WinSettings.MenuID == "X7" && sInfo.SheetIndex == 0) ? (i + 5) : ((WinSettings.MenuID == "Y6" && sInfo.SheetIndex == 1) ? ((num != 1 && num != 2) ? ((num != 3) ? (i + 1) : (i + 4)) : (i + 3)) : ((!(WinSettings.MenuID == "W1") || sInfo.SheetIndex != 1) ? (i + 1) : ((num != 1 && num != 2) ? (i + 1) : (i + 4))))) : (i + 2)) : (i + 2)) : (i + 2)) : (i + 2)));
													goto IL_03a6;
												}
												hTKKCMCRowValue.RowValue.Add(cell.Text.Trim());
											}
										}
										else
										{
											hTKKCMCRowValue.RowValue.Add((rowsInfo.Cells[num3].Controltype == "16") ? "0" : "");
										}
									}
									else
									{
										hTKKCMCRowValue.RowValue.Add((rowsInfo.Cells[num3].Controltype == "16") ? "0" : "");
									}
								}
								catch (Exception ex)
								{
									HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
								}
								num3++;
								continue;
							}
							if (GeneralFunctions.HaveToValidateRow(rowsInfo, hTKKCMCRowValue.RowValue.ToArray()))
							{
								if (dictionary.ContainsKey(num))
								{
									dictionary[num].Add(hTKKCMCRowValue);
									break;
								}
								dictionary.Add(num, new List<HTKKCMCRowValue>());
								dictionary[num].Add(hTKKCMCRowValue);
							}
							break;
						}
						break;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		status = ((status.Length < 1) ? "" : status.Substring(0, status.Length - 1));
		return dictionary;
	}

	public static bool GetDataFromExcelFileGroupF7(FpSpread ToKhaiWinFpSpread, FpSpread BangKeWinFpSpread, HTKKCMCSheetInfo ToKhaisInfo, XmlNode xmlRootNode, int DataLocationTaiBangKeRowStart, int rowStartNext, bool themDuLieu, ref string status)
	{
		int activeSheetIndex = ToKhaiWinFpSpread.ActiveSheetIndex;
		rowStartNext = DataLocationTaiBangKeRowStart;
		int num = 0;
		int num2 = 0;
		bool result = false;
		try
		{
			num = GetInfoTableCoTaiBangKe(xmlRootNode, ToKhaisInfo);
			if (num >= 0)
			{
				if (!themDuLieu)
				{
					DeleteTables1(ToKhaiWinFpSpread, ToKhaisInfo, num);
				}
				HTKKCMCRowInfo rowsInfo = ToKhaisInfo.Tables[num].RowsInfo;
				num2 = ToKhaisInfo.Tables[num].Location.DataLocationTemplateRowStart;
				for (int i = DataLocationTaiBangKeRowStart - 1; i <= BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex].NonEmptyRowCount - 1; i++)
				{
					try
					{
						string BangKeValueColMa = "";
						string BangKeValueColTen = "";
						int toKhaiColMa = -1;
						int toKhaiColTen = -1;
						GetValueColMaTenByGroup(BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex], xmlRootNode, i, ref BangKeValueColMa, ref BangKeValueColTen, ref toKhaiColMa, ref toKhaiColTen, ref status);
						if (string.IsNullOrEmpty(BangKeValueColMa) && string.IsNullOrEmpty(BangKeValueColTen))
						{
							if (!string.IsNullOrEmpty(status))
							{
								DeleteTables1(ToKhaiWinFpSpread, ToKhaisInfo, num);
								BangKeWinFpSpread.ResumeLayout();
								ToKhaiWinFpSpread.ResumeLayout();
								return true;
							}
						}
						else
						{
							if ((string.IsNullOrEmpty(BangKeValueColMa) && string.IsNullOrEmpty(BangKeValueColTen)) || (toKhaiColMa < 0 && toKhaiColTen < 0))
							{
								continue;
							}
							int num3 = GetTableInGroup(ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex], ToKhaisInfo, BangKeValueColMa, BangKeValueColTen, num2, toKhaiColMa, toKhaiColTen);
							if (num3 < 0)
							{
								bool haveRow = false;
								int rowIndex = ToKhaisInfo.Tables[num].Header.StaticCells[0].CellTemplate.RowIndex;
								if (!string.IsNullOrEmpty(BangKeValueColTen))
								{
									object value = ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex].Cells[rowIndex - 1, toKhaiColTen - 1].Value;
									haveRow = ((value != null && !string.IsNullOrEmpty(value.ToString())) ? true : false);
								}
								if (!string.IsNullOrEmpty(BangKeValueColMa))
								{
									object value = ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex].Cells[rowIndex - 1, toKhaiColMa - 1].Value;
									haveRow = ((value != null && !string.IsNullOrEmpty(value.ToString())) ? true : false);
								}
								num3 = GetCountTables(ToKhaisInfo, haveRow);
								if (num3 > 0)
								{
									InsertTable(ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex], num3 - 1, ToKhaisInfo);
									fpSpReindexingSheetTemplate(ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex], ToKhaisInfo);
									WinDoWorkEvent.UpdateFormula(ToKhaiWinFpSpread, ToKhaiWinFpSpread.ActiveSheetIndex);
									WinDoWorkEvent.KeyPress(ToKhaiWinFpSpread, ToKhaiWinFpSpread.ActiveSheetIndex, ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex].ActiveRow.Index + 1, ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex].ActiveColumn.Index + 1, null, (Keys)118);
									num2 = ToKhaisInfo.Tables[num3].Location.DataLocationTemplateRowStart;
								}
							}
							SetValueToHeader(ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex], ToKhaisInfo, num3, BangKeValueColMa, BangKeValueColTen, num2, toKhaiColMa, toKhaiColTen);
							num2 = ToKhaisInfo.Tables[num3].Location.DataLocationTemplateRowStart;
							List<string> list = new List<string>();
							int num4 = 0;
							foreach (HTKKCMCTableCellInfo cell in ToKhaisInfo.Tables[num3].RowsInfo.Cells)
							{
								try
								{
									object obj = ((num4 != 0) ? ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex].Cells[num2 - 1, cell.CellTemplate.ColIndex - 1].Value : cell.DefaultValue);
									list.Add((obj == null) ? "" : obj.ToString().Trim());
									num4++;
								}
								catch (Exception ex)
								{
									HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
								}
							}
							if (GeneralFunctions.HaveToValidateRow(rowsInfo, list.ToArray()))
							{
								num2 += ToKhaisInfo.Tables[num3].RowsValues.Count;
								GeneralFunctions.InsertTableRow(ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex], num2 - 1, -1, num2 - 1, -1, -1, ToKhaisInfo, num3);
							}
							string text = "";
							for (int j = 0; j < rowsInfo.Cells.Count; j++)
							{
								try
								{
									if (rowsInfo.Cells[j].CellBangKe != null)
									{
										string valueAndFormat = GetValueAndFormat(BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex], rowsInfo, i, j, ref status);
										if (!string.IsNullOrEmpty(status))
										{
											text = text + "Bảng kê không đúng định dạng tại vị trí " + status + "\r\n";
											break;
										}
										ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex].Cells[num2 - 1, rowsInfo.Cells[j].CellTemplate.ColIndex - 1].Value = valueAndFormat;
										result = true;
									}
								}
								catch (Exception ex)
								{
									HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
								}
							}
							if (!string.IsNullOrEmpty(text))
							{
								status = text;
								DeleteTables1(ToKhaiWinFpSpread, ToKhaisInfo, num);
								BangKeWinFpSpread.ResumeLayout();
								ToKhaiWinFpSpread.ResumeLayout();
								result = true;
								break;
							}
							continue;
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
				GeneralFunctions.fpSpReindexingSheetTemplate(ToKhaiWinFpSpread.Sheets[ToKhaiWinFpSpread.ActiveSheetIndex], ToKhaisInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static Dictionary<int, List<HTKKCMCRowValue>> GetDataFromExcelFileGroupF7LargeData(FpSpread ToKhaiWinFpSpread, FpSpread BangKeWinFpSpread, ref HTKKCMCSheetInfo ToKhaisInfo, XmlNode xmlRootNode, int DataLocationTaiBangKeRowStart, int rowStartNext, bool themDuLieu, ref string status, int sheetIndex)
	{
		BangKeWinFpSpread.SuspendLayout();
		rowStartNext = DataLocationTaiBangKeRowStart;
		int num = 0;
		Dictionary<int, List<HTKKCMCRowValue>> dictionary = new Dictionary<int, List<HTKKCMCRowValue>>();
		try
		{
			num = GetInfoTableCoTaiBangKe(xmlRootNode, ToKhaisInfo);
			if (num >= 0)
			{
				HTKKCMCRowInfo rowInfo = ToKhaisInfo.Tables[num].RowsInfo;
				for (int i = DataLocationTaiBangKeRowStart - 1; i <= BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex].NonEmptyRowCount - 1; i++)
				{
					try
					{
						string BangKeValueColMa = "";
						string BangKeValueColTen = "";
						int toKhaiColMa = -1;
						int toKhaiColTen = -1;
						GetValueColMaTenByGroup(BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex], xmlRootNode, i, ref BangKeValueColMa, ref BangKeValueColTen, ref toKhaiColMa, ref toKhaiColTen, ref status);
						if ((string.IsNullOrEmpty(BangKeValueColMa) && string.IsNullOrEmpty(BangKeValueColTen)) || (string.IsNullOrEmpty(BangKeValueColMa) && string.IsNullOrEmpty(BangKeValueColTen)) || (toKhaiColMa < 0 && toKhaiColTen < 0))
						{
							continue;
						}
						num = GetTableInGroup2(ToKhaisInfo, BangKeValueColMa, BangKeValueColTen, toKhaiColMa, toKhaiColTen);
						HTKKCMCRowValue rowValueDefault = GeneralFunctions.GetRowValueDefault(ToKhaisInfo.Tables[(num >= 0) ? num : 0].RowsInfo);
						if (ToKhaisInfo.Tables[(num >= 0) ? num : 0].RowsValues.Count == 0)
						{
							ToKhaisInfo.Tables[(num >= 0) ? num : 0].RowsValues.Add(rowValueDefault);
						}
						if (num < 0)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaisInfo.Tables[0];
							bool haveRow = false;
							int headerRow = hTKKCMCTableInfo.Header.StaticCells[0].CellTemplate.RowIndex;
							if (!string.IsNullOrEmpty(BangKeValueColTen))
							{
								string value = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == headerRow && x.CellTemplate.ColIndex == toKhaiColTen).Value;
								haveRow = ((value != null && !string.IsNullOrEmpty(value.ToString())) ? true : false);
							}
							if (!string.IsNullOrEmpty(BangKeValueColMa))
							{
								string value = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == headerRow && x.CellTemplate.ColIndex == toKhaiColMa).Value;
								haveRow = ((value != null && !string.IsNullOrEmpty(value.ToString())) ? true : false);
							}
							num = GetCountTables(ToKhaisInfo, haveRow);
							if (num > 0)
							{
								GeneralFunctions.InsertTableF7(ToKhaiWinFpSpread.Sheets[sheetIndex], ToKhaisInfo, num - 1);
							}
						}
						SetValueToHeader2(ref ToKhaisInfo, num, BangKeValueColMa, BangKeValueColTen, toKhaiColMa, toKhaiColTen);
						HTKKCMCRowValue hTKKCMCRowValue = rowValueDefault;
						int j;
						for (j = 0; j < rowInfo.Cells.Count; j++)
						{
							try
							{
								if (rowInfo.Cells[j].CellBangKe != null)
								{
									string valueAndFormat = GetValueAndFormat(BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex], rowInfo, i, j, ref status);
									int num2 = rowInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == rowInfo.Cells[j].CellTemplate.ColIndex);
									if (num2 >= 0)
									{
										hTKKCMCRowValue.RowValue[num2] = valueAndFormat;
									}
								}
							}
							catch (Exception ex)
							{
								HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
							}
						}
						if (!dictionary.ContainsKey(num))
						{
							List<HTKKCMCRowValue> list = new List<HTKKCMCRowValue>();
							list.Add(hTKKCMCRowValue);
							dictionary.Add(num, list);
						}
						else
						{
							dictionary[num].Add(hTKKCMCRowValue);
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
				BangKeWinFpSpread.ResumeLayout();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static void DeleteTables1(FpSpread ToKhaiWinFpSpread, HTKKCMCSheetInfo ToKhaisInfo, int tableIndex0)
	{
		for (int num = ToKhaisInfo.Tables.Count - 1; num >= 1; num--)
		{
			int row = ToKhaisInfo.Tables[num].Location.DataLocationTemplateRowStart - 1;
			int column = 0;
			RemoveTable(ToKhaiWinFpSpread.ActiveSheet, num, ToKhaisInfo);
			fpSpReindexingSheetTemplate(ToKhaiWinFpSpread.ActiveSheet, ToKhaisInfo);
			WinDoWorkEvent.UpdateFormula(ToKhaiWinFpSpread, ToKhaiWinFpSpread.ActiveSheetIndex);
			ToKhaiWinFpSpread.ActiveSheet.SetActiveCell(row, column);
			WinDoWorkEvent.KeyPress(ToKhaiWinFpSpread, ToKhaiWinFpSpread.ActiveSheetIndex, ToKhaiWinFpSpread.ActiveSheet.ActiveRow.Index + 1, ToKhaiWinFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, (Keys)119);
		}
		if (ToKhaisInfo.Tables[tableIndex0].RowsValues.Count > 1)
		{
			int row = ToKhaisInfo.Tables[tableIndex0].Location.DataLocationTemplateRowStart - 1;
			int column = 0;
			ToKhaiWinFpSpread.ActiveSheet.SetActiveCell(row, column);
			int frRow = ToKhaisInfo.Tables[tableIndex0].Location.DataLocationTemplateRowStart - 1;
			ToKhaiWinFpSpread.ActiveSheet.Cells[ToKhaiWinFpSpread.ActiveSheet.ActiveRow.Index, ToKhaiWinFpSpread.ActiveSheet.ActiveColumn.Index].Value = GeneralFunctions.GetValue(ToKhaiWinFpSpread.ActiveSheet, ToKhaiWinFpSpread.ActiveSheet.ActiveRow.Index + 2, ToKhaiWinFpSpread.ActiveSheet.ActiveColumn.Index + 1);
			int count = ToKhaisInfo.Tables[tableIndex0].RowsValues.Count;
			for (int num2 = count - 1; num2 >= 1; num2--)
			{
				row = ToKhaisInfo.Tables[tableIndex0].Location.DataLocationTemplateRowStart + num2 - 1;
				ToKhaiWinFpSpread.ActiveSheet.SetActiveCell(row, column);
				ToKhaisInfo.Tables[tableIndex0].RowsValues.RemoveAt(num2);
				RemoveDataRows(ToKhaiWinFpSpread.ActiveSheet, frRow, 1, ToKhaisInfo);
				fpSpReindexingTableTemplate(ToKhaiWinFpSpread.ActiveSheet, ToKhaisInfo.Tables[tableIndex0]);
				WinDoWorkEvent.UpdateFormula(ToKhaiWinFpSpread, ToKhaiWinFpSpread.ActiveSheetIndex);
				WinDoWorkEvent.KeyPress(ToKhaiWinFpSpread, ToKhaiWinFpSpread.ActiveSheetIndex, ToKhaiWinFpSpread.ActiveSheet.ActiveRow.Index + 1, ToKhaiWinFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, (Keys)117);
			}
			row = ToKhaisInfo.Tables[tableIndex0].Location.DataLocationTemplateRowStart - 1;
			ToKhaiWinFpSpread.ActiveSheet.SetActiveCell(row, column);
			for (int i = 0; i < ToKhaisInfo.Tables[tableIndex0].RowsInfo.Cells.Count; i++)
			{
				fpSpBindDataCell(ToKhaiWinFpSpread.ActiveSheet, ToKhaiWinFpSpread.ActiveSheet.ActiveRow.Index + 1, ToKhaisInfo.Tables[tableIndex0].RowsInfo.Cells[i].CellTemplate.ColIndex, ToKhaisInfo.Tables[tableIndex0].RowsInfo.Cells[i].Controltype, ToKhaisInfo.Tables[tableIndex0].RowsInfo.Cells[i].DefaultValue);
			}
		}
	}

	public static void Deletetables(SheetView ToKhaisView, HTKKCMCSheetInfo sInfo, int tableIndex0, int sheetIndex)
	{
		try
		{
			int count = sInfo.Tables.Count;
			for (int num = count - 1; num >= tableIndex0; num--)
			{
				if (num < sInfo.Tables.Count - 1 && sInfo.Tables[num] != null)
				{
					int tableTemplateRowStart = sInfo.Tables[num].Location.TableTemplateRowStart;
					GeneralFunctions.DeleteTableF8(ToKhaisView, sInfo, num, tableTemplateRowStart - 1, 0);
				}
			}
			if (sInfo.Tables[tableIndex0].RowsValues.Count > 1)
			{
				int dataLocationTemplateRowStart = sInfo.Tables[tableIndex0].Location.DataLocationTemplateRowStart;
				GeneralFunctions.RemoveDataRows(ToKhaisView, tableIndex0, dataLocationTemplateRowStart - 1, sInfo.Tables[0].RowsValues.Count - 1, sInfo);
			}
			if (sInfo.Tables[tableIndex0].RowsValues.Count > 0)
			{
				int dataLocationTemplateRowStart = sInfo.Tables[tableIndex0].Location.DataLocationTemplateRowStart;
				foreach (HTKKCMCTableCellInfo cell in sInfo.Tables[tableIndex0].RowsInfo.Cells)
				{
					ToKhaisView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
			if (sInfo.Tables[tableIndex0].Header == null || sInfo.Tables[tableIndex0].Header.StaticCells == null)
			{
				return;
			}
			foreach (HTKKCMCCellInfo staticCell in sInfo.Tables[tableIndex0].Header.StaticCells)
			{
				ToKhaisView.Cells[staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1].Value = staticCell.DefaultValue;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetValueToHeader(SheetView ToKhaisView, HTKKCMCSheetInfo sInfo, int tableIndex, string BangKeValueColMa, string BangKeValueColTen, int ToKhaiRow, int toKhaiColMa, int toKhaiColTen)
	{
		int num = sInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart - 1;
		if (toKhaiColMa > 0)
		{
			ToKhaisView.Cells[num - 1, toKhaiColMa - 1].Value = BangKeValueColMa;
		}
		if (toKhaiColTen > 0)
		{
			ToKhaisView.Cells[num - 1, toKhaiColTen - 1].Value = BangKeValueColTen;
		}
	}

	public static void SetValueToHeader2(ref HTKKCMCSheetInfo sInfo, int tableIndex, string BangKeValueColMa, string BangKeValueColTen, int toKhaiColMa, int toKhaiColTen)
	{
		int headerRow = sInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart - 1;
		sInfo.Tables[tableIndex].Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == headerRow && x.CellTemplate.ColIndex == toKhaiColTen).Value = BangKeValueColTen;
		sInfo.Tables[tableIndex].Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == headerRow && x.CellTemplate.ColIndex == toKhaiColMa).Value = BangKeValueColMa;
	}

	public static int GetCountTables(HTKKCMCSheetInfo sInfo, bool haveRow)
	{
		int num = 0;
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table.Header != null && table.Header.StaticCells != null)
			{
				num++;
			}
		}
		return (num != 1 || haveRow) ? num : 0;
	}

	public static string GetValueAndFormat(SheetView BangKesView, HTKKCMCRowInfo rowInfo, int BangKeRow, int BangKeCol, ref string status)
	{
		string result = "";
		if (rowInfo.Cells[BangKeCol].Controltype == "14")
		{
			string text = "";
			string text2 = BangKesView.Cells[BangKeRow, rowInfo.Cells[BangKeCol].CellBangKe.ColIndex - 1].Text;
			if (!string.IsNullOrEmpty(text2) && !HTKKFormater.isNullString(text2.Trim()))
			{
				if (text2.IndexOf("/") > -1)
				{
					text = text2.Trim();
				}
				else if (ConvertHelpers.ConvertToDouble(BangKesView.Cells[BangKeRow, rowInfo.Cells[BangKeCol].CellBangKe.ColIndex - 1].Value) > 0.0)
				{
					text = DateTime.FromOADate(Convert.ToDouble(BangKesView.Cells[BangKeRow, rowInfo.Cells[BangKeCol].CellBangKe.ColIndex - 1].Value, CultureInfo.InvariantCulture)).ToString("dd/MM/yyyy");
				}
			}
			result = text;
		}
		else
		{
			if (rowInfo.Cells[BangKeCol].CellBangKe.ColIndex == 0)
			{
				return result;
			}
			string text2 = BangKesView.Cells[BangKeRow, rowInfo.Cells[BangKeCol].CellBangKe.ColIndex - 1].Text;
			if (rowInfo.Cells[BangKeCol].Controltype == "16")
			{
				object value = BangKesView.Cells[BangKeRow, rowInfo.Cells[BangKeCol].CellBangKe.ColIndex - 1].Value;
				string text3 = ((value == null || value == "") ? "0" : value.ToString());
				if (text3.LastIndexOf(".0") > -1)
				{
					text3 = text3.Substring(0, text3.LastIndexOf(".0") - 1);
				}
				result = text3;
				if ((text3.ToString().IndexOf(".") > -1 || text3.ToString().IndexOf(",") > -1) && rowInfo.Cells[BangKeCol].DefaultValue.IndexOf(".") == -1 && rowInfo.Cells[BangKeCol].DefaultValue.IndexOf(",") == -1)
				{
					status = HTKKFile.ColumnIndexToLetters(rowInfo.Cells[BangKeCol].CellBangKe.ColIndex - 1) + (BangKeRow + 1);
				}
			}
			else
			{
				result = text2.Trim();
			}
		}
		return result;
	}

	public static void GetValueColMaTenByGroup(SheetView BKsView, XmlNode xmlRootNode, int BKiRow, ref string BangKeValueColMa, ref string BangKeValueColTen, ref int toKhaiColMa, ref int toKhaiColTen, ref string status)
	{
		XmlNode val = null;
		XmlNode val2 = null;
		string text = "";
		string text2 = "";
		XmlNodeList val3 = xmlRootNode.SelectNodes("Section/Header/Cell");
		if (val3 != null)
		{
			val = ((IEnumerable)val3).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["BKMaGroupByCol"] != null));
			text = ((val != null) ? ((XmlNode)val.Attributes["BKMaGroupByCol"]).Value : "");
		}
		XmlNodeList val4 = xmlRootNode.SelectNodes("Section/Header/Cell");
		if (val4 != null)
		{
			val2 = ((IEnumerable)val4).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["BKTenGroupByCol"] != null));
			text2 = ((val2 != null) ? ((XmlNode)val2.Attributes["BKTenGroupByCol"]).Value : "");
		}
		if (!string.IsNullOrEmpty(text))
		{
			object value = BKsView.Cells[BKiRow, HTKKFile.ColumnLettersToIndex(text) - 1].Value;
			BangKeValueColMa = ((value != null) ? value.ToString().Trim() : "");
		}
		if (!string.IsNullOrEmpty(text2))
		{
			object value = BKsView.Cells[BKiRow, HTKKFile.ColumnLettersToIndex(text2) - 1].Value;
			BangKeValueColTen = ((value != null) ? value.ToString().Trim() : "");
		}
		if (val != null)
		{
			string[] array = ((XmlNode)val.Attributes["CellID"]).Value.Trim().Split(new char[1] { '_' });
			toKhaiColMa = HTKKFile.ColumnLettersToIndex(array[0]);
		}
		if (val2 != null)
		{
			string[] array = ((XmlNode)val2.Attributes["CellID"]).Value.Trim().Split(new char[1] { '_' });
			toKhaiColTen = HTKKFile.ColumnLettersToIndex(array[0]);
		}
		if ((val != null || val2 != null) && BangKeValueColMa == "" && BangKeValueColTen == "")
		{
			status = "Bảng kê không đúng định dạng";
		}
	}

	public static int GetTableInGroup(SheetView ToKhaisView, HTKKCMCSheetInfo sInfo, string BangKeValueColMa, string BangKeValueColTen, int ToKhaiRow, int toKhaiColMa, int toKhaiColTen)
	{
		List<HTKKCMCTableInfo> tables = sInfo.Tables;
		for (int i = 0; i < tables.Count; i++)
		{
			int num = tables[i].Location.DataLocationTemplateRowStart - 1;
			if (!string.IsNullOrEmpty(BangKeValueColMa))
			{
				object value = ToKhaisView.Cells[num - 1, toKhaiColMa - 1].Value;
				if (value != null && BangKeValueColMa == value.ToString().Trim())
				{
					return i;
				}
			}
			if (!string.IsNullOrEmpty(BangKeValueColTen))
			{
				object value2 = ToKhaisView.Cells[num - 1, toKhaiColTen - 1].Value;
				if (value2 != null && BangKeValueColTen == value2.ToString().Trim())
				{
					return i;
				}
			}
		}
		return -1;
	}

	public static int GetTableInGroup2(HTKKCMCSheetInfo sInfo, string BangKeValueColMa, string BangKeValueColTen, int toKhaiColMa, int toKhaiColTen)
	{
		List<HTKKCMCTableInfo> tables = sInfo.Tables;
		for (int i = 0; i < tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = tables[i];
			int headerRow = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			if (!string.IsNullOrEmpty(BangKeValueColMa))
			{
				string value = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == headerRow && x.CellTemplate.ColIndex == toKhaiColMa).Value;
				if (value != null && BangKeValueColMa == value.ToString().Trim())
				{
					return i;
				}
			}
			if (!string.IsNullOrEmpty(BangKeValueColTen))
			{
				string value2 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == headerRow && x.CellTemplate.ColIndex == toKhaiColTen).Value;
				if (value2 != null && BangKeValueColTen == value2.ToString().Trim())
				{
					return i;
				}
			}
		}
		return -1;
	}

	public static int GetInfoTableCoTaiBangKe(XmlNode xmlRootNode, HTKKCMCSheetInfo sInfo)
	{
		int num = -1;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (num >= 0)
			{
				break;
			}
			XmlNodeList val = xmlRootNode.SelectNodes("Section");
			if (val.Count <= 0)
			{
				continue;
			}
			for (int j = 0; j < val.Count; j++)
			{
				XmlNode val2 = val[j];
				if (val2.Attributes["TaiBangKe"] != null && ((XmlNode)val2.Attributes["TaiBangKe"]).Value.Trim() == "1")
				{
					num = i;
					break;
				}
			}
		}
		return num;
	}

	public static List<HTKKCMCRowValue> GetDataFromExcelTmp(SheetView sViewImport, HTKKCMCRowInfo rowInfo, int DataLocationTemplateRowStart)
	{
		List<HTKKCMCRowValue> list;
		try
		{
			list = new List<HTKKCMCRowValue>();
			string text = "";
			int num = DataLocationTemplateRowStart - 1;
			while (text != "aa")
			{
				text = sViewImport.Cells[num, 1].Text;
				if (text != "aa")
				{
					HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
					for (int i = 0; i < rowInfo.Cells.Count; i++)
					{
						if (rowInfo.Cells[i].CellTemplate != null && rowInfo.Cells[i].CellTemplate.ColIndex >= 1)
						{
							string valueImportExcel = ConvertHelpers.GetValueImportExcel(sViewImport, num, rowInfo.Cells[i].CellTemplate.ColIndex - 1, rowInfo.Cells[i].Controltype);
							hTKKCMCRowValue.RowValue.Add(valueImportExcel);
						}
					}
					list.Add(hTKKCMCRowValue);
				}
				num++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return list;
	}

	public static void fpSpViewDisplayTemplate(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("B"))
			{
				if (sInfo.CoDieuChinhSoLieu >= 0)
				{
					sView.Columns[sInfo.CoDieuChinhSoLieu - 1].Visible = true;
				}
			}
			else if (sInfo.CoDieuChinhSoLieu >= 0)
			{
				sView.Columns[sInfo.CoDieuChinhSoLieu - 1].Visible = false;
			}
			if (sInfo.TaiDuLieuRatLon != "1")
			{
				for (int i = 0; i < sInfo.Tables.Count; i++)
				{
					if (sInfo.Tables[i] == null || !(sInfo.Tables[i].TableID == "1"))
					{
						continue;
					}
					for (int j = 0; j < sInfo.Tables.Count; j++)
					{
						if (sInfo.Tables[i].TableID != sInfo.Tables[j].TableID && sInfo.Tables[i].TableName == sInfo.Tables[j].TableName)
						{
							InsertTableInSpread(sView, i, sInfo);
						}
					}
				}
			}
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i] == null)
				{
					continue;
				}
				for (int j = 0; j < sInfo.Tables[i].Header.StaticCells.Count; j++)
				{
					fpSpFormatCellTemplate(sView, sInfo.Tables[i].Header.StaticCells[j]);
				}
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					sView.Cells[sInfo.Tables[i].CellTableSTT.RowIndex - 1, sInfo.Tables[i].CellTableSTT.ColIndex - 1].Value = sInfo.Tables[i].CellTableSTT.Value;
				}
				if (sInfo.Tables[i].RowsInfo.Cells.Count >= 0)
				{
					for (int j = 0; j < sInfo.Tables[i].RowsInfo.Cells.Count; j++)
					{
						fpSpFormatTableCell(sView, sInfo.Tables[i].Location.DataLocationTemplateRowStart, sInfo.Tables[i].RowsInfo.Cells[j]);
					}
					if (sInfo.Tables[i].RowsValues.Count > 0 && sInfo.Tables[i].RowsValues[0].RowDisplay == "0")
					{
						sView.Rows[sInfo.Tables[i].Location.DataLocationTemplateRowStart - 1].Visible = false;
					}
				}
				if (sInfo.TaiDuLieuLon == "1")
				{
					GC.Collect();
					GC.WaitForPendingFinalizers();
					ProcessingLargeDataOnSpread(sView, sInfo.SheetIndex, sInfo.Tables[i]);
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
				else if (sInfo.TaiDuLieuRatLon != "1" && sInfo.Tables[i].RowsValues.Count > 1)
				{
					InsertTableRowsSpreadForm(sView, sInfo.Tables[i].Location.DataLocationTemplateRowStart, -1, sInfo.Tables[i].Location.DataLocationTemplateRowStart, -1, sInfo.Tables[i].RowsValues.Count - 1, -1, sInfo, isSupportBindArray: false);
				}
				for (int j = 0; j < sInfo.Tables[i].Footer.StaticCells.Count; j++)
				{
					fpSpFormatCellTemplate(sView, sInfo.Tables[i].Footer.StaticCells[j]);
				}
			}
			for (int i = 0; i < sInfo.StaticCells.Count; i++)
			{
				fpSpFormatCellTemplate(sView, sInfo.StaticCells[i]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpViewDisplayTemplateNew(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("B"))
			{
				if (sInfo.CoDieuChinhSoLieu >= 0)
				{
					sView.Columns[sInfo.CoDieuChinhSoLieu - 1].Visible = true;
				}
			}
			else if (sInfo.CoDieuChinhSoLieu >= 0)
			{
				sView.Columns[sInfo.CoDieuChinhSoLieu - 1].Visible = false;
			}
			int i;
			for (i = 0; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i] == null || sInfo.Tables[i] == null)
				{
					continue;
				}
				if (sInfo.Tables[i].TableID == "1")
				{
					List<HTKKCMCTableInfo> list = sInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableID != "1" && x.TableName == sInfo.Tables[i].TableName).ToList();
					if (list.Count > 0)
					{
						foreach (HTKKCMCTableInfo item in list)
						{
							InsertTableInSpread(sView, i, sInfo);
						}
					}
				}
				for (int num = 0; num < sInfo.Tables[i].Header.StaticCells.Count; num++)
				{
					fpSpFormatCellTemplate(sView, sInfo.Tables[i].Header.StaticCells[num]);
				}
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					sView.Cells[sInfo.Tables[i].CellTableSTT.RowIndex - 1, sInfo.Tables[i].CellTableSTT.ColIndex - 1].Value = sInfo.Tables[i].CellTableSTT.Value;
				}
				if (sInfo.Tables[i].RowsInfo.Cells.Count >= 0)
				{
					for (int num = 0; num < sInfo.Tables[i].RowsInfo.Cells.Count; num++)
					{
						fpSpFormatTableCell(sView, sInfo.Tables[i].Location.DataLocationTemplateRowStart, sInfo.Tables[i].RowsInfo.Cells[num]);
					}
					if (sInfo.Tables[i].RowsValues.Count > 0 && sInfo.Tables[i].RowsValues[0].RowDisplay == "0")
					{
						sView.Rows[sInfo.Tables[i].Location.DataLocationTemplateRowStart - 1].Visible = false;
					}
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[sInfo.SheetIndex].TaiDuLieuLon == "1")
				{
					ProcessingLargeDataOnSpread(sView, sInfo.SheetIndex, sInfo.Tables[i]);
				}
				else if (sInfo.Tables[i].RowsValues.Count > 1)
				{
					InsertTableRowsSpreadForm(sView, sInfo.Tables[i].Location.DataLocationTemplateRowStart, -1, sInfo.Tables[i].Location.DataLocationTemplateRowStart, -1, sInfo.Tables[i].RowsValues.Count - 1, -1, sInfo, isSupportBindArray: false);
				}
				for (int num = 0; num < sInfo.Tables[i].Footer.StaticCells.Count; num++)
				{
					fpSpFormatCellTemplate(sView, sInfo.Tables[i].Footer.StaticCells[num]);
				}
			}
			for (int num2 = 0; num2 < sInfo.StaticCells.Count; num2++)
			{
				fpSpFormatCellTemplate(sView, sInfo.StaticCells[num2]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpViewDisplayReport(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				if (!(sInfo.Tables[i].TableID == "1"))
				{
					continue;
				}
				for (int j = 0; j < sInfo.Tables.Count; j++)
				{
					if (sInfo.Tables[i].TableID != sInfo.Tables[j].TableID && sInfo.Tables[i].TableName == sInfo.Tables[j].TableName)
					{
						InsertTableInSpreadReport(sView, i, sInfo);
					}
				}
			}
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i].RowsValues.Count > 1)
				{
					InsertTableRowsSpreadForm(sView, sInfo.Tables[i].Location.DataLocationReportRowStart, -1, sInfo.Tables[i].Location.DataLocationReportRowStart, -1, sInfo.Tables[i].RowsValues.Count - 1, -1, sInfo, isSupportBindArray: false);
				}
				if (sInfo.Tables[i].CellTableReportSTT != null)
				{
					sView.Cells[sInfo.Tables[i].CellTableReportSTT.RowIndex - 1, sInfo.Tables[i].CellTableReportSTT.ColIndex - 1].Value = sInfo.Tables[i].CellTableReportSTT.Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void fpSpBindDataTemplate(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			sInfo.OldSheetName = sView.SheetName;
			sView.SheetName = sInfo.SheetName;
			if (sInfo.TaiDuLieuRatLon != "1")
			{
				WinDoWorkEvent.ProcessSelectedCell(sInfo);
			}
			if (sInfo.TaiDuLieuRatLon != "1")
			{
				for (int i = 0; i < sInfo.StaticCells.Count; i++)
				{
					fpSpBindDataCell(sView, sInfo.StaticCells[i].CellTemplate.RowIndex, sInfo.StaticCells[i].CellTemplate.ColIndex, sInfo.StaticCells[i].Controltype, sInfo.StaticCells[i].Value);
				}
			}
			if (!(sInfo.TaiDuLieuLon != "1") || !(sInfo.TaiDuLieuRatLon != "1"))
			{
				return;
			}
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				for (int j = 0; j < sInfo.Tables[i].Header.StaticCells.Count; j++)
				{
					fpSpBindDataCell(sView, sInfo.Tables[i].Header.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[i].Header.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[i].Header.StaticCells[j].Controltype, sInfo.Tables[i].Header.StaticCells[j].Value);
				}
				HTKKFile.LogRun("TaiDK", "Step 1");
				if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(sInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) && sInfo.Tables[i].RowsValues.Count > 500)
				{
					BindDataLarge(sView, sInfo, i, null);
				}
				else
				{
					for (int j = 0; j < sInfo.Tables[i].RowsValues.Count; j++)
					{
						if (sInfo.Tables[i].Location.DataLocationTemplateRowStart <= 0)
						{
							continue;
						}
						if (sInfo.Tables[i].RowsValues[j].RowDisplay == "1")
						{
							sView.Rows[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1].Visible = true;
						}
						else
						{
							sView.Rows[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1].Visible = false;
						}
						for (int k = 0; k < sInfo.Tables[i].RowsInfo.Cells.Count; k++)
						{
							fpSpBindDataCell(sView, sInfo.Tables[i].Location.DataLocationTemplateRowStart + j, sInfo.Tables[i].RowsInfo.Cells[k].CellTemplate.ColIndex, sInfo.Tables[i].RowsInfo.Cells[k].Controltype, sInfo.Tables[i].RowsValues[j].RowValue[k]);
						}
						if ((WinSettings.MenuID == "01" || WinSettings.MenuID == "I5") && sInfo.SheetName == "PL 01-2/TĐ-GTGT" && !HTKKFormater.isNullString(sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, sInfo.Tables[i].iColumnSTT - 1].Text))
						{
							continue;
						}
						_ = sInfo.Tables[i].iColumnSTT;
						if (!(WinSettings.MenuID != "30"))
						{
							continue;
						}
						if (sInfo.Tables[i].CellTableSTT != null)
						{
							if (sInfo.Tables[i].CellTableSTT.Value == "abc")
							{
								string[] array = "a~b~c~d~đ~e~g~h~i~k~l~m~n~o~p~q~r~s~t~u~v~x~y".Split(new char[1] { '~' }).ToArray();
								sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, sInfo.Tables[i].iColumnSTT - 1].ResetFormula();
								sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, sInfo.Tables[i].iColumnSTT - 1].Text = array[(j >= 22) ? 22 : j];
							}
							else
							{
								sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, sInfo.Tables[i].iColumnSTT - 1].Text = sInfo.Tables[i].CellTableSTT.Value + "." + (j + 1);
							}
						}
						else
						{
							sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, sInfo.Tables[i].iColumnSTT - 1].Text = (j + 1).ToString();
						}
					}
				}
				for (int j = 0; j < sInfo.Tables[i].Footer.StaticCells.Count; j++)
				{
					fpSpBindDataCell(sView, sInfo.Tables[i].Footer.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[i].Footer.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[i].Footer.StaticCells[j].Controltype, sInfo.Tables[i].Footer.StaticCells[j].Value);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void BindDataLargeReport(SheetView sView, HTKKCMCSheetInfo sInfo, int i, List<HTKKCMCRowValue> rowValues)
	{
		try
		{
			sView.FpSpread.SuspendLayout();
			sView.AutoCalculation = false;
			int num = sInfo.Tables[i].RowsInfo.Cells[0].CellTemplate.ColIndex - sInfo.Tables[i].iColumnSTT;
			HTKKCMCRowInfo rowsInfo = sInfo.Tables[i].RowsInfo;
			int num2 = 0;
			List<HTKKCMCRowValue> list = null;
			int num3 = 0;
			if (rowValues != null)
			{
				num2 = sInfo.Tables[i].RowsValues.Count;
				list = rowValues;
				num3 = list.Count;
			}
			else
			{
				list = sInfo.Tables[i].RowsValues;
				num3 = list.Count;
			}
			int num4 = 0;
			GC.Collect();
			GC.WaitForPendingFinalizers();
			int num5 = 0;
			int nonEmptyColumnCount = sView.NonEmptyColumnCount;
			object[,] array = new object[num3, nonEmptyColumnCount];
			for (int j = 0; j < num3; j++)
			{
				List<string> rowValue = list[num4 + j].RowValue;
				int count = rowValue.Count;
				int num6 = 0;
				int indexj;
				for (indexj = num5; indexj < nonEmptyColumnCount; indexj++)
				{
					if (sInfo.Tables[i].RowsInfo.Cells.Count((HTKKCMCTableCellInfo p) => p.CellReport.ColIndex == indexj) <= 0)
					{
						continue;
					}
					HTKKCMCCellBasicInfo cellReport = sInfo.Tables[i].RowsInfo.Cells[num6].CellReport;
					if (cellReport.ColIndex > 0)
					{
						if (rowsInfo.Cells[num6].Controltype == "16")
						{
							array[j, cellReport.ColIndex - 1] = ConvertHelpers.ConvertToDouble(rowValue[num6]);
						}
						else if (rowsInfo.Cells[num6].Controltype == "2")
						{
							if (rowValue[num6].ToUpper() == "TRUE" || rowValue[num6].ToUpper() == "1" || rowValue[num6].ToUpper() == "X")
							{
								array[j, cellReport.ColIndex - 1] = "x";
							}
							else
							{
								array[j, cellReport.ColIndex - 1] = "";
							}
						}
						else
						{
							array[j, cellReport.ColIndex - 1] = rowValue[num6];
						}
					}
					num6++;
				}
				array[j, 0] = (num4 + j + 1).ToString();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
			sView.SetArray(sInfo.Tables[i].Location.DataLocationReportRowStart - 1, 0, array);
			array = null;
			HTKKFile.LogRun("TaiDK", "Out loop");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		finally
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			sView.AutoCalculation = true;
		}
	}

	public static void BindDataLarge(SheetView sView, HTKKCMCSheetInfo sInfo, int i, List<HTKKCMCRowValue> rowValues)
	{
		try
		{
			sView.FpSpread.SuspendLayout();
			sView.AutoCalculation = false;
			int num = sInfo.Tables[i].RowsInfo.Cells[0].CellTemplate.ColIndex - sInfo.Tables[i].iColumnSTT;
			HTKKCMCRowInfo rowsInfo = sInfo.Tables[i].RowsInfo;
			int num2 = 0;
			List<HTKKCMCRowValue> list = null;
			int num3 = 0;
			if (rowValues != null)
			{
				num2 = sInfo.Tables[i].RowsValues.Count;
				list = rowValues;
				num3 = list.Count;
			}
			else
			{
				list = sInfo.Tables[i].RowsValues;
				num3 = list.Count;
			}
			int num4 = 0;
			int num5 = ((num3 % 10000 > 0) ? (num3 / 10000 + 1) : (num3 / 10000));
			if (num5 == 0)
			{
				num5 = 1;
			}
			for (int j = 0; j < num5; j++)
			{
				HTKKFile.LogRun("TaiDK", "Start loop " + j + " - RowsCount: " + 10000);
				GC.Collect();
				GC.WaitForPendingFinalizers();
				int num6 = (j + 1) * 10000;
				int num7 = 0;
				if (num6 > num3)
				{
					num6 = num3;
				}
				num4 = j * 10000;
				num7 = num6 - num4;
				object[,] array = new object[num7, sInfo.Tables[i].RowsInfo.Cells.Count + num];
				for (int k = 0; k < num7; k++)
				{
					List<string> rowValue = list[num4 + k].RowValue;
					int num8 = rowValue.Count;
					if (sInfo.SheetName == "05-1BK-QTT-TNCN")
					{
						num8 = rowValue.Count - 1;
					}
					for (int l = 0; l < num8; l++)
					{
						if (rowsInfo.Cells[l].Controltype == "16")
						{
							array[k, l + num] = ConvertHelpers.ConvertToDouble(rowValue[l]);
						}
						else if (rowsInfo.Cells[l].Controltype == "2")
						{
							if (rowValue[l].ToUpper() == "TRUE" || rowValue[l].ToUpper() == "1" || rowValue[l].ToUpper() == "X")
							{
								array[k, l + num] = true;
							}
							else
							{
								array[k, l + num] = false;
							}
						}
						else
						{
							array[k, l + num] = rowValue[l];
						}
					}
					array[k, 0] = (num4 + k + 1).ToString();
				}
				sView.SetArray(sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, sInfo.Tables[i].RowsInfo.Cells[0].CellTemplate.ColIndex - 1 - num, array);
				HTKKFile.LogRun("SetArray", "Process loop " + j);
				sView.Recalculate();
				GC.Collect();
				GC.WaitForPendingFinalizers();
				array = null;
				HTKKFile.LogRun("AutoCaculation", "End loop " + j);
			}
			HTKKFile.LogRun("TaiDK", "Out loop");
			if (!(WinSettings.MenuID == "31") && !(WinSettings.MenuID == "33"))
			{
				return;
			}
			for (int m = sInfo.Tables[i].Location.DataLocationTemplateRowStart; m < sInfo.Tables[i].Location.DataLocationTemplateRowStart + num3; m++)
			{
				int num9 = 0;
				num9 = HTKKFile.ColumnLettersToIndex("D") - 1;
				string text = sView.Cells[m - 1, num9].Text;
				if (HTKKFormater.isNullString(text))
				{
					continue;
				}
				if (WinSettings.MenuID == "33")
				{
					string text2 = sView.Cells[m - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text;
					if (!HTKKFormater.isNullString(text2) && text2.Length > text.Length)
					{
						num9 = HTKKFile.ColumnLettersToIndex("C") - 1;
						text = text2;
					}
					text2 = sView.Cells[m - 1, HTKKFile.ColumnLettersToIndex("F") - 1].Text;
					if (!HTKKFormater.isNullString(text2) && text2.Length > text.Length)
					{
						num9 = HTKKFile.ColumnLettersToIndex("F") - 1;
						text = text2;
					}
				}
				Row row = sView.Rows[m - 1];
				float preferredHeight = row.GetPreferredHeight();
				if (text.Length > 20)
				{
					float num10 = preferredHeight / row.Height;
					preferredHeight += num10 * 5f;
					row.Height = preferredHeight;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		finally
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			sView.AutoCalculation = true;
		}
	}

	public static void BindDataLargeExport(SheetView sView, HTKKCMCSheetInfo sInfo, int i, List<HTKKCMCRowValue> rowValues)
	{
		try
		{
			int num = sInfo.Tables[i].RowsInfo.Cells[0].CellTemplate.ColIndex - sInfo.Tables[i].iColumnSTT;
			HTKKCMCRowInfo rowsInfo = sInfo.Tables[i].RowsInfo;
			int num2 = 0;
			List<HTKKCMCRowValue> list = null;
			int num3 = 0;
			if (rowValues != null)
			{
				num2 = sInfo.Tables[i].RowsValues.Count;
				list = rowValues;
				num3 = list.Count;
			}
			else
			{
				list = sInfo.Tables[i].RowsValues;
				num3 = list.Count;
			}
			int num4 = 0;
			int num5 = ((num3 % 10000 > 0) ? (num3 / 10000 + 1) : (num3 / 10000));
			if (num5 == 0)
			{
				num5 = 1;
			}
			for (int j = 0; j < num5; j++)
			{
				GC.Collect();
				GC.WaitForPendingFinalizers();
				int num6 = (j + 1) * 10000;
				int num7 = 0;
				if (num6 > num3)
				{
					num6 = num3;
				}
				num4 = j * 10000;
				num7 = num6 - num4;
				for (int k = 0; k < rowsInfo.Cells.Count; k++)
				{
					if (!(rowsInfo.Cells[k].Controltype == "2") && !(rowsInfo.Cells[k].Controltype == "14") && !(rowsInfo.Cells[k].Controltype == "19") && !(rowsInfo.Cells[k].Controltype == "96") && !(rowsInfo.Cells[k].Controltype == "94"))
					{
						continue;
					}
					object[,] array = new object[num7, 1];
					for (int l = 0; l < num7; l++)
					{
						List<string> rowValue = list[num4 + l].RowValue;
						switch (rowsInfo.Cells[k].Controltype)
						{
						case "2":
							if (rowValue[k].ToUpper() == "TRUE" || rowValue[k] == "1" || rowValue[k].ToUpper() == "X")
							{
								array[l, 0] = "x";
							}
							else
							{
								array[l, 0] = "";
							}
							break;
						case "14":
							array[l, 0] = HTKKFormater.toDateDDMMYYYY(rowValue[k], isMess: false);
							break;
						case "96":
							array[l, 0] = HTKKFormater.toTimeHHMMSS(rowValue[k], isMess: false);
							break;
						case "94":
							array[l, 0] = HTKKFormater.toDateAndTimeDDMMYYYYHHMMSS(rowValue[k], isMess: false);
							break;
						case "19":
							array[l, 0] = HTKKFormater.toDateMMYYYY(rowValue[k], isMess: false);
							break;
						}
					}
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowsInfo.Cells[k];
					Cell cell = sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1];
					Cell cell2 = sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1, sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num6 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1];
					switch (rowsInfo.Cells[k].Controltype)
					{
					case "2":
					case "14":
					case "96":
					case "94":
					case "19":
						cell2.CellType = new TextCellType();
						sView.SetArray(sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, sInfo.Tables[i].RowsInfo.Cells[k].CellTemplate.ColIndex - 1, array);
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void BindDataLargeAfterExport(SheetView sView, HTKKCMCSheetInfo sInfo, int i, List<HTKKCMCRowValue> rowValues)
	{
		try
		{
			int num = sInfo.Tables[i].RowsInfo.Cells[0].CellTemplate.ColIndex - sInfo.Tables[i].iColumnSTT;
			HTKKCMCRowInfo rowsInfo = sInfo.Tables[i].RowsInfo;
			int num2 = 0;
			List<HTKKCMCRowValue> list = null;
			int num3 = 0;
			if (rowValues != null)
			{
				num2 = sInfo.Tables[i].RowsValues.Count;
				list = rowValues;
				num3 = list.Count;
			}
			else
			{
				list = sInfo.Tables[i].RowsValues;
				num3 = list.Count;
			}
			int num4 = 0;
			int num5 = ((num3 % 10000 > 0) ? (num3 / 10000 + 1) : (num3 / 10000));
			if (num5 == 0)
			{
				num5 = 1;
			}
			for (int j = 0; j < num5; j++)
			{
				GC.Collect();
				GC.WaitForPendingFinalizers();
				int num6 = (j + 1) * 10000;
				int num7 = 0;
				if (num6 > num3)
				{
					num6 = num3;
				}
				num4 = j * 10000;
				num7 = num6 - num4;
				for (int k = 0; k < rowsInfo.Cells.Count; k++)
				{
					if (!(rowsInfo.Cells[k].Controltype == "2") && !(rowsInfo.Cells[k].Controltype == "14") && !(rowsInfo.Cells[k].Controltype == "19") && !(rowsInfo.Cells[k].Controltype == "96") && !(rowsInfo.Cells[k].Controltype == "94"))
					{
						continue;
					}
					object[,] array = new object[num7, 1];
					for (int l = 0; l < num7; l++)
					{
						List<string> rowValue = list[num4 + l].RowValue;
						switch (rowsInfo.Cells[k].Controltype)
						{
						case "2":
							if (rowValue[k].ToUpper() == "TRUE" || rowValue[k] == "1" || rowValue[k].ToUpper() == "X")
							{
								array[l, 0] = "True";
							}
							else
							{
								array[l, 0] = "False";
							}
							break;
						case "14":
							array[l, 0] = HTKKFormater.toDateDDMMYYYY(rowValue[k], isMess: false);
							break;
						case "96":
							array[l, 0] = HTKKFormater.toTimeHHMMSS(rowValue[k], isMess: false);
							break;
						case "94":
							array[l, 0] = HTKKFormater.toDateAndTimeDDMMYYYYHHMMSS(rowValue[k], isMess: false);
							break;
						case "19":
							array[l, 0] = HTKKFormater.toDateMMYYYY(rowValue[k], isMess: false);
							break;
						}
					}
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowsInfo.Cells[k];
					Cell cell = sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1];
					Cell cell2 = sView.Cells[sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1, sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num6 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1];
					switch (rowsInfo.Cells[k].Controltype)
					{
					case "2":
						cell2.CellType = new CheckBoxCellType();
						break;
					case "14":
					{
						MaskCellType maskCellType5 = new MaskCellType();
						maskCellType5.Mask = "00/00/0000";
						maskCellType5.MaskChar = '.';
						cell2.CellType = maskCellType5;
						break;
					}
					case "96":
					{
						MaskCellType maskCellType4 = new MaskCellType();
						maskCellType4.Mask = "00:00:00";
						maskCellType4.MaskChar = '.';
						cell2.CellType = maskCellType4;
						break;
					}
					case "94":
					{
						MaskCellType maskCellType3 = new MaskCellType();
						maskCellType3.Mask = "00/00/0000 00:00:00";
						maskCellType3.MaskChar = '.';
						cell2.CellType = maskCellType3;
						break;
					}
					case "19":
					{
						MaskCellType maskCellType2 = new MaskCellType();
						maskCellType2.Mask = "00/0000";
						maskCellType2.MaskChar = '.';
						cell2.CellType = maskCellType2;
						break;
					}
					case "109":
					{
						MaskCellType maskCellType = new MaskCellType();
						maskCellType.Mask = "0000";
						maskCellType.MaskChar = '.';
						cell2.CellType = maskCellType;
						break;
					}
					}
					sView.SetArray(sInfo.Tables[i].Location.DataLocationTemplateRowStart + num2 + num4 - 1, sInfo.Tables[i].RowsInfo.Cells[k].CellTemplate.ColIndex - 1, array);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private static void SetArrayToSpread(SheetView sView, int row, int col, object[,] arrValue)
	{
		int length = arrValue.GetLength(0);
		if (length <= 10000)
		{
			return;
		}
		int num = ((length % 10000 > 0) ? (length / 10000 + 1) : (length / 10000));
		for (int i = 0; i < num; i++)
		{
			int num2 = (i + 1) * 10000;
			if (num2 > length)
			{
				num2 = length;
			}
			row += num2 - i * 10000;
		}
	}

	public static void fpSpReindexingSheetTemplate(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					sView.Cells[sInfo.Tables[i].CellTableSTT.RowIndex - 1, sInfo.Tables[i].CellTableSTT.ColIndex - 1].Value = sInfo.Tables[i].TableID;
					sInfo.Tables[i].CellTableSTT.Value = sInfo.Tables[i].TableID;
					fpSpReindexingTableTemplate(sView, sInfo.Tables[i]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpReindexingSheetReport(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i].CellTableReportSTT != null)
				{
					sView.Cells[sInfo.Tables[i].CellTableReportSTT.RowIndex - 1, sInfo.Tables[i].CellTableReportSTT.ColIndex - 1].Value = sInfo.Tables[i].TableID;
					sInfo.Tables[i].CellTableReportSTT.Value = sInfo.Tables[i].TableID;
					fpSpReindexingTableReport(sView, sInfo.Tables[i]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpReindexingTableTemplate(SheetView sView, HTKKCMCTableInfo tableInfo)
	{
		try
		{
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				if (tableInfo.iColumnSTT < 0)
				{
					continue;
				}
				if (tableInfo.CellTableSTT != null)
				{
					if (tableInfo.CellTableSTT.Value == "abc")
					{
						string[] array = "a~b~c~d~đ~e~g~h~i~k~l~m~n~o~p~q~r~s~t~u~v~x~y".Split(new char[1] { '~' }).ToArray();
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].ResetFormula();
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = array[(i >= 22) ? 22 : i];
					}
					else
					{
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].ResetFormula();
						sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = tableInfo.CellTableSTT.Value + "." + (i + 1);
					}
				}
				else
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].ResetFormula();
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = (i + 1).ToString();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpQuickReindexingTableTemplate(SheetView sView, HTKKCMCTableInfo tableInfo, int count)
	{
		try
		{
			object[,] array = new object[count, 1];
			for (int i = 0; i < count; i++)
			{
				array[i, 0] = i + 1;
			}
			sView.Cells[tableInfo.Location.DataLocationTemplateRowStart - 1, tableInfo.iColumnSTT - 1, tableInfo.Location.DataLocationTemplateRowStart + count - 1, tableInfo.iColumnSTT - 1].Formula = "";
			sView.SetArray(tableInfo.Location.DataLocationTemplateRowStart - 1, tableInfo.iColumnSTT - 1, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpReindexingTableReport(SheetView sView, HTKKCMCTableInfo tableInfo)
	{
		try
		{
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				if (tableInfo.CellTableReportSTT != null)
				{
					sView.Cells[tableInfo.Location.DataLocationReportRowStart + i - 1, tableInfo.iColumnSTTReport - 1].Text = tableInfo.CellTableReportSTT.Value + "." + (i + 1);
				}
				else
				{
					sView.Cells[tableInfo.Location.DataLocationReportRowStart + i - 1, tableInfo.iColumnSTTReport - 1].Text = (i + 1).ToString();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpBindDataReport(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < sInfo.StaticCells.Count; i++)
			{
				if (sInfo.StaticCells[i].CellReport != null)
				{
					WinDoWorkEvent.SetCellExport(sView, sInfo.StaticCells[i].CellReport.RowIndex, sInfo.StaticCells[i].CellReport.ColIndex, sInfo.StaticCells[i].Controltype, sInfo.StaticCells[i].Value, sInfo.StaticCells[i].DefaultValue, sInfo.SheetIndex);
				}
			}
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				int num = i;
				if (sInfo.Tables[i].Header.StaticCells.Count > 0)
				{
					for (int j = 0; j < sInfo.Tables[i].Header.StaticCells.Count; j++)
					{
						int num2 = j;
						WinDoWorkEvent.SetCellExport(sView, sInfo.Tables[i].Header.StaticCells[j].CellReport.RowIndex, sInfo.Tables[i].Header.StaticCells[j].CellReport.ColIndex, sInfo.Tables[i].Header.StaticCells[j].Controltype, sInfo.Tables[i].Header.StaticCells[j].Value, sInfo.Tables[i].Header.StaticCells[j].DefaultValue, sInfo.SheetIndex);
					}
					if (sInfo.Tables[i].iColumnSTTReport > 0 && sInfo.Tables[i].Location.DataLocationReportRowStart - 1 >= 0)
					{
						sView.Cells[sInfo.Tables[i].Location.DataLocationReportRowStart - 1, sInfo.Tables[i].iColumnSTTReport - 1].Value = (i + 1).ToString();
					}
				}
				if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(sInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) && sInfo.Tables[i].RowsValues.Count > 500)
				{
					sView.FpSpread.SuspendLayout();
					sView.AutoCalculation = false;
					BindDataLargeReport(sView, sInfo, i, null);
					sView.AutoCalculation = true;
					sView.FpSpread.ResumeLayout();
				}
				else
				{
					for (int j = 0; j < sInfo.Tables[i].RowsValues.Count; j++)
					{
						num = j;
						if (sInfo.Tables[i].RowsValues[j].RowDisplay == "1")
						{
							sView.Rows[sInfo.Tables[i].Location.DataLocationReportRowStart + j - 1].Visible = true;
						}
						else
						{
							sView.Rows[sInfo.Tables[i].Location.DataLocationReportRowStart + j - 1].Visible = false;
						}
						for (int k = 0; k < sInfo.Tables[i].RowsInfo.Cells.Count; k++)
						{
							int num2 = k;
							if (ConvertHelpers.ConvertToDouble(sInfo.Tables[i].RowsInfo.Cells[k].CellReport.ColIndex) > 0.0)
							{
								WinDoWorkEvent.SetCellExport(sView, sInfo.Tables[i].Location.DataLocationReportRowStart + j, sInfo.Tables[i].RowsInfo.Cells[k].CellReport.ColIndex, sInfo.Tables[i].RowsInfo.Cells[k].Controltype, sInfo.Tables[i].RowsValues[j].RowValue[k], sInfo.Tables[i].RowsInfo.Cells[k].DefaultValue, sInfo.SheetIndex, sInfo.Tables[i]);
							}
						}
						if (WinSettings.MenuID.Equals("30") || (WinSettings.MenuID.Equals("03") && sInfo.DataFile == "03_9_TD_TNDN") || WinSettings.MenuID.Equals("J3"))
						{
							continue;
						}
						_ = sInfo.Tables[i].iColumnSTTReport;
						if (sInfo.Tables[i].iColumnSTTReport > -1)
						{
							if (sInfo.Tables[i].CellTableReportSTT != null)
							{
								sView.Cells[sInfo.Tables[i].Location.DataLocationReportRowStart + j - 1, sInfo.Tables[i].iColumnSTTReport - 1].Text = sInfo.Tables[i].CellTableReportSTT.Value + "." + (j + 1);
							}
							else
							{
								sView.Cells[sInfo.Tables[i].Location.DataLocationReportRowStart + j - 1, sInfo.Tables[i].iColumnSTTReport - 1].Value = (j + 1).ToString();
							}
						}
					}
				}
				for (int j = 0; j < sInfo.Tables[i].Footer.StaticCells.Count; j++)
				{
					WinDoWorkEvent.SetCellExport(sView, sInfo.Tables[i].Footer.StaticCells[j].CellReport.RowIndex, sInfo.Tables[i].Footer.StaticCells[j].CellReport.ColIndex, sInfo.Tables[i].Footer.StaticCells[j].Controltype, sInfo.Tables[i].Footer.StaticCells[j].Value, sInfo.Tables[i].Footer.StaticCells[j].DefaultValue, sInfo.SheetIndex, sInfo.Tables[i]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static bool BindDataToPhuLucSpread(SheetView sview, XmlNodeList elemList, CheckBox chkChonPhuLucKeKhaiAll, FpSpread winFpSpread = null)
	{
		//IL_0787: Unknown result type (might be due to invalid IL or missing references)
		//IL_082f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0961: Unknown result type (might be due to invalid IL or missing references)
		//IL_08f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_1415: Unknown result type (might be due to invalid IL or missing references)
		//IL_14ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_2415: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (elemList.Count <= 2)
			{
				return true;
			}
			int num = 0;
			int num2 = 0;
			for (int i = 1; i < elemList.Count; i++)
			{
				string text = ((elemList[i].Attributes["SheetName"] == null) ? null : ((XmlNode)elemList[i].Attributes["SheetName"]).Value);
				if (text.Contains("KHBS") || text.Contains("01-KHBS") || text.Contains("01-1/KHBS"))
				{
					continue;
				}
				CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
				checkBoxCellType.TextFalse = text;
				checkBoxCellType.TextTrue = text;
				sview.Cells.Get(i - 1, 0).CellType = checkBoxCellType;
				if (WinSettings.DicSheet.ContainsKey(text))
				{
					sview.Cells[i - 1, 0].Value = true;
					sview.Cells[i - 1, 0].Locked = true;
					num2++;
				}
				else
				{
					sview.Cells[i - 1, 0].Value = false;
					sview.Cells[i - 1, 0].Locked = false;
					num2--;
					if ("A1,H1,A3,H3,A8,H8,A2,A5,A6,A7,A9,H9,53,54,58,78,69,19,20,22,B3,B4,B5,C1,C2,C3,C4,C5,P8,P9,Q1,Q2,Q3,U3,U4,X4,X3,X2,X1,U9".Contains(WinSettings.MenuID) && WinSettings.DicSheet.Count <= 1)
					{
						sview.Cells[i - 1, 0].Value = true;
						sview.Cells[i - 1, 0].Locked = false;
						num2++;
					}
					else if (WinSettings.MenuID == "79")
					{
						if (WinSettings.TTPeriodInfo.KieuKy == "D")
						{
							sview.Cells[i - 1, 0].Value = false;
							sview.Cells[i - 1, 0].Locked = true;
							num2--;
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "1761")
							{
								if (i - 1 == 1)
								{
									sview.Cells[1, 0].Value = false;
									sview.Cells[1, 0].Locked = true;
									num2--;
								}
							}
							else if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = false;
							}
						}
					}
					else if (WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5")
					{
						if (WinSettings.TTPeriodInfo.KieuKy == "D")
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = true;
							sview.Cells[2, 0].Value = false;
							sview.Cells[2, 0].Locked = true;
							num2--;
						}
						if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "01")
						{
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = true;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Value = false;
								sview.Cells[2, 0].Locked = true;
							}
						}
						if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "02")
						{
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = true;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Value = false;
								sview.Cells[2, 0].Locked = true;
							}
						}
						if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "03" && i - 1 == 2)
						{
							sview.Cells[2, 0].Value = false;
							sview.Cells[2, 0].Locked = true;
						}
						if (!string.IsNullOrEmpty(WinSettings.HTKKHeaderInfo.mst) && WinSettings.TTPeriodInfo.TieuMucHachToanVal == "04")
						{
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Value = true;
								sview.Cells[2, 0].Locked = true;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = true;
							}
						}
					}
					else if (WinSettings.MenuID == "G5")
					{
						if (winFpSpread != null)
						{
							int[] array = GeneralFunctions.ArrayCellId("B_6");
							int[] array2 = GeneralFunctions.ArrayCellId("B_8");
							int[] array3 = GeneralFunctions.ArrayCellId("B_10");
							int[] array4 = GeneralFunctions.ArrayCellId("B_13");
							int[] array5 = GeneralFunctions.ArrayCellId("B_15");
							int[] array6 = GeneralFunctions.ArrayCellId("B_17");
							SheetView sheetView = winFpSpread.Sheets[0];
							bool flag = Convert.ToBoolean(sheetView.Cells[array[1], array[0]].Value);
							bool flag2 = Convert.ToBoolean(sheetView.Cells[array2[1], array2[0]].Value);
							bool flag3 = Convert.ToBoolean(sheetView.Cells[array3[1], array3[0]].Value);
							bool flag4 = Convert.ToBoolean(sheetView.Cells[array4[1], array4[0]].Value);
							bool flag5 = Convert.ToBoolean(sheetView.Cells[array5[1], array5[0]].Value);
							bool flag6 = Convert.ToBoolean(sheetView.Cells[array6[1], array6[0]].Value);
							if (flag3)
							{
								if (Convert.ToBoolean(sview.Cells[1, 0].Value))
								{
									frmMessage frmMessage = new frmMessage("0569", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									return false;
								}
								sview.Cells[0, 0].Value = true;
								sview.Cells[1, 0].Value = false;
							}
							else if (!flag && !flag2 && !flag3 && flag4 && !flag5 && !flag6)
							{
								if (Convert.ToBoolean(sview.Cells[0, 0].Value))
								{
									frmMessage frmMessage = new frmMessage("0568", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									return false;
								}
								sview.Cells[0, 0].Value = false;
								sview.Cells[1, 0].Value = true;
							}
							else if (!flag && !flag2 && !flag3 && flag4 && flag5 && !flag6)
							{
								if (Convert.ToBoolean(sview.Cells[0, 0].Value) && Convert.ToBoolean(sview.Cells[1, 0].Value))
								{
									frmMessage frmMessage = new frmMessage("0532", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									return false;
								}
							}
							else
							{
								if (Convert.ToBoolean(sview.Cells[0, 0].Value) || Convert.ToBoolean(sview.Cells[1, 0].Value))
								{
									frmMessage frmMessage = new frmMessage("0564", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									return false;
								}
								sview.Cells[0, 0].Value = false;
								sview.Cells[1, 0].Value = false;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Y")
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[1, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
							sview.Cells[1, 0].Locked = true;
						}
						else if (!WinSettings.DicSheet.ContainsKey(text))
						{
							sview.Cells[i - 1, 0].Value = false;
							sview.Cells[i - 1, 0].Locked = false;
						}
						DateTime dateTime = DateTime.Now;
						DateTime t = new DateTime(2022, 2, 1);
						DateTime t2 = new DateTime(2022, 12, 31);
						if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) && WinSettings.TTPeriodInfo.TuNgay.Length > 8)
						{
							dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
						}
						if (WinSettings.TTPeriodInfo.KieuKy == "Y" || WinSettings.TTPeriodInfo.Nam != "2022" || (WinSettings.TTPeriodInfo.KieuKy == "M" && WinSettings.TTPeriodInfo.Thang == "01" && WinSettings.TTPeriodInfo.Nam == "2022") || (WinSettings.TTPeriodInfo.KieuKy == "D" && (DateTime.Compare(dateTime, t) < 0 || DateTime.Compare(t2, dateTime) < 0)) || (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.TuThang == "01/2022" && WinSettings.TTPeriodInfo.DenThang == "01/2022"))
						{
							sview.Cells[2, 0].Locked = true;
							sview.Cells[2, 0].Value = false;
						}
						sview.Cells[3, 0].Locked = true;
						sview.Cells[3, 0].Value = false;
						if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if ((WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6))
							{
								sview.Cells[3, 0].Value = false;
								sview.Cells[3, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2023") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))))
						{
							sview.Cells[3, 0].Value = false;
							sview.Cells[3, 0].Locked = false;
							num2--;
						}
						sview.Cells[4, 0].Locked = true;
						sview.Cells[4, 0].Value = false;
						if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
							{
								sview.Cells[4, 0].Value = false;
								sview.Cells[4, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")))
						{
							sview.Cells[4, 0].Value = false;
							sview.Cells[4, 0].Locked = false;
							num2--;
						}
						num2--;
					}
					else if (WinSettings.MenuID == "H2" && WinSettings.TTPeriodInfo.KieuKy == "D")
					{
						sview.Cells[0, 0].Value = false;
						sview.Cells[0, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "K5")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "B8")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
						sview.Cells[1, 0].Value = true;
						sview.Cells[1, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "B9" || WinSettings.MenuID.Equals("W1"))
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
						sview.Cells[1, 0].Value = true;
						sview.Cells[1, 0].Locked = true;
						sview.Cells[2, 0].Value = true;
						sview.Cells[2, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "R7")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "P2")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "N7")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
						sview.Cells[1, 0].Value = true;
						sview.Cells[1, 0].Locked = true;
						sview.Cells[2, 0].Value = true;
						sview.Cells[2, 0].Locked = true;
					}
					else if (WinSettings.MenuID == "G7")
					{
						if (winFpSpread != null)
						{
							int[] array7 = GeneralFunctions.ArrayCellId("Q_10");
							int[] array8 = GeneralFunctions.ArrayCellId("Q_12");
							SheetView sheetView = winFpSpread.Sheets[0];
							bool flag7 = Convert.ToBoolean(sheetView.Cells[array7[1], array7[0]].Value);
							bool flag8 = Convert.ToBoolean(sheetView.Cells[array8[1], array8[0]].Value);
							if (flag7)
							{
								if (Convert.ToBoolean(sview.Cells[1, 0].Value))
								{
									frmMessage frmMessage = new frmMessage("0569", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									return false;
								}
								((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
								sview.Cells[0, 0].Value = true;
								sview.Cells[1, 0].Value = false;
							}
							if (flag8)
							{
								if (Convert.ToBoolean(sview.Cells[0, 0].Value))
								{
									frmMessage frmMessage = new frmMessage("0568", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									return false;
								}
								((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
								sview.Cells[0, 0].Value = false;
								sview.Cells[1, 0].Value = true;
							}
						}
						else if (!WinSettings.DicSheet.ContainsKey(text))
						{
							sview.Cells[i - 1, 0].Value = false;
							sview.Cells[i - 1, 0].Locked = false;
						}
						DateTime dateTime = DateTime.Now;
						DateTime dateTime2 = DateTime.Now;
						sview.Cells[2, 0].Locked = true;
						sview.Cells[2, 0].Value = false;
						if (WinSettings.TTPeriodInfo.KieuKy == "D_D")
						{
							DateTime t = new DateTime(2022, 2, 1);
							DateTime t2 = new DateTime(2022, 12, 31);
							if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) && WinSettings.TTPeriodInfo.TuNgay.Length > 8)
							{
								dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
							}
							if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.DenNgay) && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
							{
								dateTime2 = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
							}
							if (((DateTime.Compare(dateTime, t2) <= 0 && DateTime.Compare(t, dateTime) <= 0) || (DateTime.Compare(dateTime2, t2) <= 0 && DateTime.Compare(t, dateTime2) <= 0) || (DateTime.Compare(dateTime, t) <= 0 && DateTime.Compare(t2, dateTime2) <= 0)) && WinSettings.TTPeriodInfo.TuNgay.Length > 8 && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
							{
								sview.Cells[2, 0].Locked = false;
							}
						}
						if (WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if (WinSettings.TTPeriodInfo.Nam.Equals("2022") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 2 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
							{
								sview.Cells[2, 0].Locked = false;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2022"))
						{
							sview.Cells[2, 0].Locked = false;
						}
						sview.Cells[3, 0].Locked = true;
						sview.Cells[3, 0].Value = false;
						if (WinSettings.TTPeriodInfo.KieuKy == "D_D")
						{
							DateTime dateTime3 = new DateTime(2023, 7, 1);
							DateTime dateTime4 = new DateTime(2023, 12, 31);
							DateTime dateTime5 = new DateTime(2024, 6, 30);
							if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) && WinSettings.TTPeriodInfo.TuNgay.Length > 8)
							{
								dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
							}
							if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.DenNgay) && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
							{
								dateTime2 = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
							}
							if (((DateTime.Compare(dateTime, dateTime5) <= 0 && DateTime.Compare(dateTime3, dateTime) <= 0) || (DateTime.Compare(dateTime2, dateTime5) <= 0 && DateTime.Compare(dateTime3, dateTime2) <= 0) || (DateTime.Compare(dateTime, dateTime3) <= 0 && DateTime.Compare(dateTime5, dateTime2) <= 0)) && WinSettings.TTPeriodInfo.TuNgay.Length > 8 && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
							{
								sview.Cells[3, 0].Locked = false;
							}
						}
						if (WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if ((WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6))
							{
								sview.Cells[3, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2023") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))))
						{
							sview.Cells[3, 0].Locked = false;
							num2--;
						}
						sview.Cells[4, 0].Locked = true;
						sview.Cells[4, 0].Value = false;
						if (WinSettings.TTPeriodInfo.KieuKy == "D_D")
						{
							DateTime t3 = new DateTime(2024, 7, 1);
							DateTime dateTime6 = new DateTime(2024, 12, 31);
							DateTime dateTime7 = new DateTime(2025, 6, 30);
							DateTime t4 = new DateTime(2026, 12, 31);
							if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) && WinSettings.TTPeriodInfo.TuNgay.Length > 8)
							{
								dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
							}
							if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.DenNgay) && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
							{
								dateTime2 = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
							}
							if (((DateTime.Compare(dateTime, t3) <= 0 && DateTime.Compare(dateTime2, t4) >= 0) || (DateTime.Compare(dateTime2, t3) >= 0 && DateTime.Compare(dateTime2, t4) <= 0) || (DateTime.Compare(dateTime, t3) >= 0 && DateTime.Compare(dateTime, t4) <= 0)) && WinSettings.TTPeriodInfo.TuNgay.Length > 8 && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
							{
								sview.Cells[4, 0].Locked = false;
							}
						}
						if (WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
							{
								sview.Cells[4, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")))
						{
							sview.Cells[4, 0].Locked = false;
							num2--;
						}
						num2--;
					}
					else if (WinSettings.MenuID == "L8")
					{
						if (WinSettings.TTPeriodInfo.Quy == "4" || WinSettings.TTPeriodInfo.Thang == "12")
						{
							sview.Cells[0, 0].Locked = false;
							((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
						}
						else
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
							((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
						}
					}
					else if (WinSettings.MenuID == "73" || WinSettings.MenuID == "55")
					{
						if (!WinSettings.TTPeriodInfo.Nam.Equals("2020"))
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
							num2--;
						}
						else
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = false;
							num2--;
						}
						if (!WinSettings.TTPeriodInfo.Nam.Equals("2021"))
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = true;
							num2--;
						}
						else if (WinSettings.MenuID == "55" && !WinSettings.TTPeriodInfo.KieuKy.Equals("Y"))
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = false;
							num2--;
						}
						else if (WinSettings.MenuID == "73")
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = false;
							num2--;
						}
						else
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = true;
							num2--;
						}
					}
					else if (WinSettings.MenuID == "T5")
					{
						if (!WinSettings.TTPeriodInfo.Nam.Equals("2021"))
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
							num2--;
						}
						else if (WinSettings.TTPeriodInfo.KieuKy.Equals("Y"))
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = false;
							num2--;
						}
						else
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
							num2--;
						}
					}
					else if (WinSettings.MenuID == "28")
					{
						sview.Cells[1, 0].Value = false;
						sview.Cells[1, 0].Locked = true;
						num2--;
						if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if (WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) >= 3 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 4)
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = false;
							num2--;
						}
					}
					else if (WinSettings.MenuID == "G7")
					{
						if (WinSettings.TTPeriodInfo.KieuKy != "M" && (Convert.ToBoolean(sview.Cells[0, 0].Value) || Convert.ToBoolean(sview.Cells[1, 0].Value)))
						{
							frmMessage frmMessage = new frmMessage("0573", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
					else if (WinSettings.MenuID == "71")
					{
						sview.Cells[0, 0].Value = false;
						sview.Cells[0, 0].Locked = true;
						num2--;
						if (!"TH".Equals(WinSettings.TTPeriodInfo.ThuHo))
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if (WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 11 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
								{
									sview.Cells[0, 0].Value = false;
									sview.Cells[0, 0].Locked = false;
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) == 4)
							{
								sview.Cells[0, 0].Value = false;
								sview.Cells[0, 0].Locked = false;
								num2--;
							}
						}
					}
					else if (WinSettings.MenuID == "D4")
					{
						if (i - 1 == 0)
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = true;
						}
						if (i - 1 == 1)
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = true;
						}
						if (i - 1 == 2)
						{
							sview.Cells[2, 0].Value = false;
							sview.Cells[2, 0].Locked = true;
						}
						num2--;
						if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if (WinSettings.TTPeriodInfo.Nam.Equals("2022") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 2 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12 && i - 1 == 0)
							{
								sview.Cells[0, 0].Value = false;
								sview.Cells[0, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2022") && i - 1 == 0)
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = false;
							num2--;
						}
						if (WinSettings.TTPeriodInfo.ThuHo != "TH")
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if (((WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 6 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)) && i - 1 == 1)
								{
									sview.Cells[1, 0].Value = false;
									sview.Cells[1, 0].Locked = false;
									num2--;
								}
								if (((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")) && i - 1 == 2)
								{
									sview.Cells[2, 0].Value = false;
									sview.Cells[2, 0].Locked = false;
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "D")
							{
								if (((WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)) && i - 1 == 1)
								{
									sview.Cells[1, 0].Value = false;
									sview.Cells[1, 0].Locked = false;
									num2--;
								}
								if (((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")) && i - 1 == 2)
								{
									sview.Cells[2, 0].Value = false;
									sview.Cells[2, 0].Locked = false;
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
							{
								if (((WinSettings.TTPeriodInfo.Nam.Equals("2023") && WinSettings.TTPeriodInfo.Quy != "1") || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))) && i - 1 == 1)
								{
									sview.Cells[1, 0].Value = false;
									sview.Cells[1, 0].Locked = false;
									num2--;
								}
								if (((WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")) && i - 1 == 2)
								{
									sview.Cells[2, 0].Value = false;
									sview.Cells[2, 0].Locked = false;
									num2--;
								}
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if (((WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)) && i - 1 == 1)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = false;
								num2--;
							}
							if (((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")) && i - 1 == 2)
							{
								sview.Cells[2, 0].Value = false;
								sview.Cells[2, 0].Locked = false;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
						{
							if (((WinSettings.TTPeriodInfo.Nam.Equals("2023") && WinSettings.TTPeriodInfo.Quy != "1" && WinSettings.TTPeriodInfo.Quy != "2") || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))) && i - 1 == 1)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = false;
								num2--;
							}
							if (((WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")) && i - 1 == 2)
							{
								sview.Cells[2, 0].Value = false;
								sview.Cells[2, 0].Locked = false;
								num2--;
							}
						}
					}
					else if (WinSettings.MenuID == "01" || WinSettings.MenuID == "I5")
					{
						if (i - 1 == 5)
						{
							sview.Cells[5, 0].Value = false;
							sview.Cells[5, 0].Locked = true;
							num2--;
						}
						if ("07".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "1701".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if (WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 11 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12 && i - 1 == 5)
								{
									sview.Cells[5, 0].Value = false;
									sview.Cells[5, 0].Locked = false;
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) == 4 && i - 1 == 5)
							{
								sview.Cells[5, 0].Value = false;
								sview.Cells[5, 0].Locked = false;
								num2--;
							}
						}
						if (WinSettings.TTPeriodInfo.Nam.Equals("2021"))
						{
							if (i >= 1 && i <= 5)
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = true;
								num2--;
							}
							if ("00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "07".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
							{
								if (i >= 1 && i <= 5)
								{
									sview.Cells[i - 1, 0].Value = false;
									sview.Cells[i - 1, 0].Locked = false;
								}
							}
							else if ("06".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) && i - 1 == 4)
							{
								sview.Cells[4, 0].Value = false;
								sview.Cells[4, 0].Locked = false;
							}
						}
						else
						{
							if (i >= 1 && i <= 5)
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = true;
								num2--;
							}
							if (("1701".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "07".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal)) && i >= 1 && i <= 5)
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = false;
							}
						}
					}
					else if (WinSettings.MenuID == "99")
					{
						string tieuMucHachToanVal = WinSettings.TTPeriodInfo.TieuMucHachToanVal;
						switch (tieuMucHachToanVal)
						{
						default:
							if (!(tieuMucHachToanVal == "0003"))
							{
								if (i - 1 == 0)
								{
									sview.Cells[0, 0].Locked = true;
									sview.Cells[0, 0].Value = false;
								}
								break;
							}
							goto case "0039";
						case "0039":
						case "0002":
						case "0010":
						case "0004":
						case "0005":
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Value = true;
								sview.Cells[0, 0].Locked = true;
							}
							break;
						}
						if (tieuMucHachToanVal == "0002")
						{
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Value = true;
								sview.Cells[1, 0].Locked = true;
							}
						}
						else if (i - 1 == 1)
						{
							sview.Cells[1, 0].Locked = true;
							sview.Cells[1, 0].Value = false;
						}
						if (tieuMucHachToanVal == "0006")
						{
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = true;
							}
						}
						else if (i - 1 == 2)
						{
							sview.Cells[2, 0].Locked = true;
							sview.Cells[2, 0].Value = false;
						}
						if (tieuMucHachToanVal == "0006")
						{
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = false;
							}
						}
						else if (i - 1 == 3)
						{
							sview.Cells[3, 0].Locked = true;
						}
						if (tieuMucHachToanVal == "0007")
						{
							if (i - 1 == 4)
							{
								sview.Cells[4, 0].Locked = true;
								sview.Cells[4, 0].Value = true;
							}
						}
						else if (i - 1 == 4)
						{
							sview.Cells[4, 0].Locked = true;
							sview.Cells[4, 0].Value = false;
						}
					}
					else if (WinSettings.MenuID == "D1")
					{
						switch (WinSettings.TTPeriodInfo.TieuMucHachToanVal)
						{
						case "00":
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = true;
								sview.Cells[1, 0].Value = false;
								num2--;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = false;
								num2--;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = false;
								num2--;
							}
							break;
						case "01":
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = true;
								sview.Cells[0, 0].Value = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = false;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = false;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = true;
								sview.Cells[3, 0].Value = false;
							}
							num2--;
							break;
						case "06":
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = true;
								sview.Cells[1, 0].Value = false;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = false;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = true;
								sview.Cells[3, 0].Value = false;
							}
							num2--;
							break;
						default:
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = true;
								sview.Cells[0, 0].Value = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = true;
								sview.Cells[1, 0].Value = false;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = false;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = true;
								sview.Cells[3, 0].Value = false;
							}
							num2--;
							break;
						}
						if (i - 1 == 4)
						{
							sview.Cells[4, 0].Value = false;
							sview.Cells[4, 0].Locked = true;
						}
						if ("00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "06".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "05".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if (WinSettings.TTPeriodInfo.Nam.Equals("2022") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 2 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
								{
									if (i - 1 == 4)
									{
										sview.Cells[4, 0].Value = false;
										sview.Cells[4, 0].Locked = false;
									}
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2022"))
							{
								if (i - 1 == 4)
								{
									sview.Cells[4, 0].Value = false;
									sview.Cells[4, 0].Locked = false;
								}
								num2--;
							}
						}
						if (i - 1 == 5)
						{
							sview.Cells[5, 0].Value = false;
							sview.Cells[5, 0].Locked = true;
						}
						if ("00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "06".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "05".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if ((WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 6 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6))
								{
									if (i - 1 == 5)
									{
										sview.Cells[5, 0].Value = false;
										sview.Cells[5, 0].Locked = false;
									}
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2023") && WinSettings.TTPeriodInfo.Quy != "1") || (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))))
							{
								if (i - 1 == 5)
								{
									sview.Cells[5, 0].Value = false;
									sview.Cells[5, 0].Locked = false;
								}
								num2--;
							}
						}
						if (i - 1 == 6)
						{
							sview.Cells[6, 0].Value = false;
							sview.Cells[6, 0].Locked = true;
						}
						if ("00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "06".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "05".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
								{
									if (i - 1 == 6)
									{
										sview.Cells[6, 0].Value = false;
										sview.Cells[6, 0].Locked = false;
									}
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")))
							{
								if (i - 1 == 6)
								{
									sview.Cells[6, 0].Value = false;
									sview.Cells[6, 0].Locked = false;
								}
								num2--;
							}
						}
					}
					else if (WinSettings.MenuID.Equals("Z5"))
					{
						sview.Cells[4, 0].Locked = true;
						string tieuMucHachToanVal = WinSettings.TTPeriodInfo.TieuMucHachToanVal;
						switch (tieuMucHachToanVal)
						{
						case "00":
							sview.Cells[4, 0].Locked = false;
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = true;
								sview.Cells[1, 0].Value = false;
								num2--;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = false;
								num2--;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = false;
								num2--;
							}
							break;
						case "01":
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = true;
								sview.Cells[0, 0].Value = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = false;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = false;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = true;
								sview.Cells[3, 0].Value = false;
							}
							num2--;
							break;
						case "06":
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = true;
								sview.Cells[1, 0].Value = false;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = false;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = true;
								sview.Cells[3, 0].Value = false;
							}
							num2--;
							break;
						default:
							if (i - 1 == 0)
							{
								sview.Cells[0, 0].Locked = true;
								sview.Cells[0, 0].Value = false;
							}
							if (i - 1 == 1)
							{
								sview.Cells[1, 0].Locked = true;
								sview.Cells[1, 0].Value = false;
							}
							if (i - 1 == 2)
							{
								sview.Cells[2, 0].Locked = true;
								sview.Cells[2, 0].Value = false;
							}
							if (i - 1 == 3)
							{
								sview.Cells[3, 0].Locked = true;
								sview.Cells[3, 0].Value = false;
							}
							num2--;
							break;
						}
						if (tieuMucHachToanVal != "00")
						{
							sview.Cells[4, 0].Value = false;
						}
						if (i - 1 == 5)
						{
							sview.Cells[5, 0].Value = false;
							sview.Cells[5, 0].Locked = true;
						}
						if ("00".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "06".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal) || "05".Equals(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
						{
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
								{
									if (i - 1 == 5)
									{
										sview.Cells[5, 0].Value = false;
										sview.Cells[5, 0].Locked = false;
									}
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ((WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4")) || WinSettings.TTPeriodInfo.Nam.Equals("2025") || WinSettings.TTPeriodInfo.Nam.Equals("2026")))
							{
								if (i - 1 == 5)
								{
									sview.Cells[5, 0].Value = false;
									sview.Cells[5, 0].Locked = false;
								}
								num2--;
							}
						}
					}
					else if (WinSettings.MenuID == "38")
					{
						sview.Cells[0, 0].Value = false;
						sview.Cells[0, 0].Locked = true;
						num2--;
						if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) >= 3 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 4)
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = false;
							num2--;
						}
					}
					else if (WinSettings.MenuID == "31")
					{
						if (i == 2)
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = true;
							num2--;
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								if (WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
								{
									sview.Cells[1, 0].Value = false;
									sview.Cells[1, 0].Locked = false;
									num2--;
								}
							}
							else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && WinSettings.TTPeriodInfo.Nam.Equals("2021") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) >= 3 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 4)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = false;
								num2--;
							}
						}
						sview.Cells[0, 0].Value = false;
						sview.Cells[0, 0].Locked = false;
						num2--;
						if (WinSettings.TTPeriodInfo.KieuKy == "M" && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) == 12)
						{
							sview.Cells[0, 0].Value = true;
							num2++;
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) == 4)
						{
							sview.Cells[0, 0].Value = true;
							num2++;
						}
					}
					else if (WinSettings.MenuID == "30")
					{
						sview.Cells[1, 0].Value = false;
						sview.Cells[1, 0].Locked = true;
						num2--;
						if (WinSettings.TTPeriodInfo.KieuKy == "Q")
						{
							DateTime dateTime8 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy");
							DateTime dateTime9 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy");
							DateTime dateTime10 = new DateTime(2021, 7, 1);
							DateTime dateTime11 = new DateTime(2021, 12, 31);
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = false;
							num2--;
							if (dateTime9 < dateTime10 || dateTime8 > dateTime11)
							{
								sview.Cells[1, 0].Value = false;
								sview.Cells[1, 0].Locked = true;
								num2--;
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Y" && WinSettings.TTPeriodInfo.Nam.Equals("2021"))
						{
							sview.Cells[1, 0].Value = false;
							sview.Cells[1, 0].Locked = false;
							num2--;
						}
					}
					else if (WinSettings.MenuID == "29")
					{
						sview.Cells[0, 0].Value = false;
						sview.Cells[0, 0].Locked = true;
						num2--;
						if (WinSettings.TTPeriodInfo.Nam.Equals("2021"))
						{
							sview.Cells[0, 0].Value = false;
							sview.Cells[0, 0].Locked = false;
							num2--;
						}
					}
					else if (WinSettings.MenuID == "03")
					{
						if (i == 24)
						{
							if (!WinSettings.TTPeriodInfo.Nam.Equals("2020"))
							{
								sview.Cells[23, 0].Value = false;
								sview.Cells[23, 0].Locked = true;
								num2--;
							}
							else
							{
								sview.Cells[23, 0].Value = false;
								sview.Cells[23, 0].Locked = false;
								num2--;
							}
						}
						else if (i == 25)
						{
							if (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam.ToString()) < 2019)
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = true;
								num2--;
							}
							else
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = false;
								num2--;
							}
						}
						else if (i >= 20 && i <= 23)
						{
							if (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam.ToString()) < 2019)
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = true;
								num2--;
							}
							else
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = false;
								num2--;
							}
						}
						else if (i >= 16 && i <= 19)
						{
							if (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam.ToString()) > 2019 || ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam.ToString()) < 2017)
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = true;
								num2--;
							}
							else
							{
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = false;
								num2--;
							}
						}
						else
						{
							switch (i)
							{
							case 12:
								if (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam.ToString()) >= 2017)
								{
									sview.Cells[11, 0].Value = false;
									sview.Cells[11, 0].Locked = true;
									num2--;
								}
								else
								{
									sview.Cells[11, 0].Value = false;
									sview.Cells[11, 0].Locked = false;
									num2--;
								}
								break;
							case 11:
								if (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam.ToString()) >= 2016)
								{
									sview.Cells[10, 0].Value = false;
									sview.Cells[10, 0].Locked = true;
									num2--;
								}
								else
								{
									sview.Cells[10, 0].Value = false;
									sview.Cells[10, 0].Locked = false;
									num2--;
								}
								break;
							default:
								sview.Cells[i - 1, 0].Value = false;
								sview.Cells[i - 1, 0].Locked = false;
								num2--;
								break;
							}
						}
					}
					else if (WinSettings.MenuID == "Q4")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
						num2 = 1;
					}
					else if (WinSettings.MenuID == "Q5")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
						sview.Cells[1, 0].Value = true;
						sview.Cells[1, 0].Locked = true;
						num2 = 2;
					}
					else if (WinSettings.MenuID == "Q6")
					{
						sview.Cells[0, 0].Value = true;
						sview.Cells[0, 0].Locked = true;
						sview.Cells[1, 0].Value = true;
						sview.Cells[1, 0].Locked = true;
						num2 = 2;
					}
					else
					{
						sview.Cells[i - 1, 0].Value = false;
						sview.Cells[i - 1, 0].Locked = false;
						num2--;
					}
				}
				num++;
			}
			sview.Rows.Count = num;
			if (num2 == num)
			{
				chkChonPhuLucKeKhaiAll.Checked = true;
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
				if ("A1,H1,A3,H3,A8,H8,A2,A5,A6,A7,A9,H9,53,54,58,78,69,19,20,22,B3,B4,B5,C1,C2,C3,C4,C5,P8,P9,Q1,Q2,Q3,U3,U4,X4,X3,X2,X1,U9".Contains(WinSettings.MenuID) && WinSettings.DicSheet.Count <= 1)
				{
					((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
				}
			}
			else if (WinSettings.TTPeriodInfo.menuID.Equals("68") || WinSettings.TTPeriodInfo.menuID.Equals("18") || WinSettings.TTPeriodInfo.menuID.Equals("Q9") || WinSettings.TTPeriodInfo.menuID == "Y7")
			{
				if (num > 1 && num2 > -2)
				{
					sview.Cells[0, 0].Locked = true;
					sview.Cells[1, 0].Locked = true;
					chkChonPhuLucKeKhaiAll.Checked = false;
					((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
				}
			}
			else
			{
				chkChonPhuLucKeKhaiAll.Checked = false;
				((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void fpSpDisplayTemplate(FpSpread winFpSpread, List<HTKKCMCSheetInfo> sheetInfos)
	{
		try
		{
			WinDoWorkEvent.PreBindData(winFpSpread);
			for (short num = 0; num < winFpSpread.Sheets.Count - 1; num++)
			{
				if (sheetInfos[num] != null)
				{
					if (WinFnProcess.IsWorking(sheetInfos[num].SheetName))
					{
						ProcessSheet(num, winFpSpread, sheetInfos);
					}
					else
					{
						winFpSpread.Sheets[num].Visible = false;
					}
				}
				else
				{
					winFpSpread.Sheets[num].Visible = false;
				}
			}
			if (winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].SheetName.ToLower() == "header")
			{
				winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Visible = false;
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("B") && WinSettings.ToKhaiInfo.KieuKHBS.Equals("") && winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].SheetName.ToLower() == "khbs")
			{
				winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].Visible = false;
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("B") && !WinSettings.ToKhaiInfo.KieuKHBS.Equals(""))
			{
				winFpSpread.Sheets[0].SheetName = "Tờ khai điều chỉnh";
				if (WinSettings.ToKhaiInfo.KieuKHBS != "3")
				{
					if (WinSettings.MenuID == "21")
					{
						if (winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("P_53")[1], GeneralFunctions.ArrayCellId("P_53")[0]].Text == WinSettings.TTinChung.mst)
						{
							WinSettings.TTPeriodInfo.SoNgayNopChamKHBS = (int)(DateTime.ParseExact(WinSettings.TTPeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(6) - DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None)).TotalDays;
						}
						else
						{
							WinSettings.TTPeriodInfo.SoNgayNopChamKHBS = 10;
						}
					}
					winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].Cells[4, HTKKFile.ColumnLettersToIndex("BG") - 1].Value = WinSettings.TTPeriodInfo.NgayLapKHBS;
					if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.Sheets.Count - 2] != null && !WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.Sheets.Count - 2].IsReadDataFile && !WinSettings.isReadDataXML)
					{
						WinDoWorkEvent.SetUpDataKhbs(winFpSpread, WinSettings.TTPeriodInfo, isClickTongHopSNNC: true);
					}
					else if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.Sheets.Count - 2].IsReadDataFile)
					{
						int num2 = GeneralFunctions.SoNgayNopCham(WinSettings.TTPeriodInfo);
						winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].Cells[4, HTKKFile.ColumnLettersToIndex("BD") - 1].Value = num2.ToString();
						HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.Sheets.Count - 2].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/PhatChamNop/soNgayNopCham"));
						winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].ResetFormula();
						winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num2;
					}
				}
				else
				{
					int index = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
					int[] array = GeneralFunctions.ArrayCellId("AA_36");
					HTKKCMCCellInfo hTKKCMCCellInfo2 = WinSettings.ToKhaiInfo.SheetInfos[index].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_1"));
					HTKKCMCCellInfo hTKKCMCCellInfo3 = WinSettings.ToKhaiInfo.SheetInfos[index].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_2"));
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[index];
					if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.NgayLapKHBS))
					{
						winFpSpread.Sheets[index].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = WinSettings.TTPeriodInfo.NgayLapKHBS;
					}
					if (WinSettings.ToKhaiInfo.SheetInfos[index] != null && !WinSettings.ToKhaiInfo.SheetInfos[index].IsReadDataFile && !WinSettings.isReadDataXML)
					{
						WinDoWorkEvent.SetUpDataKhbs(winFpSpread, WinSettings.TTPeriodInfo, isClickTongHopSNNC: true, WinSettings.HTKKHeaderInfo);
						winFpSpread.Sheets[index].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = winFpSpread.Sheets[index].Cells[array[1], array[0]].Value;
					}
					else if (WinSettings.ToKhaiInfo.SheetInfos[index].IsReadDataFile)
					{
						string text = winFpSpread.Sheets[index].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
						if (!string.IsNullOrEmpty(text))
						{
							WinSettings.TTPeriodInfo.NgayLapKHBS = text;
						}
					}
				}
			}
			WinSetupSpread(winFpSpread);
			winFpSpread.TabStripInsertTab = false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void ProcessSheet(int sheetIndex, FpSpread winFpSpread, List<HTKKCMCSheetInfo> sheetInfos)
	{
		try
		{
			WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
			WinSettings.ToKhaiInfo.TaiDuLieuLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuLon");
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			sheetView.AutoCalculation = false;
			winFpSpread.SuspendLayout();
			if (WinSettings.MenuID != "G9")
			{
				fpSpViewDisplayTemplate(sheetView, sheetInfos[sheetIndex]);
			}
			else
			{
				fpSpViewDisplayTemplateNew(sheetView, sheetInfos[sheetIndex]);
			}
			WinSetupSpreadSheet(winFpSpread.Sheets[sheetIndex]);
			fpSpBindDataTemplate(winFpSpread.Sheets[sheetIndex], sheetInfos[sheetIndex]);
			if (sheetInfos[sheetIndex].TaiDuLieuRatLon == "1")
			{
				sheetInfos[sheetIndex] = WinFnProcess.GetDataFromDatafileLarge(sheetIndex, getOnlyTemplate: false);
			}
			WinDoWorkEvent.AutoBindData(winFpSpread, sheetIndex);
			setSpread_WH(winFpSpread.Sheets[sheetIndex], sheetInfos[sheetIndex]);
			HTKKFile.LogRun("TaiDK", "Start calculation");
			if (WinSettings.MenuID == "73")
			{
				if (sheetIndex == 1)
				{
					winFpSpread.Sheets[0].Cells[72, HTKKFile.ColumnLettersToIndex("AW") - 1].Formula = "BF73";
				}
				if (sheetIndex == 2)
				{
					winFpSpread.Sheets[0].Cells[72, HTKKFile.ColumnLettersToIndex("AW") - 1].Formula = "BF82";
				}
			}
			HTKKFile.LogRun("TaiDK", "End calculation");
		}
		catch (Exception ex)
		{
			GC.Collect();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		finally
		{
			if (!((Control)winFpSpread).InvokeRequired)
			{
				winFpSpread.ResumeLayout();
				winFpSpread.Sheets[sheetIndex].AutoCalculation = true;
			}
		}
	}

	public static void fpSpDisplayReport(FpSpread winFpSpread, List<HTKKCMCSheetInfo> SheetInfos)
	{
		try
		{
			winFpSpread.SuspendLayout();
			if (WinSettings.MenuID == "57")
			{
				List<HTKKCMCSheetInfo> SheetInfos2 = SheetInfos;
				fpSpDisplayReportHoTroIn05QTT(winFpSpread, ref SheetInfos2);
				for (short num = 0; num < winFpSpread.Sheets.Count; num++)
				{
					if (SheetInfos2.Count > num && SheetInfos2[num] != null)
					{
						winFpSpread.Sheets[num].AutoCalculation = false;
						fpSpViewDisplayReport(winFpSpread.Sheets[num], SheetInfos2[num]);
						WinSetupSpreadReportSheet(winFpSpread.Sheets[num]);
						fpSpBindDataReport(winFpSpread.Sheets[num], SheetInfos2[num]);
						winFpSpread.Sheets[num].RecalculateAll();
						winFpSpread.Sheets[num].AutoCalculation = true;
					}
				}
				winFpSpread.ActiveSheetIndex = 0;
				WinSetupSpread(winFpSpread);
				WinDoWorkEvent.SetFPSpreadExp(winFpSpread, SheetInfos2);
				return;
			}
			for (short num = 0; num < winFpSpread.Sheets.Count; num++)
			{
				if (SheetInfos.Count > num && SheetInfos[num] != null)
				{
					fpSpViewDisplayReport(winFpSpread.Sheets[num], SheetInfos[num]);
					WinSetupSpreadReportSheet(winFpSpread.Sheets[num]);
					fpSpBindDataReport(winFpSpread.Sheets[num], SheetInfos[num]);
					winFpSpread.Sheets[num].RecalculateAll();
				}
			}
			winFpSpread.ActiveSheetIndex = 0;
			WinSetupSpread(winFpSpread);
			WinDoWorkEvent.SetFPSpreadExp(winFpSpread, SheetInfos);
		}
		catch (Exception ex)
		{
			winFpSpread.ResumeLayout();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void fpSpDisplayReportHoTroIn05QTT(FpSpread winFpSpread, ref List<HTKKCMCSheetInfo> SheetInfos)
	{
		try
		{
			WinSettings.DicSheet = new Dictionary<string, int>();
			int num = (ConvertHelpers.ConvertToDateTime(WinSettings.TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(WinSettings.TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(WinSettings.TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(WinSettings.TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
			int num2 = 0;
			List<HTKKCMCSheetInfo> list = SheetInfos;
			SheetInfos = new List<HTKKCMCSheetInfo>();
			HTKKCMCTableInfo hTKKCMCTableInfo = list[0].Tables[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				List<string> arrValueHtIn = hTKKCMCTableInfo.RowsValues[i].RowValue;
				if (!ConvertHelpers.ConvertToBool(arrValueHtIn[0]))
				{
					continue;
				}
				num2++;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				HTKKCMCSheetInfo hTKKCMCSheetInfo2 = new HTKKCMCSheetInfo();
				string infaceTemplatePath = WinSettings.ToKhaiInfo.InterfaceTemplateXml + "05_MAP_02_QTT_TNCN_216.xml";
				string xMLFileName = GeneralFunctions.GetXMLFileName(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst, string.Format("05_1_QTT_TNCN_{0}_{1}_{2}", WinSettings.TTPeriodInfo.Nam, WinSettings.TTPeriodInfo.TuThang.Replace("/", ""), WinSettings.TTPeriodInfo.DenThang.Replace("/", "")), "242", "xml");
				hTKKCMCSheetInfo.GetDetailSheetInfo(infaceTemplatePath);
				hTKKCMCSheetInfo2.GetDetailSheetInfo(xMLFileName);
				List<HTKKCMCRowValue> list2 = hTKKCMCSheetInfo2.Tables[0].RowsValues.Where((HTKKCMCRowValue p) => p.RowValue[2].Equals(arrValueHtIn[2])).ToList();
				if (list2 != null && list2.Count > 0)
				{
					List<string> rowValue = list2[0].RowValue;
					double num3 = ConvertHelpers.ConvertToDouble1(arrValueHtIn[6]);
					for (int num4 = 0; num4 < hTKKCMCSheetInfo.StaticCells.Count; num4++)
					{
						switch (HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.StaticCells[num4].CellReport.ColIndex) + "_" + hTKKCMCSheetInfo.StaticCells[num4].CellReport.RowIndex)
						{
						case "D_6":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[1];
							break;
						case "E_6":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[2];
							break;
						case "W_42":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[4];
							break;
						case "W_43":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[4];
							break;
						case "W_44":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[5];
							break;
						case "W_45":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[6];
							break;
						case "W_47":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[7];
							break;
						case "W_48":
							hTKKCMCSheetInfo.StaticCells[num4].Value = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToDouble1(arrValueHtIn[8]) + ConvertHelpers.ConvertToDouble(arrValueHtIn[9]) + ConvertHelpers.ConvertToDouble(arrValueHtIn[10]) + ConvertHelpers.ConvertToDouble(arrValueHtIn[11]) + ConvertHelpers.ConvertToDouble(arrValueHtIn[12]));
							break;
						case "W_49":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[8];
							break;
						case "W_50":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[9];
							break;
						case "W_51":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[10];
							break;
						case "W_52":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[11];
							break;
						case "W_53":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[12];
							break;
						case "W_54":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[13];
							break;
						case "W_55":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[14];
							break;
						case "W_56":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[15];
							break;
						case "W_57":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[15];
							break;
						case "W_62":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[16];
							break;
						case "W_64":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[17];
							break;
						case "W_66":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[18];
							break;
						case "W_67":
							hTKKCMCSheetInfo.StaticCells[num4].Value = arrValueHtIn[18];
							break;
						default:
							hTKKCMCSheetInfo.StaticCells[num4].Value = "";
							break;
						case "W_61":
						case "W_65":
							break;
						}
					}
				}
				hTKKCMCSheetInfo.SheetName = winFpSpread.Sheets[0].SheetName + num2;
				WinSettings.DicSheet.Add(hTKKCMCSheetInfo.SheetName, 1);
				SheetInfos.Add(hTKKCMCSheetInfo);
			}
			WinSettings.ToKhaiInfo.SheetInfos = SheetInfos;
			SheetView sheetView = winFpSpread.Sheets[0].Clone();
			SheetView value = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Clone();
			winFpSpread.Sheets.Count = SheetInfos.Count + 1;
			winFpSpread.Sheets[0].SheetName = SheetInfos[0].SheetName;
			int[] array = GeneralFunctions.ArrayCellId("R_13");
			for (int i = 1; i <= SheetInfos.Count; i++)
			{
				SheetView sheetView2 = winFpSpread.Sheets[0].Clone();
				if (i < SheetInfos.Count)
				{
					sheetView2.SheetName = SheetInfos[i].SheetName;
					winFpSpread.Sheets[i] = sheetView2;
					winFpSpread.Sheets[i].Cells[array[1], array[0]].Formula = "";
					winFpSpread.Sheets[i].Cells[array[1], array[0]].Text = WinSettings.TTPeriodInfo.Nam;
				}
				else
				{
					winFpSpread.Sheets[i] = value;
				}
			}
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Text = WinSettings.TTPeriodInfo.Nam;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void fpSpFormatCellTemplate(SheetView sView, HTKKCMCCellInfo cellInfo)
	{
		try
		{
			int rowIndex = cellInfo.CellTemplate.RowIndex;
			int colIndex = cellInfo.CellTemplate.ColIndex;
			string controltype = cellInfo.Controltype;
			string value = cellInfo.Value;
			string defaultValue = cellInfo.DefaultValue;
			if (cellInfo.Validate.ColorDefault == Color.Black)
			{
				cellInfo.Validate.ColorDefault = sView.Cells[rowIndex - 1, colIndex - 1].BackColor;
			}
			int num = 0;
			if (rowIndex == 141 && colIndex == HTKKFile.ColumnLettersToIndex("R"))
			{
				num = rowIndex;
			}
			if (sView.Cells[rowIndex - 1, colIndex - 1].Formula.ToUpper().IndexOf("HEADER") > -1)
			{
				cellInfo.Validate.IsReset = false;
			}
			HTKKCMCValidate validate = cellInfo.Validate;
			switch (controltype)
			{
			case "11":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.DMNganhNgheCoDoanhThuCaoNhat);
				return;
			case "8":
				GeneralFunctions.InsertCombobox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.getAllTextCqtCucMoi(), deletText: false);
				return;
			case "9":
				GeneralFunctions.InsertCombobox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.getAllTextCqtChiCucMoi(), deletText: false);
				return;
			case "23":
				WinInsertCombox(sView, rowIndex, colIndex, DMTaiNguyen());
				return;
			case "2":
				InsertCheckbox(sView, rowIndex, colIndex, value);
				return;
			case "29":
				WinInsertComboxKey(sView, rowIndex, colIndex, Functions.getDictionary("Catalogue_DM_Lydo_GH_TT16.xml"));
				return;
			case "18":
				WinInsertComboxKey(sView, rowIndex, colIndex, Functions.getDictionary("Catalogue_FirstDayQuarter.xml"));
				return;
			case "59":
				WinInsertCombox(sView, rowIndex, colIndex, GetDanhMuc("GetDMMucLePhiMonBaiCT22"));
				return;
			case "10":
				WinInsertComboxKey(sView, rowIndex, colIndex, Functions.getDictionary("Catalogue_DM_Lydo_GH_TT16.xml"));
				return;
			case "25":
				WinInsertComboxKey(sView, rowIndex, colIndex, DMLoaiTemRuou());
				return;
			case "16":
				InsertNumberCell(sView, rowIndex, colIndex, value, validate, defaultValue);
				return;
			case "33":
				InsertQYear(sView, rowIndex, colIndex, value, validate, 0);
				return;
			case "15":
				InsertButton(sView, rowIndex, colIndex, value);
				return;
			case "5":
				if (WinSettings.MenuID == "G1" || WinSettings.MenuID == "G4" || WinSettings.MenuID == "N5" || WinSettings.MenuID == "H5" || WinSettings.MenuID == "O7" || WinSettings.MenuID == "P5" || "U8~".Contains(WinSettings.MenuID))
				{
					WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.TinhTP_VN);
				}
				else
				{
					WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.TinhTP);
				}
				return;
			case "103":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.dicThueTinh.Keys.ToArray());
				return;
			case "100":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.dicTinhTPMoi.Keys.ToArray());
				return;
			case "6":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP["ALL"]);
				return;
			case "7":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen["ALL"]);
				return;
			case "101":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(""));
				return;
			case "102":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.coQuanThus.Select((CoQuanThu x) => x.Ten).ToArray());
				return;
			case "54":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.GioiTinh);
				return;
			case "37":
				WinInsertCombox(sView, rowIndex, colIndex, GetDanhMuc("GetDMNganhNgheKD"));
				return;
			case "38":
				WinInsertCombox(sView, rowIndex, colIndex, GetDanhMuc("GetDMNhomNganhNgheKD"));
				return;
			case "53":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMYKienKiemToan());
				return;
			case "56":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMQuocGia());
				return;
			case "67":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMQuocTichFull());
				return;
			case "4":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMQuocGiaVietNamKhac());
				return;
			case "61":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMNoiCap());
				return;
			case "70":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMCacNuocLangGieng());
				return;
			case "80":
				WinInsertCombox(sView, rowIndex, colIndex, DMThueSuat());
				return;
			case "36":
				WinInsertCombox(sView, rowIndex, colIndex, GetDanhMuc("GetDMThan"));
				return;
			case "71":
				WinInsertCombox(sView, rowIndex, colIndex, GetDanhMuc("GetDMXangDau"));
				return;
			case "72":
				WinInsertCombox(sView, rowIndex, colIndex, DMMo());
				return;
			case "86":
				WinInsertCombox(sView, rowIndex, colIndex, DMPHLPNG());
				return;
			case "81":
				WinInsertCombox(sView, rowIndex, colIndex, DMLoaiDat());
				return;
			case "82":
				WinInsertCombox(sView, rowIndex, colIndex, DMHangDat());
				return;
			case "87":
				WinInsertCombox(sView, rowIndex, colIndex, DMNnkdNtnn());
				return;
			case "89":
				WinInsertCombox(sView, rowIndex, colIndex, DMLoaiTienDk());
				return;
			case "74":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMLoaiDiaDiem());
				return;
			case "75":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMMienPhiMonBai());
				return;
			case "79":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMMaChiTieuKHBS());
				return;
			case "77":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMToKhaiPhuLucKHBS());
				return;
			case "78":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMTieuMucKHBS(null));
				return;
			case "57":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMPPLK());
				return;
			case "65":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMPPLKND132());
				return;
			case "66":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMTM());
				return;
			case "83":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMHangHoaGTGT());
				return;
			case "84":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMThueSuatGTGT("71"));
				return;
			case "85":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMThueSuatGTGT("01"));
				return;
			case "48":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.LoaiHoanThue);
				return;
			case "91":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMNganhCoThuNhapCap());
				return;
			case "92":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMChuong());
				return;
			case "93":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMNganHang());
				return;
			case "106":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMDynamic("DM_NganHang_Moi.xml"), "106");
				return;
			case "107":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMDynamic("DM_Kho_Bac.xml"));
				return;
			case "108":
				WinInsertCombox(sView, rowIndex, colIndex, GetDMDynamic("DM_CQHQ.xml"));
				return;
			case "98":
				WinInsertComboxEdit(sView, rowIndex, colIndex, GetDMNganHangKBNN());
				return;
			case "49":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.LyDoHoanThue);
				return;
			case "105":
				WinInsertCombox(sView, rowIndex, colIndex, HTKKWinDictionaryDiaDanh.ThueSuatHoanThue);
				return;
			case "26":
			case "0":
			{
				string text = ((cellInfo.Path != null) ? cellInfo.Path : string.Empty);
				if (!string.IsNullOrEmpty(text) && text.ToUpper().Contains("LYDOKHAC"))
				{
					InsertTextCellKHBS(sView, rowIndex, colIndex, value, validate);
				}
				else
				{
					InsertTextCell(sView, rowIndex, colIndex, value, validate);
				}
				return;
			}
			}
			if (controltype != null && controltype.Length > 3)
			{
				if (controltype.Substring(0, 3) == "DC_" && controltype.IndexOf("_M_") < 0)
				{
					WinInsertCombox(sView, rowIndex, colIndex, GetDMDungChung(controltype));
				}
				if (controltype.Substring(0, 3) == "DC_" && controltype.IndexOf("_M_") >= 0)
				{
					WinInsertComboxCheckboxEdit(sView, rowIndex, colIndex, GetDMDungChung(controltype), controltype);
				}
			}
			else
			{
				DefaultCellFomat(sView, cellInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpFormatTableCellTemplate(SheetView sView, int iRow, int iCol, string style, string values, HTKKCMCValidate validate, string defaultValue)
	{
		try
		{
			switch (style)
			{
			case "11":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.DMNganhNgheCoDoanhThuCaoNhat);
				return;
			case "8":
				GeneralFunctions.InsertCombobox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.getAllTextCqtCucMoi(), deletText: false);
				return;
			case "9":
				GeneralFunctions.InsertCombobox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.getAllTextCqtChiCucMoi(), deletText: false);
				return;
			case "4":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.QuocTich);
				return;
			case "70":
				WinInsertCombox(sView, iRow, iCol, GetDMCacNuocLangGieng());
				return;
			case "74":
				WinInsertCombox(sView, iRow, iCol, GetDMLoaiDiaDiem());
				return;
			case "75":
				WinInsertCombox(sView, iRow, iCol, GetDMMienPhiMonBai());
				return;
			case "79":
				WinInsertCombox(sView, iRow, iCol, GetDMMaChiTieuKHBS());
				return;
			case "77":
				WinInsertCombox(sView, iRow, iCol, GetDMToKhaiPhuLucKHBS());
				return;
			case "78":
				WinInsertCombox(sView, iRow, iCol, GetDMTieuMucKHBS(null));
				return;
			case "3":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.QuanHe);
				return;
			case "10":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.LyDoGiaHan);
				return;
			case "54":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.GioiTinh);
				return;
			case "55":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan);
				return;
			case "64":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.LinhVucHD);
				return;
			case "61":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.NoiCap);
				return;
			case "20":
				WinInsertCombox(sView, iRow, iCol, DMKhoangSan());
				return;
			case "62":
				WinInsertCombox(sView, iRow, iCol, DMKhoangSan());
				return;
			case "53":
				WinInsertCombox(sView, iRow, iCol, GetDMYKienKiemToan());
				return;
			case "56":
				WinInsertCombox(sView, iRow, iCol, GetDMQuocGia());
				return;
			case "67":
				WinInsertCombox(sView, iRow, iCol, GetDMQuocTichFull());
				return;
			case "57":
				WinInsertCombox(sView, iRow, iCol, GetDMPPLK());
				return;
			case "65":
				WinInsertCombox(sView, iRow, iCol, GetDMPPLKND132());
				return;
			case "66":
				WinInsertCombox(sView, iRow, iCol, GetDMTM());
				return;
			case "83":
				WinInsertCombox(sView, iRow, iCol, GetDMHangHoaGTGT());
				return;
			case "84":
				WinInsertCombox(sView, iRow, iCol, GetDMThueSuatGTGT("71"));
				return;
			case "85":
				WinInsertCombox(sView, iRow, iCol, GetDMThueSuatGTGT("01"));
				return;
			case "105":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.ThueSuatHoanThue);
				return;
			case "23":
				WinInsertCombox(sView, iRow, iCol, DMTaiNguyen());
				return;
			case "21":
			{
				Dictionary<string, string> lstDonVi = new Dictionary<string, string>();
				DMKhoangSanDonVi(ref lstDonVi);
				WinInsertCombox(sView, iRow, iCol, lstDonVi);
				return;
			}
			case "22":
				WinInsertCombox(sView, iRow, iCol, DMPhiLephi());
				return;
			case "88":
				WinInsertCombox(sView, iRow, iCol, DMPhiLephiTt63());
				return;
			case "86":
				WinInsertCombox(sView, iRow, iCol, DMPHLPNG());
				return;
			case "51":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.KyPhatSinhHoanThue);
				return;
			case "90":
				WinInsertCombox(sView, iRow, iCol, GetMaTieuMuc());
				return;
			case "91":
				WinInsertCombox(sView, iRow, iCol, GetDMNganhCoThuNhapCap());
				return;
			case "92":
				WinInsertCombox(sView, iRow, iCol, GetDMChuong());
				return;
			case "93":
				WinInsertCombox(sView, iRow, iCol, GetDMNganHang());
				return;
			case "106":
				WinInsertCombox(sView, iRow, iCol, GetDMDynamic("DM_NganHang_Moi.xml"), "106");
				return;
			case "107":
				WinInsertCombox(sView, iRow, iCol, GetDMDynamic("DM_Kho_Bac.xml"));
				return;
			case "108":
				WinInsertCombox(sView, iRow, iCol, GetDMDynamic("DM_CQHQ.xml"));
				return;
			case "98":
				WinInsertComboxEdit(sView, iRow, iCol, GetDMNganHangKBNN());
				return;
			case "50":
				WinInsertCombox(sView, iRow, iCol, HTKKWinDictionaryDiaDanh.KhoanNopNSSHoanThue.OrderBy((KeyValuePair<string, string> p) => p.Key).ToDictionary((KeyValuePair<string, string> p) => p.Key, (KeyValuePair<string, string> p) => p.Value));
				return;
			case "2":
				InsertCheckbox(sView, iRow, iCol, values);
				return;
			case "14":
				InsertDateCell(sView, iRow, iCol, values);
				return;
			case "96":
				InsertTimeCell(sView, iRow, iCol, values);
				return;
			case "94":
				InsertDateTimeCell(sView, iRow, iCol, values);
				return;
			case "19":
				InsertMonthYearCell(sView, iRow, iCol, values);
				return;
			case "109":
				InsertYearCell(sView, iRow, iCol, values);
				return;
			case "16":
				InsertNumberCell(sView, iRow, iCol, values, validate, defaultValue);
				return;
			case "33":
				InsertQYear(sView, iRow, iCol, values, validate, 0);
				return;
			case "27":
				InsertNumberCell(sView, iRow, iCol, values, validate, defaultValue);
				return;
			case "15":
				InsertButton(sView, iRow, iCol, values);
				return;
			case "26":
			case "17":
				CellTextFomat(sView, iRow, iCol, values, validate);
				return;
			case "104":
				InsertDayMonthCell(sView, iRow, iCol, values);
				return;
			}
			if (style != null && style.Length > 3)
			{
				if (style.Substring(0, 3) == "DC_" && style.IndexOf("_M_") < 0)
				{
					WinInsertCombox(sView, iRow, iCol, GetDMDungChung(style));
				}
				if (style.Substring(0, 3) == "DC_" && style.IndexOf("_M_") >= 0)
				{
					WinInsertComboxCheckboxEdit(sView, iRow, iCol, GetDMDungChung(style), style);
				}
			}
			else
			{
				DefaultTableCellFomat(sView, iRow, iCol, values, validate);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpFormatCellReport(SheetView sView, int iRow, int iCol, string style, string values)
	{
		try
		{
			switch (style)
			{
			case "2":
				InsertCheckboxReport(sView, iRow, iCol, values);
				break;
			case "14":
				InsertDateCell(sView, iRow, iCol, values);
				break;
			case "96":
				InsertTimeCell(sView, iRow, iCol, values);
				break;
			case "94":
				InsertDateTimeCell(sView, iRow, iCol, values);
				break;
			case "104":
				InsertDayMonthCell(sView, iRow, iCol, values);
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpBindDataCell(SheetView sView, int iRow, int iCol, string style, string values)
	{
		try
		{
			if (!string.IsNullOrEmpty(values))
			{
				values = values.Trim();
			}
			if (iRow == 57 && iCol == 18)
			{
				int num = 0;
			}
			if (iRow <= 0 || iCol <= 0)
			{
				return;
			}
			switch (style)
			{
			case "2":
				if (values.ToUpper() == "TRUE" || values.ToUpper() == "X" || values == "1")
				{
					sView.Cells[iRow - 1, iCol - 1].Value = true;
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = false;
				}
				break;
			case "7":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "101":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "102":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "6":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "103":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "5":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "100":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				break;
			case "14":
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				InsertDateCell(sView, iRow, iCol, values);
				break;
			case "96":
				InsertTimeCell(sView, iRow, iCol, values);
				break;
			case "94":
				InsertDateTimeCell(sView, iRow, iCol, values);
				break;
			case "33":
				if (!string.IsNullOrEmpty(values))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble(values);
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = values;
				}
				break;
			case "19":
				InsertMonthYearCell(sView, iRow, iCol, values);
				break;
			case "104":
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				InsertDayMonthCell(sView, iRow, iCol, values);
				break;
			case "109":
				InsertYearCell(sView, iRow, iCol, values);
				break;
			default:
				if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
				{
					sView.Cells[iRow - 1, iCol - 1].ResetFormula();
				}
				if (sView.Cells[iRow - 1, iCol - 1].CellType != null && (sView.Cells[iRow - 1, iCol - 1].CellType.GetType().Name == "NumberCellType" || sView.Cells[iRow - 1, iCol - 1].CellType.GetType().Name == "CurrencyCellType"))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble(values);
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Text = values;
				}
				break;
			}
			if (!string.IsNullOrEmpty(values))
			{
				Row row = sView.Rows[iRow - 1];
				float preferredHeight = row.GetPreferredHeight();
				if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
				{
					row.Height = preferredHeight;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpBindDataCellReport(SheetView sView, int iRow, int iCol, string style, string values)
	{
		try
		{
			if (iRow <= 0 || iCol <= 0)
			{
				return;
			}
			if (style != null && style == "2")
			{
				if (values.ToLower() == "true")
				{
					sView.Cells[iRow - 1, iCol - 1].Value = "x";
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = "";
				}
			}
			else
			{
				sView.Cells[iRow - 1, iCol - 1].Value = values;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void fpSpBindDataCellControl(SheetView sView, int iRow, int iCol, string style, string values)
	{
		try
		{
			sView.Cells[iRow - 1, iCol - 1].Value = values;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpFormatTableCell(SheetView sView, int iDataLocationTemplateRowStart, HTKKCMCTableCellInfo cellInfo)
	{
		try
		{
			int colIndex = cellInfo.CellTemplate.ColIndex;
			string controltype = cellInfo.Controltype;
			string value = cellInfo.CellTemplate.Value;
			string defaultValue = cellInfo.DefaultValue;
			HTKKCMCValidate validate = cellInfo.Validate;
			int num = 0;
			if ((iDataLocationTemplateRowStart == 39 || iDataLocationTemplateRowStart == 40) && colIndex == HTKKFile.ColumnLettersToIndex("BE"))
			{
				num = 1;
			}
			if (cellInfo.Validate.ColorDefault == Color.Black)
			{
				cellInfo.Validate.ColorDefault = sView.Cells[iDataLocationTemplateRowStart - 1, colIndex - 1].BackColor;
			}
			switch (controltype)
			{
			case "3":
				WinInsertComboxKey(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.QuanHe);
				break;
			case "4":
				WinInsertComboxKey(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.QuocTich);
				break;
			case "5":
				if (WinSettings.MenuID == "G1" || WinSettings.MenuID == "G4" || WinSettings.MenuID == "N5" || WinSettings.MenuID == "H5" || WinSettings.MenuID == "O7" || WinSettings.MenuID == "P5" || "U8~".Contains(WinSettings.MenuID))
				{
					WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.TinhTP_VN);
				}
				else
				{
					WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.TinhTP);
				}
				break;
			case "103":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.dicThueTinh.Keys.ToArray());
				break;
			case "100":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.dicTinhTPMoi.Keys.ToArray());
				break;
			case "6":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP["ALL"]);
				break;
			case "7":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen["ALL"]);
				break;
			case "101":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(""));
				break;
			case "102":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.coQuanThus.Select((CoQuanThu x) => x.Ten).ToArray());
				break;
			case "8":
				GeneralFunctions.InsertCombobox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.getAllTextCqtCucMoi(), deletText: false);
				break;
			case "9":
				GeneralFunctions.InsertCombobox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.getAllTextCqtChiCucMoi(), deletText: false);
				break;
			case "10":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.LyDoGiaHan);
				break;
			case "54":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.GioiTinh);
				break;
			case "55":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan);
				break;
			case "64":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.LinhVucHD);
				break;
			case "61":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.NoiCap);
				break;
			case "11":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.DMNganhNgheCoDoanhThuCaoNhat);
				break;
			case "12":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, HTKKWinDictionaryDiaDanh.PhuongPhapXacDinhGia);
				break;
			case "25":
				WinInsertComboxKey(sView, iDataLocationTemplateRowStart, colIndex, DMLoaiTemRuou());
				break;
			case "28":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMGhiChuTemRuou());
				break;
			case "73":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMLoaiTienTe());
				break;
			case "80":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMThueSuat());
				break;
			case "81":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMLoaiDat());
				break;
			case "82":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMHangDat());
				break;
			case "87":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMNnkdNtnn());
				break;
			case "89":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMLoaiTienDk());
				break;
			case "20":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMKhoangSan"));
				break;
			case "62":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMKSBVMT"));
				break;
			case "21":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMDonVi"));
				break;
			case "41":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMTTDBDonVi"));
				break;
			case "32":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMHoaDon"));
				break;
			case "76":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMHoaDonDT"));
				break;
			case "34":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetHTHoaDon"));
				break;
			case "35":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMThueBVMT"));
				break;
			case "36":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMThan"));
				break;
			case "71":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMXangDau"));
				break;
			case "37":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMNganhNgheKD"));
				break;
			case "38":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMNhomNganhNgheKD"));
				break;
			case "39":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMThueTieuThuDB"));
				break;
			case "42":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDichVuThueTieuThuDB"));
				break;
			case "40":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMBienLaiPhi"));
				break;
			case "46":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMHopDong"));
				break;
			case "68":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMHopDong1"));
				break;
			case "69":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetPhuongXaTheoCQT"));
				break;
			case "72":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMMo());
				break;
			case "86":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, DMPHLPNG());
				break;
			case "44":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMMucDich"));
				break;
			case "45":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMTrangThai"));
				break;
			case "43":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMLoaiThongTinThayDoi"));
				break;
			case "58":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMLoaiThongTinThayDoiBLP"));
				break;
			case "59":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMMucLePhiMonBaiCT22"));
				break;
			case "60":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMMucLePhiMonBaiCT23"));
				break;
			case "47":
				WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDanhMuc("GetDMGhiChu"));
				break;
			case "0":
				InsertTextCell(sView, iDataLocationTemplateRowStart, colIndex, value, validate);
				break;
			case "2":
				fpSpFormatTableCellTemplate(sView, iDataLocationTemplateRowStart, colIndex, controltype, value, validate, defaultValue);
				break;
			default:
				if (controltype != null && controltype.Length > 3)
				{
					if (controltype.Substring(0, 3) == "DC_" && controltype.IndexOf("_M_") < 0)
					{
						WinInsertCombox(sView, iDataLocationTemplateRowStart, colIndex, GetDMDungChung(controltype));
					}
					if (controltype.Substring(0, 3) == "DC_" && controltype.IndexOf("_M_") >= 0)
					{
						WinInsertComboxCheckboxEdit(sView, iDataLocationTemplateRowStart, colIndex, GetDMDungChung(controltype), controltype);
					}
				}
				else
				{
					fpSpFormatTableCellTemplate(sView, iDataLocationTemplateRowStart, colIndex, controltype, value, validate, defaultValue);
				}
				break;
			}
			Row row = sView.Rows[iDataLocationTemplateRowStart];
			float preferredHeight = row.GetPreferredHeight();
			if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
			{
				row.Height = preferredHeight;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertButton(SheetView sView, int iRow, int iCol, string strCaption)
	{
		try
		{
			ButtonCellType buttonCellType = new ButtonCellType();
			buttonCellType.TwoState = false;
			strCaption = ((!strCaption.Contains("TaiBangKe")) ? sView.Cells[iRow - 1, iCol - 1].Text : strCaption.Replace("TaiBangKe", ""));
			buttonCellType.Text = strCaption;
			buttonCellType.ShadowSize = 0;
			buttonCellType.TextColor = Color.Blue;
			sView.Cells[iRow - 1, iCol - 1].CellType = buttonCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertCheckbox(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			CheckBoxCellType cellType = new CheckBoxCellType();
			sView.Cells[iRow - 1, iCol - 1].CellType = cellType;
			sView.Cells[iRow - 1, iCol - 1].Value = value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertCheckboxReport(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			if (value == "True")
			{
				sView.Cells[iRow - 1, iCol - 1].Text = "x";
			}
			else
			{
				sView.Cells[iRow - 1, iCol - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertDateCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toDateDDMMYYYY(value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTimeCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toTimeHHMMSS(value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertDateTimeCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toDateAndTimeDDMMYYYYHHMMSS(value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000 00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00/0000 00:00:00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertDayMonthCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			if (!HTKKFormater.isNullString(value))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/00";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTextCell(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate)
	{
		try
		{
			TextCellType textCellType = new TextCellType();
			int num = (textCellType.MaxLength = GeneralFunctions.getMaxLen(validate.MaxLen));
			if (num != 7)
			{
				textCellType.WordWrap = true;
				textCellType.Multiline = true;
			}
			sView.Cells[iRow - 1, iCol - 1].CellType = textCellType;
			sView.Cells[iRow - 1, iCol - 1].Text = "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTextCellKHBS(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate)
	{
		try
		{
			TextCellType textCellType = new TextCellType();
			int maxLen = GeneralFunctions.getMaxLen(validate.MaxLen);
			textCellType.MaxLength = maxLen;
			textCellType.WordWrap = true;
			textCellType.Multiline = true;
			textCellType.ScrollBars = (ScrollBars)2;
			sView.Cells[iRow - 1, iCol - 1].CellType = textCellType;
			sView.Cells[iRow - 1, iCol - 1].Text = "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertMonthYearCell(SheetView sView, int iRow, int iCol, string value)
	{
		try
		{
			string text = HTKKFormater.toDateMMYYYY(string.IsNullOrEmpty(value) ? "" : value, isMess: false);
			if (!HTKKFormater.isNullString(text))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = text;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "00/0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertYearCell(SheetView sView, int iRow, int iCol, string dateValue)
	{
		try
		{
			if (!HTKKFormater.isNullString(dateValue))
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = dateValue;
			}
			else
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "0000";
				maskCellType.MaskChar = '.';
				sView.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void DefaultTableCellFomat(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate)
	{
		try
		{
			ICellType cellType = sView.Cells[iRow - 1, iCol - 1].CellType;
			if (cellType == null)
			{
				TextCellType textCellType = new TextCellType();
				textCellType.WordWrap = true;
				textCellType.MaxLength = GeneralFunctions.getMaxLen(validate.MaxLen);
				sView.Cells[iRow - 1, iCol - 1].CellType = textCellType;
			}
			else if (cellType.ToString() == "GeneralCellType")
			{
				GeneralCellType generalCellType = (GeneralCellType)sView.Cells[iRow - 1, iCol - 1].CellType;
				NumberCellType numberCellType = new NumberCellType();
				if (generalCellType.NumberFormat != null)
				{
					if (generalCellType.FormatString.IndexOf(',') != -1)
					{
						numberCellType.ShowSeparator = true;
					}
					else
					{
						numberCellType.ShowSeparator = false;
					}
					numberCellType.NegativeFormat = NegativeFormat.Parentheses;
					numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(validate.MaxValue);
					numberCellType.MinimumValue = GeneralFunctions.getMinNumber(validate.MinValue);
					numberCellType.DecimalPlaces = generalCellType.NumberFormat.NumberDecimalDigits;
					sView.Cells[iRow - 1, iCol - 1].CellType = numberCellType;
				}
				else if (generalCellType.DateTimeFormat != null)
				{
					DateTimeCellType cellType2 = new DateTimeCellType();
					sView.Cells[iRow - 1, iCol - 1].CellType = cellType2;
				}
			}
			else if (cellType.ToString() == "DateTimeCellType")
			{
				DateTimeCellType cellType2 = new DateTimeCellType();
				sView.Cells[iRow - 1, iCol - 1].CellType = cellType2;
			}
			else
			{
				TextCellType textCellType = new TextCellType();
				textCellType.WordWrap = true;
				textCellType.MaxLength = GeneralFunctions.getMaxLen(validate.MaxLen);
				sView.Cells[iRow - 1, iCol - 1].CellType = textCellType;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellTextFomat(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate)
	{
		try
		{
			TextCellType textCellType = new TextCellType();
			textCellType.WordWrap = true;
			textCellType.MaxLength = GeneralFunctions.getMaxLen(validate.MaxLen);
			sView.Cells[iRow - 1, iCol - 1].CellType = textCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void DefaultCellFomat(SheetView sView, HTKKCMCCellInfo cellInfo)
	{
		try
		{
			int rowIndex = cellInfo.CellTemplate.RowIndex;
			int colIndex = cellInfo.CellTemplate.ColIndex;
			if (rowIndex == 251)
			{
				int num = 0;
			}
			string controltype = cellInfo.Controltype;
			string value = cellInfo.Value;
			ICellType cellType = sView.Cells[rowIndex - 1, colIndex - 1].CellType;
			if (cellType != null)
			{
				if (cellType.ToString() == "GeneralCellType" || cellType.ToString() == "NumberCellType")
				{
					GeneralCellType generalCellType = null;
					if (cellType.ToString() == "GeneralCellType")
					{
						generalCellType = (GeneralCellType)sView.Cells[rowIndex - 1, colIndex - 1].CellType;
					}
					NumberCellType numberCellType = new NumberCellType();
					if (generalCellType != null && generalCellType.NumberFormat != null)
					{
						if (generalCellType.FormatString.IndexOf(',') != -1)
						{
							numberCellType.ShowSeparator = true;
						}
						else
						{
							numberCellType.ShowSeparator = false;
						}
						numberCellType.NegativeFormat = NegativeFormat.Parentheses;
						numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(cellInfo.Validate.MaxValue);
						numberCellType.MinimumValue = GeneralFunctions.getMinNumber(cellInfo.Validate.MinValue);
						numberCellType.DecimalPlaces = generalCellType.NumberFormat.NumberDecimalDigits;
						sView.Cells[rowIndex - 1, colIndex - 1].CellType = numberCellType;
						cellInfo.Formula = sView.Cells[rowIndex - 1, colIndex - 1].Formula;
					}
					else if (generalCellType != null && generalCellType.DateTimeFormat != null)
					{
						DateTimeCellType dateTimeCellType = new DateTimeCellType();
						dateTimeCellType.UserDefinedFormat = "dd/MM/yyyy";
						sView.Cells[rowIndex - 1, colIndex - 1].CellType = dateTimeCellType;
					}
				}
				else if (cellType.ToString() == "DateTimeCellType")
				{
					DateTimeCellType dateTimeCellType = new DateTimeCellType();
					dateTimeCellType.UserDefinedFormat = "dd/MM/yyyy";
					sView.Cells[rowIndex - 1, colIndex - 1].CellType = dateTimeCellType;
				}
				else
				{
					TextCellType textCellType = new TextCellType();
					textCellType.WordWrap = true;
					textCellType.MaxLength = GeneralFunctions.getMaxLen(cellInfo.Validate.MaxLen);
					sView.Cells[rowIndex - 1, colIndex - 1].CellType = textCellType;
				}
			}
			else
			{
				TextCellType textCellType = new TextCellType();
				textCellType.WordWrap = true;
				textCellType.MaxLength = GeneralFunctions.getMaxLen(cellInfo.Validate.MaxLen);
				sView.Cells[rowIndex - 1, colIndex - 1].CellType = textCellType;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static long getNumberByMaxLen(string MaxLen)
	{
		long num = 0L;
		return MaxLen switch
		{
			"1" => 9L, 
			"2" => 99L, 
			"3" => 999L, 
			"4" => 9999L, 
			"5" => 99999L, 
			"6" => 999999L, 
			"7" => 9999999L, 
			"8" => 99999999L, 
			"9" => 999999999L, 
			"10" => 9999999999L, 
			"11" => 99999999999L, 
			"12" => 999999999999L, 
			"13" => 9999999999999L, 
			"14" => 99999999999999L, 
			_ => long.MaxValue, 
		};
	}

	public static void InsertNumberCell(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate, string defaultValue)
	{
		try
		{
			int num = 0;
			if (iRow == 23 && iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				num = iRow;
			}
			if ((!HTKKFormater.isNullString(validate.MaxValue) && validate.MaxValue.ToNumber() > 999999999999999m) || (!HTKKFormater.isNullString(validate.MinValue) && validate.MinValue.ToNumber() < -999999999999999m))
			{
				InsertCurrencyCell(sView, iRow, iCol, value, validate, defaultValue);
				return;
			}
			int decimalPlaces = 0;
			NumberCellType numberCellType = new NumberCellType();
			numberCellType.ShowSeparator = true;
			numberCellType.NegativeFormat = NegativeFormat.Parentheses;
			if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
			{
				decimalPlaces = 0;
				numberCellType.DecimalSeparator = ",";
				numberCellType.Separator = ".";
			}
			else if (defaultValue.IndexOf(".") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
				numberCellType.DecimalSeparator = ".";
				numberCellType.Separator = ",";
			}
			else if (defaultValue.IndexOf(",") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
				numberCellType.DecimalSeparator = ",";
				numberCellType.Separator = ".";
			}
			numberCellType.MinimumValue = GeneralFunctions.getMinNumber(validate.MinValue);
			numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(validate.MaxValue);
			numberCellType.DecimalPlaces = decimalPlaces;
			sView.Cells[iRow - 1, iCol - 1].CellType = numberCellType;
			if (!HTKKFormater.isNullString(sView.Cells[iRow - 1, iCol - 1].Formula))
			{
				if (!HTKKFormater.isNullString(value))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(value).ToString();
				}
				else if (!HTKKFormater.isNullString(defaultValue))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(defaultValue).ToString();
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = "0";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertCurrencyCell(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate, string defaultValue)
	{
		try
		{
			int num = 0;
			if (iRow == 23 && iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				num = iRow;
			}
			int decimalPlaces = 0;
			CurrencyCellType currencyCellType = new CurrencyCellType();
			currencyCellType.ShowSeparator = true;
			currencyCellType.ShowCurrencySymbol = false;
			if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
			{
				decimalPlaces = 0;
				currencyCellType.DecimalSeparator = ",";
				currencyCellType.Separator = ".";
			}
			else if (defaultValue.IndexOf(".") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
				currencyCellType.DecimalSeparator = ".";
				currencyCellType.Separator = ",";
			}
			else if (defaultValue.IndexOf(",") > -1)
			{
				decimalPlaces = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
				currencyCellType.DecimalSeparator = ",";
				currencyCellType.Separator = ".";
			}
			currencyCellType.MinimumValue = GeneralFunctions.getMinNumber(validate.MinValue);
			currencyCellType.MaximumValue = GeneralFunctions.getMaxNumber(validate.MaxValue);
			currencyCellType.DecimalPlaces = decimalPlaces;
			sView.Cells[iRow - 1, iCol - 1].CellType = currencyCellType;
			if (!HTKKFormater.isNullString(sView.Cells[iRow - 1, iCol - 1].Formula))
			{
				if (!HTKKFormater.isNullString(value))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(value).ToString();
				}
				else if (!HTKKFormater.isNullString(defaultValue))
				{
					sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble1(defaultValue).ToString();
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = "0";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertQYear(SheetView sView, int iRow, int iCol, string value, HTKKCMCValidate validate, int decimalPlaces)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			numberCellType.ShowSeparator = false;
			numberCellType.NegativeFormat = NegativeFormat.Parentheses;
			numberCellType.DecimalSeparator = "";
			numberCellType.Separator = "....";
			numberCellType.MinimumValue = GeneralFunctions.getMinNumber(validate.MinValue);
			numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(validate.MaxValue);
			numberCellType.DecimalPlaces = decimalPlaces;
			sView.Cells[iRow - 1, iCol - 1].CellType = numberCellType;
			sView.Cells[iRow - 1, iCol - 1].Value = value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinInsertCombox(SheetView sView, int iRow, int iCol, Dictionary<string, string> dicValues, string controltype = "", bool deleteText = true, HTKKCMCSheetInfo sInfo = null)
	{
		try
		{
			string[] array = new string[dicValues.Count];
			string[] array2 = new string[dicValues.Count];
			int num = 0;
			foreach (KeyValuePair<string, string> dicValue in dicValues)
			{
				string[] array4;
				string[] array3 = (array4 = array);
				int num2 = num;
				nint num3 = num2;
				array3[num2] = array4[num3] + dicValue.Value;
				string[] array5 = (array4 = array2);
				int num4 = num;
				num3 = num4;
				array5[num4] = array4[num3] + dicValue.Key;
				num++;
			}
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			comboBoxCellType.MaxLength = 1000;
			comboBoxCellType.Items = array;
			comboBoxCellType.ItemData = array2;
			int autoWidth = GeneralFunctions.GetAutoWidth(array);
			comboBoxCellType.ListWidth = ((autoWidth > 1000) ? 1000 : autoWidth);
			if (!string.IsNullOrEmpty(controltype))
			{
				string[] array6 = "106~".Split(new char[1] { '~' });
				comboBoxCellType.AcceptsArrowKeys = AcceptsArrowKeys.AllArrows;
				comboBoxCellType.Editable = true;
				comboBoxCellType.ListAlignment = ListAlignment.Left;
				comboBoxCellType.ListOffset = 20;
				comboBoxCellType.ListWidth = 0;
			}
			sView.Cells[iRow - 1, iCol - 1].CellType = comboBoxCellType;
			if (!deleteText)
			{
				return;
			}
			sView.Cells[iRow - 1, iCol - 1].Text = "";
			if (sInfo == null)
			{
				return;
			}
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				HTKKCMCTableCellInfo cell = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo y) => y.Controltype == controltype);
				if (cell != null)
				{
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo y) => y.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(cell.indexColumnSelectedValue));
					if (hTKKCMCTableCellInfo != null)
					{
						sView.Cells[iRow - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text = "";
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo y) => y.CellTemplate.RowIndex == iRow && y.Controltype == controltype);
			if (hTKKCMCCellInfo != null)
			{
				sView.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Text = "";
				hTKKCMCCellInfo.CellSelectedValue.Value = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinInsertComboxEdit(SheetView sView, int iRow, int iCol, Dictionary<string, string> dicValues)
	{
		try
		{
			string[] array = new string[dicValues.Count];
			string[] array2 = new string[dicValues.Count];
			int num = 0;
			foreach (KeyValuePair<string, string> dicValue in dicValues)
			{
				string[] array4;
				string[] array3 = (array4 = array);
				int num2 = num;
				nint num3 = num2;
				array3[num2] = array4[num3] + dicValue.Value;
				string[] array5 = (array4 = array2);
				int num4 = num;
				num3 = num4;
				array5[num4] = array4[num3] + dicValue.Key;
				num++;
			}
			int autoWidth = GeneralFunctions.GetAutoWidth(array);
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			comboBoxCellType.MaxLength = 1000;
			comboBoxCellType.Editable = true;
			comboBoxCellType.Items = array;
			comboBoxCellType.ItemData = array2;
			comboBoxCellType.MaxDrop = 10;
			comboBoxCellType.ListWidth = autoWidth;
			sView.Cells[iRow - 1, iCol - 1].CellType = comboBoxCellType;
			sView.Cells[iRow - 1, iCol - 1].Text = "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinInsertComboxCheckboxEdit(SheetView sView, int iRow, int iCol, Dictionary<string, string> dicValues, string controltype = "", bool deleteText = true, HTKKCMCSheetInfo sInfo = null)
	{
		try
		{
			CheckedComboBoxCellType checkedComboBoxCellType = new CheckedComboBoxCellType();
			checkedComboBoxCellType.Options = dicValues.Values.ToArray();
			CheckedComboBoxCellType cellType = checkedComboBoxCellType;
			sView.Cells[iRow - 1, iCol - 1].CellType = cellType;
			string key = controltype + "_" + (iCol - 1);
			if (!GeneralFunctions.dicDms.ContainsKey(key))
			{
				GeneralFunctions.dicDms.Add(key, dicValues);
			}
			if (!deleteText)
			{
				return;
			}
			sView.Cells[iRow - 1, iCol - 1].Text = "";
			if (sInfo == null)
			{
				return;
			}
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				HTKKCMCTableCellInfo cell = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo y) => y.Controltype == controltype);
				if (cell != null)
				{
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo y) => y.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(cell.indexColumnSelectedValue));
					if (hTKKCMCTableCellInfo != null)
					{
						sView.Cells[iRow - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text = "";
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo y) => y.CellTemplate.RowIndex == iRow && y.Controltype == controltype);
			if (hTKKCMCCellInfo != null)
			{
				sView.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Text = "";
				hTKKCMCCellInfo.CellSelectedValue.Value = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinInsertComboxKey(SheetView sView, int iRow, int iCol, Dictionary<string, string> dicValues)
	{
		try
		{
			bool flag = false;
			if (WinSettings.ToKhaiInfo.FolderName == "HEADER")
			{
				flag = true;
			}
			string[] array;
			if (dicValues != null)
			{
				array = new string[1];
				int num = 1;
				if (flag && dicValues.Count > 4)
				{
					foreach (KeyValuePair<string, string> dicValue in dicValues)
					{
						Array.Resize(ref array, num);
						string[] array3;
						string[] array2 = (array3 = array);
						int num2 = num - 1;
						nint num3 = num2;
						array2[num2] = array3[num3] + dicValue.Key;
						num++;
					}
				}
				else
				{
					foreach (KeyValuePair<string, string> dicValue2 in dicValues)
					{
						Array.Resize(ref array, num);
						string[] array3;
						string[] array4 = (array3 = array);
						int num4 = num - 1;
						nint num3 = num4;
						array4[num4] = array3[num3] + dicValue2.Key;
						num++;
					}
				}
			}
			else
			{
				array = new string[1] { "Khác" };
			}
			TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
			taxComboBoxCellType.Items = array;
			taxComboBoxCellType.Items = array;
			taxComboBoxCellType.MaxDrop = 10;
			taxComboBoxCellType.ListWidth = 500;
			sView.Cells[iRow - 1, iCol - 1].CellType = taxComboBoxCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinInsertAutoCompleteKey(SheetView sView, int iRow, int iCol, Dictionary<string, string> dicValues)
	{
		//IL_0181: Unknown result type (might be due to invalid IL or missing references)
		//IL_0188: Expected O, but got Unknown
		bool flag = false;
		if (WinSettings.ToKhaiInfo.FolderName == "HEADER")
		{
			flag = true;
		}
		string[] array;
		if (dicValues != null)
		{
			array = new string[1];
			int num = 1;
			if (flag && dicValues.Count > 4)
			{
				foreach (KeyValuePair<string, string> dicValue in dicValues)
				{
					if (dicValue.Key.IndexOf(" - ") == 3)
					{
						Array.Resize(ref array, num);
						string[] array3;
						string[] array2 = (array3 = array);
						int num2 = num - 1;
						nint num3 = num2;
						array2[num2] = array3[num3] + dicValue.Key;
						num++;
					}
				}
			}
			else
			{
				foreach (KeyValuePair<string, string> dicValue2 in dicValues)
				{
					if (dicValue2.Key.IndexOf(" - ") != 3)
					{
						Array.Resize(ref array, num);
						string[] array3;
						string[] array4 = (array3 = array);
						int num4 = num - 1;
						nint num3 = num4;
						array4[num4] = array3[num3] + dicValue2.Key;
						num++;
					}
				}
			}
		}
		else
		{
			array = new string[1] { "Khác" };
		}
		GeneralCellType generalCellType = new GeneralCellType();
		generalCellType.AutoCompleteMode = (AutoCompleteMode)3;
		generalCellType.AutoCompleteSource = (AutoCompleteSource)64;
		AutoCompleteStringCollection val = new AutoCompleteStringCollection();
		val.AddRange(array);
		generalCellType.AutoCompleteCustomSource = val;
		generalCellType.AutoFillAutoCompleteCustomSource = true;
		generalCellType.AutoCompleteMode = (AutoCompleteMode)3;
		sView.Cells[iRow - 1, iCol - 1].CellType = generalCellType;
	}

	public static void WinInsertAutoComplete(SheetView sView, int iRow, int iCol, Dictionary<string, string> dicValues)
	{
		//IL_0184: Unknown result type (might be due to invalid IL or missing references)
		//IL_018b: Expected O, but got Unknown
		bool flag = false;
		if (WinSettings.ToKhaiInfo.FolderName == "HEADER")
		{
			flag = true;
		}
		string[] array;
		if (dicValues != null)
		{
			array = new string[1];
			int num = 1;
			if (flag && dicValues.Count > 4)
			{
				foreach (KeyValuePair<string, string> dicValue in dicValues)
				{
					if (dicValue.Key.IndexOf(" - ") > 0)
					{
						Array.Resize(ref array, num);
						string[] array3;
						string[] array2 = (array3 = array);
						int num2 = num - 1;
						nint num3 = num2;
						array2[num2] = array3[num3] + dicValue.Value;
						num++;
					}
				}
			}
			else
			{
				foreach (KeyValuePair<string, string> dicValue2 in dicValues)
				{
					if (dicValue2.Key.IndexOf(" - ") == -1)
					{
						Array.Resize(ref array, num);
						string[] array3;
						string[] array4 = (array3 = array);
						int num4 = num - 1;
						nint num3 = num4;
						array4[num4] = array3[num3] + dicValue2.Value;
						num++;
					}
				}
			}
		}
		else
		{
			array = new string[1] { "Khác" };
		}
		GeneralCellType generalCellType = new GeneralCellType();
		generalCellType.AutoCompleteMode = (AutoCompleteMode)3;
		generalCellType.AutoCompleteSource = (AutoCompleteSource)64;
		AutoCompleteStringCollection val = new AutoCompleteStringCollection();
		val.AddRange(array);
		generalCellType.AutoCompleteCustomSource = val;
		generalCellType.AutoFillAutoCompleteCustomSource = true;
		generalCellType.AutoCompleteMode = (AutoCompleteMode)3;
		sView.Cells[iRow - 1, iCol - 1].CellType = generalCellType;
	}

	public static void WinInsertCombox(SheetView sView, int iRow, int iCol, string[] Values)
	{
		try
		{
			if (iCol >= 1 && iRow >= 1)
			{
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = Values;
				comboBoxCellType.MaxDrop = 10;
				comboBoxCellType.ListWidth = 500;
				sView.Cells[iRow - 1, iCol - 1].CellType = comboBoxCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static Dictionary<string, string> DMTaiNguyen()
	{
		//IL_0349: Unknown result type (might be due to invalid IL or missing references)
		//IL_0350: Expected O, but got Unknown
		//IL_037d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0384: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = "";
			int month = 0;
			if (WinSettings.TTPeriodInfo.KieuKy == "M")
			{
				month = int.Parse(WinSettings.TTPeriodInfo.Thang);
			}
			else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
			{
				month = int.Parse(WinSettings.TTPeriodInfo.Quy) * 3 - 2;
			}
			else if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				month = 12;
			}
			int year = int.Parse(WinSettings.TTPeriodInfo.Nam);
			DateTime dtPeriod = new DateTime(year, month, 1);
			if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				string nam = WinSettings.TTPeriodInfo.Nam;
				text = (nam.Equals("2008") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_052008_01", "CatalogueQTTN_2008_01") : (nam.Equals("2010") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072010_01", "CatalogueQTTN_2010_01") : (nam.Equals("2014") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_022014_01", "CatalogueQTTN_2014_01") : (nam.Equals("2016") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072016_01", "CatalogueQTTN_2016_01") : ((!nam.Equals("2019")) ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0] : GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_012019_01", "Catalogue_092019_01"))))));
			}
			else
			{
				text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
			}
			if (isGetDmMoi(dtPeriod, "23") || (WinSettings.TTPeriodInfo.Nam.Equals("2019") && WinSettings.TTPeriodInfo.KieuKy == "Y"))
			{
				HTKKSerializer.Serializer serializer = new HTKKSerializer.Serializer();
				string input = File.ReadAllText(text);
				danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
				if (danhmuc != null && danhmuc.items.Count > 0)
				{
					List<item> items = danhmuc.items;
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
						string text2 = array[0];
						string text3 = array[1];
						string value = array[2];
						if (text3.Equals("V5"))
						{
							dictionary.Add(text3, value);
						}
					}
				}
			}
			else
			{
				XmlDocument val = new XmlDocument();
				val.Load(text);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item2 in elementsByTagName)
				{
					XmlNode val2 = item2;
					string value2 = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value2, "###");
					string text2 = array[0];
					string text3 = array[1];
					string value = array[2];
					if (array[^1].Equals("1") && ((ConvertHelpers.ConvertToInt(text2) >= 404 && text3.Substring(0, 2).Equals("07")) || ConvertHelpers.ConvertToInt(text2) == 469 || (ConvertHelpers.ConvertToInt(text2) < 404 && text3.Substring(0, 2).Equals("09"))))
					{
						dictionary.Add(text3, value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static DateTime[] getTuNgayDenNgay(string kieuKy)
	{
		DateTime[] array = new DateTime[2];
		try
		{
			string text = "";
			string nam = WinSettings.TTPeriodInfo.Nam;
			switch (kieuKy)
			{
			case "M":
			{
				string thang = WinSettings.TTPeriodInfo.Thang;
				DateTime dateTime = DateTime.ParseExact("01/" + thang + "/" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				array[0] = dateTime;
				dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				array[1] = dateTime;
				break;
			}
			case "Y":
			{
				text = WinSettings.TTPeriodInfo.TuNgay;
				DateTime dateTime;
				if (!HTKKFormater.isNullString(text))
				{
					dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				}
				else
				{
					text = WinSettings.TTPeriodInfo.TuThang;
					dateTime = (HTKKFormater.isNullString(text) ? DateTime.ParseExact("01/01" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture) : DateTime.ParseExact("01/" + text, "dd/MM/yyyy", CultureInfo.CurrentCulture));
				}
				array[0] = dateTime;
				text = WinSettings.TTPeriodInfo.DenNgay;
				if (!HTKKFormater.isNullString(text))
				{
					dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				}
				else
				{
					text = WinSettings.TTPeriodInfo.DenThang;
					if (!HTKKFormater.isNullString(text))
					{
						dateTime = DateTime.ParseExact("01/" + text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
					}
					else
					{
						dateTime = DateTime.ParseExact("01/12" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
					}
				}
				array[1] = dateTime;
				break;
			}
			case "D":
			{
				string ngay = WinSettings.TTPeriodInfo.Ngay;
				string thang = WinSettings.TTPeriodInfo.Thang;
				DateTime dateTime = DateTime.ParseExact(ngay + "/" + thang + "/" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				array[0] = dateTime;
				array[1] = dateTime;
				break;
			}
			case "Q":
				if (!HTKKFormater.isNullString(WinSettings.TTPeriodInfo.TuNgay) && !HTKKFormater.isNullString(WinSettings.TTPeriodInfo.DenNgay))
				{
					DateTime dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					array[0] = dateTime;
					dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					array[1] = dateTime;
				}
				else
				{
					string quy = WinSettings.TTPeriodInfo.Quy;
					DateTime dateTime = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1);
					array[0] = dateTime;
					dateTime = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1);
					array[1] = dateTime;
				}
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return array;
	}

	public static List<PathDm> GetCatalogueFileNameNew(DateTime dtPeriod, string menuId, string kieuKy, int indexValidityChild)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Expected O, but got Unknown
		try
		{
			List<PathDm> list = new List<PathDm>();
			string text = "Menu.xml";
			int year = dtPeriod.Year;
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("Root/Menu[@ID='" + menuId + "']");
			DateTime[] tuNgayDenNgay = getTuNgayDenNgay(kieuKy);
			if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("DATE_KY_KE_KHAI"))
			{
				WinSettings.TTPeriodInfo.dicVar.Remove("DATE_KY_KE_KHAI");
			}
			WinSettings.TTPeriodInfo.dicVar.Add("DATE_KY_KE_KHAI", tuNgayDenNgay);
			if (tuNgayDenNgay == null)
			{
				return null;
			}
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				string value = ((XmlNode)val3.Attributes["StartDate"]).Value;
				string text2 = ((val3.Attributes["EndDate"] != null) ? ((XmlNode)val3.Attributes["EndDate"]).Value : null);
				if (HTKKFormater.isNullString(text2))
				{
					continue;
				}
				DateTime dateTime = DateTime.ParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (WinSettings.MenuID == "77" && WinSettings.TTPeriodInfo.Nam.ToNumber() == (decimal)DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture).Year && dateTime >= DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture))
				{
					string value2 = ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["InterfaceIni"]).Value;
					string pSaveInt = "../InterfaceIni/" + ((XmlNode)val3.Attributes["DataFile"]).Value;
					List<string> pathDmGiaTinh = GetPathDmGiaTinh(DateTime.ParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture), "100_14", kieuKy, indexValidityChild);
					string pSaveDataFile = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + ((XmlNode)val3.Attributes["DataFile"]).Value;
					list.Add(new PathDm(value2, pSaveInt, pSaveDataFile, pathDmGiaTinh));
				}
				else
				{
					if ((!(tuNgayDenNgay[0] >= dateTime) || !(tuNgayDenNgay[0] <= dateTime2)) && (!(dateTime >= tuNgayDenNgay[0]) || !(dateTime <= tuNgayDenNgay[1])))
					{
						continue;
					}
					if (dateTime >= DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						string value2 = ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["InterfaceIni"]).Value;
						string pSaveInt = "../InterfaceIni/" + ((XmlNode)val3.Attributes["DataFile"]).Value;
						List<string> pathDmGiaTinh = GetPathDmGiaTinh(DateTime.ParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture), "100_14", kieuKy, indexValidityChild);
						string pSaveDataFile = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + ((XmlNode)val3.Attributes["DataFile"]).Value;
						list.Add(new PathDm(value2, pSaveInt, pSaveDataFile, pathDmGiaTinh));
					}
					else if (WinSettings.MenuID == "77")
					{
						if (WinSettings.TTPeriodInfo.Nam.ToNumber() < (decimal)DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture).Year)
						{
							string value3 = ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["InterfaceIni"]).Value;
							string pSaveDataFileOld = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["DataFile"]).Value + ".xml";
							list.Add(new PathDm(value3, pSaveDataFileOld));
						}
					}
					else
					{
						string value3 = ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["InterfaceIni"]).Value;
						string pSaveDataFileOld = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["DataFile"]).Value + ".xml";
						list.Add(new PathDm(value3, pSaveDataFileOld));
					}
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static List<string> GetPathDmGiaTinh(DateTime dateStart, string menuId, string kieuKy, int indexValidityChild)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c6: Expected O, but got Unknown
		List<string> list = new List<string>();
		try
		{
			string text = "Menu.xml";
			if (menuId.Equals("100_14"))
			{
				text = "HieuLucBangGia.xml";
			}
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("Root/Menu[@ID='" + menuId + "']");
			if (menuId.Equals("100_14"))
			{
				val2 = ((XmlNode)val).SelectSingleNode("Root/Map[@ID='" + WinSettings.HTKKHeaderInfo.maCQTCapCuc.Substring(0, 3) + "']");
			}
			DateTime[] tuNgayDenNgay = getTuNgayDenNgay(kieuKy);
			if (tuNgayDenNgay == null)
			{
				return null;
			}
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				DateTime dateTime = default(DateTime);
				string value = ((XmlNode)val3.Attributes["StartDate"]).Value;
				string text2 = ((val3.Attributes["EndDate"] != null) ? ((XmlNode)val3.Attributes["EndDate"]).Value : null);
				if (!HTKKFormater.isNullString(text2))
				{
					DateTime dateTime2 = DateTime.ParseExact(value, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime3 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (((tuNgayDenNgay[0] >= dateTime2 && tuNgayDenNgay[0] <= dateTime3) || (dateTime2 >= tuNgayDenNgay[0] && dateTime2 <= tuNgayDenNgay[1])) && val3.ChildNodes.Count > indexValidityChild)
					{
						string value2 = ((XmlNode)val3.ChildNodes[indexValidityChild].Attributes["InterfaceIni"]).Value;
						list.Add(value2.Replace(".xml", "_" + WinSettings.HTKKHeaderInfo.maCQTCapCuc.Substring(0, 3) + "_" + value.Replace("/", "") + ".xml"));
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return list;
	}

	public static Dictionary<string, string> GetDanhMuc(string nameFunction)
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = "";
			int month = 1;
			string text2 = "";
			if (WinSettings.TTPeriodInfo.KieuKy == "D")
			{
				month = int.Parse(WinSettings.TTPeriodInfo.Thang);
				text2 = WinSettings.TTPeriodInfo.TuNgay;
			}
			else if (WinSettings.TTPeriodInfo.KieuKy == "M")
			{
				month = int.Parse(WinSettings.TTPeriodInfo.Thang);
			}
			else if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				month = 12;
			}
			else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && !HTKKFormater.isNullString(WinSettings.TTPeriodInfo.Quy))
			{
				month = GeneralFunctions.QuarterToFirstMonth(ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy));
			}
			int year = int.Parse(WinSettings.TTPeriodInfo.Nam);
			DateTime dtPeriod = (HTKKFormater.isNullString(text2) ? new DateTime(year, month, 1) : DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture));
			if (isGetDmMoi(dtPeriod, nameFunction))
			{
				if (!WinSettings.TTPeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
				{
					WinSettings.TTPeriodInfo.dicVar.Add("IS_DM_MOI", "1");
				}
				string menuId = "100_7";
				int indexValidityChild = 0;
				if (nameFunction == "GetDMKSBVMT" || WinSettings.MenuID == "86" || WinSettings.MenuID == "87")
				{
					menuId = "100_13";
				}
				if (nameFunction == "GetDMThueTieuThuDB" || nameFunction == "GetDichVuThueTieuThuDB")
				{
					indexValidityChild = 1;
				}
				List<PathDm> catalogueFileNameNew = GetCatalogueFileNameNew(dtPeriod, menuId, WinSettings.TTPeriodInfo.KieuKy, indexValidityChild);
				dictionary = WinDoWorkEvent.GetCatalogueNew(catalogueFileNameNew, nameFunction);
			}
			else
			{
				string text3 = "";
				switch (nameFunction)
				{
				case "GetDMHoaDon":
					text = HTKKConfig.INTERFACE_INI + "\\Catalogue_loai_HD.xml";
					break;
				case "GetDMHoaDonDT":
					text = HTKKConfig.INTERFACE_INI + "\\Catalogue_loai_HDDT.xml";
					break;
				default:
					if (!(nameFunction == "GetDMMo"))
					{
						switch (nameFunction)
						{
						case "GetDMNganhNgheKD":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_nganhnghe.xml";
							break;
						case "GetDMBienLaiPhi":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_loai_BLP.xml";
							break;
						case "GetDMDonVi":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_29122007_2.xml";
							break;
						case "GetDMTTDBDonVi":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_29122007.xml";
							break;
						case "GetDMLoaiThongTinThayDoi":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_loai_ThongTinThayDoi.xml";
							break;
						case "GetDMLoaiThongTinThayDoiBLP":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_loai_ThongTinThayDoiBLP.xml";
							break;
						case "GetDMMucLePhiMonBaiCT22":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Muc_Phi_Mon_Bai_ct22.xml";
							break;
						case "GetDMMucLePhiMonBaiCT23":
							text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Muc_Phi_Mon_Bai_ct23.xml";
							break;
						default:
							if (!(nameFunction == "GetDichVuThueTieuThuDB"))
							{
								if (WinSettings.TTPeriodInfo.KieuKy == "Y")
								{
									string nam = WinSettings.TTPeriodInfo.Nam;
									if (nam.Equals("2008"))
									{
										text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_052008_01", "CatalogueQTTN_2008_01");
									}
									else if (nam.Equals("2010"))
									{
										text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072010_01", "CatalogueQTTN_2010_01");
									}
									else if (nam.Equals("2014"))
									{
										text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_022014_01", "CatalogueQTTN_2014_01");
									}
									else if (nam.Equals("2016"))
									{
										text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072016_01", "CatalogueQTTN_2016_01");
									}
									else if (nam.Equals("2019"))
									{
										if (WinSettings.MenuID == "77" || WinSettings.MenuID == "87")
										{
											string denNgay = WinSettings.TTPeriodInfo.DenNgay;
											DateTime dateTime;
											if (!HTKKFormater.isNullString(denNgay))
											{
												dateTime = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
											}
											else
											{
												denNgay = WinSettings.TTPeriodInfo.DenThang;
												dateTime = DateTime.ParseExact("01/" + denNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
												dateTime = new DateTime(dateTime.Year, dateTime.Month + 1, 1).AddDays(-1.0);
											}
											DateTime dateTime2 = DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture);
											if (dateTime.Year >= dateTime2.Year && dateTime <= dateTime2)
											{
												text = "../InterfaceIni/Catalogue_012019_01.xml";
											}
										}
									}
									else
									{
										text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
									}
								}
								else
								{
									text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
								}
								break;
							}
							goto case "GetDMThueTieuThuDB";
						case "GetDMThueTieuThuDB":
							text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[1];
							break;
						}
						break;
					}
					goto case "GetHTHoaDon";
				case "GetHTHoaDon":
				case "GetDMThueBVMT":
				case "GetDMThan":
				case "GetDMXangDau":
				case "GetDMNhomNganhNgheKD":
				case "GetDMHopDong":
				case "GetDMHopDong1":
				case "GetDMMucDich":
				case "GetDMTrangThai":
				case "GetDMGhiChu":
				case "GetPhuongXaTheoCQT":
					text = "";
					break;
				}
				if (text.EndsWith(".xml") && text.IndexOf("\\") > -1)
				{
					text3 = text.Split(new char[1] { '\\' })[text.Split(new char[1] { '\\' }).Length - 1];
				}
				string text4 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text3;
				dictionary = ((!File.Exists(text4)) ? WinDoWorkEvent.GetCatalogue(text, nameFunction) : WinDoWorkEvent.GetCatalogue(text4, nameFunction));
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static bool isGetDmMoi(DateTime dtPeriod, string nameFunction)
	{
		bool result = false;
		if (nameFunction != "GetDMKhoangSan" && nameFunction != "GetDMKSBVMT" && nameFunction != "GetDMThueTieuThuDB" && nameFunction != "GetDichVuThueTieuThuDB" && nameFunction != "23")
		{
			return false;
		}
		DateTime dateTime = DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture);
		string denThang = WinSettings.TTPeriodInfo.DenThang;
		denThang = ((!HTKKFormater.isNullString(denThang)) ? ("01/" + denThang) : WinSettings.TTPeriodInfo.DenNgay);
		if (HTKKFormater.isNullString(denThang))
		{
			denThang = "01/01/" + WinSettings.TTPeriodInfo.Nam;
		}
		DateTime dateTime2 = DateTime.ParseExact(denThang, "dd/MM/yyyy", CultureInfo.CurrentCulture);
		if (WinSettings.TTPeriodInfo.KieuKy == "Y")
		{
			dateTime2 = new DateTime(dateTime2.Year, dateTime2.Month, 1).AddMonths(1).AddDays(-1.0);
			if (WinSettings.MenuID == "77" && dtPeriod.Year == dateTime.Year)
			{
				return true;
			}
			if (dtPeriod.Year > dateTime.Year || (dtPeriod.Year >= dateTime.Year && dateTime2 >= dateTime))
			{
				return true;
			}
		}
		else if (WinSettings.TTPeriodInfo.KieuKy == "M")
		{
			if (dtPeriod.Year > dateTime.Year || (dtPeriod.Year >= dateTime.Year && dtPeriod.Month >= dateTime.Month && dateTime2 >= dateTime))
			{
				return true;
			}
		}
		else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
		{
			DateTime dateTime3 = DateTime.ParseExact(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if ((dateTime3 <= dateTime && dateTime2 >= dateTime) || dateTime3 > dateTime)
			{
				return true;
			}
		}
		else if (dtPeriod >= DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture))
		{
			return true;
		}
		return result;
	}

	public static Dictionary<string, string> DMKhoangSan(string nameFunction = "")
	{
		//IL_01e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ee: Expected O, but got Unknown
		//IL_0218: Unknown result type (might be due to invalid IL or missing references)
		//IL_021f: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (!HTKKFormater.isNullString(nameFunction))
			{
				dictionary = GetDanhMuc(nameFunction);
			}
			else
			{
				string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_042010_01.xml";
				int month = 0;
				if (WinSettings.TTPeriodInfo.KieuKy == "M")
				{
					month = int.Parse(WinSettings.TTPeriodInfo.Thang);
				}
				else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
				{
					month = int.Parse(WinSettings.TTPeriodInfo.Quy) * 3 - 2;
				}
				else if (WinSettings.TTPeriodInfo.KieuKy == "Y")
				{
					month = 12;
				}
				int year = int.Parse(WinSettings.TTPeriodInfo.Nam);
				DateTime dtPeriod = new DateTime(year, month, 1);
				if (WinSettings.TTPeriodInfo.KieuKy == "Y")
				{
					string nam = WinSettings.TTPeriodInfo.Nam;
					text = (nam.Equals("2008") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_052008_01", "CatalogueQTTN_2008_01") : (nam.Equals("2010") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072010_01", "CatalogueQTTN_2010_01") : (nam.Equals("2014") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_022014_01", "CatalogueQTTN_2014_01") : ((!nam.Equals("2016")) ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0] : GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072016_01", "CatalogueQTTN_2016_01")))));
				}
				else
				{
					text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
				}
				XmlDocument val = new XmlDocument();
				val.Load(text);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item in elementsByTagName)
				{
					XmlNode val2 = item;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					string key = array[1];
					string value2 = array[2];
					if (array[^1] == "1")
					{
						dictionary.Add(key, value2);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMPHLPNG(string nameFunction = "")
	{
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Expected O, but got Unknown
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = "";
			text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\Catalogue_PHLPNG_012022_01_CF.xml";
			if (!File.Exists(text))
			{
				text = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLPNG_012022_01.xml";
			}
			XmlDocument val = new XmlDocument();
			val.Load(text);
			HTKKSerializer.Serializer serializer = new HTKKSerializer.Serializer();
			if (((XmlNode)val.DocumentElement).Name.IndexOf("danhmuc") < 0)
			{
				dmconfig dmconfig = serializer.Deserialize<dmconfig>(((XmlNode)val).InnerXml);
				string text2 = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLPNG_012022_01.xml";
				XmlDocument val2 = new XmlDocument();
				val2.Load(text2);
				danhmuc danhmuc = serializer.Deserialize<danhmuc>(((XmlNode)val2).InnerXml);
				if (danhmuc != null && danhmuc.items.Count > 0 && dmconfig != null && dmconfig.dmcfs.Count > 0)
				{
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
							{
								dictionary.Add(item.id, item.ten);
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							dictionary.Add(item.id, item.ten);
						}
					}
				}
			}
			else
			{
				danhmuc danhmuc = serializer.Deserialize<danhmuc>(((XmlNode)val).InnerXml);
				if (danhmuc != null && danhmuc.items.Count > 0)
				{
					foreach (item item2 in danhmuc.items)
					{
						dictionary.Add(item2.id, item2.ten);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static item GetItemDMPHLPNGById(string id)
	{
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Expected O, but got Unknown
		HTKKSerializer.Serializer serializer = new HTKKSerializer.Serializer();
		string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLPNG_012022_01.xml";
		XmlDocument val = new XmlDocument();
		val.Load(text);
		danhmuc danhmuc = serializer.Deserialize<danhmuc>(((XmlNode)val).InnerXml);
		return danhmuc.items.FirstOrDefault((item x) => x.id == id);
	}

	public static Dictionary<string, string> DMKhoangSanDonVi(ref Dictionary<string, string> lstDonVi, string nameFunction = "")
	{
		//IL_0296: Unknown result type (might be due to invalid IL or missing references)
		//IL_029d: Expected O, but got Unknown
		//IL_02ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d1: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (!HTKKFormater.isNullString(nameFunction))
			{
				List<item> danhMucItem = GetDanhMucItem(nameFunction);
				if (danhMucItem != null && danhMucItem.Count > 0)
				{
					foreach (item item in danhMucItem)
					{
						if (!dictionary.ContainsKey(item.id))
						{
							dictionary.Add(item.id, item.dvt);
						}
						if (!lstDonVi.ContainsValue(item.dvt))
						{
							lstDonVi.Add(item.id, item.dvt);
						}
					}
				}
			}
			else
			{
				string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_042010_01.xml";
				int month = 0;
				if (WinSettings.TTPeriodInfo.KieuKy == "M")
				{
					month = int.Parse(WinSettings.TTPeriodInfo.Thang);
				}
				else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
				{
					month = int.Parse(WinSettings.TTPeriodInfo.Quy) * 3 - 2;
				}
				else if (WinSettings.TTPeriodInfo.KieuKy == "Y")
				{
					month = 12;
				}
				int year = int.Parse(WinSettings.TTPeriodInfo.Nam);
				DateTime dtPeriod = new DateTime(year, month, 1);
				if (WinSettings.TTPeriodInfo.KieuKy == "Y")
				{
					string nam = WinSettings.TTPeriodInfo.Nam;
					text = (nam.Equals("2008") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_052008_01", "CatalogueQTTN_2008_01") : (nam.Equals("2010") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072010_01", "CatalogueQTTN_2010_01") : (nam.Equals("2014") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_022014_01", "CatalogueQTTN_2014_01") : ((!nam.Equals("2016")) ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0] : GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072016_01", "CatalogueQTTN_2016_01")))));
				}
				else
				{
					text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
				}
				XmlDocument val = new XmlDocument();
				val.Load(text);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item2 in elementsByTagName)
				{
					XmlNode val2 = item2;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					string key = array[1];
					string value2 = array[3];
					dictionary.Add(key, value2);
					if (!lstDonVi.ContainsValue(value2) && array[^1] == "1")
					{
						lstDonVi.Add(key, value2);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMPhiLephi()
	{
		//IL_01d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01db: Expected O, but got Unknown
		//IL_0210: Unknown result type (might be due to invalid IL or missing references)
		//IL_0217: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012026#022025#012025#102024#012024#072023#012021#012020#012018#012017#022015#012014#102013#012012#112011#012011#012009";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int num = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang);
			if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				if (WinSettings.MenuID != "C7")
				{
					month = 12;
				}
				else
				{
					DateTime dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
					month = dateTime.Month;
					num = dateTime.Year;
				}
			}
			DateTime t2 = new DateTime(num, month, 1).AddMonths(1).AddDays(-1.0);
			for (int i = 0; i < array.Length; i++)
			{
				DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
				if (DateTime.Compare(t2, dateTime2) >= 0 && DateTime.Compare(dateTime2, t) >= 0)
				{
					t = dateTime2;
					text = array[i];
				}
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\Catalogue_PHLP_" + text + "_01.xml";
			if (!File.Exists(text3))
			{
				text3 = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLP_" + text + "_01.xml";
			}
			XmlDocument val = new XmlDocument();
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			string text4 = string.Empty;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string text5 = value.Substring(value.Length - 1, 1);
				if (!(text5 == "1"))
				{
					continue;
				}
				string[] array2;
				string text6;
				string value2;
				if (WinSettings.TTPeriodInfo.KieuKy == "Y")
				{
					array2 = Regex.Split(value, "###");
					text6 = array2[1];
					value2 = array2[2];
					if (num >= 2021)
					{
						if (!(array2[0] == "2") || ConvertHelpers.ConvertToInt(array2[1]) < 27502751 || ConvertHelpers.ConvertToInt(array2[1]) == 84002618 || ConvertHelpers.ConvertToInt(array2[1]) == 84002622 || array2[1].Equals("49004949"))
						{
							if (!array2[^2].Equals("0000") && !dictionary.ContainsKey(text6))
							{
								dictionary.Add(text6, value2);
							}
							else
							{
								text4 = text4 + text6 + "-";
							}
						}
					}
					else if (!array2[^2].Equals("0000") && !dictionary.ContainsKey(text6))
					{
						dictionary.Add(text6, value2);
					}
					else
					{
						text4 = text4 + text6 + "-";
					}
					continue;
				}
				array2 = Regex.Split(value, "###");
				text6 = array2[1];
				value2 = array2[2];
				string text7 = "";
				if (WinSettings.MenuID == "I7")
				{
					text7 = "~51~52~53~54~56~57~60~";
					if (text7.IndexOf("~" + array2[0].Trim() + "~") > -1)
					{
						if (!array2[^2].Equals("0000") && !dictionary.ContainsKey(text6))
						{
							dictionary.Add(text6, value2);
						}
						else
						{
							text4 = text4 + text6 + "-";
						}
					}
				}
				else if (WinSettings.MenuID == "I4")
				{
					text7 = "~1~2~3~4~5~6~7~8~9~10~11~12~57~61~";
					if (text7.IndexOf("~" + array2[0].Trim() + "~") > -1)
					{
						if (!array2[^2].Equals("0000") && !dictionary.ContainsKey(text6))
						{
							dictionary.Add(text6, value2);
						}
						else
						{
							text4 = text4 + text6 + "-";
						}
					}
				}
				else if (!array2[^2].Equals("0000") && !dictionary.ContainsKey(text6))
				{
					dictionary.Add(text6, value2);
				}
				else
				{
					text4 = text4 + text6 + "-";
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMPhiLephiTt63()
	{
		//IL_01d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01db: Expected O, but got Unknown
		//IL_0210: Unknown result type (might be due to invalid IL or missing references)
		//IL_0217: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "072026#042026#072025#012025#072024#012024#122023#012009";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang);
			if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				if (WinSettings.MenuID != "C7")
				{
					month = 12;
				}
				else
				{
					DateTime dateTime = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
					month = dateTime.Month;
					year = dateTime.Year;
				}
			}
			DateTime t2 = new DateTime(year, month, 1).AddMonths(1).AddDays(-1.0);
			for (int i = 0; i < array.Length; i++)
			{
				DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
				if (DateTime.Compare(t2, dateTime2) >= 0 && DateTime.Compare(dateTime2, t) >= 0)
				{
					t = dateTime2;
					text = array[i];
				}
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\Catalogue_PHLP_TT63_" + text + "_01.xml";
			if (!File.Exists(text3))
			{
				text3 = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLP_TT63_" + text + "_01.xml";
			}
			XmlDocument val = new XmlDocument();
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			string text4 = string.Empty;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string text5 = value.Substring(value.Length - 1, 1);
				if (!(text5 == "1"))
				{
					continue;
				}
				string[] array2;
				string text6;
				string value2;
				string text7;
				if (WinSettings.TTPeriodInfo.KieuKy == "Y")
				{
					array2 = Regex.Split(value, "###");
					text6 = array2[1];
					value2 = array2[2];
					text7 = "~1~";
					if (text7.IndexOf("~" + array2[0].Trim() + "~") > -1)
					{
						if (!dictionary.ContainsKey(text6))
						{
							dictionary.Add(text6, value2);
							continue;
						}
						text4 = text4 + text6 + "-";
						dictionary.Add(text6 + "-", value2);
					}
					continue;
				}
				array2 = Regex.Split(value, "###");
				text6 = array2[1];
				value2 = array2[2];
				text7 = "";
				if (WinSettings.MenuID == "I7")
				{
					text7 = "~2~";
					if (text7.IndexOf("~" + array2[0].Trim() + "~") > -1)
					{
						if (!dictionary.ContainsKey(text6))
						{
							dictionary.Add(text6, value2);
							continue;
						}
						text4 = text4 + text6 + "-";
						dictionary.Add(text6 + "-", value2);
					}
				}
				else if (WinSettings.MenuID == "I4")
				{
					text7 = "~1~";
					if (text7.IndexOf("~" + array2[0].Trim() + "~") > -1)
					{
						if (!dictionary.ContainsKey(text6))
						{
							dictionary.Add(text6, value2);
							continue;
						}
						text4 = text4 + text6 + "-";
						dictionary.Add(text6 + "-", value2);
					}
				}
				else if (!dictionary.ContainsKey(text6))
				{
					dictionary.Add(text6, value2);
				}
				else
				{
					text4 = text4 + text6 + "-";
					dictionary.Add(text6 + "-", value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMPhiLephi_TieuMuc()
	{
		//IL_0134: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Expected O, but got Unknown
		//IL_0166: Unknown result type (might be due to invalid IL or missing references)
		//IL_016d: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012026#022025#012025#102024#012024#072023#012021#012020#012018#012017#022015#012014#102013#012012#112011#012011#012009";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang);
			if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				month = 12;
			}
			DateTime t2 = new DateTime(year, month, 1).AddMonths(1).AddDays(-1.0);
			for (int i = 0; i < array.Length; i++)
			{
				DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
				if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
				{
					t = dateTime;
					text = array[i];
				}
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text3 = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLP_" + text + "_01.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string key = array2[1];
				string value2 = array2[4];
				if (!array2[^1].Equals("0000"))
				{
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMPhiLephiTt63_TieuMuc()
	{
		//IL_0134: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Expected O, but got Unknown
		//IL_0170: Unknown result type (might be due to invalid IL or missing references)
		//IL_0177: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "072026#042026#072025#012025#072024#012024#122023#012009";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang);
			if (WinSettings.TTPeriodInfo.KieuKy == "Y")
			{
				month = 12;
			}
			DateTime t2 = new DateTime(year, month, 1).AddMonths(1).AddDays(-1.0);
			for (int i = 0; i < array.Length; i++)
			{
				DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
				if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
				{
					t = dateTime;
					text = array[i];
				}
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text3 = HTKKConfig.INTERFACE_INI + "\\Catalogue_PHLP_TT63_" + text + "_01.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			string text4 = "";
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string text5 = array2[1];
				string value2 = array2[4];
				if (!array2[^1].Equals("0000"))
				{
					if (!dictionary.ContainsKey(text5))
					{
						dictionary.Add(text5, value2);
						continue;
					}
					text4 = text4 + text5 + "-";
					dictionary.Add(text5 + "-", value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMYKienKiemToan()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_Y_KIEN_KIEM_TOAN.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetMaTieuMuc()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_MaTM.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["MaTM"]).Value;
				string value2 = ((XmlNode)val2.Attributes["TenTM"]).Value;
				dictionary.Add(value, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMNganhCoThuNhapCap()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_NganhCoDoanhThuCao.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Ma"]).Value;
				string value2 = ((XmlNode)val2.Attributes["Ten"]).Value;
				dictionary.Add(value, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMChuong()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_Chuong.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Ma"]).Value;
				string value2 = ((XmlNode)val2.Attributes["Ten"]).Value;
				dictionary.Add(value, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMNganHang()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_NganHang.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Ma"]).Value;
				string value2 = ((XmlNode)val2.Attributes["Ten"]).Value;
				dictionary.Add(value, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMNganHangKBNN()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_Tinh_NganHang.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["MaNH"]).Value;
				string value2 = ((XmlNode)val2.Attributes["TenNH"]).Value;
				dictionary.Add(value, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMLoaiDiaDiem()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Loai_Dia_Diem.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMMienPhiMonBai()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_MienPhiMonBai.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					if (!dictionary.ContainsKey(key))
					{
						dictionary.Add(key, value2);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMMaChiTieuKHBS()
	{
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Expected O, but got Unknown
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
			if (mCT == null)
			{
				return dictionary;
			}
			foreach (XmlNode childNode in mCT.ChildNodes)
			{
				XmlNode val = childNode;
				int num3 = ((val.Attributes["Dynamic"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value));
				int num4 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num : num2));
				if (num != num4)
				{
					continue;
				}
				if (num3 == 0)
				{
					string value = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value.ToString().Trim());
					string text = ((val.Attributes["MCT"] == null) ? "" : ((XmlNode)val.Attributes["MCT"]).Value.ToString().Trim());
					if (!dictionary.ContainsKey(text) && !string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(text, value);
					}
				}
				if (num3 == 1)
				{
					string value = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value.ToString().Trim());
					string str = ((val.Attributes["TablePath"] == null) ? "" : ((XmlNode)val.Attributes["TablePath"]).Value.ToString().Trim());
					string text = ((val.Attributes["MCT"] == null) ? "" : ((XmlNode)val.Attributes["MCT"]).Value.ToString().Trim());
					string text2 = Functions.FomartKey(str, text);
					if (!dictionary.ContainsKey(text2) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(text2, value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMTieuMucKHBS(FpSpread winFpForm)
	{
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT_TieuMuc.xml", WinSettings.ToKhaiInfo.MenuID);
			if (mCT == null)
			{
				return dictionary;
			}
			foreach (XmlNode childNode in mCT.ChildNodes)
			{
				XmlNode val = childNode;
				string text = ((val.Attributes["MaTieuMuc"] == null) ? "" : ((XmlNode)val.Attributes["MaTieuMuc"]).Value.Trim());
				if (!string.IsNullOrEmpty(text))
				{
					string value = ((val.Attributes["TenTieuMuc"] == null) ? "" : ((XmlNode)val.Attributes["TenTieuMuc"]).Value.Trim());
					string text2 = text;
					if (!dictionary.ContainsKey(text2) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(text2, value);
					}
				}
				else
				{
					if (winFpForm == null)
					{
						continue;
					}
					string maTieuMucPath = ((val.Attributes["MaTieuMucPath"] == null) ? "" : ((XmlNode)val.Attributes["MaTieuMucPath"]).Value.Trim());
					string tablePath = ((val.Attributes["TablePath"] == null) ? "" : ((XmlNode)val.Attributes["TablePath"]).Value.Trim());
					int num = ((val.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value));
					SheetView sheetView = winFpForm.Sheets[num];
					int num2 = -1;
					if (!string.IsNullOrEmpty(maTieuMucPath) && !string.IsNullOrEmpty(tablePath) && num >= 0)
					{
						num2 = WinSettings.ToKhaiInfo.SheetInfos[num].Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == tablePath);
					}
					if (num2 < 0)
					{
						continue;
					}
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num].Tables[num2];
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
					{
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTieuMucPath);
						if (hTKKCMCTableCellInfo != null)
						{
							string text3 = sheetView.Cells[dataLocationTemplateRowStart + num3 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
							string value = text3;
							if (!dictionary.ContainsKey(text3) && !string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(value))
							{
								dictionary.Add(text3, value);
							}
						}
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMToKhaiPhuLucKHBS()
	{
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Expected O, but got Unknown
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
			if (mCT != null)
			{
				foreach (XmlNode childNode in mCT.ChildNodes)
				{
					XmlNode val = childNode;
					int num3 = ((val.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value));
					string value = ((val.Attributes["SheetName"] == null) ? "" : ((XmlNode)val.Attributes["SheetName"]).Value);
					int num4 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num : num2));
					if (num == num4 && !dictionary.ContainsKey(num3.ToString()) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(num3.ToString(), value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMQuocGia()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Gia_full.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					if (!dictionary.ContainsKey(key))
					{
						dictionary.Add(key, value2);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMNoiCap()
	{
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			Dictionary<string, string> noiCap = HTKKWinDictionaryDiaDanh.NoiCap;
			noiCap.Remove("Cả nước");
			foreach (KeyValuePair<string, string> item in noiCap)
			{
				dictionary.Add(item.Value, item.Key);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMQuocGiaVietNamKhac()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Gia.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMCacNuocLangGieng()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Gia_Lang_Gieng.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMQuocTichFull()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Quoc_Tich_full.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMPPLK()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_PPDG_TNDN_ND20.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMPPLKND132()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_PPDG_TNDN_ND132.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMTM()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_NNKD.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1 && array[0].Equals("01_GTGT"))
				{
					string key = array[1];
					string value2 = array[2];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMDungChung(string controlType)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_005b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Expected O, but got Unknown
		//IL_00a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00af: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_DUNG_CHUNG.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Map");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Controltype"]).Value;
				if (!(controlType == value))
				{
					continue;
				}
				foreach (XmlNode childNode in val2.ChildNodes)
				{
					XmlNode val3 = childNode;
					string value2 = ((XmlNode)val3.Attributes["Value"]).Value;
					string[] array = Regex.Split(value2, "###");
					string key = array[0];
					string value3 = array[1];
					dictionary.Add(key, value3);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMHangHoaGTGT()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\DM_loai_HH_Dich_Vu.xml";
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMThueSuatGTGT(string ID)
	{
		//IL_005c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Expected O, but got Unknown
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Expected O, but got Unknown
		try
		{
			string text = "";
			if (ID.Equals("01"))
			{
				text = HTKKConfig.INTERFACE_INI + "\\DM_thue_suat_01_GTGT.xml";
			}
			else if (ID.Equals("71"))
			{
				text = HTKKConfig.INTERFACE_INI + "\\DM_thue_suat_GTGT.xml";
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary.Add(key, value2);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMLoaiTemRuou()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_loai_Tem_ruou.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[2];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMGhiChuTemRuou()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_ghi_chu_tem_ruou.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMLoaiTienTe()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_don_vi_tien.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMLoaiTienTeNew()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_don_vi_tien_new.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMThueSuat()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_thue_suat_HDDT.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[2];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMMo()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\DM_MO_KHOANG_SAN.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMLoaiDat()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_LoaiDat.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMHangDat()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_HangDat.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMNnkdNtnn()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\DM_NNKD_NTNN.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> DMLoaiTienDk()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\DM_LOAI_TIEN_DK.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void InsertTableRows(SheetView sView, int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount)
	{
		if (rowCount != 0)
		{
			sView.Rows.Add(toRow, rowCount);
			int num = (int)Math.Floor(Math.Log(rowCount, 2.0));
			for (int i = 0; i < num; i++)
			{
				sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, i) - 1.0), toCol, (int)Math.Pow(2.0, i), colCount, dataOnly: false);
				string text = ((int)((double)toRow + Math.Pow(2.0, i) - 1.0)).ToString();
				string text2 = ((int)Math.Pow(2.0, i)).ToString();
			}
			sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, num) - 1.0), toCol, rowCount - (int)Math.Pow(2.0, num) + 1, colCount, dataOnly: false);
		}
	}

	public static void InsertTableRows(SheetView sView, int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount, HTKKCMCSheetInfo sCMCInfo)
	{
		if (rowCount == 0)
		{
			return;
		}
		sView.Rows.Add(toRow, rowCount);
		int num = (int)Math.Floor(Math.Log(rowCount, 2.0));
		for (int i = 0; i < num; i++)
		{
			sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, i) - 1.0), toCol, (int)Math.Pow(2.0, i), colCount, dataOnly: false);
			string text = ((int)((double)toRow + Math.Pow(2.0, i) - 1.0)).ToString();
			string text2 = ((int)Math.Pow(2.0, i)).ToString();
		}
		sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, num) - 1.0), toCol, rowCount - (int)Math.Pow(2.0, num) + 1, colCount, dataOnly: false);
		sCMCInfo.IncrementStaticCell(frRow, rowCount);
		foreach (HTKKCMCTableInfo table in sCMCInfo.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > frRow)
			{
				table.Increment(rowCount);
			}
		}
	}

	public static void InsertTableRowsLargeData(SheetView sView, int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount, HTKKCMCSheetInfo sCMCInfo)
	{
		if (rowCount == 0)
		{
			return;
		}
		sView.Rows.Add(toRow, rowCount);
		sCMCInfo.IncrementStaticCell(frRow, rowCount);
		foreach (HTKKCMCTableInfo table in sCMCInfo.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > frRow)
			{
				table.Increment(rowCount);
			}
		}
	}

	public static void InsertTableRowsSpreadForm(SheetView sView, int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount, HTKKCMCSheetInfo sCMCInfo, bool isSupportBindArray)
	{
		if (rowCount == 0)
		{
			return;
		}
		sView.Rows.Add(toRow, rowCount);
		if (!"05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(sCMCInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) || rowCount <= 500 || !isSupportBindArray)
		{
			int num = (int)Math.Floor(Math.Log(rowCount, 2.0));
			for (int i = 0; i < num; i++)
			{
				sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, i) - 1.0), toCol, (int)Math.Pow(2.0, i), colCount, dataOnly: false);
				string text = ((int)((double)toRow + Math.Pow(2.0, i) - 1.0)).ToString();
				string text2 = ((int)Math.Pow(2.0, i)).ToString();
			}
			sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, num) - 1.0), toCol, rowCount - (int)Math.Pow(2.0, num) + 1, colCount, dataOnly: false);
		}
	}

	public static void InsertTableRow(SheetView sView, int frRow, int frCol, int toRow, int toCol, int colCount, HTKKCMCSheetInfo sInfo, int iCurTable)
	{
		try
		{
			sInfo.IncrementStaticCell(frRow, 1);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow)
				{
					table.Increment(1);
				}
				table.IncrementFooter(frRow, 1);
			}
			sView.Rows.Add(frRow, 1);
			sView.CopyRange(frRow - 1, frCol, toRow, toCol, 1, colCount, dataOnly: false);
			for (int i = 0; i < sInfo.Tables[iCurTable].RowsInfo.Cells.Count; i++)
			{
				if (sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].BackColor == Color.Pink || sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].BackColor == Color.Yellow)
				{
					sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].ResetBackColor();
				}
				sView.Cells[frRow, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1].ResetNote();
				fpSpBindDataCell(sView, frRow + 1, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[iCurTable].RowsInfo.Cells[i].Controltype, sInfo.Tables[iCurTable].RowsInfo.Cells[i].DefaultValue);
			}
			sInfo.Tables[iCurTable].addNewDefaultRow(frRow - sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableRowData(SheetView sView, int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount)
	{
		sView.Rows.Add(frRow, rowCount);
		int num = (int)Math.Floor(Math.Log(rowCount, 2.0));
		for (int i = 0; i < num; i++)
		{
			sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, i) - 1.0), toCol, (int)Math.Pow(2.0, i), colCount, dataOnly: false);
			string text = ((int)((double)toRow + Math.Pow(2.0, i) - 1.0)).ToString();
			string text2 = ((int)Math.Pow(2.0, i)).ToString();
		}
		sView.CopyRange(frRow - 1, frCol, (int)((double)toRow + Math.Pow(2.0, num) - 1.0), toCol, rowCount - (int)Math.Pow(2.0, num) + 1, colCount, dataOnly: false);
	}

	public static void ClearData(FpSpread winFpSpread, int activeSheet, bool isShowMess)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		try
		{
			string text = "";
			text = ((winFpSpread.ActiveSheetIndex != 0) ? "0035" : "0030");
			frmMessage frmMessage = new frmMessage(text, "5", "2", "Thông báo");
			if (isShowMess)
			{
				winFpSpread.ResumeLayout();
				((Form)frmMessage).ShowDialog();
			}
			if (!winFpSpread.IsLayoutSuspended)
			{
				winFpSpread.SuspendLayout();
			}
			if (!frmMessage.Yes && isShowMess)
			{
				return;
			}
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			if (WinSettings.ToKhaiInfo.SheetInfos[activeSheet].TaiDuLieuRatLon == "1")
			{
				WinSettings.ToKhaiInfo.SheetInfos[activeSheet] = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val, activeSheet, getTemplate: true, WinSettings.TTPeriodInfo);
				WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
				return;
			}
			ResetColorSheetDefault(winFpSpread, WinSettings.ToKhaiInfo, activeSheet);
			while (true)
			{
				List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[activeSheet].Tables;
				winFpSpread.Sheets[activeSheet].SetActiveCell(0, 0);
				int num = 0;
				while (true)
				{
					if (num < tables.Count)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = tables[num];
						if (hTKKCMCTableInfo.Header == null)
						{
							if (num == 0)
							{
								RemoveDataRows(winFpSpread.Sheets[activeSheet], hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count - 1, WinSettings.ToKhaiInfo.SheetInfos[activeSheet]);
								hTKKCMCTableInfo.RowsValues.Clear();
							}
							if (num > 0)
							{
								if (hTKKCMCTableInfo.Header != null && hTKKCMCTableInfo.Header.StaticCells.Count > 0)
								{
									break;
								}
								RemoveDataRows(winFpSpread.Sheets[activeSheet], hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count - 1, WinSettings.ToKhaiInfo.SheetInfos[activeSheet]);
								hTKKCMCTableInfo.RowsValues.Clear();
							}
						}
						num++;
						continue;
					}
					for (num = tables.Count - 1; num >= 0; num--)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = tables[num];
						if (hTKKCMCTableInfo.Header != null && hTKKCMCTableInfo.Header.StaticCells.Count > 0)
						{
							if (hTKKCMCTableInfo.TableID != "1")
							{
								int tableTemplateRowStart = hTKKCMCTableInfo.Location.TableTemplateRowStart;
								GeneralFunctions.DeleteTableF8(winFpSpread.ActiveSheet, WinSettings.ToKhaiInfo.SheetInfos[activeSheet], num, tableTemplateRowStart - 1, 0, batBuocRemove: true);
							}
							else
							{
								GeneralFunctions.RemoveDataRows(winFpSpread.ActiveSheet, num, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, WinSettings.ToKhaiInfo.SheetInfos[activeSheet]);
							}
						}
						else if (hTKKCMCTableInfo.Header != null && hTKKCMCTableInfo.Header.StaticCells.Count == 0)
						{
							GeneralFunctions.RemoveDataRows(winFpSpread.ActiveSheet, num, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, WinSettings.ToKhaiInfo.SheetInfos[activeSheet]);
						}
					}
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[activeSheet];
					HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val, activeSheet, getTemplate: true, WinSettings.TTPeriodInfo);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet] = hTKKSheetInfo;
					for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells.Count; i++)
					{
						int rowIndex = WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].CellTemplate.RowIndex;
						int colIndex = WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].CellTemplate.ColIndex;
						if (!(WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].Controltype == "0") && !(WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].Controltype == "14"))
						{
							continue;
						}
						bool flag = false;
						for (int j = 0; j < hTKKCMCSheetInfo.StaticCells.Count; j++)
						{
							if (hTKKCMCSheetInfo.StaticCells[j].Path == WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].Path)
							{
								if (!hTKKCMCSheetInfo.StaticCells[j].Validate.IsReset)
								{
									WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].Validate = hTKKCMCSheetInfo.StaticCells[j].Validate;
									WinSettings.ToKhaiInfo.SheetInfos[activeSheet].StaticCells[i].Value = hTKKCMCSheetInfo.StaticCells[j].Value;
								}
								flag = true;
								break;
							}
						}
						if (!flag && winFpSpread.Sheets[activeSheet].Rows[rowIndex - 1].Visible && !winFpSpread.Sheets[activeSheet].Cells[rowIndex - 1, colIndex - 1].Locked && (!(WinSettings.MenuID == "53") || (colIndex != HTKKFile.ColumnLettersToIndex("P") && colIndex != HTKKFile.ColumnLettersToIndex("H"))))
						{
							winFpSpread.Sheets[activeSheet].Cells[rowIndex - 1, colIndex - 1].ResetFormula();
							winFpSpread.Sheets[activeSheet].Cells[rowIndex - 1, colIndex - 1].Value = "";
						}
					}
					ProcessSheet(activeSheet, winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
					WinDoWorkEvent.ClearData(winFpSpread, activeSheet);
					return;
				}
				RemoveTable(winFpSpread.ActiveSheet, num, WinSettings.ToKhaiInfo.SheetInfos[activeSheet]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveTable(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			sView.Rows.Remove(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			for (int i = itable + 1; i < sInfo.Tables.Count; i++)
			{
				sInfo.Tables[i].TableID = i.ToString();
				if (sInfo.Tables[i].Header.StaticCells.Count > 0 && ConvertHelpers.ConvertToString(sInfo.Tables[i].Header.StaticCells[0].DefaultValue).Equals("NA"))
				{
					sInfo.Tables[i].Header.StaticCells[0].Value = sInfo.Tables[i].TableID;
				}
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					sInfo.Tables[i].CellTableSTT.Value = i.ToString();
				}
				sInfo.Tables[i].Increment((hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
				sInfo.Tables[i].IncrementFooter((hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, (hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count) * -1);
			sInfo.Tables.RemoveAt(itable);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableInSpread(SheetView sView, int itable, HTKKCMCSheetInfo sCMCInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sCMCInfo.Tables[itable];
			if (hTKKCMCTableInfo.RowsInfo.Cells.Count > 0)
			{
				sView.Rows.Add(hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength, hTKKCMCTableInfo.TableLength + 1);
				sView.CopyRange(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, -1, hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength, -1, hTKKCMCTableInfo.TableLength + 1, -1, dataOnly: false);
			}
			else
			{
				sView.Rows.Add(hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength - 1, hTKKCMCTableInfo.TableLength);
				sView.CopyRange(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, -1, hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength - 1, -1, hTKKCMCTableInfo.TableLength, -1, dataOnly: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableInSpreadReport(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			if (hTKKCMCTableInfo.RowsInfo.Cells.Count > 0)
			{
				sView.Rows.Add(hTKKCMCTableInfo.Location.TableReportRowStart + hTKKCMCTableInfo.TableLength, hTKKCMCTableInfo.TableLength + 1);
				sView.CopyRange(hTKKCMCTableInfo.Location.TableReportRowStart - 1, -1, hTKKCMCTableInfo.Location.TableReportRowStart + hTKKCMCTableInfo.TableLength, -1, hTKKCMCTableInfo.TableLength + 1, -1, dataOnly: false);
			}
			else
			{
				sView.Rows.Add(hTKKCMCTableInfo.Location.TableReportRowStart + hTKKCMCTableInfo.TableLength - 1, hTKKCMCTableInfo.TableLength);
				sView.CopyRange(hTKKCMCTableInfo.Location.TableReportRowStart - 1, -1, hTKKCMCTableInfo.Location.TableReportRowStart + hTKKCMCTableInfo.TableLength - 1, -1, hTKKCMCTableInfo.TableLength, -1, dataOnly: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTable(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			sView.Rows.Add(hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			sView.CopyRange(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, -1, hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1, -1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count, -1, dataOnly: false);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = new HTKKCMCTableInfo(hTKKCMCTableInfo, 1);
			sInfo.Tables.Insert(itable + 1, hTKKCMCTableInfo2);
			if (WinSettings.MenuID == "G9")
			{
				string tableID = sInfo.Tables[itable].TableID;
				sInfo.Tables[itable + 1].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
				if (sInfo.Tables[itable + 1].CellTableSTT != null)
				{
					sInfo.Tables[itable + 1].CellTableSTT.Value = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
				}
				if (sInfo.Tables[itable + 1].CellTableReportSTT != null)
				{
					sInfo.Tables[itable + 1].CellTableReportSTT.Value = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
				}
			}
			else
			{
				sInfo.Tables[itable + 1].TableID = (itable + 2).ToString();
				if (sInfo.Tables[itable + 1].CellTableSTT != null)
				{
					sInfo.Tables[itable + 1].CellTableSTT.Value = (itable + 2).ToString();
				}
				if (sInfo.Tables[itable + 1].CellTableReportSTT != null)
				{
					sInfo.Tables[itable + 1].CellTableReportSTT.Value = (itable + 2).ToString();
				}
			}
			hTKKCMCTableInfo2.Increment(hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			hTKKCMCTableInfo2.IncrementFooter(hTKKCMCTableInfo.TableLength + 1);
			if (sInfo.Tables[itable + 1].RowsInfo.RowDefaultDisplay == "0")
			{
				sView.Rows[sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart - 1].Visible = false;
				sInfo.Tables[itable + 1].RowsValues[0].RowDisplay = "0";
			}
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				sView.Rows.Remove(sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count - 1);
			}
			string tableName = sInfo.Tables[itable + 1].TableName;
			for (int i = itable + 2; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					if (WinSettings.MenuID == "G9")
					{
						string tableName2 = sInfo.Tables[i].TableName;
						if (tableName2 == tableName)
						{
							string tableID = sInfo.Tables[i - 1].TableID;
							sInfo.Tables[i].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
						}
					}
					else
					{
						sInfo.Tables[i].TableID = (i + 1).ToString();
					}
				}
				if (WinSettings.MenuID == "G9")
				{
					string tableName2 = sInfo.Tables[i].TableName;
					if (tableName2 == tableName && sInfo.Tables[i].CellTableSTT != null)
					{
						string tableID = sInfo.Tables[i - 1].TableID;
						sInfo.Tables[i].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
					}
				}
				else if (sInfo.Tables[i].CellTableSTT != null)
				{
					sInfo.Tables[i].CellTableSTT.Value = (i + 1).ToString();
				}
				sInfo.Tables[i].Increment(hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			if (sInfo.Tables[itable + 1].Header != null)
			{
				for (int j = 0; j < sInfo.Tables[itable + 1].Header.StaticCells.Count; j++)
				{
					if (sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue == "NA")
					{
						sInfo.Tables[itable + 1].Header.StaticCells[j].Value = (ConvertHelpers.ConvertToDouble(sInfo.Tables[itable + 1].Header.StaticCells[0].Value) + 1.0).ToString();
						continue;
					}
					sInfo.Tables[itable + 1].Header.StaticCells[j].Value = sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue;
					fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Header.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[itable + 1].Header.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[itable + 1].Header.StaticCells[j].Controltype, sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue ?? "");
				}
			}
			for (int i = 0; i < sInfo.Tables[itable + 1].RowsInfo.Cells.Count; i++)
			{
				fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart, sInfo.Tables[itable + 1].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[itable + 1].RowsInfo.Cells[i].Controltype, sInfo.Tables[itable + 1].RowsInfo.Cells[i].DefaultValue);
			}
			if (sInfo.Tables[itable + 1].Footer == null)
			{
				return;
			}
			for (int j = 0; j < sInfo.Tables[itable + 1].Footer.StaticCells.Count; j++)
			{
				if (sInfo.Tables[itable + 1].Footer.StaticCells[j].DefaultValue == "NA")
				{
					sInfo.Tables[itable + 1].Footer.StaticCells[j].Value = (ConvertHelpers.ConvertToDouble(sInfo.Tables[itable + 1].Footer.StaticCells[0].Value) + 1.0).ToString();
					continue;
				}
				sInfo.Tables[itable + 1].Footer.StaticCells[j].Value = sInfo.Tables[itable + 1].Footer.StaticCells[j].DefaultValue;
				fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Footer.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[itable + 1].Footer.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[itable + 1].Footer.StaticCells[j].Controltype, sInfo.Tables[itable + 1].Footer.StaticCells[j].DefaultValue ?? "");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InsertTableNew(SheetView sView, int itable, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[itable];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = new HTKKCMCTableInfo(hTKKCMCTableInfo, 1);
			sInfo.Tables.Insert(itable + 1, hTKKCMCTableInfo2);
			string tableID = sInfo.Tables[itable].TableID;
			sInfo.Tables[itable + 1].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
			if (sInfo.Tables[itable + 1].CellTableSTT != null)
			{
				sInfo.Tables[itable + 1].CellTableSTT.Value = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
			}
			if (sInfo.Tables[itable + 1].CellTableReportSTT != null)
			{
				sInfo.Tables[itable + 1].CellTableReportSTT.Value = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
			}
			hTKKCMCTableInfo2.Increment(hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			hTKKCMCTableInfo2.IncrementFooter(hTKKCMCTableInfo.TableLength + 1);
			string tableName = sInfo.Tables[itable + 1].TableName;
			string text = "";
			for (int i = itable + 2; i < sInfo.Tables.Count; i++)
			{
				if (sInfo.Tables[i].CellTableSTT != null)
				{
					text = sInfo.Tables[i].TableName;
					if (text == tableName)
					{
						tableID = sInfo.Tables[i - 1].TableID;
						sInfo.Tables[i].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
					}
				}
				text = sInfo.Tables[i].TableName;
				if (text == tableName && sInfo.Tables[i].CellTableSTT != null)
				{
					tableID = sInfo.Tables[i - 1].TableID;
					sInfo.Tables[i].TableID = string.Concat(ConvertHelpers.ConvertToInt(tableID) + 1);
				}
				sInfo.Tables[i].Increment(hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			}
			sInfo.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo2.TableLength + hTKKCMCTableInfo2.RowsValues.Count);
			sView.Rows.Add(hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count);
			sView.CopyRange(hTKKCMCTableInfo.Location.TableTemplateRowStart - 1, -1, hTKKCMCTableInfo.Location.TableTemplateRowStart + hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1, -1, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count, -1, dataOnly: false);
			if (sInfo.Tables[itable + 1].RowsInfo.RowDefaultDisplay == "0")
			{
				sView.Rows[sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart - 1].Visible = false;
				sInfo.Tables[itable + 1].RowsValues[0].RowDisplay = "0";
			}
			if (sInfo.Tables[itable].RowsValues.Count > 1)
			{
				GeneralFunctions.RemoveRowOnSView(sView, sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart, sInfo.Tables[itable].RowsValues.Count - 1);
			}
			if (sInfo.Tables[itable + 1].Header != null)
			{
				for (int j = 0; j < sInfo.Tables[itable + 1].Header.StaticCells.Count; j++)
				{
					if (sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue == "NA")
					{
						sInfo.Tables[itable + 1].Header.StaticCells[j].Value = (ConvertHelpers.ConvertToDouble(sInfo.Tables[itable + 1].Header.StaticCells[0].Value) + 1.0).ToString();
						continue;
					}
					sInfo.Tables[itable + 1].Header.StaticCells[j].Value = sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue;
					fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Header.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[itable + 1].Header.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[itable + 1].Header.StaticCells[j].Controltype, sInfo.Tables[itable + 1].Header.StaticCells[j].DefaultValue ?? "");
				}
			}
			for (int i = 0; i < sInfo.Tables[itable + 1].RowsInfo.Cells.Count; i++)
			{
				fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Location.DataLocationTemplateRowStart, sInfo.Tables[itable + 1].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[itable + 1].RowsInfo.Cells[i].Controltype, sInfo.Tables[itable + 1].RowsInfo.Cells[i].DefaultValue);
			}
			if (sInfo.Tables[itable + 1].Footer == null)
			{
				return;
			}
			for (int j = 0; j < sInfo.Tables[itable + 1].Footer.StaticCells.Count; j++)
			{
				if (sInfo.Tables[itable + 1].Footer.StaticCells[j].DefaultValue == "NA")
				{
					sInfo.Tables[itable + 1].Footer.StaticCells[j].Value = (ConvertHelpers.ConvertToDouble(sInfo.Tables[itable + 1].Footer.StaticCells[0].Value) + 1.0).ToString();
					continue;
				}
				sInfo.Tables[itable + 1].Footer.StaticCells[j].Value = sInfo.Tables[itable + 1].Footer.StaticCells[j].DefaultValue;
				fpSpBindDataCell(sView, sInfo.Tables[itable + 1].Footer.StaticCells[j].CellTemplate.RowIndex, sInfo.Tables[itable + 1].Footer.StaticCells[j].CellTemplate.ColIndex, sInfo.Tables[itable + 1].Footer.StaticCells[j].Controltype, sInfo.Tables[itable + 1].Footer.StaticCells[j].DefaultValue ?? "");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetValueRow(SheetView sView, HTKKCMCTableInfo tableInfo, int iRow, HTKKCMCRowValue RowValue)
	{
		for (int i = 0; i < tableInfo.RowsInfo.Cells.Count; i++)
		{
			if (sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + iRow - 1, tableInfo.RowsInfo.Cells[i].CellTemplate.ColIndex - 1].CellType.ToString() == "NumberCellType")
			{
				sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + iRow - 1, tableInfo.RowsInfo.Cells[i].CellTemplate.ColIndex - 1].Value = ConvertHelpers.ConvertToDouble(RowValue.RowValue[i]);
			}
			else
			{
				sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + iRow - 1, tableInfo.RowsInfo.Cells[i].CellTemplate.ColIndex - 1].Value = RowValue.RowValue[i];
			}
		}
		Row row = sView.Rows[tableInfo.Location.DataLocationTemplateRowStart + iRow - 1];
		float preferredHeight = row.GetPreferredHeight();
		if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
		{
			row.Height = preferredHeight;
		}
	}

	public static void RemoveDataRows(SheetView sView, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (rowCount <= 0)
			{
				return;
			}
			sView.Rows.Remove(frRow, rowCount);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow + 1)
				{
					table.Increment(rowCount * -1);
				}
				if (table.Footer != null && table.Footer.StaticCells.Count > 0)
				{
					table.IncrementFooter(frRow, rowCount * -1);
				}
			}
			sInfo.IncrementStaticCell(frRow, rowCount * -1);
		}
		catch (Exception)
		{
		}
	}

	public static void RemoveDataRows(SheetView sView, int tableIndex, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (rowCount <= 0)
			{
				return;
			}
			sView.Rows.Remove(frRow, rowCount);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow + 1)
				{
					table.Increment(rowCount * -1);
				}
			}
			sInfo.IncrementStaticCell(frRow, rowCount * -1);
			int num = sInfo.Tables[tableIndex].RowsValues.Count - 1;
			for (int num2 = num; num2 > num - rowCount; num2--)
			{
				sInfo.Tables[tableIndex].RowsValues.RemoveAt(num2);
			}
		}
		catch (Exception)
		{
		}
	}

	public static string WriteContentTableDatafile(SheetView sView, XmlNode SectionNode, HTKKCMCTableInfo tableInfo)
	{
		string result = "";
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = null;
		}
		return result;
	}

	public static void fpSpValidateDataCell(SheetView sView, HTKKCMCTableCellInfo cellInfo)
	{
		try
		{
			if (cellInfo.Validate != null)
			{
				if (cellInfo.Validate.Bang != null && sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].Value.ToString() != cellInfo.Validate.Bang.ToString())
				{
					sView.SetNote(cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1, cellInfo.Validate.MessageID);
					sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].NoteStyle = NoteStyle.PopupNote;
				}
				if (cellInfo.Validate.LonHon != null && !ValidateBase.ValidateLonHon(sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].Value.ToString(), cellInfo.Validate.Bang.ToString()))
				{
					sView.SetNote(cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1, cellInfo.Validate.MessageID);
					sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].NoteStyle = NoteStyle.PopupNote;
				}
				if (cellInfo.Validate.MinValue != null && !ValidateBase.ValidateMinValue(sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].Value.ToString(), cellInfo.Validate.Bang.ToString()))
				{
					sView.SetNote(cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1, cellInfo.Validate.MessageID);
					sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].NoteStyle = NoteStyle.PopupNote;
				}
				if (cellInfo.Validate.MaxLen != null && !ValidateBase.ValidateMaxLen(sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].Value.ToString(), cellInfo.Validate.MaxLen))
				{
					sView.SetNote(cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1, cellInfo.Validate.MessageID);
					sView.Cells[cellInfo.CellTemplate.RowIndex - 1, cellInfo.CellTemplate.ColIndex - 1].NoteStyle = NoteStyle.PopupNote;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void setSpread_WH(SheetView sView, HTKKCMCSheetInfo sheetInfo)
	{
		try
		{
			if (sheetInfo == null)
			{
				return;
			}
			int num = 0;
			int num2 = 0;
			num2 = sView.NonEmptyColumnCount;
			num = sView.NonEmptyRowCount;
			sView.GrayAreaBackColor = ColorTranslator.FromHtml("#F0F0F0");
			sView.Columns.Count = num2;
			sView.Rows.Count = num;
			sView.Columns[num2 - 1].BackColor = ColorTranslator.FromHtml("#2E2E2E");
			sView.Columns[num2 - 1].Width = 1f;
			sView.Columns[num2 - 1].Locked = true;
			sView.Rows[num - 1].BackColor = ColorTranslator.FromHtml("#2E2E2E");
			sView.Rows[num - 1].Height = 1f;
			sView.Rows[num - 1].Locked = true;
			for (int num3 = num - 1; num3 >= 0; num3--)
			{
				if (sView.Cells[num3, num2 - 1].Text == "eof")
				{
					sView.Cells[num3, num2 - 1].Value = "";
					break;
				}
			}
			sView.SetActiveCell(0, 0);
			sView.Protect = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void fpSpDisplayExp(FpSpread winFpForm, FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos)
	{
		try
		{
			for (short num = 0; num < winFpForm.Sheets.Count - 1; num++)
			{
				if (sheetInfos[num] != null)
				{
					if (WinFnProcess.IsWorking(sheetInfos[num].SheetName))
					{
						winFpExp.Sheets[num].SheetName = sheetInfos[num].SheetName;
						WinSetupSpreadSheet(winFpExp.Sheets[num]);
						setSpreadExp(winFpExp.Sheets[num], sheetInfos[num]);
						setValueReport(winFpForm.Sheets[num], winFpExp.Sheets[num], sheetInfos[num]);
					}
					else
					{
						winFpExp.Sheets[num].Visible = false;
					}
				}
				else
				{
					winFpExp.Sheets[num].Visible = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void setSpreadExp(SheetView sView, HTKKCMCSheetInfo sheetInfos)
	{
		try
		{
			if (sheetInfos == null)
			{
				return;
			}
			int num = 0;
			int num2 = 0;
			int nonEmptyColumnCount = sView.NonEmptyColumnCount;
			int nonEmptyRowCount = sView.NonEmptyRowCount;
			int num3 = 0;
			string text = "";
			List<Cell> list = new List<Cell>();
			for (int num4 = nonEmptyRowCount; num4 > 5; num4--)
			{
				for (int i = 0; i < nonEmptyColumnCount; i++)
				{
					text = sView.Cells[num4, i].Text;
					if (text != null)
					{
						if (text.Trim() == "eof")
						{
							list.Add(sView.Cells[num4, i]);
							num3++;
						}
						if (num3 == 2)
						{
							break;
						}
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j].Column.Index > num2)
				{
					num2 = list[j].Column.Index;
				}
				if (list[j].Row.Index > num)
				{
					num = list[j].Row.Index;
				}
			}
			sView.Columns.Count = num2;
			sView.Rows.Count = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void setValueReport(SheetView sViewForm, SheetView sViewReport, HTKKCMCSheetInfo sheetInfos)
	{
		try
		{
			List<HTKKCMCCellInfo> staticCells = sheetInfos.StaticCells;
			for (int i = 0; i < staticCells.Count; i++)
			{
				HTKKCMCCellBasicInfo cellTemplate = staticCells[i].CellTemplate;
				HTKKCMCCellBasicInfo cellReport = staticCells[i].CellReport;
				fpSpBindDataCell(sViewReport, cellReport.RowIndex, cellReport.ColIndex, staticCells[i].Controltype, staticCells[i].Value);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool setValueSpreadImpToForm(FpSpread fpSpreadForm, FpSpread fpSpreadImp, List<HTKKCMCSheetInfo> sheetInfos)
	{
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Expected O, but got Unknown
		try
		{
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[i] != null && WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName))
				{
					WinSettings.DicSheet.Remove(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName);
				}
				WinSettings.ToKhaiInfo.SheetInfos[i] = null;
			}
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			List<HTKKSheetInMenu> allSheetInMenu = WinDoWorkEvent.getAllSheetInMenu(val);
			for (int i = 0; i < fpSpreadImp.Sheets.Count - 1; i++)
			{
				if (fpSpreadImp.Sheets[i].Visible)
				{
					HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val, i, getTemplate: true, WinSettings.TTPeriodInfo);
					WinSettings.ToKhaiInfo.SheetInfos[i] = hTKKSheetInfo;
					WinSettings.DicSheet.Add(((XmlNode)val[i].Attributes["SheetName"]).Value, 2);
					setValueImportExcel(fpSpreadForm.Sheets[i], fpSpreadImp.Sheets[i], WinSettings.ToKhaiInfo.SheetInfos[i]);
					fpSpreadForm.Sheets[i].Visible = true;
				}
				else
				{
					fpSpreadForm.Sheets[i].Visible = false;
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void setValueImportExcel(SheetView sViewForm, SheetView sViewImport, HTKKCMCSheetInfo sheetInfo)
	{
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_012b: Unknown result type (might be due to invalid IL or missing references)
		//IL_016c: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (WinSettings.MenuID == "39")
			{
				string text = sViewImport.Cells[8, HTKKFile.ColumnLettersToIndex("D") - 1].Text.Trim();
				if (HTKKFormater.isNullString(text))
				{
					frmMessage frmMessage = new frmMessage("0141", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				if (text.Length == 13)
				{
					text.Replace("-", "");
				}
				if (text != WinSettings.TTinChung.mst)
				{
					frmMessage frmMessage = new frmMessage("0141", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				string text2 = sViewImport.Cells[4, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				if (HTKKFormater.isNullString(text2))
				{
					frmMessage frmMessage = new frmMessage("0142", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				if (text2 != WinSettings.TTinChung.kyKKhai)
				{
					frmMessage frmMessage = new frmMessage("0142", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
			}
			List<HTKKCMCTableInfo> tables = sheetInfo.Tables;
			if (tables.Count > 0)
			{
				for (int i = 0; i < tables.Count; i++)
				{
					TaiBangKe(sViewForm, sViewImport, sheetInfo, i);
				}
			}
			List<HTKKCMCCellInfo> staticCells = sheetInfo.StaticCells;
			for (int i = 0; i < staticCells.Count; i++)
			{
				HTKKCMCCellBasicInfo cellTemplate = staticCells[i].CellTemplate;
				try
				{
					string valueImport = GeneralFunctions.GetValueImport(sViewImport, staticCells[i].Controltype, cellTemplate.RowIndex, cellTemplate.ColIndex, staticCells[i].DefaultValue);
					if (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9")
					{
						if (staticCells[i].Controltype == "2")
						{
							if (valueImport == null)
							{
								staticCells[i].Value = "False";
								continue;
							}
							switch (valueImport.ToString())
							{
							case "x":
								staticCells[i].Value = "True";
								break;
							case "0":
								staticCells[i].Value = "True";
								break;
							case "":
								staticCells[i].Value = "False";
								break;
							case "X":
								staticCells[i].Value = "False";
								break;
							}
						}
						else if (staticCells[i].Controltype != "15")
						{
							staticCells[i].Value = valueImport;
						}
					}
					else if (!string.IsNullOrEmpty(valueImport))
					{
						staticCells[i].Value = valueImport;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void TaiBangKe(SheetView sViewForm, SheetView sViewImport, HTKKCMCSheetInfo sheetInfos, int iTable)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sheetInfos.Tables[iTable];
		hTKKCMCTableInfo.RowsValues.Clear();
		List<HTKKCMCRowValue> list = new List<HTKKCMCRowValue>();
		list = GetDataFromExcelTmp(sViewImport, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart);
		hTKKCMCTableInfo.RowsValues.AddRange(list);
		sheetInfos.IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, list.Count - 1);
		foreach (HTKKCMCTableInfo table in sheetInfos.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > hTKKCMCTableInfo.Location.DataLocationTemplateRowStart)
			{
				table.Increment(list.Count - 1);
			}
		}
	}

	public static void DelNullRow(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count - 1; i++)
			{
				for (int j = 0; j < sInfo.Tables.Count; j++)
				{
					HTKKCMCRowInfo rowsInfo = sInfo.Tables[j].RowsInfo;
					List<HTKKCMCRowValue> rowsValues = sInfo.Tables[j].RowsValues;
					for (int k = 0; k < rowsValues.Count; k++)
					{
						List<string> rowValue = rowsValues[k].RowValue;
						bool flag = true;
						int num = sInfo.Tables[j].Location.DataLocationTemplateRowStart - 1 + k;
						for (int l = 0; l < rowValue.Count; l++)
						{
							string text = rowsInfo.Cells[l].DefaultValue;
							if (!(text == "NA"))
							{
								string text2 = ConvertHelpers.ConvertToString(sView.Cells[num, rowsInfo.Cells[l].CellTemplate.ColIndex - 1].Value);
								if (rowsInfo.Cells[l].Controltype == "16")
								{
									text = ConvertHelpers.ConvertToDouble(rowsInfo.Cells[l].DefaultValue).ToString();
									text2 = ConvertHelpers.ConvertToDouble(text2).ToString();
								}
								if (!text2.Equals(text))
								{
									flag = false;
									break;
								}
							}
						}
						if (flag)
						{
							if (sInfo.Tables[j].RowsValues.Count > 1)
							{
								sInfo.Tables[j].RowsValues.RemoveAt(k);
								RemoveDataRows(sView, num, 1, sInfo);
								fpSpReindexingTableTemplate(sView, sInfo.Tables[j]);
								k--;
							}
							if (sInfo.Tables[j].TableID != "1")
							{
								RemoveTable(sView, j, sInfo);
								fpSpReindexingSheetTemplate(sView, sInfo);
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void InserRowTKDkNPT(SheetView sView, HTKKCMCSheetInfo sInfo, List<HTKKCMCRowValue> rowsValues)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			InsertTableRows(sView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, -1, hTKKCMCTableInfo.RowsValues.Count + hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, -1, rowsValues.Count, -1, sInfo);
			RemoveDataRows(sView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count, sInfo);
			hTKKCMCTableInfo.RowsValues.Clear();
			for (int i = 0; i < rowsValues.Count; i++)
			{
				for (int j = 0; j < hTKKCMCTableInfo.RowsInfo.Cells.Count; j++)
				{
					try
					{
						fpSpBindDataCell(sView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i + hTKKCMCTableInfo.RowsValues.Count, hTKKCMCTableInfo.RowsInfo.Cells[j].CellTemplate.ColIndex, hTKKCMCTableInfo.RowsInfo.Cells[j].Controltype, rowsValues[i].RowValue[j]);
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
			}
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				hTKKCMCTableInfo.RowsValues.InsertRange(hTKKCMCTableInfo.RowsValues.Count - 1, rowsValues);
			}
			else
			{
				hTKKCMCTableInfo.RowsValues.InsertRange(hTKKCMCTableInfo.RowsValues.Count, rowsValues);
			}
			DelNullRow(sView, sInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int GetLastDataRow(int numCell, HTKKCMCSheetInfo sheetInfo)
	{
		int result = -1;
		try
		{
			List<HTKKCMCCellInfo> staticCells = sheetInfo.StaticCells;
			HTKKCMCCellInfo hTKKCMCCellInfo = null;
			HTKKCMCCellBasicInfo hTKKCMCCellBasicInfo = null;
			if (staticCells.Count > numCell)
			{
				hTKKCMCCellInfo = staticCells[staticCells.Count - numCell];
			}
			if (hTKKCMCCellInfo != null)
			{
				hTKKCMCCellBasicInfo = hTKKCMCCellInfo.CellReport;
			}
			if (hTKKCMCCellBasicInfo != null)
			{
				result = hTKKCMCCellBasicInfo.RowIndex;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static int[] GetRowG1G2(HTKKCMCSheetInfo sheetInfo)
	{
		try
		{
			int[] array = new int[2];
			List<HTKKCMCCellInfo> staticCells = sheetInfo.StaticCells;
			HTKKCMCCellInfo hTKKCMCCellInfo = null;
			HTKKCMCCellBasicInfo hTKKCMCCellBasicInfo = null;
			if (staticCells.Count > 10)
			{
				hTKKCMCCellInfo = staticCells[staticCells.Count - 26];
			}
			if (hTKKCMCCellInfo != null)
			{
				hTKKCMCCellBasicInfo = hTKKCMCCellInfo.CellReport;
			}
			if (hTKKCMCCellBasicInfo != null)
			{
				array[0] = hTKKCMCCellBasicInfo.RowIndex;
			}
			if (staticCells.Count > 23)
			{
				hTKKCMCCellInfo = staticCells[staticCells.Count - 12];
			}
			if (hTKKCMCCellInfo != null)
			{
				hTKKCMCCellBasicInfo = hTKKCMCCellInfo.CellReport;
			}
			if (hTKKCMCCellBasicInfo != null)
			{
				array[1] = hTKKCMCCellBasicInfo.RowIndex;
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void ResetColorCellDefault(SheetView sView, Color color, int iRow, int iCol)
	{
		try
		{
			sView.Cells[iRow - 1, iCol - 1].BackColor = color;
			sView.Cells[iRow - 1, iCol - 1].ResetNote();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ResetColorSheetDefault(FpSpread winFpSpread, HTKKCMCToKhaiInfo tKhaiInfo, int activeSheet)
	{
		try
		{
			if (winFpSpread.Sheets[activeSheet].SheetName == "KHBS")
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = tKhaiInfo.SheetInfos[activeSheet].Tables;
			for (int i = 0; i < tKhaiInfo.SheetInfos[activeSheet].StaticCells.Count; i++)
			{
				ResetColorCellDefault(winFpSpread.Sheets[activeSheet], tKhaiInfo.SheetInfos[activeSheet].StaticCells[i].Validate.ColorDefault, tKhaiInfo.SheetInfos[activeSheet].StaticCells[i].CellTemplate.RowIndex, tKhaiInfo.SheetInfos[activeSheet].StaticCells[i].CellTemplate.ColIndex);
			}
			for (int i = 0; i < tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = tables[i];
				for (int j = 0; j < hTKKCMCTableInfo.Header.StaticCells.Count; j++)
				{
					ResetColorCellDefault(winFpSpread.Sheets[activeSheet], hTKKCMCTableInfo.Header.StaticCells[j].Validate.ColorDefault, hTKKCMCTableInfo.Header.StaticCells[j].CellTemplate.RowIndex, hTKKCMCTableInfo.Header.StaticCells[j].CellTemplate.ColIndex);
				}
				for (int j = 0; j < hTKKCMCTableInfo.RowsInfo.Cells.Count; j++)
				{
					ResetColorCellDefault(winFpSpread.Sheets[activeSheet], hTKKCMCTableInfo.RowsInfo.Cells[j].Validate.ColorDefault, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsInfo.Cells[j].CellTemplate.ColIndex);
				}
			}
			winFpSpread.TabStrip[activeSheet].BackColor = Color.White;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void reloadBc26AcPl310(FpSpread winFpSpread)
	{
		try
		{
			if (!(WinSettings.MenuID == "68") || !winFpSpread.Sheets[2].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[2];
			RemoveDataRows(winFpSpread.Sheets[2], hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart, hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1, hTKKCMCSheetInfo);
			hTKKCMCSheetInfo.Tables[0].RowsValues.Clear();
			List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
			{
				HTKKCMCRowValue hTKKCMCRowValue = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues[i];
				HTKKCMCRowInfo rowsInfo = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsInfo;
				HTKKCMCRowInfo rowsInfo2 = WinSettings.ToKhaiInfo.SheetInfos[2].Tables[0].RowsInfo;
				List<string> list2 = new List<string>();
				string text = "";
				text = winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[0].CellTemplate.ColIndex - 1].Text;
				text = text + "~" + winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[1].CellTemplate.ColIndex - 1].Text;
				text = text + "~" + winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[2].CellTemplate.ColIndex - 1].Text;
				text = text + "~" + winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[3].CellTemplate.ColIndex - 1].Text;
				string text2 = "";
				text2 = winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[21].CellTemplate.ColIndex - 1].Text;
				text2 = text2 + "~" + winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[22].CellTemplate.ColIndex - 1].Text;
				text2 = text2 + "~" + winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[23].CellTemplate.ColIndex - 1].Value.ToString();
				string[] array = new string[6];
				if (WinSettings.ToKhaiInfo.SheetInfos[0].dicVar.ContainsKey(text))
				{
					array = WinSettings.ToKhaiInfo.SheetInfos[0].dicVar[text].ToString().Split(new char[1] { '~' });
					WinSettings.ToKhaiInfo.SheetInfos[0].dicVar.Remove(text);
					string[] array2 = text2.Split(new char[1] { '~' });
					if (array[0] == "#" || array[1] == "#" || array[2] == "#")
					{
						string text3 = text2;
						text2 = text3 + "~" + array[3] + "~" + array[4] + "~" + array[5];
					}
					else if (array2[0] != array[0] || array2[1] != array[1])
					{
						text2 = text2 + "~" + text2;
					}
					else
					{
						string text3 = text2;
						text2 = text3 + "~" + array[3] + "~" + array[4] + "~" + array[5];
					}
				}
				else
				{
					text2 = text2 + "~" + text2;
				}
				WinSettings.ToKhaiInfo.SheetInfos[0].dicVar.Add(text, text2);
				for (int j = 0; j < rowsInfo2.Cells.Count; j++)
				{
					switch (j)
					{
					case 0:
						list2.Add(winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[0].CellTemplate.ColIndex - 1].Text);
						break;
					case 1:
						list2.Add(winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[1].CellTemplate.ColIndex - 1].Text);
						break;
					case 12:
						list2.Add(winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[2].CellTemplate.ColIndex - 1].Text);
						break;
					case 16:
						list2.Add(winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[3].CellTemplate.ColIndex - 1].Text);
						break;
					case 18:
						list2.Add(text2.Split(new char[1] { '~' })[5]);
						break;
					case 19:
						list2.Add(text2.Split(new char[1] { '~' })[3]);
						break;
					case 22:
						list2.Add(text2.Split(new char[1] { '~' })[4]);
						break;
					case 24:
						list2.Add(text2.Split(new char[1] { '~' })[5]);
						break;
					default:
						list2.Add("");
						break;
					}
				}
				if (winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[21].CellTemplate.ColIndex - 1].Value.ToNumber() > 0m || winFpSpread.Sheets[0].Cells[WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1, rowsInfo.Cells[22].CellTemplate.ColIndex - 1].Value.ToNumber() > 0m)
				{
					hTKKCMCSheetInfo.Tables[0].RowsValues.Add(new HTKKCMCRowValue
					{
						RowValue = list2,
						RowDisplay = "1"
					});
				}
			}
			if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count == 0)
			{
				List<string> list2 = new List<string>();
				HTKKCMCRowInfo rowsInfo2 = WinSettings.ToKhaiInfo.SheetInfos[2].Tables[0].RowsInfo;
				for (int j = 0; j < rowsInfo2.Cells.Count; j++)
				{
					list2.Add(rowsInfo2.Cells[j].DefaultValue);
				}
				hTKKCMCSheetInfo.Tables[0].RowsValues.Add(new HTKKCMCRowValue
				{
					RowValue = list2,
					RowDisplay = "1"
				});
			}
			hTKKCMCSheetInfo.IncrementStaticCell(hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart, hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1);
			ProcessSheet(2, winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
			int num = WinSettings.ToKhaiInfo.InterfaceTemplateExcel.LastIndexOf('\\');
			string text4 = WinSettings.ToKhaiInfo.InterfaceTemplateExcel.Substring(0, num + 1);
			text4 += "Bangke_BC26_Mau_3_10.xls";
			GeneralFunctions.InsertHyperLink(winFpSpread.Sheets[2], 2, HTKKFile.ColumnLettersToIndex("B"), "Mẫu tải bảng kê", text4);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<item> GetDanhMucItem(string nameFunction, bool all = true)
	{
		List<item> list = new List<item>();
		try
		{
			return WinDoWorkEvent.getItemsDm(nameFunction);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static bool CheckNghiaVuNTTNCN(FpSpread winFpSpread)
	{
		try
		{
			bool result = true;
			if (WinSettings.MenuID == "37")
			{
				string obj = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "M_65"));
				if (ConvertHelpers.ConvertToDouble(obj) <= 50000.0 && ConvertHelpers.ConvertToDouble(obj) > 0.0)
				{
					result = false;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public static void ProcessingLargeDataOnSpread(SheetView sView, int sheetIndex, HTKKCMCTableInfo table)
	{
		Dictionary<string, string> mapTaiDuLieuLon = GetMapTaiDuLieuLon(sheetIndex);
		List<int> list = new List<int>();
		if (mapTaiDuLieuLon != null && mapTaiDuLieuLon.Count > 0 && mapTaiDuLieuLon.ContainsKey("ColsFormula"))
		{
			string text = mapTaiDuLieuLon["ColsFormula"];
			if (!string.IsNullOrEmpty(text))
			{
				string[] array = text.Split(new char[1] { ',' });
				foreach (string column in array)
				{
					int item = HTKKFile.ColumnLettersToIndex(column);
					list.Add(item);
				}
			}
		}
		int num = -1;
		num = ((mapTaiDuLieuLon == null || mapTaiDuLieuLon.Count <= 0 || !mapTaiDuLieuLon.ContainsKey("ColBoderRight") || string.IsNullOrEmpty(mapTaiDuLieuLon["ColBoderRight"])) ? (table.RowsInfo.Cells[table.RowsInfo.Cells.Count - 1].CellTemplate.ColIndex + 1) : Convert.ToInt32(HTKKFile.ColumnLettersToIndex(mapTaiDuLieuLon["ColBoderRight"])));
		GeneralFunctions.BindDataToSviewFromTable(sView, table, list, num);
	}

	public static Dictionary<string, string> GetMapTaiDuLieuLon(int sheetIndex)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		//IL_0085: Unknown result type (might be due to invalid IL or missing references)
		//IL_008c: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load("MapTaiDuLieuLon.xml");
			XmlNodeList source = ((XmlNode)val.DocumentElement).SelectNodes("/Root/Map");
			XmlNode val2 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["ID"]).Value == WinSettings.MenuID));
			if (val2 == null)
			{
				return null;
			}
			foreach (XmlNode childNode in val2.ChildNodes)
			{
				XmlNode val3 = childNode;
				string text = ((val3.Attributes["SheetIndex"] == null) ? "" : ((XmlNode)val3.Attributes["SheetIndex"]).Value.Trim());
				string value = ((val3.Attributes["ColsFormula"] == null) ? "" : ((XmlNode)val3.Attributes["ColsFormula"]).Value.Trim());
				string value2 = ((val3.Attributes["ColBoderRight"] == null) ? "" : ((XmlNode)val3.Attributes["ColBoderRight"]).Value.Trim());
				if (!string.IsNullOrEmpty(text) && sheetIndex.ToString() == text)
				{
					if (!dictionary.ContainsKey("ColsFormula"))
					{
						dictionary.Add("ColsFormula", value);
					}
					else
					{
						dictionary["ColsFormula"] = value;
					}
					if (!dictionary.ContainsKey("ColBoderRight"))
					{
						dictionary.Add("ColBoderRight", value2);
					}
					else
					{
						dictionary["ColBoderRight"] = value2;
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void InsertTableRowsLargeDataNoSView(int frRow, int frCol, int toRow, int toCol, int rowCount, int colCount, HTKKCMCSheetInfo sCMCInfo)
	{
		if (rowCount == 0)
		{
			return;
		}
		sCMCInfo.IncrementStaticCell(frRow, rowCount);
		foreach (HTKKCMCTableInfo table in sCMCInfo.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > frRow)
			{
				table.Increment(rowCount);
			}
		}
	}

	public static Dictionary<string, string> GetDMDynamic(string fileName)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\" + fileName;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Ma"]).Value;
				string value2 = ((XmlNode)val2.Attributes["Ten"]).Value;
				dictionary.Add(value, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetDMDynamicByParent(string parentCode, string fileName)
	{
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Expected O, but got Unknown
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.INTERFACE_INI + "\\" + fileName;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (string.IsNullOrEmpty(parentCode))
			{
				return dictionary;
			}
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["MaCha"]).Value;
				if (value == parentCode)
				{
					string value2 = ((XmlNode)val2.Attributes["Ma"]).Value;
					string value3 = ((XmlNode)val2.Attributes["Ten"]).Value;
					dictionary.Add(value2, value3);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static HTKKCMCSheetInfo GetDataFromExcelLarge1(FpSpread BangKeWinFpSpread, HTKKCMCSheetInfo sInfo, ref string status)
	{
		try
		{
			List<string> list = new List<string>();
			if (WinSettings.MenuID == "P2")
			{
				list.Add("II");
			}
			if (WinSettings.MenuID == "X6")
			{
				list.Add("II");
				list.Add("Thông tin cá nhân nước ngoài");
			}
			int num = 0;
			int num2 = ((sInfo.Tables[num].Location.DataLocationTaiBangKeRowStart >= 0) ? sInfo.Tables[num].Location.DataLocationTaiBangKeRowStart : 0);
			for (int i = num2 - 1; i <= BangKeWinFpSpread.Sheets[BangKeWinFpSpread.ActiveSheetIndex].NonEmptyRowCount - 1; i++)
			{
				try
				{
					HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
					while (true)
					{
						IL_00a6:
						HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
						HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
						int num3 = 0;
						while (true)
						{
							if (num3 < rowsInfo.Cells.Count)
							{
								try
								{
									if (rowsInfo.Cells[num3].CellBangKe != null)
									{
										if (rowsInfo.Cells[num3].CellBangKe.ColIndex > 0)
										{
											Cell cell = BangKeWinFpSpread.ActiveSheet.Cells[i, rowsInfo.Cells[num3].CellBangKe.ColIndex - 1];
											if (rowsInfo.Cells[num3].Controltype == "14")
											{
												string text = "";
												if (!HTKKFormater.isNullString(cell.Text.Trim()))
												{
													if (cell.Text.IndexOf("/") > -1)
													{
														text = cell.Text.Trim();
													}
													else if (ConvertHelpers.ConvertToDouble(cell.Value) > 0.0)
													{
														text = DateTime.FromOADate(Convert.ToDouble(cell.Value, CultureInfo.InvariantCulture)).ToString("dd/MM/yyyy");
													}
												}
												hTKKCMCRowValue.RowValue.Add(GeneralFunctions.RemoveSpecialCharactersUnicode(text.Replace("/", "")));
											}
											else if (rowsInfo.Cells[num3].Controltype == "16")
											{
												if (cell.Value != null)
												{
													if (!GeneralFunctions.IsNumeric(cell.Value.ToString()))
													{
														status = status + HTKKFile.ColumnIndexToLetters(rowsInfo.Cells[num3].CellBangKe.ColIndex) + (i + 1) + ",";
														hTKKCMCRowValue.RowValue.Add("0");
													}
													else
													{
														hTKKCMCRowValue.RowValue.Add(GeneralFunctions.RemoveSpecialCharactersUnicode(cell.Value.ToString()));
													}
												}
												else
												{
													hTKKCMCRowValue.RowValue.Add("0");
												}
											}
											else if (rowsInfo.Cells[num3].Controltype == "2")
											{
												hTKKCMCRowValue.RowValue.Add((!string.IsNullOrEmpty(cell.Text.Trim()) && cell.Text.Trim().ToLower() == "x") ? "1" : "0");
											}
											else
											{
												if (num3 == 0 && list.Count > 0 && list.Contains(cell.Text.Trim()))
												{
													i = ((!(WinSettings.MenuID == "X6") || num != 0) ? (i + 1) : (i + 5));
													num++;
													goto IL_00a6;
												}
												hTKKCMCRowValue.RowValue.Add(GeneralFunctions.RemoveSpecialCharactersUnicode(cell.Text.Trim()));
											}
										}
										else
										{
											hTKKCMCRowValue.RowValue.Add((rowsInfo.Cells[num3].Controltype == "16") ? "0" : "");
										}
									}
									else
									{
										hTKKCMCRowValue.RowValue.Add((rowsInfo.Cells[num3].Controltype == "16") ? "0" : "");
									}
								}
								catch (Exception ex)
								{
									HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
								}
								num3++;
								continue;
							}
							if (GeneralFunctions.HaveToValidateRow(rowsInfo, hTKKCMCRowValue.RowValue.ToArray()))
							{
								hTKKCMCTableInfo.RowsValues.Add(hTKKCMCRowValue);
								int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
								if (hTKKCMCTableInfo.RowsValues.Count > 1)
								{
									GeneralFunctions.InsertTableRowNoSView(dataLocationTemplateRowStart, 1, sInfo);
								}
							}
							break;
						}
						break;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		status = ((status.Length < 1) ? "" : status.Substring(0, status.Length - 1));
		return sInfo;
	}
}
