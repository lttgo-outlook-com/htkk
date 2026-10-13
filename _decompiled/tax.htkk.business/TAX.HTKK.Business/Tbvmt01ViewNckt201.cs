using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tbvmt01ViewNckt201
{
	private Dictionary<string, string[]> dicDMThueBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueBVMT4 = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMThueBVMTTmp = new Dictionary<string, string>();

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_Q => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_U => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxCol_AB => HTKKFile.ColumnLettersToIndex("AC") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			string text2 = "#010108#0106#020109#020120#020121#020119#020112#020205#010102#020206#";
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sView.Cells[num + i, idxCol_AB].Text;
				if (text2.Contains("#" + text + "#"))
				{
					sView.Cells[num + i, idxCol_Q].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int num3 = HTKKFile.ColumnLettersToIndex("U");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("AC");
			string text = "";
			string text2 = "#010108#0106#020109#020120#020121#020119#020112#020205#010102#020206#";
			text = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
			string text3 = "ROUND((L###*Q###),0)";
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + rowsValues.Count && ((int)iCol == num || (int)iCol == num2))
			{
				sheetView.Cells[(int)iRow - 1, num3 - 1].Formula = text3.Replace("###", iRow.ToString());
			}
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + rowsValues.Count && ((int)iCol == num || (int)iCol == num4))
			{
				if (text2.Contains("#" + text + "#"))
				{
					sheetView.Cells[(int)iRow - 1, num4 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num4 - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			string text = "#010108#0106#020109#020120#020121#020119#020112#020205#010102#020206#";
			if (iCol == num)
			{
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length == 4)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[3];
				}
				if (text.Contains("#" + array[0] + "#"))
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("M");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("D");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			int num5 = HTKKFile.ColumnLettersToIndex("Z");
			if (iCol == num)
			{
				string text = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc[sheetView.Cells[(int)iRow - 1, num - 1].Text.ToString().Trim()];
				text = text.Substring(0, 3);
				string[] array = new string[1];
				int num6 = 1;
				foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy)
				{
					if (item.Key.IndexOf(text) == 0)
					{
						string[] array3;
						string[] array2 = (array3 = array);
						int num7 = num6 - 1;
						nint num8 = num7;
						array2[num7] = array3[num8] + item.Value;
						num6++;
						Array.Resize(ref array, num6);
					}
				}
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = array;
				sheetView.Cells[(int)iRow - 1, num2 - 1].CellType = comboBoxCellType;
			}
			if (iCol == num3)
			{
				string[] array4 = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array4.Length == 4)
				{
					sheetView.Cells[(int)iRow - 1, num4 - 1].Text = array4[3];
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = array4[0];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, idxCol_C].Value);
				array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, idxCol_J].Value);
				array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, idxCol_L].Value);
				array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, idxCol_Q].Value);
				array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, idxCol_U].Value);
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_C);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_C, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_C);
				}
				if (!string.IsNullOrEmpty(array[0]) && ConvertHelpers.ConvertToDouble(array[2]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, idxCol_L, message);
					value = 1;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_L);
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = (winFpSpread.Sheets[1].Visible ? winFpSpread.Sheets[1] : null);
			if (sheetView == null)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<int> list = new List<int>();
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("L");
			int num4 = HTKKFile.ColumnLettersToIndex("M");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = HTKKFile.ColumnLettersToIndex("S");
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("V");
			int num10 = HTKKFile.ColumnLettersToIndex("W");
			int num11 = 0;
			double num12 = 0.0;
			bool flag = true;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					num11 = dataLocationTemplateRowStart + j;
					string text6 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Value);
					string str = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Value);
					string valueCot4 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Value);
					string text7 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Value);
					string text8 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Value);
					string text9 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Value);
					string text10 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num7 - 1].Value);
					string obj = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Value);
					string text11 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Value);
					string text12 = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Value);
					string text13 = "";
					if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, rowsValues[j].RowValue.ToArray()))
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						continue;
					}
					flag = false;
					if (HTKKFormater.isNullString(str))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text4);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						list.Add(0);
					}
					int num13 = rowsValues.Count((HTKKCMCRowValue p) => p.RowValue[3].Equals(valueCot4));
					if (HTKKFormater.isNullString(valueCot4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text5);
						list.Add(1);
					}
					else if (!ValidateBase.CheckMaSoThue(valueCot4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text2);
						list.Add(1);
					}
					else if (num13 > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						list.Add(0);
					}
					num12 += ConvertHelpers.ConvertToDouble(obj);
				}
			}
			if (num12 != 100.0 && !flag)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num11 + 2 - 1, num8 - 1, text);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, num11 + 2 - 1, num8 - 1);
				value = 0;
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					value = 2;
					break;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMThueBVMT(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Expected O, but got Unknown
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMThueBVMTTmp.Count > 0)
			{
				return dicDMThueBVMTTmp;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
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
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string key = array2[0];
				string value2 = array2[1];
				dictionary2.Add(key, value2);
				dictionary.Add(array2[1], array2);
			}
			dicDMThueBVMT = dictionary;
			dicDMThueBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThan(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Expected O, but got Unknown
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMThueBVMT4.Count > 0)
			{
				return dicDMThueBVMT4;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
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
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string text4 = array2[0];
				string value2 = array2[1];
				if (text4.Substring(0, 4).Equals("0102"))
				{
					dictionary2.Add(text4, value2);
				}
			}
			dicDMThueBVMT4 = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			string text2 = "#010108#0106#020109#020120#020121#020119#020112#020205#010102#020206#";
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num2 - 1].Text;
				if (text2.Contains("#" + text + "#"))
				{
					sheetView.Cells[num + i, idxCol_Q].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
