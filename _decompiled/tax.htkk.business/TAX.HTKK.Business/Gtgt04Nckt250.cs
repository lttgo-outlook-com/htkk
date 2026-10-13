using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt04Nckt250
{
	public void AutoBindData_0(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) < 2021)
			{
				return;
			}
			if (PeriodInfo.ThuHo.Equals("TH"))
			{
				sView.Cells[4, HTKKFile.ColumnLettersToIndex("K") - 1].Value = true;
				sView.Cells[42, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = "0";
				sView.Cells[42, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = true;
				sView.Cells[44, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = "0";
				sView.Cells[44, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = true;
				sView.Cells[45, HTKKFile.ColumnLettersToIndex("V") - 1].Text = "10%";
				sView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q46*10/100,0)";
				sView.Cells[45, HTKKFile.ColumnLettersToIndex("X") - 1].Text = "[29]=[28]x10%";
				return;
			}
			CheckBoxCellType cellType = new CheckBoxCellType();
			sView.Cells[4, HTKKFile.ColumnLettersToIndex("K") - 1].CellType = cellType;
			sView.Cells[4, HTKKFile.ColumnLettersToIndex("K") - 1].Value = false;
			sView.Cells[42, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = false;
			sView.Cells[44, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = false;
			sView.Cells[45, HTKKFile.ColumnLettersToIndex("V") - 1].Text = "2%";
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				sView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q46*2/100,0)";
			}
			sView.Cells[45, HTKKFile.ColumnLettersToIndex("X") - 1].Text = "[29]=[28]x2%";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				sView.FpSpread.Sheets[0].Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
				sView.FpSpread.Sheets[0].Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
				sView.FpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
				sView.FpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
			}
			else
			{
				sView.FpSpread.Sheets[0].Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				sView.FpSpread.Sheets[0].Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				sView.FpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				sView.FpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					continue;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
						continue;
					}
					break;
				}
				winFpSpread.TabStrip[key].BackColor = Color.White;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKFile.RemoveValidateBackGround(sheetView, 46, HTKKFile.ColumnLettersToIndex("Z") - 1);
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int column = HTKKFile.ColumnLettersToIndex("G") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("I") - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			string message = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_20"));
			string message2 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_26"));
			string message3 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_27"));
			string message4 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_28"));
			string message5 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_29"));
			int num4 = 0;
			double num5 = 0.0;
			double num6 = 0.0;
			double num7 = 0.0;
			double num8 = 0.0;
			for (int i = 0; i < tables.Count; i++)
			{
				num4 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				num5 = 0.0;
				num6 = 0.0;
				num7 = 0.0;
				num8 = 0.0;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView2, num4 + j, num);
					HTKKFile.RemoveValidate(sheetView2, num4 + j, num2);
					HTKKFile.RemoveValidate(sheetView2, num4 + j, num3);
					string text = sheetView2.Cells[num4 + j, num].Text;
					double num9 = double.Parse(sheetView2.Cells[num4 + j, num2].Value.ToString());
					string text2 = sheetView2.Cells[num4 + j, num3].Text;
					string text3 = sheetView2.Cells[num4 + j, column2].Text;
					string obj = sheetView2.Cells[num4 + j, column].Value.ToString();
					switch (text3)
					{
					case "1":
						num5 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "2":
						num6 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "3":
						num7 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "5":
						num8 += ConvertHelpers.ConvertToDouble(obj);
						break;
					}
					if (string.IsNullOrEmpty(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num4 + j, num, message);
						value = 2;
					}
					if (num9 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num4 + j, num2, message);
						value = 2;
					}
					if (string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num4 + j, num3, message);
						value = 2;
					}
				}
			}
			List<HTKKCMCCellInfo> staticCells = ToKhaiInfo.SheetInfos[1].StaticCells;
			double num10 = ConvertHelpers.ConvertToDouble(sheetView.Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num11 = ConvertHelpers.ConvertToDouble(sheetView.Cells[42, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num13 = ConvertHelpers.ConvertToDouble(sheetView.Cells[43, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num14 = ConvertHelpers.ConvertToDouble(sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num16 = ConvertHelpers.ConvertToDouble(sheetView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num17 = ConvertHelpers.ConvertToDouble(sheetView.Cells[45, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			HTKKFile.RemoveValidateBackGround(sheetView, 42, HTKKFile.ColumnLettersToIndex("Z") - 1);
			HTKKFile.RemoveValidateBackGround(sheetView, 43, HTKKFile.ColumnLettersToIndex("Z") - 1);
			HTKKFile.RemoveValidateBackGround(sheetView, 44, HTKKFile.ColumnLettersToIndex("Z") - 1);
			HTKKFile.RemoveValidateBackGround(sheetView, 45, HTKKFile.ColumnLettersToIndex("Z") - 1);
			if (num10 != num11 - num5)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 42, HTKKFile.ColumnLettersToIndex("Z") - 1, message2);
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			if (num12 != num13 - num8)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 43, HTKKFile.ColumnLettersToIndex("Z") - 1, message3);
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			if (num14 != num15 - num7)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 44, HTKKFile.ColumnLettersToIndex("Z") - 1, message4);
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			if (num16 != num17 - num6)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 45, HTKKFile.ColumnLettersToIndex("Z") - 1, message5);
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("Q");
				int num2 = HTKKFile.ColumnLettersToIndex("Z");
				if (iCol == num && iRow == 43)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Formula = "ROUND(Q43*1/100,0)";
				}
				if (iCol == num && iRow == 44)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Formula = "ROUND(Q44*5/100,0)";
				}
				if (iCol == num && iRow == 45)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Formula = "ROUND(Q45*3/100,0)";
				}
				if (iCol == num && iRow == 46)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Formula = "ROUND(Q46*2/100,0)";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			if (hTKKCMCSheetInfo2 == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo2.Tables;
			int column = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int column3 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int num2 = 0;
			for (int i = 0; i < tables.Count; i++)
			{
				num2 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sheetView2.Cells[num2 + j, column].Text;
					string text2 = sheetView2.Cells[num2 + j, column3].Text;
					sheetView2.Cells[num2 + j, column2].Text = getTenThueSuat(text2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private string getTenThueSuat(string maThueSuat)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		string result = "";
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_thue_suat_GTGT.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 2 && maThueSuat == array[0])
			{
				result = array[1];
				break;
			}
		}
		return result;
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[0];
			if (hTKKCMCSheetInfo != null)
			{
				sheetView.Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
				sheetView.Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
				sheetView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = false;
			}
			else
			{
				sheetView.Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				sheetView.Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				sheetView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
