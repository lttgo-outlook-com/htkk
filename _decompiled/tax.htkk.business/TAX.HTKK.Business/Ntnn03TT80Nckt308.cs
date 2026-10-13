using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn03TT80Nckt308
{
	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.KieuKy == "D")
			{
				int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
				int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				for (int i = num; i <= num2; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.KieuKy == "D")
			{
				int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
				int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				for (int i = num; i <= num2; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("M") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("AD") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("AL") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("AR") - 1;
			int num7 = HTKKFile.ColumnLettersToIndex("BE") - 1;
			int num8 = HTKKFile.ColumnLettersToIndex("BK") - 1;
			int num9 = HTKKFile.ColumnLettersToIndex("CR") - 1;
			int num10 = HTKKFile.ColumnLettersToIndex("BU") - 1;
			int num11 = HTKKFile.ColumnLettersToIndex("BX") - 1;
			int column = HTKKFile.ColumnLettersToIndex("CQ") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value2 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num2].Value);
				string value3 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num3].Value);
				string text = ConvertHelpers.ConvertToString(sView.Cells[num + i, num4].Value);
				string value4 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num5].Value);
				string text2 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num6].Value);
				string text3 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num7].Value);
				string obj = ConvertHelpers.ConvertToString(sView.Cells[num + i, num8].Value);
				string obj2 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num10].Value);
				string obj3 = ConvertHelpers.ConvertToString(sView.Cells[num + i, num11].Value);
				string value5 = ConvertHelpers.ConvertToString(sView.Cells[num + i, column].Value);
				HTKKFile.RemoveValidate(sView, num + i, num2);
				HTKKFile.RemoveValidate(sView, num + i, num3);
				HTKKFile.RemoveValidate(sView, num + i, num4);
				HTKKFile.RemoveValidate(sView, num + i, num5);
				HTKKFile.RemoveValidate(sView, num + i, num6);
				HTKKFile.RemoveValidate(sView, num + i, num7);
				HTKKFile.RemoveValidate(sView, num + i, num11);
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else if (string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30")));
					value = 2;
				}
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (!string.IsNullOrEmpty(text) & !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				if (string.IsNullOrEmpty(value4))
				{
				}
				if (!string.IsNullOrEmpty(text2))
				{
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(text2.Replace("/", ""), "ddMMyyyy");
					if (dateTime == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_28")));
						value = 2;
					}
					else if (dateTime.CompareTo(DateTime.Today) >= 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text3))
				{
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text3.Replace("/", ""), "ddMMyyyy");
					if (!ValidateBase.ValidateFullDate(text3, "ddMMyyyy"))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					else if (dateTime2 > DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture) || dateTime2 < DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
						value = 2;
					}
				}
				if (ConvertHelpers.ConvertToDouble(obj3) > Math.Round(ConvertHelpers.ConvertToDouble(obj2) * ConvertHelpers.ConvertToDouble(obj) / 100.0, 0))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					value = 2;
				}
				HTKKFile.RemoveValidate(sView, num + i, num8);
				HTKKFile.RemoveValidate(sView, num + i, num10);
				if (ConvertHelpers.ConvertToDouble(obj) == 0.0)
				{
				}
				if (ConvertHelpers.ConvertToDouble(obj2) == 0.0)
				{
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private string getThueSuat(string maNNKD)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Expected O, but got Unknown
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		try
		{
			string result = "";
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
				if (array.Length == 4 && maNNKD.Equals(array[0]))
				{
					result = array[2] + "_" + array[3];
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
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
			int column = HTKKFile.ColumnLettersToIndex("CR") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("BU") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, column].Value);
				text = text.Replace('.', ',');
				sheetView.Cells[num + i, column2].Text = GeneralFunctions.getThuSuatByMa(text, "DC_97");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
