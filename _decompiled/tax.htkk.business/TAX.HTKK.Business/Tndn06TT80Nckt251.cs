using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn06TT80Nckt251
{
	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("I_66");
			int[] array2 = GeneralFunctions.ArrayCellId("BL_66");
			int[] array3 = GeneralFunctions.ArrayCellId("AL_66");
			int[] array4 = GeneralFunctions.ArrayCellId("BM_66");
			int[] array5 = GeneralFunctions.ArrayCellId("I_68");
			int[] array6 = GeneralFunctions.ArrayCellId("BL_68");
			string value = sheetView.Cells[array3[1], array3[1]].Text.Trim();
			if (!string.IsNullOrEmpty(value))
			{
				return;
			}
			KeyValuePair<string, string> keyValuePair = default(KeyValuePair<string, string>);
			string maCQTCapCucDefault = ((!string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc)) ? PeriodInfo.maCQTCapCuc : TTinChung.maCQTCapCuc);
			if (!string.IsNullOrEmpty(maCQTCapCucDefault))
			{
				keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCucDefault.Substring(0, 3));
				sheetView.Cells[array[1], array[0]].Text = ((!string.IsNullOrEmpty(keyValuePair.Key)) ? keyValuePair.Key : "");
				sheetView.Cells[array2[1], array2[0]].Text = ((!string.IsNullOrEmpty(keyValuePair.Value)) ? keyValuePair.Value : "");
			}
			string value2 = ((!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy)) ? PeriodInfo.maCQTQuanLy : TTinChung.maCQTQuanLy);
			if (!string.IsNullOrEmpty(value2) || !string.IsNullOrEmpty(maCQTCapCucDefault))
			{
			}
			if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(maCQTCapCucDefault))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("I_66");
			int[] array2 = GeneralFunctions.ArrayCellId("BL_66");
			int[] array3 = GeneralFunctions.ArrayCellId("AL_66");
			int[] array4 = GeneralFunctions.ArrayCellId("BM_66");
			int[] array5 = GeneralFunctions.ArrayCellId("I_68");
			int[] array6 = GeneralFunctions.ArrayCellId("BL_68");
			int[] array7 = GeneralFunctions.ArrayCellId("I_71");
			string text = sView.Cells[array7[1], array7[0]].Text.Trim();
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (!hTKKCMCSheetInfo.IsReadDataFile && PeriodInfo.LoaiToKhai == "C")
			{
				string value = ((!string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc)) ? PeriodInfo.maCQTCapCuc : TTinChung.maCQTCapCuc);
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == PeriodInfo.MaTinh);
				sView.Cells[array[1], array[0]].Value = ((!string.IsNullOrEmpty(keyValuePair.Key)) ? keyValuePair.Key : "");
				sView.Cells[array2[1], array2[0]].Value = ((!string.IsNullOrEmpty(keyValuePair.Value)) ? keyValuePair.Value : "");
				string value2 = ((!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy)) ? PeriodInfo.maCQTQuanLy : TTinChung.maCQTQuanLy);
				if (!string.IsNullOrEmpty(value2) || !string.IsNullOrEmpty(value))
				{
				}
				if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(value))
				{
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
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			int[] array = GeneralFunctions.ArrayCellId("AL_57");
			int[] array2 = GeneralFunctions.ArrayCellId("AC_76");
			int[] array3 = GeneralFunctions.ArrayCellId("AH_78");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43"));
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			string text = sheetView.Cells[array[1], array[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				if (text.Length == 10 || text.Length == 13)
				{
					if (!ValidateBase.CheckMaSoThue(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message2);
						value = 2;
					}
				}
				else if (text.Length != 9 && text.Length != 12)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message3);
					value = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message);
				value = 2;
			}
			string text2 = sheetView.Cells[array2[1], array2[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			if (!string.IsNullOrEmpty(text2))
			{
				if (!HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out var dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], HTKKFile.GetMessage("0071"));
					value = 2;
				}
				else if (dtDate.CompareTo(DateTime.Today) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message4);
					value = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message);
				value = 2;
			}
			string text3 = sheetView.Cells[array3[1], array3[0]].Text.Trim();
			HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
			if (!string.IsNullOrEmpty(text3))
			{
				if (!HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out var dtDate2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], HTKKFile.GetMessage("0071"));
					value = 2;
				}
				else if (dtDate2.CompareTo(DateTime.Today) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message4);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
