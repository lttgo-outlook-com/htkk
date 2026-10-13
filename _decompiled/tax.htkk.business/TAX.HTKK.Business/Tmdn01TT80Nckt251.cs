using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tmdn01TT80Nckt251
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("AO_8");
			sView.Cells[array[1], array[0]].Value = PeriodInfo.LanXuatBan;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("G");
			int[] array = GeneralFunctions.ArrayCellId("G_48");
			int[] array2 = GeneralFunctions.ArrayCellId("AO_48");
			int num2 = 37;
			int num3 = 45;
			if (iCol == num && iRow <= num3 && iRow >= num2)
			{
				sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
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
			int[] array = GeneralFunctions.ArrayCellId("R_130");
			int[] array2 = GeneralFunctions.ArrayCellId("K_22");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_45"));
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			string text = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				if (!HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out var dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message3);
					value = 2;
				}
				else if (DateTime.Compare(dtDate, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message2);
					value = 2;
				}
			}
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			string text2 = sheetView.Cells[array[1], array[0]].Text.Trim();
			DateTime dtDate2;
			if (string.IsNullOrEmpty(text2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message);
				value = 2;
			}
			else if (!HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out dtDate2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message3);
				value = 2;
			}
			else if (DateTime.Compare(dtDate2, DateTime.Now) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message2);
				value = 2;
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
