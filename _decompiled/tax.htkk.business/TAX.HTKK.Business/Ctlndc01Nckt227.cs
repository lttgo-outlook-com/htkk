using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ctlndc01Nckt227
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				if (!string.IsNullOrEmpty(PeriodInfo.NgayChiaCTLN))
				{
					sView.Cells[GeneralFunctions.ArrayCellId("M_5")[1], GeneralFunctions.ArrayCellId("M_5")[0]].Text = PeriodInfo.NgayChiaCTLN;
				}
				sView.Cells[GeneralFunctions.ArrayCellId("I_7")[1], GeneralFunctions.ArrayCellId("I_7")[0]].Value = PeriodInfo.CTDC;
				sView.Cells[GeneralFunctions.ArrayCellId("U_7")[1], GeneralFunctions.ArrayCellId("U_7")[0]].Value = PeriodInfo.LNDC;
			}
			DateTime dateTime = new DateTime(2020, 12, 5);
			string[] array = PeriodInfo.NgayChiaCTLN.Split(new char[1] { '/' });
			DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
			if (dateTime2 < dateTime)
			{
				sView.Cells[GeneralFunctions.ArrayCellId("C_5")[1], GeneralFunctions.ArrayCellId("C_5")[0]].Text = HTKKFile.GetMessage("0512");
				return;
			}
			sView.Cells[GeneralFunctions.ArrayCellId("C_5")[1], GeneralFunctions.ArrayCellId("C_5")[0]].Text = HTKKFile.GetMessage("0513");
			sView.Cells[GeneralFunctions.ArrayCellId("B_52")[1], GeneralFunctions.ArrayCellId("B_52")[0]].Text = HTKKFile.GetMessage("0514");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("I_7");
			int[] array2 = GeneralFunctions.ArrayCellId("U_7");
			if (activeRowIndex == array[1] && activeColumnIndex == array[0])
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[array[1], array[0]].Value))
				{
					sheetView.Cells[array2[1], array2[0]].Value = false;
				}
			}
			else if (activeRowIndex == array2[1] && activeColumnIndex == array2[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[array2[1], array2[0]].Value))
			{
				sheetView.Cells[array[1], array[0]].Value = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if ((iRow - 1 == GeneralFunctions.ArrayCellId("U_45")[1] && iCol - 1 == GeneralFunctions.ArrayCellId("U_45")[0]) || (iRow - 1 == GeneralFunctions.ArrayCellId("U_46")[1] && iCol - 1 == GeneralFunctions.ArrayCellId("U_46")[0]))
			{
				int[] array = GeneralFunctions.ArrayCellId("U_47");
				sheetView.Cells[array[1], array[0]].Formula = "ROUND(U45*(U46/100), 0)";
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
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int num = 0;
		if (levelWarning.Count > 0)
		{
			num = levelWarning[0];
		}
		SheetView sheetView = winFpSpread.Sheets[0];
		int[] array = GeneralFunctions.ArrayCellId("M_3");
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
