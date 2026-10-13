using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnTkqt04Nckt251
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
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
			int[] array = GeneralFunctions.ArrayCellId("F_124");
			int[] array2 = GeneralFunctions.ArrayCellId("AB_126");
			int[] array3 = GeneralFunctions.ArrayCellId("F_130");
			int[] array4 = GeneralFunctions.ArrayCellId("AB_132");
			SheetView sheetView = winFpSpread.Sheets[0];
			double num = Convert.ToDouble(sheetView.Cells[array[1], array[0]].Value);
			if (iRow == array[1] + 1)
			{
				string value = MoneyHelpers.MoneyToText(num.ToString());
				sheetView.Cells[array2[1], array2[0]].Value = value;
			}
			double num2 = Convert.ToDouble(sheetView.Cells[array3[1], array3[0]].Value);
			if (iRow == array3[1] + 1)
			{
				string value = MoneyHelpers.MoneyToText(num2.ToString());
				sheetView.Cells[array4[1], array4[0]].Value = value;
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
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
