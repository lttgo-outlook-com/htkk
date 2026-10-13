using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ctlndc01TT80Nckt251
{
	private int[] cellCt12 = GeneralFunctions.ArrayCellId("U_73");

	private int[] cellCt13 = GeneralFunctions.ArrayCellId("U_74");

	private int[] cellCt14 = GeneralFunctions.ArrayCellId("U_75");

	private int[] cellCt14Formula = GeneralFunctions.ArrayCellId("AF_75");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("U_75");
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				if (!string.IsNullOrEmpty(PeriodInfo.NgayChiaCTLN))
				{
					sView.Cells[GeneralFunctions.ArrayCellId("Q_7")[1], GeneralFunctions.ArrayCellId("Q_7")[0]].Text = PeriodInfo.NgayChiaCTLN;
				}
				sView.Cells[GeneralFunctions.ArrayCellId("I_25")[1], GeneralFunctions.ArrayCellId("I_25")[0]].Value = PeriodInfo.CTDC;
				sView.Cells[GeneralFunctions.ArrayCellId("U_25")[1], GeneralFunctions.ArrayCellId("U_25")[0]].Value = PeriodInfo.LNDC;
			}
			GeneralFunctions.SetRowHeight(sView, array[1] + 1, 40f);
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
			int[] array = GeneralFunctions.ArrayCellId("I_25");
			int[] array2 = GeneralFunctions.ArrayCellId("U_25");
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
			if (iRow == cellCt12[1] + 1 || iRow == cellCt13[1] + 1)
			{
				sheetView.Cells[cellCt14[1], cellCt14[0]].Value = sheetView.Cells[cellCt14Formula[1], cellCt14Formula[0]].Value;
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
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
