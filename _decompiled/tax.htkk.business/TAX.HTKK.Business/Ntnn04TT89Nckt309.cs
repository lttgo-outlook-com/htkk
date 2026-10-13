using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn04TT89Nckt309
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("BD_10");
			int[] array2 = GeneralFunctions.ArrayCellId("BE_10");
			sView.Cells[array[1], array[0]].Text = PeriodInfo.MaQuyetToan;
			sView.Cells[array2[1], array2[0]].Text = PeriodInfo.TenQuyetToan;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			capNhatLaiCt4_GtHD_Tk(winFpSpread);
			capNhatLaiCt4_DoanhThu_Tk(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("BI"))
			{
				capNhatLaiCt4_GtHD_Tk(winFpSpread);
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("BX"))
			{
				capNhatLaiCt4_DoanhThu_Tk(winFpSpread);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			capNhatLaiCt4_GtHD_Tk(winFpSpread);
			capNhatLaiCt4_DoanhThu_Tk(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int actionType)
	{
		try
		{
			capNhatLaiCt4_GtHD_Tk(winFpSpread);
			capNhatLaiCt4_DoanhThu_Tk(winFpSpread);
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
			checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			checkPL4_1(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
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
		}
	}

	private void checkPL4_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void capNhatLaiCt4_GtHD_Tk(FpSpread winFpSpread)
	{
		int[] array = GeneralFunctions.ArrayCellId("AB_32");
		int num = HTKKFile.ColumnLettersToIndex("BD");
		SheetView sheetView = winFpSpread.Sheets[0];
		sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array[1], num - 1].Value;
	}

	public void capNhatLaiCt4_DoanhThu_Tk(FpSpread winFpSpread)
	{
		int[] array = GeneralFunctions.ArrayCellId("AB_33");
		int num = HTKKFile.ColumnLettersToIndex("BD");
		SheetView sheetView = winFpSpread.Sheets[0];
		sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array[1], num - 1].Value;
	}
}
