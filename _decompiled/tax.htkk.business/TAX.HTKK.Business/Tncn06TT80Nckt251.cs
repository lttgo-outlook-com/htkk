using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tncn06TT80Nckt251
{
	private int[] cellTemPhuLuc = GeneralFunctions.ArrayCellId("AB_27");

	private int[] cellKhaiThay = GeneralFunctions.ArrayCellId("F_7");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
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
			if (iRow == cellKhaiThay[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
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
			int num = HTKKFile.ColumnLettersToIndex("AE");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			int num4 = HTKKFile.ColumnLettersToIndex("M");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			if (iCol == num2 && iRow >= dataLocationTemplateRowStart && iRow < dataLocationTemplateRowStart2)
			{
				sheetView.Cells[(int)iRow - 1, num3 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
			if (iCol == num4 && iRow >= dataLocationTemplateRowStart2)
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 6 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_06_1_TNCN_TT80.xls");
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

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
