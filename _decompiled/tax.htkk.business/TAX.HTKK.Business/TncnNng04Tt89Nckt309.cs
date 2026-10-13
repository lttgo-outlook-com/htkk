using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnNng04Tt89Nckt309
{
	private int colCongThuc = HTKKFile.ColumnLettersToIndex("AN");

	private int[] cellCt31 = GeneralFunctions.ArrayCellId("U_129");

	private int[] cellCt32 = GeneralFunctions.ArrayCellId("U_130");

	private int[] cellCt35 = GeneralFunctions.ArrayCellId("U_137");

	private int[] cellCt36 = GeneralFunctions.ArrayCellId("U_138");

	private int[] cellCt39 = GeneralFunctions.ArrayCellId("U_145");

	private int[] cellCt40 = GeneralFunctions.ArrayCellId("U_146");

	private int[] cellCt43 = GeneralFunctions.ArrayCellId("U_153");

	private int[] cellCt44 = GeneralFunctions.ArrayCellId("U_154");

	private int[] cellTenLoaiTN = GeneralFunctions.ArrayCellId("AL_20");

	private int[] cellMaLoaiTN = GeneralFunctions.ArrayCellId("AM_20");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string tieuMucHachToanVal = PeriodInfo.TieuMucHachToanVal;
			anHienMuc(sView, PeriodInfo, ToKhaiInfo.SheetInfos[0], tieuMucHachToanVal);
			sView.Cells[cellTenLoaiTN[1], cellTenLoaiTN[0]].Text = PeriodInfo.NganhNgheKinhDoanh;
			sView.Cells[cellMaLoaiTN[1], cellMaLoaiTN[0]].Text = tieuMucHachToanVal;
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
			if (iRow == cellCt31[1] + 1)
			{
				sheetView.Cells[cellCt32[1], cellCt32[0]].Value = sheetView.Cells[cellCt32[1], colCongThuc - 1].Value;
			}
			if (iRow == cellCt35[1] + 1)
			{
				sheetView.Cells[cellCt36[1], cellCt36[0]].Value = sheetView.Cells[cellCt36[1], colCongThuc - 1].Value;
			}
			if (iRow == cellCt39[1] + 1)
			{
				sheetView.Cells[cellCt40[1], cellCt40[0]].Value = sheetView.Cells[cellCt40[1], colCongThuc - 1].Value;
			}
			if (iRow == cellCt43[1] + 1)
			{
				sheetView.Cells[cellCt44[1], cellCt44[0]].Value = sheetView.Cells[cellCt44[1], colCongThuc - 1].Value;
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

	public void anHienMuc(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCSheetInfo sInfo, string nganhNgheKD)
	{
		int startRow = 109;
		int num = 115;
		int num2 = 121;
		int num3 = 123;
		int startRow2 = 129;
		int endRow = 131;
		int startRow3 = 137;
		int endRow2 = 139;
		int startRow4 = 145;
		int endRow3 = 147;
		int startRow5 = 153;
		int endRow4 = 155;
		int num4 = 161;
		int endRow5 = 163;
		int[] array = GeneralFunctions.ArrayCellId("U_109");
		int[] array2 = GeneralFunctions.ArrayCellId("C_111");
		int[] array3 = GeneralFunctions.ArrayCellId("U_113");
		int[] array4 = GeneralFunctions.ArrayCellId("C_115");
		int[] array5 = GeneralFunctions.ArrayCellId("U_121");
		int[] array6 = GeneralFunctions.ArrayCellId("U_123");
		int[] array7 = GeneralFunctions.ArrayCellId("U_161");
		int[] array8 = GeneralFunctions.ArrayCellId("U_163");
		GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, startRow, endRow5, isLock: true, isDeleteData: true);
		switch (nganhNgheKD)
		{
		case "01":
			GeneralFunctions.LockUnLockCol(sView, array[1] + 1, array[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, array3[1] + 1, array3[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, array2[1] + 1, array2[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, array4[1] + 1, array4[0] + 1, locked: false, deleteData: false);
			break;
		case "02":
			GeneralFunctions.LockUnLockCol(sView, array5[1] + 1, array5[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, array6[1] + 1, array6[0] + 1, locked: false, deleteData: false);
			break;
		case "03":
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, startRow2, endRow, isLock: false, isDeleteData: false);
			break;
		case "04":
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, startRow3, endRow2, isLock: false, isDeleteData: false);
			break;
		case "05":
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, startRow4, endRow3, isLock: false, isDeleteData: false);
			break;
		case "06":
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, startRow5, endRow4, isLock: false, isDeleteData: false);
			break;
		default:
			GeneralFunctions.LockUnLockCol(sView, array7[1] + 1, array7[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, array8[1] + 1, array8[0] + 1, locked: false, deleteData: false);
			break;
		}
	}
}
