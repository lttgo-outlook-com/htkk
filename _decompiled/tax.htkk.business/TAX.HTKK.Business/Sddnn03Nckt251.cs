using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Sddnn03Nckt251
{
	private int[] cellTenTinh = GeneralFunctions.ArrayCellId("G_103");

	private int[] cellMaTinh = GeneralFunctions.ArrayCellId("AO_103");

	private int[] cellTenHuyen = GeneralFunctions.ArrayCellId("G_105");

	private int[] cellMaHuyen = GeneralFunctions.ArrayCellId("AO_105");

	private int[] cellCt11 = GeneralFunctions.ArrayCellId("I_77");

	private int[] cellCt11Formular = GeneralFunctions.ArrayCellId("AO_77");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("AO_8");
			sView.Cells[array[1], array[0]].Value = PeriodInfo.LanXuatBan;
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, cellTenTinh, cellMaTinh, cellTenHuyen, cellMaHuyen, null, null, null, null, null, PhuThuocTruSoChinh: false);
			if (string.IsNullOrEmpty(sView.Cells[cellCt11[1], cellCt11[0]].Text.Trim()))
			{
				sView.Cells[cellCt11[1], cellCt11[0]].Value = sView.Cells[cellCt11Formular[1], cellCt11Formular[0]].Value;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("P_149");
			int[] array2 = GeneralFunctions.ArrayCellId("P_151");
			int[] array3 = GeneralFunctions.ArrayCellId("P_153");
			int[] array4 = GeneralFunctions.ArrayCellId("P_155");
			int[] array5 = GeneralFunctions.ArrayCellId("P_159");
			int[] array6 = GeneralFunctions.ArrayCellId("P_161");
			int[] array7 = GeneralFunctions.ArrayCellId("AO_161");
			if ((iCol == array[0] + 1 || iCol == array2[0] + 1 || iCol == array3[0] + 1 || iCol == array4[0] + 1 || iCol == array5[0] + 1) && (iRow == array[1] + 1 || iRow == array2[1] + 1 || iRow == array3[1] + 1 || iRow == array4[1] + 1 || iRow == array5[1] + 1))
			{
				sheetView.Cells[array6[1], array6[0]].Value = sheetView.Cells[array7[1], array7[0]].Value;
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
			int[] array = GeneralFunctions.ArrayCellId("M_9");
			string text = sheetView.Cells[array[1], array[0]].Text;
			if (string.IsNullOrEmpty(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
				value = 2;
				return;
			}
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
			if (dateTime == DateTime.MinValue)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
				value = 2;
			}
			else if (dateTime.CompareTo(DateTime.Today) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48")));
				value = 2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
