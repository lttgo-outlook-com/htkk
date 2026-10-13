using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn02TT89Nckt309
{
	private int[] cellGiaTriHDCt4_tk = GeneralFunctions.ArrayCellId("AF_34");

	private int[] cellDoanhThuChuaThueCt4_tk = GeneralFunctions.ArrayCellId("AF_35");

	private int colBO_tk = HTKKFile.ColumnLettersToIndex("BO");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("BL_8");
			int[] array2 = GeneralFunctions.ArrayCellId("BK_8");
			int[] array3 = GeneralFunctions.ArrayCellId("AE_9");
			int[] array4 = GeneralFunctions.ArrayCellId("AR_9");
			HTKKFile.RemoveValidate(sView, array3[1], array3[0]);
			HTKKFile.RemoveValidate(sView, array4[1], array4[0]);
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[array[1], array[0]].Text = PeriodInfo.TenQuyetToan;
				sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaQuyetToan;
			}
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[array3[1], array3[0]].Locked = true;
				sView.Cells[array4[1], array4[0]].Locked = true;
			}
			string[] array5 = PeriodInfo.DenNgay.Split(new char[1] { '/' });
			if (array5.Length == 3 && PeriodInfo.Nam != array5[2])
			{
				PeriodInfo.Nam = array5[2];
			}
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
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
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("BY"))
			{
				tinhGiaTriHDCt4_tk(winFpSpread);
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("CJ"))
			{
				tinhDoanhThuChuaThueCt4_tk(winFpSpread);
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
			tinhGiaTriHDCt4_tk(winFpSpread);
			tinhDoanhThuChuaThueCt4_tk(winFpSpread);
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
			tinhGiaTriHDCt4_tk(winFpSpread);
			tinhDoanhThuChuaThueCt4_tk(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("BI"))
			{
				tinhGiaTriHDCt4_tk(winFpSpread);
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("BX"))
			{
				tinhDoanhThuChuaThueCt4_tk(winFpSpread);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			tinhGiaTriHDCt4_tk(winFpSpread);
			tinhDoanhThuChuaThueCt4_tk(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int actionType)
	{
		try
		{
			tinhGiaTriHDCt4_tk(winFpSpread);
			tinhDoanhThuChuaThueCt4_tk(winFpSpread);
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

	public void tinhGiaTriHDCt4_tk(FpSpread winFpSpread)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		sheetView.Cells[cellGiaTriHDCt4_tk[1], cellGiaTriHDCt4_tk[0]].Value = sheetView.Cells[cellGiaTriHDCt4_tk[1], colBO_tk - 1].Value;
	}

	public void tinhDoanhThuChuaThueCt4_tk(FpSpread winFpSpread)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		sheetView.Cells[cellDoanhThuChuaThueCt4_tk[1], cellDoanhThuChuaThueCt4_tk[0]].Value = sheetView.Cells[cellDoanhThuChuaThueCt4_tk[1], colBO_tk - 1].Value;
	}
}
