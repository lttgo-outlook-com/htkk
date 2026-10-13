using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn03TT80Nckt294
{
	private static Color lockCellBackColor = Color.FromArgb(255, 192, 192, 192);

	private static Color editableCellBackColor = Color.FromArgb(0, 0, 0, 0);

	private int idxToKhai = 0;

	private int idxPL03_1a = 1;

	private int idxPL03_1b = 2;

	private int idxPL03_1c = 3;

	private int idxPL03_2 = 4;

	private int idxPL03_3a = 5;

	private int idxPL03_3b = 6;

	private int idxPL03_3c = 7;

	private int idxPL03_3d = 8;

	private int idxPL03_4 = 9;

	private int idxPL03_5 = 10;

	private int idxPL03_6 = 11;

	private int idxPL03_8 = 12;

	private int idxPL03_8a = 13;

	private int idxPL03_8b = 14;

	private int idxPL03_8c = 15;

	private int idxPL03_9 = 16;

	private int idxPL03_01GDLK_ND132 = 17;

	private int idxPL03_02GDLK_ND132 = 18;

	private int idxPL03_03GDLK_ND132 = 19;

	private int idxPL03_04GDLK_ND132 = 20;

	private int idxPL03_ND92 = 21;

	private string UU_DAI_KHAC_TINH = "02";

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private HTKKCMCTTinChungInfo _TTinChung = null;

	private HTKKPeriodInfo _PeriodInfo = null;

	private List<string> _listKeyAndRowTo_Pl01_8C = new List<string>();

	private List<string> _listRowAndKeyOld_Pl01_8C = new List<string>();

	public bool ValidatePeriod_axfpsPhuLuc(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		return true;
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtA1_TK(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtA1_TK(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_3(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtA1_TK(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_4(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtD2_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_5(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt16PL03_3a(winFpSpread.Sheets[idxPL03_3a], _ToKhaiInfo.SheetInfos[idxPL03_3a]);
			TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtD7_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_6(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt16PL03_3b(winFpSpread.Sheets[idxPL03_3b], _ToKhaiInfo.SheetInfos[idxPL03_3b]);
			TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_7(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt21PL03_3c(winFpSpread.Sheets[idxPL03_3c], _ToKhaiInfo.SheetInfos[idxPL03_3c]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_8(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt16PL03_3d(winFpSpread.Sheets[idxPL03_3d], _ToKhaiInfo.SheetInfos[idxPL03_3d]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_9(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtC16_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_10(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtB15_TK(winFpSpread.Sheets[0]);
			TongHopCtD2_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_11(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtE6_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			TinhCtE5_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_12(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCtG1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtG2_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_13(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCtG4_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtG3_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_14(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_15(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCtG1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtG2_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_16(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_17(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_18(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_19(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_20(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_21(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtC15_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			SetPhuLucND92(winFpSpread);
			LockUnlockMucBUuDaiCungTinh_PL03_3a(winFpSpread.Sheets[idxPL03_3a], _ToKhaiInfo.SheetInfos[idxPL03_3a]);
			LockUnlockMucBUuDaiCungTinh_PL03_3b(winFpSpread.Sheets[idxPL03_3b], _ToKhaiInfo.SheetInfos[idxPL03_3b]);
			LockUnlockMucBUuDaiCungTinh_PL03_3c(winFpSpread.Sheets[idxPL03_3c], _ToKhaiInfo.SheetInfos[idxPL03_3c]);
			LockUnlockMucBUuDaiCungTinh_PL03_3d(winFpSpread.Sheets[idxPL03_3d], _ToKhaiInfo.SheetInfos[idxPL03_3d]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_3(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_4(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtD2_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_5(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt16PL03_3a(winFpSpread.Sheets[idxPL03_3a], _ToKhaiInfo.SheetInfos[idxPL03_3a]);
			TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtD7_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_6(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt16PL03_3b(winFpSpread.Sheets[idxPL03_3b], _ToKhaiInfo.SheetInfos[idxPL03_3b]);
			TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_7(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt21PL03_3c(winFpSpread.Sheets[idxPL03_3c], _ToKhaiInfo.SheetInfos[idxPL03_3c]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_8(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCt16PL03_3d(winFpSpread.Sheets[idxPL03_3d], _ToKhaiInfo.SheetInfos[idxPL03_3d]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_9(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtC16_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_10(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtB15_TK(winFpSpread.Sheets[0]);
			TongHopCtD2_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_11(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtE6_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			TinhCtE5_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_12(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCtG1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtG2_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_13(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCtG4_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtG3_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_14(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_15(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TinhCtG1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtG2_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			_listKeyAndRowTo_Pl01_8C.Clear();
			_listRowAndKeyOld_Pl01_8C.Clear();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_16(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_17(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_18(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_19(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_20(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_21(FpSpread winFpSpread)
	{
		try
		{
			SetAllPhuLuc(winFpSpread);
			TongHopCtC15_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			SetPhuLucND92(winFpSpread);
			LockUnlockMucBUuDaiCungTinh_PL03_3a(winFpSpread.Sheets[idxPL03_3a], _ToKhaiInfo.SheetInfos[idxPL03_3a]);
			LockUnlockMucBUuDaiCungTinh_PL03_3b(winFpSpread.Sheets[idxPL03_3b], _ToKhaiInfo.SheetInfos[idxPL03_3b]);
			LockUnlockMucBUuDaiCungTinh_PL03_3c(winFpSpread.Sheets[idxPL03_3c], _ToKhaiInfo.SheetInfos[idxPL03_3c]);
			LockUnlockMucBUuDaiCungTinh_PL03_3d(winFpSpread.Sheets[idxPL03_3d], _ToKhaiInfo.SheetInfos[idxPL03_3d]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ClearData_25(FpSpread winFpSpread)
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
		winFpSpread.ResumeLayout();
		checkTKChinh(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL3_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3A(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3B(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3C(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3D(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_4(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_6(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_8(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_8a(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_8b(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_8c(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_9(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_1_ND132(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPLND114(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int[] array = ConvertPathToCellId(hTKKCMCSheetInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE2");
			if (!CheckE2_TK(winFpSpread, ToKhaiInfo.SheetInfos[0]))
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_212")));
				if (value != 2)
				{
					value = 1;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				int colCheck = HTKKFile.ColumnLettersToIndex("C");
				int num = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sheetView, hTKKCMCTableInfo, colCheck, "Không được chọn trùng loại thu nhập miễn thuế giữa các dòng");
				if (num > 0)
				{
					value = num;
				}
			}
			SheetView sheetView2 = winFpSpread.Sheets[idxPL03_2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[idxPL03_2];
			int[] array2 = GeneralFunctions.ArrayCellId("O_53");
			int[] array3 = GeneralFunctions.ArrayCellId("O_54");
			double num2 = Convert.ToDouble(sheetView.Cells[array2[1], array2[0]].Value);
			double num3 = Convert.ToDouble(sheetView.Cells[array3[1], array3[0]].Value);
			if (hTKKCMCSheetInfo2 != null && num2 > num3)
			{
				int[] array4 = ConvertPathToCellId(hTKKCMCSheetInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3a");
				int[] array5 = GeneralFunctions.ArrayCellId("M_33");
				double num4 = Convert.ToDouble(sheetView2.Cells[array5[1], array5[0]].Value);
				double num5 = Convert.ToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				if (num5 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array4[1], array4[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_227")));
					value = 2;
				}
				int[] array6 = ConvertPathToCellId(hTKKCMCSheetInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3b");
				int[] array7 = GeneralFunctions.ArrayCellId("M_41");
				double num6 = Convert.ToDouble(sheetView2.Cells[array7[1], array7[0]].Value);
				double num7 = Convert.ToDouble(sheetView.Cells[array6[1], array6[0]].Value);
				if (num7 > num6)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_228")));
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

	private void checkPL3_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_2, out value);
			SheetView sheetView = winFpSpread.Sheets[idxPL03_2];
			if (!sheetView.Visible)
			{
				return;
			}
			int[] array = new int[10] { 27, 28, 29, 30, 31, 35, 36, 37, 38, 39 };
			bool flag = false;
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("D_" + array[i]);
				int[] array3 = GeneralFunctions.ArrayCellId("H_" + array[i]);
				int[] array4 = GeneralFunctions.ArrayCellId("M_" + array[i]);
				HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
				double num = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
				double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				if (num < num2 + num3)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
					value = 2;
				}
				if (num != 0.0)
				{
					flag = true;
				}
			}
			for (int i = 0; i < array.Length; i++)
			{
				if (!flag)
				{
					int[] array2 = GeneralFunctions.ArrayCellId("D_" + array[i]);
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_2].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_2].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_3A(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_3a, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3a];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3a];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[64, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_205"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_74"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_76"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_213"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_214"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_215"));
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int num5 = HTKKFile.ColumnLettersToIndex("L");
			int num6 = HTKKFile.ColumnLettersToIndex("AZ");
			int num7 = HTKKFile.ColumnLettersToIndex("Q");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("W");
			int num10 = HTKKFile.ColumnLettersToIndex("AZ");
			int colNam = HTKKFile.ColumnLettersToIndex("U");
			int colNam2 = HTKKFile.ColumnLettersToIndex("W");
			if (hTKKCMCSheetInfo != null)
			{
				List<string> list = new List<string>();
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num11 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
					HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
					HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
					HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
					HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1);
					bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
					string text4 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					string text5 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text;
					string text6 = sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text;
					string item = text5 + text6 + flag;
					if (list.Contains(item))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, message8);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, message8);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message8);
						value = 2;
					}
					else
					{
						list.Add(item);
					}
					if (string.IsNullOrEmpty(text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (text4 == UU_DAI_KHAC_TINH)
					{
						if (CheckBatNhapCQT(sheetView, hTKKCMCCellInfo4, hTKKCMCCellInfo5, hTKKCMCCellInfo6, hTKKCMCCellInfo7))
						{
							value = 2;
						}
						if (CheckCQTKhac_TSC(sheetView, TTinChung.maCQTCapCuc, hTKKCMCCellInfo4, hTKKCMCCellInfo5))
						{
							value = 2;
						}
					}
					if (text4 == UU_DAI_KHAC_TINH && CheckBatNhapTinh(sheetView, hTKKCMCCellInfo8, cellMaTinh, hTKKCMCCellInfo9, cellMaHuyen, hTKKCMCCellInfo10, cellMaXa))
					{
						value = 2;
					}
					for (int num12 = 0; num12 < hTKKCMCTableInfo.RowsValues.Count; num12++)
					{
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num9 - 1);
						string text7 = sheetView.Cells[num11 + num12, num - 1].Text;
						string text8 = sheetView.Cells[num11 + num12, num2 - 1].Text;
						string text9 = sheetView.Cells[num11 + num12, num3 - 1].Text;
						string text10 = sheetView.Cells[num11 + num12, num4 - 1].Text;
						string text11 = sheetView.Cells[num11 + num12, num5 - 1].Text;
						string text12 = sheetView.Cells[num11 + num12, num6 - 1].Text;
						string text13 = sheetView.Cells[num11 + num12, num7 - 1].Text;
						string text14 = sheetView.Cells[num11 + num12, num8 - 1].Text;
						string text15 = sheetView.Cells[num11 + num12, num9 - 1].Text;
						string text16 = sheetView.Cells[num11 + num12, num10 - 1].Text;
						if (CheckKhongDuocLonHonNamHienTai(sheetView, num11 + num12 + 1, colNam))
						{
							value = 2;
						}
						if (CheckKhongDuocLonHonNamHienTai(sheetView, num11 + num12 + 1, colNam2))
						{
							value = 2;
						}
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num - 1, text);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text8))
						{
							if (!string.IsNullOrEmpty(text8) && text4 != UU_DAI_KHAC_TINH)
							{
								if (!ValidateBase.CheckMaSoThue(text8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message2);
									value = 2;
								}
							}
							else if (text8.Length == 5)
							{
								if (!ValidateBase.CheckMaSoThue5So(text8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message4);
									value = 2;
								}
							}
							else if (text8.Length == 13 || text8.Length == 14)
							{
								if (!ValidateBase.CheckMaSoThue(text8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message);
									value = 2;
								}
								else if (text8.Substring(0, 10) != TTinChung.mst.Substring(0, 10))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message5);
									value = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message3);
								value = 2;
							}
						}
						if (string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num4 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, text);
							value = 2;
						}
						else if (string.IsNullOrEmpty(text12))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, message10);
							value = 2;
						}
						else if (flag)
						{
							if (text16 != "0247")
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, message6);
								value = 2;
							}
						}
						else if (text16 == "0247")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, message7);
							value = 2;
						}
						if (string.IsNullOrEmpty(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num7 - 1, text);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_3a].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_3a].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_3a].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_3B(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_3b, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3b];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3b];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_205"));
			string text3 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[64, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_74"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_76"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_214"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_215"));
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int num5 = HTKKFile.ColumnLettersToIndex("L");
			int num6 = HTKKFile.ColumnLettersToIndex("Q");
			int num7 = HTKKFile.ColumnLettersToIndex("R");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("X");
			int num10 = HTKKFile.ColumnLettersToIndex("BA");
			int colNam = HTKKFile.ColumnLettersToIndex("V");
			int colNam2 = HTKKFile.ColumnLettersToIndex("X");
			if (hTKKCMCSheetInfo != null)
			{
				List<string> list = new List<string>();
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num11 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
					HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
					HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
					HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1);
					string text4 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					string text5 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text;
					string text6 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text;
					string item = text5 + text6;
					if (list.Contains(item))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message8);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, message8);
						value = 2;
					}
					else
					{
						list.Add(item);
					}
					if (string.IsNullOrEmpty(text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (text4 == UU_DAI_KHAC_TINH)
					{
						if (CheckBatNhapCQT(sheetView, hTKKCMCCellInfo3, hTKKCMCCellInfo4, hTKKCMCCellInfo5, hTKKCMCCellInfo6))
						{
							value = 2;
						}
						if (CheckCQTKhac_TSC(sheetView, TTinChung.maCQTCapCuc, hTKKCMCCellInfo3, hTKKCMCCellInfo4))
						{
							value = 2;
						}
					}
					if (text4 == UU_DAI_KHAC_TINH && CheckBatNhapTinh(sheetView, hTKKCMCCellInfo7, cellMaTinh, hTKKCMCCellInfo8, cellMaHuyen, hTKKCMCCellInfo9, cellMaXa))
					{
						value = 2;
					}
					for (int num12 = 0; num12 < hTKKCMCTableInfo.RowsValues.Count; num12++)
					{
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, num11 + num12, num9 - 1);
						string text7 = sheetView.Cells[num11 + num12, num - 1].Text;
						string text8 = sheetView.Cells[num11 + num12, num2 - 1].Text;
						string text9 = sheetView.Cells[num11 + num12, num3 - 1].Text;
						string text10 = sheetView.Cells[num11 + num12, num4 - 1].Text;
						string text11 = sheetView.Cells[num11 + num12, num5 - 1].Text;
						string text12 = sheetView.Cells[num11 + num12, num6 - 1].Text;
						string text13 = sheetView.Cells[num11 + num12, num7 - 1].Text;
						string text14 = sheetView.Cells[num11 + num12, num8 - 1].Text;
						string text15 = sheetView.Cells[num11 + num12, num9 - 1].Text;
						string text16 = sheetView.Cells[num11 + num12, num10 - 1].Text;
						if (CheckKhongDuocLonHonNamHienTai(sheetView, num11 + num12 + 1, colNam))
						{
							value = 2;
						}
						if (CheckKhongDuocLonHonNamHienTai(sheetView, num11 + num12 + 1, colNam2))
						{
							value = 2;
						}
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num - 1, text);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text8))
						{
							if (!string.IsNullOrEmpty(text8) && text4 != UU_DAI_KHAC_TINH)
							{
								if (!ValidateBase.CheckMaSoThue(text8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message);
									value = 2;
								}
							}
							else if (text8.Length == 5)
							{
								if (!ValidateBase.CheckMaSoThue5So(text8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message4);
									value = 2;
								}
							}
							else if (text8.Length == 13 || text8.Length == 14)
							{
								if (!ValidateBase.CheckMaSoThue(text8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message2);
									value = 2;
								}
								else if (text8.Substring(0, 10) != TTinChung.mst.Substring(0, 10))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message5);
									value = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num2 - 1, message3);
								value = 2;
							}
						}
						if (string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num4 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, text);
							value = 2;
						}
						else if (string.IsNullOrEmpty(text16))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, message9);
							value = 2;
						}
						else if (text16 == "0117")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num5 - 1, message7);
							value = 2;
						}
						if (string.IsNullOrEmpty(text12))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num6 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num11 + num12, num7 - 1, text);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_3b].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_3b].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_3b].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_3C(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_3c, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3c];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3c];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[81, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[76, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[77, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[78, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[174, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_214"));
			if (hTKKCMCSheetInfo != null)
			{
				List<string> list = new List<string>();
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
					HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
					HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
					HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_sdld_nu"));
					HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("sold_nu"));
					HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("so_xn_sdld_nu"));
					HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ngay_xn_sdld_nu"));
					HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_sdld_dtoc"));
					HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("sold_dtoc"));
					HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("so_xn_sdld_dtoc"));
					HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ngay_xn_sdld_dtoc"));
					HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct14"));
					HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct15"));
					HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct16"));
					HTKKCMCCellInfo hTKKCMCCellInfo21 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct17"));
					HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct18"));
					HTKKCMCCellInfo hTKKCMCCellInfo23 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct19"));
					HTKKCMCCellInfo hTKKCMCCellInfo24 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct20"));
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo20.CellTemplate.RowIndex - 1, hTKKCMCCellInfo20.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo22.CellTemplate.RowIndex - 1, hTKKCMCCellInfo22.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, hTKKCMCCellInfo24.CellTemplate.ColIndex - 1);
					bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value);
					double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value);
					bool flag2 = ConvertHelpers.ConvertToBool(sheetView.Cells[hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1].Value);
					double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Value);
					double num4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1].Value);
					double num5 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1].Value);
					double num6 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo20.CellTemplate.RowIndex - 1, hTKKCMCCellInfo20.CellTemplate.ColIndex - 1].Value);
					double num7 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo21.CellTemplate.RowIndex - 1, hTKKCMCCellInfo21.CellTemplate.ColIndex - 1].Value);
					double num8 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo22.CellTemplate.RowIndex - 1, hTKKCMCCellInfo22.CellTemplate.ColIndex - 1].Value);
					double num9 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo23.CellTemplate.RowIndex - 1, hTKKCMCCellInfo23.CellTemplate.ColIndex - 1].Value);
					double num10 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, hTKKCMCCellInfo24.CellTemplate.ColIndex - 1].Value);
					string text7 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					string text8 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text;
					string text9 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text;
					string item = text8 + text9;
					if (list.Contains(item))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, message);
						value = 2;
					}
					else
					{
						list.Add(item);
					}
					if (string.IsNullOrEmpty(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (text7 == UU_DAI_KHAC_TINH)
					{
						if (CheckBatNhapCQT(sheetView, hTKKCMCCellInfo3, hTKKCMCCellInfo4, hTKKCMCCellInfo5, hTKKCMCCellInfo6))
						{
							value = 2;
						}
						if (CheckCQTKhac_TSC(sheetView, TTinChung.maCQTCapCuc, hTKKCMCCellInfo3, hTKKCMCCellInfo4))
						{
							value = 2;
						}
					}
					if (text7 == UU_DAI_KHAC_TINH && CheckBatNhapTinh(sheetView, hTKKCMCCellInfo7, cellMaTinh, hTKKCMCCellInfo8, cellMaHuyen, hTKKCMCCellInfo9, cellMaXa))
					{
						value = 2;
					}
					if (!Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value) && !Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1].Value))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, text6);
						value = 2;
					}
					if (flag && num2 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (flag2 && num3 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if ((num4 != 0.0 || num5 != 0.0) && num5 >= num4)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, text2);
						value = 2;
					}
					if (num6 > num5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo20.CellTemplate.RowIndex - 1, hTKKCMCCellInfo20.CellTemplate.ColIndex - 1, text3);
						value = 2;
					}
					if (num8 > num7)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo22.CellTemplate.RowIndex - 1, hTKKCMCCellInfo22.CellTemplate.ColIndex - 1, text4);
						value = 2;
					}
					if (num10 > num9)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, hTKKCMCCellInfo24.CellTemplate.ColIndex - 1, text5);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_3c].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_3c].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_3c].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_3D(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_3d, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3d];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3d];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_205"));
			string text3 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[64, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[173, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[176, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[177, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[178, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[152, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_74"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_215"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_76"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_214"));
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int num5 = HTKKFile.ColumnLettersToIndex("AG");
			int num6 = HTKKFile.ColumnLettersToIndex("AQ");
			int num7 = HTKKFile.ColumnLettersToIndex("O");
			int num8 = HTKKFile.ColumnLettersToIndex("BA");
			if (hTKKCMCSheetInfo != null)
			{
				List<string> list = new List<string>();
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num9 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
					HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10a_ten"));
					HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10a_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
					HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
					HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
					HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
					HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10"));
					HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct11"));
					HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct12"));
					HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct14"));
					HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct15"));
					HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct17"));
					HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct18"));
					HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_chuyenGiao"));
					HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_khoaHocCN"));
					HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("noiCap"));
					Reset1KhoangLienTiep(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo17.CellTemplate.RowIndex, hTKKCMCCellInfo16.CellTemplate.RowIndex);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1);
					string text9 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text;
					string text10 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text;
					string item = text9 + text10;
					if (list.Contains(item))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message9);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, message9);
						value = 2;
					}
					else
					{
						list.Add(item);
					}
					bool flag = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1].Value);
					bool flag2 = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1].Value);
					if (!flag && !flag2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, text7);
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, text7);
						value = 2;
					}
					if (flag)
					{
						int num10 = CheckTrong1KhoangGiaTriBatNhap(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo12.CellTemplate.RowIndex, hTKKCMCCellInfo13.CellTemplate.RowIndex);
						if (num10 > 0)
						{
							value = 2;
						}
					}
					if (flag2)
					{
						int num10 = CheckTrong1KhoangGiaTriBatNhap(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo18.CellTemplate.RowIndex + 1, hTKKCMCCellInfo19.CellTemplate.RowIndex + 1);
						if (num10 > 0)
						{
							value = 2;
						}
						int num11 = CheckTrong1KhoangGiaTriBatNhap(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo14.CellTemplate.RowIndex, hTKKCMCCellInfo16.CellTemplate.RowIndex);
						if (num11 > 0)
						{
							value = 2;
						}
					}
					double num12 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Value);
					double num13 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1].Value);
					if (num12 != 0.0 && num13 != 0.0 && num13 >= num12)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1, text5);
						value = 2;
					}
					double num14 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Value);
					if (num14 > 100.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, text6);
						value = 2;
					}
					string text11 = sheetView.Cells[hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text11) && !GeneralFunctions.IsNumeric(text11))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, text8);
						value = 2;
					}
					string text12 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					string text13 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
					if (string.IsNullOrEmpty(text12) || string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (text12 == UU_DAI_KHAC_TINH)
					{
						if (CheckBatNhapCQT(sheetView, hTKKCMCCellInfo3, hTKKCMCCellInfo4, hTKKCMCCellInfo5, hTKKCMCCellInfo6))
						{
							value = 2;
						}
						if (CheckCQTKhac_TSC(sheetView, TTinChung.maCQTCapCuc, hTKKCMCCellInfo3, hTKKCMCCellInfo4))
						{
							value = 2;
						}
					}
					if (text12 == UU_DAI_KHAC_TINH && CheckBatNhapTinh(sheetView, hTKKCMCCellInfo7, cellMaTinh, hTKKCMCCellInfo8, cellMaHuyen, hTKKCMCCellInfo9, cellMaXa))
					{
						value = 2;
					}
					for (int num15 = 0; num15 < hTKKCMCTableInfo.RowsValues.Count; num15++)
					{
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num15, num6 - 1);
						string value2 = sheetView.Cells[num9 + num15, num7 - 1].Text.Trim();
						string value3 = sheetView.Cells[num9 + num15, num - 1].Text.Trim();
						string text14 = sheetView.Cells[num9 + num15, num2 - 1].Text.Trim();
						string text15 = sheetView.Cells[num9 + num15, num3 - 1].Text.Trim();
						string value4 = sheetView.Cells[num9 + num15, num4 - 1].Text.Trim();
						int num16 = Convert.ToInt32(sheetView.Cells[num9 + num15, num5 - 1].Value);
						int num17 = Convert.ToInt32(sheetView.Cells[num9 + num15, num6 - 1].Value);
						if (flag2 && !string.IsNullOrEmpty(value2) && CheckKhongDuocLonHonNamHienTai(sheetView, num9 + num15 + 1, num7))
						{
							value = 2;
						}
						if (num17 > num16)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num6 - 1, text4);
							value = 2;
						}
						if (flag2 && string.IsNullOrEmpty(value3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num - 1, text);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text14))
						{
							if (!string.IsNullOrEmpty(text14) && text12 != UU_DAI_KHAC_TINH)
							{
								if (!ValidateBase.CheckMaSoThue(text14))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num2 - 1, message);
									value = 2;
								}
							}
							else if (text14.Length == 5)
							{
								if (!ValidateBase.CheckMaSoThue5So(text14))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num2 - 1, message4);
									value = 2;
								}
							}
							else if (text14.Length == 13 || text14.Length == 14)
							{
								if (!ValidateBase.CheckMaSoThue(text14))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num2 - 1, message2);
									value = 2;
								}
								else if (text14.Substring(0, 10) != TTinChung.mst.Substring(0, 10))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num2 - 1, message5);
									value = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num2 - 1, message3);
								value = 2;
							}
						}
						if (flag2 && string.IsNullOrEmpty(value4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num15, num4 - 1, text);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_3d].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_3d].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_3d].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_4, out value);
			SheetView sheetView = winFpSpread.Sheets[idxPL03_4];
			if (!sheetView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_4];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			string text2 = "";
			string text3 = "";
			int num2 = GeneralFunctions.ArrayCellId("E_27")[0];
			int num3 = GeneralFunctions.ArrayCellId("O_27")[0];
			int num4 = GeneralFunctions.ArrayCellId("Q_27")[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num2].Text;
				text2 = sheetView.Cells[num + i, num3].Text;
				text3 = sheetView.Cells[num + i, num4].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, num2);
				HTKKFile.RemoveValidate(sheetView, num + i, num3);
				HTKKFile.RemoveValidate(sheetView, num + i, num4);
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_4].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_4].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_4].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_6, out value);
			SheetView sheetView = winFpSpread.Sheets[idxPL03_6];
			if (!sheetView.Visible)
			{
				return;
			}
			int[] array = GeneralFunctions.ArrayCellId("N_10");
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			object value2 = sheetView.Cells[array[1], array[0]].Value;
			if (Convert.ToInt32(PeriodInfo.Nam) < 2025 && Convert.ToInt32(value2) > 10)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], "Được nhập tối đa 10%");
				value = 2;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_6];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			double num2 = 0.0;
			double num3 = ConvertHelpers.ConvertToDouble(PeriodInfo.Nam);
			int num4 = GeneralFunctions.ArrayCellId("D_26")[0];
			int num5 = GeneralFunctions.ArrayCellId("X_26")[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num4].Text;
				num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, num5].Value);
				HTKKFile.RemoveValidate(sheetView, num + i, num4);
				HTKKFile.RemoveValidateBackGround(sheetView, num + i, num5);
				if (!string.IsNullOrEmpty(text))
				{
					if (!HTKKFormater.IsNumber(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
						value = 2;
					}
					else if (ConvertHelpers.ConvertToDouble(text) > num3 || ConvertHelpers.ConvertToDouble(text) < num3 - 5.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_58")));
						value = 2;
					}
				}
				if (num2 < 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
			}
			hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text2 = "";
			num3 = ConvertHelpers.ConvertToDouble(PeriodInfo.Nam);
			num4 = GeneralFunctions.ArrayCellId("D_39")[0];
			int column = GeneralFunctions.ArrayCellId("AH_39")[0];
			int num6 = GeneralFunctions.ArrayCellId("O_39")[0];
			int num7 = GeneralFunctions.ArrayCellId("Q_39")[0];
			int num8 = GeneralFunctions.ArrayCellId("R_39")[0];
			int num9 = GeneralFunctions.ArrayCellId("U_39")[0];
			int num10 = GeneralFunctions.ArrayCellId("V_39")[0];
			int num11 = GeneralFunctions.ArrayCellId("X_39")[0];
			int num12 = GeneralFunctions.ArrayCellId("Z_39")[0];
			int[] array2 = new int[7] { num6, num7, num8, num9, num10, num11, num12 };
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num4].Text;
				text2 = sheetView.Cells[num + i, column].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, num4);
				for (int j = 0; j < array2.Length; j++)
				{
					HTKKFile.RemoveValidate(sheetView, num + i, array2[j]);
				}
				if (!string.IsNullOrEmpty(text))
				{
					if (!HTKKFormater.IsNumber(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
						value = 2;
					}
					else if (ConvertHelpers.ConvertToDouble(text) > num3)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
						value = 2;
					}
				}
				if (text2 == "0")
				{
					for (int j = 0; j < array2.Length; j++)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, array2[j], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
					}
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_6].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_6].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_6].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_6B(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPLND114(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_8, out value);
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8];
			if (!sheetView.Visible)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_80"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_67"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_183"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			string text6 = "";
			string text7 = "";
			int[] array = GeneralFunctions.ArrayCellId("AR_21");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("AR");
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("AT");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("AS");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("AU");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AV");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text2 = sheetView.Cells[num + i, num3 - 1].Text;
				text = sheetView.Cells[num + i, num2 - 1].Text;
				text3 = sheetView.Cells[num + i, num4 - 1].Text;
				text4 = sheetView.Cells[num + i, num5 - 1].Text;
				text5 = sheetView.Cells[num + i, num6 - 1].Text;
				text6 = sheetView.Cells[num + i, num7 - 1].Text;
				text7 = sheetView.Cells[num + i, num9 - 1].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num7 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num9 - 1);
				if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, message);
					value = 2;
				}
				else if (Convert.ToInt32(sheetView.Cells[array[1], array[0]].Value) > 1 && sheetView.Cells[num + i, num3 - 1].Text.Trim() == "01")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, message7);
					value = 2;
				}
				if ((text2 == "02" || text2 == "03") && text2 != "03")
				{
					GeneralFunctions.CheckTinhTheoCQTDangBangKhac_TSC(TTinChung.maCQTCapCuc, TTinChung.maCQTNoiNop, sheetView, num6, colMaTinh, num5, colMaHuyen, -1, -1, -1, num + i + 1);
				}
				if (string.IsNullOrEmpty(text5) && !sheetView.Cells[num + i, num6 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num6 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text4) && !sheetView.Cells[num + i, num5 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text7) && !sheetView.Cells[num + i, num9 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text3))
				{
					continue;
				}
				if (text2 == "02")
				{
					if (text3.Length == 13 || text3.Length == 14)
					{
						if (!ValidateBase.CheckMaSoThue(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message5);
							value = 2;
						}
					}
					else if (text3.Length == 5)
					{
						if (!ValidateBase.CheckMaSoThue5So(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message4);
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message3);
						value = 2;
					}
				}
				else
				{
					if (!(text2 == "03"))
					{
						continue;
					}
					if (text3.Length == 10 && !ValidateBase.CheckMaSoThue(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message2);
						value = 2;
					}
					if (text3.Length == 13 || text3.Length == 14)
					{
						if (!ValidateBase.CheckMaSoThue(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message5);
							value = 2;
						}
					}
					else if (text3.Length == 5)
					{
						if (!ValidateBase.CheckMaSoThue5So(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message4);
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, message2);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_8].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_8].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_8].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_8a(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_8a, out value);
			bool flag = false;
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8a];
			if (!sheetView.Visible)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_184"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_199"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8a];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			string text6 = "";
			string text7 = "";
			int[] array = GeneralFunctions.ArrayCellId("AU_2");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("AU");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("T");
			int num7 = HTKKFile.ColumnLettersToIndex("R");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AY");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("BB");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("AV");
			int num11 = Convert.ToInt32(sheetView.Cells[array[1], array[0]].Value);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text8 = sheetView.Cells[num + i, num2 - 1].Text;
				text2 = sheetView.Cells[num + i, num3 - 1].Text;
				text = sheetView.Cells[num + i, num4 - 1].Text;
				text3 = sheetView.Cells[num + i, num5 - 1].Text;
				text4 = sheetView.Cells[num + i, num6 - 1].Text;
				text5 = sheetView.Cells[num + i, num7 - 1].Text;
				text6 = sheetView.Cells[num + i, num8 - 1].Text;
				text7 = sheetView.Cells[num + i, num9 - 1].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num7 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num9 - 1);
				if (num11 > 1 && text2 == "01")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, message3);
					value = 2;
				}
				if (string.IsNullOrEmpty(text8))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				if (text2 == "02" || text2 == "03" || text2 == "04")
				{
					GeneralFunctions.CheckTinhTheoCQTDangBangKhac_TSC(TTinChung.maCQTCapCuc, TTinChung.maCQTNoiNop, sheetView, num7, colMaTinh, num6, colMaHuyen, -1, -1, -1, num + i + 1);
				}
				if (string.IsNullOrEmpty(text5) && !sheetView.Cells[num + i, num7 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text4) && !sheetView.Cells[num + i, num6 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num6 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text7) && !sheetView.Cells[num + i, num9 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text3) && (text2 == "02" || text2 == "03"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
					value = 2;
				}
				else if (text2 == "02")
				{
					if (text3.Length == 13 || text3.Length == 14)
					{
						if (!ValidateBase.CheckMaSoThue(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_80")));
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_74")));
						value = 2;
					}
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num5, text3);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
						value = 2;
					}
				}
				else if (text2 == "03")
				{
					if (text3.Length != 5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66")));
						value = 2;
					}
					if (text3.Length == 5 && !ValidateBase.CheckMaSoThue5So(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_81")));
						value = 2;
					}
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num5, text3);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
						value = 2;
					}
				}
				if (!(text2 == "04"))
				{
					continue;
				}
				string text9 = sheetView.Cells[num + i, num10 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text9))
				{
					List<int> list = new List<int>();
					list.Add(num3);
					list.Add(num10);
					List<string> list2 = new List<string>();
					list2.Add(text2);
					list2.Add(text9);
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
					if (rowIndexInSViewByCompares.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, message2);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_8a].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_8a].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_8a].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_8b(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_8b, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8b];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8b];
			int index = 1;
			int index2 = 2;
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[79, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[64, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[175, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_68"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_71"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_72"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_207"));
			int[] array = GeneralFunctions.ArrayCellId("AU_7");
			int[] array2 = GeneralFunctions.ArrayCellId("AU_8");
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("T");
			int num5 = HTKKFile.ColumnLettersToIndex("BE");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("AV");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("AX");
			int num10 = HTKKFile.ColumnLettersToIndex("AA");
			int num11 = HTKKFile.ColumnLettersToIndex("AY");
			int num12 = HTKKFile.ColumnLettersToIndex("AF");
			int num13 = HTKKFile.ColumnLettersToIndex("AI");
			int num14 = HTKKFile.ColumnLettersToIndex("AM");
			int num15 = HTKKFile.ColumnLettersToIndex("AU");
			if (hTKKCMCSheetInfo != null)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num16 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array3 = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					int rowIndex = hTKKCMCTableInfo.Header.StaticCells[0].CellTemplate.RowIndex;
					int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[index].CellTemplate.RowIndex;
					int colIndex = hTKKCMCTableInfo.Header.StaticCells[index].CellTemplate.ColIndex;
					int rowIndex3 = hTKKCMCTableInfo.Header.StaticCells[index2].CellTemplate.RowIndex;
					int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[index2].CellTemplate.ColIndex;
					HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num13 - 1);
					HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, rowIndex3 - 1, colIndex2 - 1);
					string text6 = sheetView.Cells[rowIndex2 - 1, colIndex - 1].Text;
					string text7 = sheetView.Cells[rowIndex3 - 1, colIndex2 - 1].Text;
					string text8 = sheetView.Cells[rowIndex3 - 1, num15 - 1].Text;
					if (!string.IsNullOrEmpty(text7) && !ValidateBase.CheckMaSoThue(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex3 - 1, colIndex2 - 1, text3);
						value = 2;
					}
					double num17 = Convert.ToDouble(sheetView.Cells[rowIndex - 1, num13 - 1].Value);
					if (num17 != 100.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num13 - 1, message5);
						value = 2;
					}
					List<int> rowIndexInSViewByCompareHeader = GeneralFunctions.GetRowIndexInSViewByCompareHeader(sheetView, hTKKCMCSheetInfo.Tables, num2, text6);
					if (string.IsNullOrEmpty(text6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num2 - 1, text);
						value = 2;
					}
					else if (rowIndexInSViewByCompareHeader.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num2 - 1, message8);
						value = 2;
					}
					else if (hTKKCMCTableInfo.RowsValues.Count < 2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num2 - 1, message6);
						value = 2;
					}
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						HTKKFile.RemoveValidate(sheetView, num16 + j, num - 1);
						HTKKFile.RemoveValidate(sheetView, num16 + j, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num16 + j, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, num16 + j, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, num16 + j, num10 - 1);
						HTKKFile.RemoveValidate(sheetView, num16 + j, num - 1);
						string text9 = sheetView.Cells[num16 + j, num15 - 1].Text;
						string text10 = sheetView.Cells[num16 + j, num - 1].Text;
						string text11 = sheetView.Cells[num16 + j, num2 - 1].Text;
						string text12 = sheetView.Cells[num16 + j, num3 - 1].Text;
						string text13 = sheetView.Cells[num16 + j, num4 - 1].Text;
						string text14 = sheetView.Cells[num16 + j, num5 - 1].Text;
						string text15 = sheetView.Cells[num16 + j, num6 - 1].Text;
						string text16 = sheetView.Cells[num16 + j, num8 - 1].Text;
						string text17 = sheetView.Cells[num16 + j, num10 - 1].Text;
						double num18 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num16 + j, num12 - 1].Value);
						double num19 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num16 + j, num13 - 1].Value);
						bool flag = true;
						if (!string.IsNullOrEmpty(text10) || !string.IsNullOrEmpty(text11) || !string.IsNullOrEmpty(text12) || !string.IsNullOrEmpty(text13) || !string.IsNullOrEmpty(text15) || !string.IsNullOrEmpty(text16) || num18 != 0.0 || num19 != 0.0)
						{
							flag = false;
						}
						if (!flag && string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num - 1, text);
							value = 2;
						}
						List<int> list = new List<int>();
						list.Add(num4);
						list.Add(num6);
						list.Add(num8);
						list.Add(num10);
						List<string> list2 = new List<string>();
						list2.Add(text13);
						list2.Add(text15);
						list2.Add(text16);
						list2.Add(text17);
						List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
						if (rowIndexInSViewByCompares.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num4 - 1, text5);
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num6 - 1, text5);
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num8 - 1, text5);
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num10 - 1, text5);
							value = 2;
						}
						if (string.IsNullOrEmpty(text13) && !sheetView.Cells[num16 + j, num4 - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num4 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text15) && !sheetView.Cells[num16 + j, num6 - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num6 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text16) && !sheetView.Cells[num16 + j, num8 - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num8 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text17) && !sheetView.Cells[num16 + j, num10 - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num10 - 1, text);
							value = 2;
						}
						List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num15, "01");
						if (rowIndexInSViewByCompare.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num16 + j, num - 1, message7);
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_8b].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_8b].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_8b].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_8c(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_8c, out value);
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8c];
			if (!sheetView.Visible)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_68"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_196"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_197"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8c];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			string text6 = "";
			string text7 = "";
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("AP");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("J");
			int num6 = HTKKFile.ColumnLettersToIndex("Q");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("R");
			int num9 = HTKKFile.ColumnLettersToIndex("T");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("AS");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("AI");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text8 = sheetView.Cells[num + i, num2 - 1].Text;
				text2 = sheetView.Cells[num + i, num3 - 1].Text;
				text = sheetView.Cells[num + i, num4 - 1].Text;
				text3 = sheetView.Cells[num + i, num5 - 1].Text;
				text4 = sheetView.Cells[num + i, num6 - 1].Text;
				text5 = sheetView.Cells[num + i, num7 - 1].Text;
				text6 = sheetView.Cells[num + i, num8 - 1].Text;
				text7 = sheetView.Cells[num + i, num9 - 1].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num7 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, num9 - 1);
				if (string.IsNullOrEmpty(text8) || string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text5) && !sheetView.Cells[num + i, num7 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text4) && !sheetView.Cells[num + i, num6 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num6 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text7) && !sheetView.Cells[num + i, num9 - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9 - 1, message);
					value = 2;
				}
				if (text2 == "01" || text2 == "02" || text2 == "03")
				{
					if (GeneralFunctions.CheckTinhTheoCQTDangBangKhac_TSC(TTinChung.maCQTCapCuc, TTinChung.maCQTNoiNop, sheetView, num7, colMaTinh, num6, colMaHuyen, -1, -1, -1, num + i + 1))
					{
						value = 2;
					}
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num7, text5);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7 - 1, message2);
						value = 2;
					}
				}
				if (text2 == "01")
				{
					if (string.IsNullOrEmpty(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
						value = 2;
					}
					else if (text3.Length == 13 || text3.Length == 14)
					{
						if (!ValidateBase.CheckMaSoThue(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_80")));
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_74")));
						value = 2;
					}
				}
				else if (text2 == "02")
				{
					if (string.IsNullOrEmpty(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66")));
						value = 2;
					}
					else
					{
						if (text3.Length != 5)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66")));
							value = 2;
						}
						if (text3.Length == 5 && !ValidateBase.CheckMaSoThue5So(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_81")));
							value = 2;
						}
					}
				}
				if (text2 == "01")
				{
					List<int> list = new List<int>();
					list.Add(num3);
					list.Add(num6);
					list.Add(num7);
					list.Add(num8);
					list.Add(num9);
					List<string> list2 = new List<string>();
					list2.Add(text2);
					list2.Add(text4);
					list2.Add(text5);
					list2.Add(text6);
					list2.Add(text7);
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, message3);
						value = 2;
					}
				}
				if (text2 == "02")
				{
					List<int> list = new List<int>();
					list.Add(num3);
					list.Add(num6);
					list.Add(num7);
					list.Add(num8);
					list.Add(num9);
					List<string> list2 = new List<string>();
					list2.Add(text2);
					list2.Add(text4);
					list2.Add(text5);
					list2.Add(text6);
					list2.Add(text7);
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2 - 1, message4);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[idxPL03_8c].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[idxPL03_8c].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[idxPL03_8c].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_9(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(idxPL03_9, out value);
			SheetView sheetView = winFpSpread.Sheets[idxPL03_9];
			if (sheetView.Visible)
			{
				switch (value)
				{
				case 2:
					winFpSpread.TabStrip[idxPL03_9].BackColor = Color.Pink;
					break;
				case 1:
					winFpSpread.TabStrip[idxPL03_9].BackColor = Color.Yellow;
					break;
				default:
					winFpSpread.TabStrip[idxPL03_9].BackColor = Color.White;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_1_ND132(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = idxPL03_01GDLK_ND132;
			if (winFpSpread.Sheets[num] == null || !winFpSpread.Sheets[num].Visible)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			bool flag = false;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num2 = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num12 = num11 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (!sheetView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[num];
			int num13 = ToKhaiInfo.SheetInfos[num].Tables[0].RowsValues.Count - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
			int num14 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("I") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("U") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("W") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("X") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AA") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AC") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("I") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				else
				{
					if (string.IsNullOrEmpty(array[13]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("I") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_122")));
						value = 2;
					}
					if (array[1] == "Việt Nam" && !string.IsNullOrEmpty(array[2]) && !ValidateBase.CheckMaSoThue(array[2]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
						value = 2;
					}
				}
				if (!ConvertHelpers.ConvertToBool(array[3]) && !ConvertHelpers.ConvertToBool(array[4]) && !ConvertHelpers.ConvertToBool(array[5]) && !ConvertHelpers.ConvertToBool(array[6]) && !ConvertHelpers.ConvertToBool(array[7]) && !ConvertHelpers.ConvertToBool(array[8]) && !ConvertHelpers.ConvertToBool(array[9]) && !ConvertHelpers.ConvertToBool(array[10]) && !ConvertHelpers.ConvertToBool(array[11]) && !ConvertHelpers.ConvertToBool(array[12]) && !ConvertHelpers.ConvertToBool(array[13]) && !ConvertHelpers.ConvertToBool(array[14]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("W") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AA") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AE") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num14 + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
				{
					HTKKFile.RemoveValidate(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
					HTKKFile.RemoveValidate(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
					HTKKFile.RemoveValidate(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				HTKKFile.RemoveValidate(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				HTKKFile.RemoveValidate(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_II/ct1"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_II/ct2a"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_II/ct2c"));
			if (!ConvertHelpers.ConvertToBool(hTKKCMCCellInfo.Value))
			{
				for (int i = 1; i < 9; i++)
				{
					for (int num15 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num15 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num15++)
					{
						string[] array = hTKKCMCSheetInfo.Tables[i].RowsValues[num15 - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + 1].RowValue.ToArray();
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[i].RowsInfo, array))
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AG") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AH") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AI") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AJ") - 1);
							continue;
						}
						if (array[3].ToIntOrDefault() < 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_117")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("M") - 1);
						}
						if (array[7].ToIntOrDefault() > 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_118")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("Y") - 1);
						}
						if (array[1].ToIntOrDefault() > 0 && array[4].IsNullOrEmpty())
						{
							if (!hTKKCMCCellInfo2.Value.ToBoolOrDefault() && !hTKKCMCCellInfo3.Value.ToBoolOrDefault())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("N") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("N") - 1);
						}
						if (array[5].ToIntOrDefault() > 0 && array[8].IsNullOrEmpty())
						{
							if (!hTKKCMCCellInfo2.Value.ToBoolOrDefault() && !hTKKCMCCellInfo3.Value.ToBoolOrDefault())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AB") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("AB") - 1);
						}
					}
				}
				bool flag2 = hTKKCMCSheetInfo.StaticCells[3].Value.ToBoolOrDefault();
				bool flag3 = hTKKCMCSheetInfo.StaticCells[4].Value.ToBoolOrDefault();
				bool flag4 = hTKKCMCSheetInfo.StaticCells[5].Value.ToBoolOrDefault();
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_2/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_4/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_6/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_7/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_9/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_2/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, testColor: true);
				if (flag2)
				{
					value = CheckTongCot6MucIVP1P3_ND132(value, sheetView, hTKKCMCSheetInfo);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1, testColor: true);
				}
				if (flag3)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/khong"));
					HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_2/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_2/khong"));
					HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/khong"));
					if (!hTKKCMCCellInfo14.Value.ToBoolOrDefault() && !hTKKCMCCellInfo15.Value.ToBoolOrDefault() && !hTKKCMCCellInfo16.Value.ToBoolOrDefault() && !hTKKCMCCellInfo17.Value.ToBoolOrDefault() && !hTKKCMCCellInfo18.Value.ToBoolOrDefault() && !hTKKCMCCellInfo19.Value.ToBoolOrDefault())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, testColor: true);
					}
				}
				for (int i = 9; i < 12; i++)
				{
					for (int num15 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num15 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num15++)
					{
						string[] array = hTKKCMCSheetInfo.Tables[i].RowsValues[num15 - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + 1].RowValue.ToArray();
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[i].RowsInfo, array) && (!flag4 || i != 9))
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("D") - 1);
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("V") - 1);
						}
						else if (string.IsNullOrEmpty(sheetView.Cells[num15, HTKKFile.ColumnLettersToIndex("D") - 1].Text) && !sheetView.Cells[num15, HTKKFile.ColumnLettersToIndex("D") - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num15, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num15, HTKKFile.ColumnLettersToIndex("D") - 1);
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static int CheckTongCot6MucIVP1P3_ND132(int iwarningColor, SheetView sView, HTKKCMCSheetInfo sheetInfo)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_4/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_6/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_7/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_1/Dong_9/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_1/ct06"));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_IV/IV_3/Dong_2/ct06"));
			bool flag = false;
			if (sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m)
			{
				flag = true;
			}
			if (!flag)
			{
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, testColor: true);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, testColor: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return iwarningColor;
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year - 1;
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("//tieuMucHachToan");
			if ((val != null) & !string.IsNullOrEmpty(PeriodInfo.TieuMucHachToanVal))
			{
				val.InnerText = PeriodInfo.TieuMucHachToanVal;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportTableF7XML_8(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[8];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			XmlNodeList val = xmlNode.SelectNodes("//PLuc_03_3D_TNDN/NoiDungPL03_3D");
			string text = "//congNghe_sanPham";
			int num = 0;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				XmlNode val2 = val[i];
				for (int num2 = hTKKCMCSheetInfo.Tables[i].Header.StaticCells.Count - 1; num2 > 0; num2--)
				{
					num = 0;
					if (hTKKCMCSheetInfo.Tables[i].Header.StaticCells[num2].Path.IndexOf("congNghe_sanPham/ten") > -1 && !hTKKCMCSheetInfo.Tables[i].Header.StaticCells[num2].Path.Equals("congNghe_sanPham/ten"))
					{
						string text2 = winFpSpread.Sheets[8].Cells[hTKKCMCSheetInfo.Tables[i].Header.StaticCells[num2].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.Tables[i].Header.StaticCells[num2].CellTemplate.ColIndex - 1].Text;
						if (text2 != "")
						{
							if (val2.SelectNodes("congNghe_sanPham")[0].Attributes["id"] != null)
							{
								((XmlNode)val2.SelectNodes("congNghe_sanPham")[0].Attributes["id"]).Value = "ID_1";
							}
							if (val2.SelectNodes("congNghe_sanPham")[0].Attributes["ID"] != null)
							{
								((XmlNode)val2.SelectNodes("congNghe_sanPham")[0].Attributes["ID"]).Value = "ID_1";
							}
							XmlNode val3 = val2.SelectNodes("congNghe_sanPham")[0].Clone();
							string path = hTKKCMCSheetInfo.Tables[i].Header.StaticCells[num2].Path;
							val3.SelectSingleNode("ten").InnerText = text2;
							if (val3.Attributes["id"] != null)
							{
								((XmlNode)val3.Attributes["id"]).Value = "ID_" + (ConvertHelpers.ConvertToInt(path.Substring(path.Length - 1, 1)) + 1);
							}
							if (val3.Attributes["ID"] != null)
							{
								((XmlNode)val3.Attributes["ID"]).Value = "ID_" + (ConvertHelpers.ConvertToInt(path.Substring(path.Length - 1, 1)) + 1);
							}
							val2.InsertAfter(val3, val2.SelectNodes("congNghe_sanPham")[0]);
							num++;
						}
					}
					if (num >= 5)
					{
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_17(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			int index = idxPL03_01GDLK_ND132;
			winFpSpread.SuspendLayout();
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0])
			{
				bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value);
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = flag;
			}
			if ((idxRow == 42 + count || idxRow == 41 + count || idxRow == 40 + count || idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count) && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				bool flag2 = ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value);
				bool flag3 = ConvertHelpers.ConvertToBool(sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value);
				bool flag4 = ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value);
				bool flag5 = ConvertHelpers.ConvertToBool(sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value);
				bool flag6 = ConvertHelpers.ConvertToBool(sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value);
				bool flag7 = ConvertHelpers.ConvertToBool(sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value);
				sheetView.Cells[39 + count, GeneralFunctions.ArrayCellId("N_41")[0]].Value = flag2 || flag3 || flag4 || flag5 || flag6 || flag7;
			}
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0])
			{
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value))
			{
				sheetView.Cells[39 + count, GeneralFunctions.ArrayCellId("N_41")[0]].Value = false;
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
			}
			if (idxRow == 40 + count && idxCol == GeneralFunctions.ArrayCellId("N_42")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 41 + count && idxCol == GeneralFunctions.ArrayCellId("N_43")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 42 + count && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
				{
					sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
					sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
					sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
					sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
					AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
				}
				else
				{
					sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
					sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
					sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				}
			}
			if ((idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count) && idxCol == GeneralFunctions.ArrayCellId("N_44")[0] && (ConvertHelpers.ConvertToBool(sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value)))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[37 + count, GeneralFunctions.ArrayCellId("N_39")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = true;
				AnMucIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow >= 38 + count && idxRow <= 45 + count && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				XoaDLIII_IV_ND132(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count, ToKhaiInfo);
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AG") - 1].Formula = $"AG{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AG{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AH") - 1].Formula = $"AH{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AH{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AI") - 1].Formula = $"AI{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AI{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AJ") - 1].Formula = $"AJ{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AJ{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					for (int num12 = 1; num12 < 9; num12++)
					{
						for (int num13 = hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1; num13 < hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[num12].RowsValues.Count; num13++)
						{
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
						}
					}
					bool flag8 = ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && idxRow == 40 + count)
					{
						for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
						{
							if (num12 != 109 + num8 && num12 != 111 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							}
						}
						for (int num12 = 106 + num8; num12 < 132 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value) && (idxRow == 42 + count || idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count))
					{
						for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
						for (int num12 = 106 + num8; num12 < 132 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
					}
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && idxRow == 40 + count)
					{
						for (int num12 = 220 + num10; num12 < 234 + num10; num12++)
						{
							if (num12 != 224 + num10 && num12 != 225 + num10 && num12 != 226 + num10 && num12 != 233 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							}
						}
						for (int num12 = 175 + num10; num12 < 234 + num10; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
					}
					if (!ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
					{
						for (int num12 = 220 + num10; num12 < 234 + num10; num12++)
						{
							if (num12 != 224 + num10 && num12 != 225 + num10 && num12 != 226 + num10 && num12 != 233 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							}
						}
						for (int num12 = 175 + num10; num12 < 234 + num10; num12++)
						{
							if (num12 == 175 + num10 || num12 == 176 + num10 || num12 == 190 + num10 || num12 == 191 + num10 || num12 == 206 + num10 || num12 == 211 + num10 || num12 == 219 + num10 || num12 == 225 + num10 || num12 == 226 + num10 || num12 == 227 + num10 || num12 == 233 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (num12 + 1);
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
					}
				}
				else
				{
					sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
					sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AG") - 1].Formula = $"AG{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AG{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AH") - 1].Formula = $"AH{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AH{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AI") - 1].Formula = $"AI{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AI{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AJ") - 1].Formula = $"AJ{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AJ{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					for (int num12 = 1; num12 < 9; num12++)
					{
						for (int num13 = hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1; num13 < hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[num12].RowsValues.Count; num13++)
						{
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = string.Format("J{0}-H{0}", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = string.Format("V{0}-R{0}", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = string.Format("ABS(M{0})+ABS(Y{0})", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
						}
					}
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						}
					}
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 == 109 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 == 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (num12 + 1).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = 121 + num8; num12 < 133 + num8; num12++)
					{
						if (num12 == 121 + num8 || num12 == 122 + num8 || num12 == 123 + num8 || num12 == 124 + num8 || num12 == 125 + num8 || num12 == 131 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (num12 + 1);
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					for (int num12 = 220 + num10; num12 < 234 + num10; num12++)
					{
						if (num12 != 224 + num10 && num12 != 225 + num10 && num12 != 226 + num10 && num12 != 233 + num10)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						}
					}
					for (int num12 = 175 + num10; num12 < 234 + num10; num12++)
					{
						if (num12 == 175 + num10 || num12 == 176 + num10 || num12 == 190 + num10 || num12 == 191 + num10 || num12 == 206 + num10 || num12 == 211 + num10 || num12 == 219 + num10 || num12 == 225 + num10 || num12 == 226 + num10 || num12 == 227 + num10 || num12 == 233 + num10)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (num12 + 1);
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
			}
			if (idxRow == 100 + num8 && (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] || idxCol == GeneralFunctions.ArrayCellId("U_102")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value = false;
				}
			}
			if (idxRow == 139 + num9 && (idxCol == GeneralFunctions.ArrayCellId("X_141")[0] || idxCol == GeneralFunctions.ArrayCellId("U_141")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_141")[0] && sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("U_141")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_141")[0] && sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("U_141")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value = false;
				}
			}
			if (idxRow == 169 + num10 && (idxCol == GeneralFunctions.ArrayCellId("X_171")[0] || idxCol == GeneralFunctions.ArrayCellId("U_171")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_171")[0] && sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("U_171")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_171")[0] && sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("U_171")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value = false;
				}
			}
			if (idxRow == 100 + num8 && (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] || idxCol == GeneralFunctions.ArrayCellId("U_102")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value = false;
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = 126 + num8; num12 < 132 + num8; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						}
					}
					for (int num12 = 126 + num8; num12 < 132 + num8; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (idxRow == 139 + num9 && (idxCol == GeneralFunctions.ArrayCellId("X_141")[0] || idxCol == GeneralFunctions.ArrayCellId("U_141")[0]))
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value))
				{
					for (int num12 = 145 + num9; num12 < 160 + num9; num12++)
					{
						if (num12 != 147 + num9 && num12 != 150 + num9 && num12 != 156 + num9)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else
				{
					for (int num12 = 145 + num9; num12 < 160 + num9; num12++)
					{
						if (num12 != 147 + num9 && num12 != 150 + num9 && num12 != 156 + num9)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (idxRow != 169 + num10 || (idxCol != GeneralFunctions.ArrayCellId("X_171")[0] && idxCol != GeneralFunctions.ArrayCellId("U_171")[0]))
			{
				return;
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value))
			{
				for (int num12 = 220 + num10; num12 < 225 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = 228 + num10; num12 < 233 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int num12 = 220 + num10; num12 < 224 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = 228 + num10; num12 < 233 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int num12 = 220 + num10; num12 < 224 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int num12 = 228 + num10; num12 < 233 + num10; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		finally
		{
			winFpSpread.ResumeLayout();
		}
	}

	private static void AnMucIII_IV_ND132(int idxRow, int idxCol, SheetView sView, int rowCountTable1)
	{
		try
		{
			if (ConvertHelpers.ConvertToBool(sView.Cells[idxRow, idxCol].Value))
			{
				for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
				{
					if (!(sView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x"))
					{
						continue;
					}
					for (int j = 0; j < sView.NonEmptyColumnCount; j++)
					{
						Cell cell = sView.Cells[i, j];
						if (!cell.Locked && cell.CellType != null)
						{
							if (cell.CellType.ToString() == "NumberCellType")
							{
								cell.Value = 0;
							}
							else if (cell.CellType.ToString() == "CheckBoxCellType")
							{
								cell.Value = false;
							}
							else
							{
								cell.Value = "";
							}
						}
					}
					sView.Rows[i].Visible = false;
				}
				return;
			}
			for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
			{
				if (sView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x")
				{
					sView.Rows[i].Visible = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void XoaDLIII_IV_ND132(int idxRow, int idxCol, SheetView sView, int rowCountTable1, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
			{
				if (!(sView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x"))
				{
					continue;
				}
				for (int j = 0; j < sView.NonEmptyColumnCount; j++)
				{
					Cell cell = sView.Cells[i, j];
					if (cell.Locked && !cell.Note.IsNulOrEmpty())
					{
						HTKKFile.RemoveValidateBackGround(sView, i, j);
					}
					else if (!cell.Locked && !cell.Note.IsNulOrEmpty())
					{
						HTKKFile.RemoveValidate(sView, i, j);
					}
					if (!cell.Locked && cell.CellType != null)
					{
						if (cell.CellType.ToString() == "NumberCellType")
						{
							cell.Value = 0;
						}
						else if (cell.CellType.ToString() == "CheckBoxCellType")
						{
							cell.Value = false;
						}
						else
						{
							cell.Value = "";
						}
					}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			int[] array = GeneralFunctions.ArrayCellId("O_50");
			int[] array2 = GeneralFunctions.ArrayCellId("O_51");
			int[] array3 = GeneralFunctions.ArrayCellId("O_53");
			int[] array4 = GeneralFunctions.ArrayCellId("O_54");
			int[] array5 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3");
			int[] array6 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3a");
			int[] array7 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3b");
			int[] array8 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC4");
			int[] array9 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC6");
			int[] array10 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7_tenThueSuat");
			int[] array11 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC8");
			int[] array12 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC8a");
			int[] array13 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC11");
			int[] array14 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC12");
			int[] array15 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC13");
			int[] array16 = GeneralFunctions.ArrayCellId("P_53");
			int[] array17 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC5");
			if (iRow == array13[1] + 1 || iRow == array14[1] + 1 || iRow == array15[1] + 1 || iRow == array10[1] + 1 || iRow == array11[1] + 1 || iRow == array12[1] + 1)
			{
				TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iRow == array[1] + 1)
			{
				sheetView.Cells[array3[1], array3[0]].Value = sheetView.Cells[array16[1], array16[0]].Value;
			}
			if (iRow == array17[1] + 1 || iRow == array3[1] + 1 || iRow == array4[1] + 1 || iRow == array[1] + 1 || iRow == array2[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iRow == array3[1] + 1 || iRow == array4[1] + 1)
			{
				LockUnLockC3bTK(sheetView, sInfo);
			}
			if (iRow == array3[1] + 1 || iRow == array4[1] + 1)
			{
				double num = Convert.ToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num2 = Convert.ToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				if (num - num2 <= 0.0)
				{
					sheetView.Cells[array5[1], array5[0]].Value = 0;
					sheetView.Cells[array6[1], array6[0]].Value = 0;
					sheetView.Cells[array7[1], array7[0]].Value = 0;
					sheetView.Cells[array8[1], array8[0]].Value = 0;
					sheetView.Cells[array17[1], array17[0]].Value = 0;
					sheetView.Cells[array9[1], array9[0]].Value = 0;
				}
			}
			SheetView sheetView2 = winFpSpread.Sheets[idxPL03_3d];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3d];
			int num3 = HTKKFile.ColumnLettersToIndex("AG");
			int num4 = HTKKFile.ColumnLettersToIndex("BA");
			if (hTKKCMCSheetInfo != null && iRow == array10[1] + 1)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num5 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						sheetView2.Cells[num5 + j, num3 - 1].Value = sheetView2.Cells[num5 + j, num4 - 1].Value;
					}
				}
			}
			if (iRow >= array3[1] + 1 && iRow < array5[1])
			{
				tinhCtC3a_tk(winFpSpread, ToKhaiInfo);
				lockUnlockCtC3a(sheetView, sInfo);
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
			int num = HTKKFile.ColumnLettersToIndex("E");
			if (iCol == num)
			{
				TongHopCtA1_TK(winFpSpread.Sheets[0]);
			}
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
			int num = HTKKFile.ColumnLettersToIndex("E");
			if (iCol == num)
			{
				TongHopCtA1_TK(winFpSpread.Sheets[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("E");
			if (iCol == num)
			{
				TongHopCtA1_TK(winFpSpread.Sheets[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("M");
			int[] array = GeneralFunctions.ArrayCellId("M_33");
			int[] array2 = GeneralFunctions.ArrayCellId("M_41");
			if (iCol == num && iRow <= array[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				LockUnLockC3bTK(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num && iRow <= array2[1] + 1 && iRow >= array[1] + 1)
			{
				TongHopCtD2_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num && iRow < array[1] + 1)
			{
				tinhCtC3a_tk(winFpSpread, ToKhaiInfo);
				lockUnlockCtC3a(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3a];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3a];
			int num = HTKKFile.ColumnLettersToIndex("K");
			int num2 = HTKKFile.ColumnLettersToIndex("X");
			int num3 = HTKKFile.ColumnLettersToIndex("Y");
			int num4 = HTKKFile.ColumnLettersToIndex("Z");
			int num5 = HTKKFile.ColumnLettersToIndex("AA");
			int num6 = HTKKFile.ColumnLettersToIndex("AE");
			int num7 = HTKKFile.ColumnLettersToIndex("BD");
			int num8 = HTKKFile.ColumnLettersToIndex("AF");
			int num9 = HTKKFile.ColumnLettersToIndex("AK");
			int num10 = HTKKFile.ColumnLettersToIndex("I");
			int num11 = HTKKFile.ColumnLettersToIndex("AX");
			int num12 = HTKKFile.ColumnLettersToIndex("K");
			int num13 = HTKKFile.ColumnLettersToIndex("AY");
			int num14 = HTKKFile.ColumnLettersToIndex("AG");
			int num15 = HTKKFile.ColumnLettersToIndex("AN");
			int num16 = HTKKFile.ColumnLettersToIndex("K");
			int num17 = HTKKFile.ColumnLettersToIndex("AW");
			int num18 = HTKKFile.ColumnLettersToIndex("C");
			int num19 = HTKKFile.ColumnLettersToIndex("BE");
			int num20 = HTKKFile.ColumnLettersToIndex("BF");
			if (iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5)
			{
				sheetView.Cells[(int)iRow - 1, num6 - 1].Value = sheetView.Cells[(int)iRow - 1, num7 - 1].Value;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num21 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
				HTKKCMCCellInfo cellTenTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
				HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
				HTKKCMCCellInfo cellTenHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
				HTKKCMCCellInfo cellTenXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct16"));
				bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				string text = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
				string text2 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				string text3 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
				string text4 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
				string text5 = sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text.Trim();
				if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex || iCol == num6 || iCol == num8)
				{
					TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
					TinhCtD7_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
				}
				if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex || iCol == num6 || iCol == num8 || iCol == num9)
				{
					TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
					TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
				}
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo4.CellTemplate.RowIndex)
				{
					LockTinhCQT((text == UU_DAI_KHAC_TINH) ? true : false, sheetView, cellTenTinh, cellMaTinh, cellTenHuyen, cellMaHuyen, cellTenXa, cellMaXa, null, hTKKCMCCellInfo4, hTKKCMCCellInfo6);
					LockUnlockMucBUuDaiCungTinh_PL03_3a(sheetView, hTKKCMCSheetInfo);
				}
				if (iRow == hTKKCMCCellInfo4.CellTemplate.RowIndex)
				{
					if (iCol == hTKKCMCCellInfo4.CellTemplate.ColIndex)
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: false, deleteData: true);
						sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text = "";
						text5 = sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text.Trim();
					}
					UpdateCacHangSauLuonGiongHangTruoc_PL03_3a(winFpSpread, ToKhaiInfo);
				}
				TinhLaiCQTTheoUuDai_PL03_3A(sheetView, TTinChung, PeriodInfo, hTKKCMCTableInfo);
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo4.CellTemplate.RowIndex)
				{
					UpdateCacHangSauLuonGiongHangTruoc_PL03_3a(winFpSpread, ToKhaiInfo);
				}
			}
			if (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num8 || iCol == num9)
			{
				TinhCt16PL03_3a(sheetView, hTKKCMCSheetInfo);
			}
			if (iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num8)
			{
				TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num8 || iCol == num9)
			{
				TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num8 || iCol == num9)
			{
				TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num8)
			{
				TinhCtD7_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num18 || iCol == num16 || (iCol >= num2 && iCol <= num15))
			{
				TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_23(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3b];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3b];
			int num = HTKKFile.ColumnLettersToIndex("Y");
			int num2 = HTKKFile.ColumnLettersToIndex("Z");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			int num4 = HTKKFile.ColumnLettersToIndex("AB");
			int num5 = HTKKFile.ColumnLettersToIndex("AF");
			int num6 = HTKKFile.ColumnLettersToIndex("BE");
			int num7 = HTKKFile.ColumnLettersToIndex("AG");
			int num8 = HTKKFile.ColumnLettersToIndex("AL");
			int num9 = HTKKFile.ColumnLettersToIndex("I");
			int num10 = HTKKFile.ColumnLettersToIndex("AY");
			int num11 = HTKKFile.ColumnLettersToIndex("K");
			int num12 = HTKKFile.ColumnLettersToIndex("AZ");
			int num13 = HTKKFile.ColumnLettersToIndex("AH");
			int num14 = HTKKFile.ColumnLettersToIndex("AO");
			int num15 = HTKKFile.ColumnLettersToIndex("K");
			int num16 = HTKKFile.ColumnLettersToIndex("AW");
			if (iCol == num || iCol == num2 || iCol == num3 || iCol == num4)
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Value = sheetView.Cells[(int)iRow - 1, num6 - 1].Value;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num17 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
				HTKKCMCCellInfo cellTenTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
				HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
				HTKKCMCCellInfo cellTenHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
				HTKKCMCCellInfo cellTenXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct16"));
				string text = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
				string text2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				string text3 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				string text4 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
				string text5 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
				{
					LockTinhCQT((text == UU_DAI_KHAC_TINH) ? true : false, sheetView, cellTenTinh, cellMaTinh, cellTenHuyen, cellMaHuyen, cellTenXa, cellMaXa, null, hTKKCMCCellInfo3, hTKKCMCCellInfo5);
					LockUnlockMucBUuDaiCungTinh_PL03_3b(sheetView, hTKKCMCSheetInfo);
				}
				if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
				{
					if (iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex)
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: true);
						sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
						text5 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
					}
					UpdateCacHangSauLuonGiongHangTruoc_PL03_3b(winFpSpread, ToKhaiInfo);
				}
				TinhLaiCQTTheoUuDai_PL03_3B(sheetView, TTinChung, PeriodInfo, hTKKCMCTableInfo);
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
				{
					UpdateCacHangSauLuonGiongHangTruoc_PL03_3b(winFpSpread, ToKhaiInfo);
				}
			}
			if (iCol == num15 || iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num7 || iCol == num8)
			{
				TinhCt16PL03_3b(sheetView, hTKKCMCSheetInfo);
			}
			if (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num7)
			{
				TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num7 || iCol == num8)
			{
				TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num7 || iCol == num8)
			{
				TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num15 || (iCol >= num && iCol <= num14))
			{
				TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_7(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3c];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3c];
			int num = HTKKFile.ColumnLettersToIndex("T");
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num3 = HTKKFile.ColumnLettersToIndex("U");
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
				HTKKCMCCellInfo cellTenTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
				HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
				HTKKCMCCellInfo cellTenHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
				HTKKCMCCellInfo cellTenXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_sdld_nu"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("sold_nu"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("so_xn_sdld_nu"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ngay_xn_sdld_nu"));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_sdld_dtoc"));
				HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("sold_dtoc"));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("so_xn_sdld_dtoc"));
				HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ngay_xn_sdld_dtoc"));
				HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("giam_TNDN_ldNu"));
				HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("giam_TNDN_ldDToc"));
				HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct14"));
				HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct17"));
				HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct19"));
				HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/ct10");
				HTKKCMCCellInfo hTKKCMCCellInfo21 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/ct15");
				HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/ct18");
				HTKKCMCCellInfo hTKKCMCCellInfo23 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/ct20");
				HTKKCMCCellInfo hTKKCMCCellInfo24 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct21");
				HTKKCMCCellInfo hTKKCMCCellInfo25 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct22");
				HTKKCMCCellInfo hTKKCMCCellInfo26 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct23");
				HTKKCMCCellInfo hTKKCMCCellInfo27 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct24");
				HTKKCMCCellInfo hTKKCMCCellInfo28 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct25");
				if (iCol == num2 || iRow == hTKKCMCCellInfo21.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo22.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo23.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, hTKKCMCCellInfo24.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, num - 1].Value;
				}
				if (iRow == hTKKCMCCellInfo22.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo23.CellTemplate.RowIndex)
				{
					TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
				}
				if (iRow >= hTKKCMCCellInfo20.CellTemplate.RowIndex && iRow < hTKKCMCCellInfo21.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo21.CellTemplate.RowIndex - 1, hTKKCMCCellInfo21.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo21.CellTemplate.RowIndex - 1, num3 - 1].Value;
				}
				string text = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
				string text2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				string text3 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				string text4 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
				string text5 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
				{
					LockTinhCQT((text == UU_DAI_KHAC_TINH) ? true : false, sheetView, cellTenTinh, cellMaTinh, cellTenHuyen, cellMaHuyen, cellTenXa, cellMaXa, null, hTKKCMCCellInfo3, hTKKCMCCellInfo5);
				}
				if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex)
				{
					GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: true);
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
					text5 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
				}
				bool flag = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value);
				bool flag2 = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value);
				if (iRow == hTKKCMCCellInfo7.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Value = flag;
					if (flag)
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex, hTKKCMCCellInfo8.CellTemplate.ColIndex, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex, hTKKCMCCellInfo9.CellTemplate.ColIndex, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex, hTKKCMCCellInfo10.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex, hTKKCMCCellInfo8.CellTemplate.ColIndex, locked: true, deleteData: true);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex, hTKKCMCCellInfo9.CellTemplate.ColIndex, locked: true, deleteData: true);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex, hTKKCMCCellInfo10.CellTemplate.ColIndex, locked: true, deleteData: true);
					}
				}
				if (iRow == hTKKCMCCellInfo11.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1].Value = flag2;
					if (flag2)
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo12.CellTemplate.RowIndex, hTKKCMCCellInfo12.CellTemplate.ColIndex, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo13.CellTemplate.RowIndex, hTKKCMCCellInfo13.CellTemplate.ColIndex, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex, hTKKCMCCellInfo14.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo12.CellTemplate.RowIndex, hTKKCMCCellInfo12.CellTemplate.ColIndex, locked: true, deleteData: true);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo13.CellTemplate.RowIndex, hTKKCMCCellInfo13.CellTemplate.ColIndex, locked: true, deleteData: true);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex, hTKKCMCCellInfo14.CellTemplate.ColIndex, locked: true, deleteData: true);
					}
				}
				if (iRow == hTKKCMCCellInfo7.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo11.CellTemplate.RowIndex)
				{
					KhoaCt17Ct19PL_03_3c(sheetView, hTKKCMCSheetInfo);
				}
				LockUnlockMucBUuDaiCungTinh_PL03_3c(sheetView, ToKhaiInfo.SheetInfos[idxPL03_3c]);
				if (iRow >= hTKKCMCCellInfo20.CellTemplate.RowIndex && iRow < hTKKCMCCellInfo17.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, num - 1].Value;
				}
				if (iRow == hTKKCMCCellInfo21.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo22.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo23.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
				{
					TinhCt21PL03_3c(sheetView, hTKKCMCSheetInfo);
					TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3d];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3d];
			int num = HTKKFile.ColumnLettersToIndex("I");
			int num2 = HTKKFile.ColumnLettersToIndex("AW");
			int num3 = HTKKFile.ColumnLettersToIndex("K");
			int num4 = HTKKFile.ColumnLettersToIndex("AX");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("Q");
			int num7 = HTKKFile.ColumnLettersToIndex("R");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("AF");
			int num10 = HTKKFile.ColumnLettersToIndex("AG");
			int num11 = HTKKFile.ColumnLettersToIndex("BA");
			int num12 = HTKKFile.ColumnLettersToIndex("BB");
			int num13 = HTKKFile.ColumnLettersToIndex("BA");
			int num14 = HTKKFile.ColumnLettersToIndex("AL");
			int num15 = HTKKFile.ColumnLettersToIndex("AQ");
			int num16 = HTKKFile.ColumnLettersToIndex("K");
			int num17 = HTKKFile.ColumnLettersToIndex("AW");
			if (iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8)
			{
				sheetView.Cells[(int)iRow - 1, num9 - 1].Value = sheetView.Cells[(int)iRow - 1, num12 - 1].Value;
			}
			if (iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8 || iCol == num9)
			{
				sheetView.Cells[(int)iRow - 1, num10 - 1].Value = sheetView.Cells[(int)iRow - 1, num11 - 1].Value;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num18 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10a_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10a_ma"));
				HTKKCMCCellInfo cellTenTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
				HTKKCMCCellInfo cellMaTinh = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ma"));
				HTKKCMCCellInfo cellTenHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo cellMaHuyen = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ma"));
				HTKKCMCCellInfo cellTenXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo cellMaXa = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct11"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct12"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct13"));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct14"));
				HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct15"));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct16"));
				HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct17"));
				HTKKCMCCellInfo toRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct18"));
				HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_chuyenGiao"));
				HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_khoaHocCN"));
				HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("so_gcn"));
				HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ngayCap"));
				HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("noiCap"));
				HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct23"));
				if (iRow == hTKKCMCCellInfo9.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo10.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, num17 - 1].Value;
				}
				bool flag = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Value);
				bool flag2 = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1].Value);
				if (iRow == hTKKCMCCellInfo15.CellTemplate.RowIndex)
				{
					if (flag)
					{
						sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value = 50;
						LockUnlockTrong1Khoang(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo7, hTKKCMCCellInfo9, null, isLock: false, isDelete: false);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex, hTKKCMCCellInfo11.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value = 0;
						LockUnlockTrong1Khoang(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo7, hTKKCMCCellInfo9, null, isLock: true, isDelete: true);
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex, hTKKCMCCellInfo11.CellTemplate.ColIndex, locked: true, deleteData: true);
					}
				}
				if (iRow == hTKKCMCCellInfo16.CellTemplate.RowIndex)
				{
					if (flag2)
					{
						LockUnlockTrong1Khoang(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo7, toRow, hTKKCMCCellInfo16, isLock: false, isDelete: false);
					}
					else
					{
						LockUnlockTrong1Khoang(sheetView, hTKKCMCTableInfo, hTKKCMCCellInfo7, toRow, hTKKCMCCellInfo16, isLock: true, isDelete: true);
					}
				}
				if (iRow == hTKKCMCCellInfo11.CellTemplate.RowIndex)
				{
					TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
				}
				if (iRow == hTKKCMCCellInfo7.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo8.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, num11 - 1].Value;
				}
				if (iRow == hTKKCMCCellInfo12.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo13.CellTemplate.RowIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, num13 - 1].Value;
				}
				string text = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
				string text2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				string text3 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				string text4 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
				string text5 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
				{
					LockTinhCQT((text == UU_DAI_KHAC_TINH) ? true : false, sheetView, cellTenTinh, cellMaTinh, cellTenHuyen, cellMaHuyen, cellTenXa, cellMaXa, null, hTKKCMCCellInfo3, hTKKCMCCellInfo5);
					LockUnlockMucBUuDaiCungTinh_PL03_3d(sheetView, hTKKCMCSheetInfo);
				}
				if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
				{
					if (iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex)
					{
						GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: true);
						sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
						text5 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
					}
					UpdateCacHangSauLuonGiongHangTruoc_PL03_3d(winFpSpread, ToKhaiInfo);
				}
				TinhLaiCQTTheoUuDai_PL03_3D(sheetView, TTinChung, PeriodInfo, hTKKCMCTableInfo);
				if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
				{
					UpdateCacHangSauLuonGiongHangTruoc_PL03_3d(winFpSpread, ToKhaiInfo);
				}
			}
			if (iCol == num16 || iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8 || iCol == num9 || iCol == num10 || iCol == num14)
			{
				TinhCt16PL03_3d(sheetView, hTKKCMCSheetInfo);
			}
			if (iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8 || iCol == num9 || iCol == num10 || iCol == num14)
			{
				TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8 || iCol == num9 || iCol == num10 || iCol == num14)
			{
				TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num16 || (iCol >= num5 && iCol <= num15))
			{
				TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_9(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AL");
			if (iCol == num)
			{
				TongHopCtC16_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
				TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_24(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_10(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("E_30");
			if (iCol == num && iRow <= num2 + 1)
			{
				TongHopCtB15_TK(winFpSpread.Sheets[0]);
				TongHopCtD2_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				LockUnLockC3bTK(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_11(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("N_12");
			if (iRow == array[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			int num = HTKKFile.ColumnLettersToIndex("X");
			int num2 = HTKKFile.ColumnLettersToIndex("Z");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AD");
			if (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5)
			{
				TongHopCtE6_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
				TinhCtE5_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_12(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8];
			int[] array = GeneralFunctions.ArrayCellId("I_21");
			int[] array2 = GeneralFunctions.ArrayCellId("I_23");
			int num = HTKKFile.ColumnLettersToIndex("P");
			int colTenHuyen = HTKKFile.ColumnLettersToIndex("T");
			int num2 = HTKKFile.ColumnLettersToIndex("R");
			int num3 = HTKKFile.ColumnLettersToIndex("V");
			int colHide = HTKKFile.ColumnLettersToIndex("AA");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("BA");
			int num4 = HTKKFile.ColumnLettersToIndex("AS");
			int num5 = HTKKFile.ColumnLettersToIndex("AU");
			int num6 = HTKKFile.ColumnLettersToIndex("AV");
			int num7 = HTKKFile.ColumnLettersToIndex("AF");
			int num8 = HTKKFile.ColumnLettersToIndex("AI");
			int num9 = HTKKFile.ColumnLettersToIndex("AY");
			int num10 = HTKKFile.ColumnLettersToIndex("AM");
			int num11 = HTKKFile.ColumnLettersToIndex("AZ");
			int num12 = HTKKFile.ColumnLettersToIndex("AN");
			int num13 = HTKKFile.ColumnLettersToIndex("AO");
			int num14 = HTKKFile.ColumnLettersToIndex("E");
			int num15 = HTKKFile.ColumnLettersToIndex("AR");
			string text = sheetView.Cells[(int)iRow - 1, num15 - 1].Text.Trim();
			if (text == "01" && iCol == num14)
			{
				sheetView.Cells[(int)iRow - 1, num - 1].Value = TTinChung.mst;
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num, locked: true, deleteData: false);
				ThayDoiTinhTheoCQTDangBang(TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy, sheetView, hTKKCMCSheetInfo, num2, num4, colTenHuyen, colMaHuyen, (int)iRow, isLock: false);
			}
			if (iCol == num2)
			{
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
				string text2 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim();
				if (!(text != "01") && !(headerInfo.maCQTCapCuc != "82500"))
				{
				}
			}
			if (text == "02" || text == "03")
			{
				if (iCol == num14)
				{
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num2, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide, locked: false, deleteData: true);
				}
				if (iCol == num2)
				{
					SetLockUnlockBang_P_03_8(sheetView, hTKKCMCSheetInfo, TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
				}
			}
			if (iCol == num7)
			{
				sheetView.Cells[(int)iRow - 1, num8 - 1].Value = sheetView.Cells[(int)iRow - 1, num9 - 1].Value;
			}
			if (iRow == array2[1] + 1)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num8 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num9 - 1].Value;
				}
			}
			if (iCol == num7 || iCol == num8)
			{
				sheetView.Cells[(int)iRow - 1, num10 - 1].Value = sheetView.Cells[(int)iRow - 1, num11 - 1].Value;
			}
			if (iRow == array[1] + 1 || iRow == array2[1] + 1)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num10 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Value;
				}
			}
			if (iCol == num12)
			{
				TinhCtG1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num13)
			{
				TinhCtG2_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num14)
			{
				SetLockUnlockBang_P_03_8(sheetView, hTKKCMCSheetInfo, TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_13(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8a];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[idxPL03_8a];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("P");
			int num3 = HTKKFile.ColumnLettersToIndex("T");
			int num4 = HTKKFile.ColumnLettersToIndex("R");
			int num5 = HTKKFile.ColumnLettersToIndex("V");
			int colHide = HTKKFile.ColumnLettersToIndex("AA");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("BB");
			int num6 = HTKKFile.ColumnLettersToIndex("AV");
			int num7 = HTKKFile.ColumnLettersToIndex("AX");
			int num8 = HTKKFile.ColumnLettersToIndex("AY");
			int num9 = HTKKFile.ColumnLettersToIndex("E");
			int num10 = HTKKFile.ColumnLettersToIndex("AU");
			int num11 = HTKKFile.ColumnLettersToIndex("AM");
			int num12 = HTKKFile.ColumnLettersToIndex("AN");
			string text = sheetView.Cells[(int)iRow - 1, num10 - 1].Text.Trim();
			if (text == "01" && iCol == num9)
			{
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = TTinChung.mst;
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num2, locked: true, deleteData: false);
				ThayDoiTinhTheoCQTDangBang(TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy, sheetView, sInfo, num4, num6, num3, colMaHuyen, (int)iRow, isLock: false);
				SetLockUnlockBang_P_03_8A(sheetView, sInfo, TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
			}
			if (text != "01" && iCol == num9)
			{
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num2, locked: false, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num4, locked: false, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num3, locked: false, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide, locked: false, deleteData: true);
			}
			if (iCol == num4)
			{
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
				string text2 = sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim();
				if (!(text != "01") && !(headerInfo.maCQTCapCuc != "82500"))
				{
				}
			}
			if (iCol == num12)
			{
				TinhCtG4_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num11)
			{
				TinhCtG3_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (text == "04" && iCol == num9)
			{
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num2, locked: true, deleteData: true);
			}
			if (iCol == num9)
			{
				SetLockUnlockBang_P_03_8A(sheetView, sInfo, TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_14(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8b];
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8b];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("S");
			int num5 = HTKKFile.ColumnLettersToIndex("AW");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("AV");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("AX");
			int num10 = HTKKFile.ColumnLettersToIndex("AA");
			int num11 = HTKKFile.ColumnLettersToIndex("AY");
			int num12 = HTKKFile.ColumnLettersToIndex("AF");
			int num13 = HTKKFile.ColumnLettersToIndex("AI");
			int num14 = HTKKFile.ColumnLettersToIndex("AM");
			int num15 = HTKKFile.ColumnLettersToIndex("BI");
			int num16 = HTKKFile.ColumnLettersToIndex("AP");
			int num17 = HTKKFile.ColumnLettersToIndex("BJ");
			int colCheck = HTKKFile.ColumnLettersToIndex("AU");
			int colChilden = HTKKFile.ColumnLettersToIndex("BB");
			int colChilden2 = HTKKFile.ColumnLettersToIndex("BC");
			int colChilden3 = HTKKFile.ColumnLettersToIndex("BD");
			int num18 = HTKKFile.ColumnLettersToIndex("BE");
			int num19 = HTKKFile.ColumnLettersToIndex("BF");
			int num20 = HTKKFile.ColumnLettersToIndex("BG");
			int num21 = HTKKFile.ColumnLettersToIndex("AZ");
			int num22 = HTKKFile.ColumnLettersToIndex("BA");
			int[] array = GeneralFunctions.ArrayCellId("AU_7");
			int[] array2 = GeneralFunctions.ArrayCellId("AU_8");
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[72, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			if (iCol == num || iCol == num12)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, num12, colChilden);
			}
			if (iCol == num || iCol == num3)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, num3, colChilden2);
			}
			if (iCol == num || iCol == num2)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, num2, colChilden3);
			}
			if (iCol == num12 || iCol == num13)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					int rowIndex = table.Header.StaticCells[0].CellTemplate.RowIndex;
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					double num23 = GeneralFunctions.CalculateColSumTable(sheetView, table, num13);
					sheetView.Cells[dataLocationTemplateRowStart - 2, num13 - 1].Value = num23;
				}
			}
			int index = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[index], colCheck, "01");
			if (iCol == num && rowIndexInSViewByCompare.Count > 1)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num - 1, msg);
			}
			if (iCol == num)
			{
				SetLockUnlockBang_P_03_8B(sheetView, hTKKCMCSheetInfo, TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
			}
			if (iCol == num)
			{
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[index]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_15(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8c];
			SheetView sheetView = winFpSpread.Sheets[idxPL03_8c];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int[] array = GeneralFunctions.ArrayCellId("J_16");
			int[] array2 = GeneralFunctions.ArrayCellId("J_40");
			int num2 = HTKKFile.ColumnLettersToIndex("AP");
			int colHide = HTKKFile.ColumnLettersToIndex("J");
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("AI");
			int num5 = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("AK");
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int num8 = HTKKFile.ColumnLettersToIndex("AL");
			int num9 = HTKKFile.ColumnLettersToIndex("C");
			int num10 = HTKKFile.ColumnLettersToIndex("AL");
			int num11 = HTKKFile.ColumnLettersToIndex("U");
			int num12 = HTKKFile.ColumnLettersToIndex("V");
			int num13 = HTKKFile.ColumnLettersToIndex("AR");
			int num14 = HTKKFile.ColumnLettersToIndex("Y");
			int num15 = HTKKFile.ColumnLettersToIndex("Z");
			int num16 = HTKKFile.ColumnLettersToIndex("AA");
			int num17 = HTKKFile.ColumnLettersToIndex("AB");
			int num18 = HTKKFile.ColumnLettersToIndex("AC");
			if (iCol == num11)
			{
				sheetView.Cells[(int)iRow - 1, num12 - 1].Value = sheetView.Cells[(int)iRow - 1, num13 - 1].Value;
			}
			if (iRow == array2[1] + 1)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num12 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num13 - 1].Value;
				}
			}
			if (iCol == num15)
			{
				TinhCtG1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			if (iCol == num16)
			{
				TinhCtG2_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			}
			string text = sheetView.Cells[(int)iRow - 1, num2 - 1].Text.Trim();
			if (iCol == num)
			{
				if (text == "03")
				{
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide, locked: true, deleteData: true);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide, locked: false, deleteData: false);
				}
			}
			if (iRow == array[1] + 1 || iRow == array2[1] + 1)
			{
				TinhTongTheoTinh_Pl03_8C(winFpSpread, ToKhaiInfo, -1);
			}
			else
			{
				TinhTongTheoTinh_Pl03_8C(winFpSpread, ToKhaiInfo, (int)iCol);
			}
			if (iCol == num)
			{
				SetLockUnlockBang_P_03_8C(sheetView, hTKKCMCSheetInfo, TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_16(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_9];
			SheetView sheetView = winFpSpread.Sheets[idxPL03_9];
			int num = HTKKFile.ColumnLettersToIndex("R");
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			int num4 = HTKKFile.ColumnLettersToIndex("AV");
			int num5 = HTKKFile.ColumnLettersToIndex("AX");
			int num6 = HTKKFile.ColumnLettersToIndex("AY");
			int num7 = HTKKFile.ColumnLettersToIndex("AF");
			int num8 = HTKKFile.ColumnLettersToIndex("AI");
			int num9 = HTKKFile.ColumnLettersToIndex("AM");
			if (iCol == num7 || iCol == num8 || iCol == num9)
			{
				TongHopCtG5_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_17(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_01GDLK_ND132];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_01GDLK_ND132];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1_1.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo21 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo23 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo24 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo25 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo26 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo27 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo28 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo29 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo30 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo31 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo32 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.3/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo33 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo34 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo35 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo36 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo37 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo38 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo39 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo40 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo41 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo42 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo43 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo44 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.4_2.4.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo45 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo46 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo47 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo48 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2_2.5/ct12d"));
			if (iRow == hTKKCMCCellInfo5.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo17.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo5.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo6.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo7.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo8.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo9.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo13.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo9.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo13.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo10.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo10.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo14.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo11.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo11.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo15.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo12.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo12.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo16.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo21.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo25.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo29.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo33.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo45.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo21.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo22.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo23.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo24.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo37.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo41.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo41.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo37.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo41.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo42.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo38.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo42.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo43.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo39.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo43.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo44.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo40.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo44.CellTemplate.RowIndex}";
				}
			}
			for (int num = 1; num < 9; num++)
			{
				if (iRow < hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart || iRow > hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count)
				{
					continue;
				}
				if (num == 1 || num == 2)
				{
					if (iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo9.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo13.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo10.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo10.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo14.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo11.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo11.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo15.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo12.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo12.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo16.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
				else if (num == 3 || num == 4 || num == 5 || num == 8)
				{
					if (iCol == hTKKCMCCellInfo21.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo22.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo23.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo24.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
				else if (num == 6 || num == 7)
				{
					if (iCol == hTKKCMCCellInfo41.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo37.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo41.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo42.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo38.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo42.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo43.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo39.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo43.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo44.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AJ{hTKKCMCCellInfo40.CellTemplate.RowIndex}+AJ{hTKKCMCCellInfo44.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AJ{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
			}
			string pathCt6 = "PL_GDLK2-01/Muc_IV/IV_1/Dong_5/ct06";
			HTKKCMCCellInfo hTKKCMCCellInfo49 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt6);
			if (iRow < hTKKCMCCellInfo49.CellTemplate.RowIndex && iRow >= hTKKCMCCellInfo49.CellTemplate.RowIndex - 11)
			{
				sheetView.Cells[hTKKCMCCellInfo49.CellTemplate.RowIndex - 1, hTKKCMCCellInfo49.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo49.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_21(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_ND92];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_ND92];
			int[] array = GeneralFunctions.ArrayCellId("E_23");
			if (iRow == array[1] + 1)
			{
				TongHopCtC15_TK(winFpSpread.Sheets[0], _ToKhaiInfo.SheetInfos[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static double SumChiTieuCellTable(SheetView sView, HTKKCMCSheetInfo sInfo, int cellIndex)
	{
		double num = 0.0;
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				num += double.Parse(GeneralFunctions.GetValue(sView, sInfo.Tables[i].Header.StaticCells[cellIndex].CellTemplate.RowIndex, sInfo.Tables[i].Header.StaticCells[cellIndex].CellTemplate.ColIndex));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception)
		{
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			_ToKhaiInfo = ToKhaiInfo;
			_TTinChung = TTinChung;
			_PeriodInfo = PeriodInfo;
			int[] array = GeneralFunctions.ArrayCellId("F_29");
			int[] array2 = GeneralFunctions.ArrayCellId("E_29");
			int[] array3 = GeneralFunctions.ArrayCellId("C_29");
			int[] array4 = GeneralFunctions.ArrayCellId("D_29");
			SetAllPhuLuc(winFpSpread);
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sheetView.Cells[array[1], array[0]].Text = PeriodInfo.TenQuyetToan;
				sheetView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaQuyetToan;
			}
			if (!string.IsNullOrEmpty(HeaderInfo.NgayBatDauNamTaiChinh))
			{
				sheetView.Cells[array3[1], array3[0]].Text = HeaderInfo.NgayBatDauNamTaiChinh;
				sheetView.Cells[array4[1], array4[0]].Text = GeneralFunctions.GetNgayTaiChinhDen(HeaderInfo.NgayBatDauNamTaiChinh);
			}
			LockUnLockC3bTK(sheetView, ToKhaiInfo.SheetInfos[0]);
			Dictionary<string, string> dictionary = SetLaiGiaTriCombobox(sheetView, PeriodInfo, hTKKCMCTableInfo, HTKKFile.ColumnLettersToIndex("C"), 20);
			string pathCtC7_tenThueSuat = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7_tenThueSuat";
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtC7_tenThueSuat);
			if (Convert.ToInt32(PeriodInfo.Nam) < 2025)
			{
				GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: false);
			}
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + num - 1, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				if (!string.IsNullOrEmpty(text) && !dictionary.ContainsKey(text))
				{
					sheetView.Cells[dataLocationTemplateRowStart + num - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Value = "";
					sheetView.Cells[dataLocationTemplateRowStart + num - 1, HTKKFile.ColumnLettersToIndex("P") - 1].Value = "";
				}
			}
			LockCtD2_TK(sheetView, hTKKCMCSheetInfo);
			lockUnlockCtC3a(sheetView, hTKKCMCSheetInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (!ToKhaiInfo.SheetInfos[idxPL03_2].IsReadDataFile)
			{
				SheetView sheetView = winFpSpread.Sheets[idxPL03_2];
				int num = int.Parse(PeriodInfo.Nam);
				int num2 = num - 5;
				for (int i = 0; i < 5; i++)
				{
					sheetView.Cells[26 + i, 2].Value = num2 + i;
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.DecimalPlaces = 0;
					numberCellType.ShowSeparator = false;
					sheetView.Cells[26 + i, 2].CellType = numberCellType;
					sheetView.Cells[34 + i, 2].Value = num2 + i;
					numberCellType = new NumberCellType();
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.DecimalPlaces = 0;
					numberCellType.ShowSeparator = false;
					sheetView.Cells[34 + i, 2].CellType = numberCellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3a];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[idxPL03_3a].Tables[0];
			SetPhuLucND92(winFpSpread);
			LockUnlockMucBUuDaiCungTinh_PL03_3a(sheetView, ToKhaiInfo.SheetInfos[idxPL03_3a]);
			Dictionary<string, string> dictionary = SetLaiGiaTriCombobox(sheetView, PeriodInfo, hTKKCMCTableInfo, HTKKFile.ColumnLettersToIndex("L"), 208);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("AZ") - 1].Text;
				if (!string.IsNullOrEmpty(text) && !dictionary.ContainsKey(text))
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Value = "";
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("AZ") - 1].Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3b];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[idxPL03_3b].Tables[0];
			SetPhuLucND92(winFpSpread);
			LockUnlockMucBUuDaiCungTinh_PL03_3b(sheetView, ToKhaiInfo.SheetInfos[idxPL03_3b]);
			Dictionary<string, string> dictionary = SetLaiGiaTriCombobox(sheetView, PeriodInfo, hTKKCMCTableInfo, HTKKFile.ColumnLettersToIndex("L"), 208);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("BA") - 1].Text;
				if (!string.IsNullOrEmpty(text) && !dictionary.ContainsKey(text))
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Value = "";
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("BA") - 1].Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_7(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SetPhuLucND92(winFpSpread);
			KhoaCt17Ct19PL_03_3c(winFpSpread.Sheets[idxPL03_3c], ToKhaiInfo.SheetInfos[idxPL03_3c]);
			LockUnlockMucBUuDaiCungTinh_PL03_3c(winFpSpread.Sheets[idxPL03_3c], ToKhaiInfo.SheetInfos[idxPL03_3c]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SetPhuLucND92(winFpSpread);
			LockUnlockMucBUuDaiCungTinh_PL03_3d(winFpSpread.Sheets[idxPL03_3d], ToKhaiInfo.SheetInfos[idxPL03_3d]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_9(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void PreBindData(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_12(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SetLockUnlockBang_P_03_8B(sView, ToKhaiInfo.SheetInfos[idxPL03_8], TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_13(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SetLockUnlockBang_P_03_8A(sView, ToKhaiInfo.SheetInfos[idxPL03_8a], TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_14(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SetLockUnlockBang_P_03_8B(sView, ToKhaiInfo.SheetInfos[idxPL03_8b], TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_15(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SetLockUnlockBang_P_03_8C(sView, ToKhaiInfo.SheetInfos[idxPL03_8c], TTinChung.maCQTCapCuc, TTinChung.maCQTQuanLy);
			UpdateChoTongTheoTinhDuLieuCu_Pl03_8C(sView, ToKhaiInfo.SheetInfos[idxPL03_8c]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_17(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int index = idxPL03_01GDLK_ND132;
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			List<int> list = new List<int>();
			list.Add(HTKKFile.ColumnLettersToIndex("AF"));
			if (Convert.ToInt32(PeriodInfo.Nam) >= 2024)
			{
				GeneralFunctions.LockUnLockColsInTable(sheetView, hTKKCMCSheetInfo.Tables[0], list, locked: false, deleteData: false);
			}
			else
			{
				GeneralFunctions.LockUnLockColsInTable(sheetView, hTKKCMCSheetInfo.Tables[0], list, locked: true, deleteData: true);
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value))
			{
				for (int i = 51 + count; i < sheetView.NonEmptyRowCount; i++)
				{
					if (sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x")
					{
						sheetView.Rows[i].Visible = false;
					}
				}
			}
			else
			{
				for (int i = 51 + count; i < sheetView.NonEmptyRowCount; i++)
				{
					if (sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text == "x")
					{
						sheetView.Rows[i].Visible = true;
					}
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
				for (int i = 1; i < 9; i++)
				{
					for (int num12 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
					}
				}
				bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 106 + num8; i < 123 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
					for (int i = 106 + num8; i < 132 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
				{
					for (int i = 106 + num8; i < 123 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
					for (int i = 106 + num8; i < 132 + num8; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 220 + num10; i < 234 + num10; i++)
					{
						if (i != 224 + num10 && i != 225 + num10 && i != 226 + num10 && i != 233 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int i = 175 + num10; i < 234 + num10; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
				if (!ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 220 + num10; i < 234 + num10; i++)
					{
						if (i != 224 + num10 && i != 225 + num10 && i != 226 + num10 && i != 233 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						}
					}
					for (int i = 175 + num10; i < 234 + num10; i++)
					{
						if (i == 175 + num10 || i == 176 + num10 || i == 190 + num10 || i == 191 + num10 || i == 206 + num10 || i == 211 + num10 || i == 219 + num10 || i == 225 + num10 || i == 226 + num10 || i == 227 + num10 || i == 233 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (i + 1);
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
			}
			else
			{
				sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
				sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_1/ct12a"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_GDLK2-01/Muc_III/Dong_II_2/ct12a"));
				for (int i = 1; i < 9; i++)
				{
					for (int num12 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = string.Format("J{0}-H{0}", (num12 + 1).ToString());
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = string.Format("V{0}-R{0}", (num12 + 1).ToString());
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = string.Format("ABS(M{0})+ABS(Y{0})", (num12 + 1).ToString());
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("AK") - 1].Locked = false;
					}
				}
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					}
				}
				string pathCt6 = "PL_GDLK2-01/Muc_IV/IV_1/Dong_5/ct06";
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt6);
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i == 109 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
						if (i != hTKKCMCCellInfo3.CellTemplate.RowIndex - 1)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					else if (i == 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
						if (i != hTKKCMCCellInfo3.CellTemplate.RowIndex - 1)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					else
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (i + 1).ToString());
						if (i != hTKKCMCCellInfo3.CellTemplate.RowIndex - 1)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
				}
				for (int i = 121 + num8; i < 133 + num8; i++)
				{
					if (i == 121 + num8 || i == 122 + num8 || i == 123 + num8 || i == 124 + num8 || i == 125 + num8 || i == 131 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (i + 1);
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
				}
				for (int i = 220 + num10; i < 234 + num10; i++)
				{
					if (i != 224 + num10 && i != 225 + num10 && i != 226 + num10 && i != 233 + num10)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					}
				}
				for (int i = 175 + num10; i < 234 + num10; i++)
				{
					if (i == 175 + num10 || i == 176 + num10 || i == 190 + num10 || i == 191 + num10 || i == 206 + num10 || i == 211 + num10 || i == 219 + num10 || i == 225 + num10 || i == 226 + num10 || i == 227 + num10 || i == 233 + num10)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "AN" + (i + 1);
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = 126 + num8; i < 132 + num8; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
				for (int i = 126 + num8; i < 132 + num8; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[139 + num9, GeneralFunctions.ArrayCellId("X_141")[0]].Value))
			{
				for (int i = 145 + num9; i < 160 + num9; i++)
				{
					if (i != 147 + num9 && i != 150 + num9 && i != 156 + num9)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 145 + num9; i < 160 + num9; i++)
				{
					if (i != 147 + num9 && i != 150 + num9 && i != 156 + num9)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[169 + num10, GeneralFunctions.ArrayCellId("X_171")[0]].Value))
			{
				for (int i = 220 + num10; i < 224 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = 228 + num10; i < 233 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int i = 220 + num10; i < 224 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = 228 + num10; i < 233 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 220 + num10; i < 224 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int i = 228 + num10; i < 233 + num10; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_21(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SetPhuLucND92(winFpSpread);
		LockUnlockMucBUuDaiCungTinh_PL03_3a(winFpSpread.Sheets[idxPL03_3a], _ToKhaiInfo.SheetInfos[idxPL03_3a]);
		LockUnlockMucBUuDaiCungTinh_PL03_3b(winFpSpread.Sheets[idxPL03_3b], _ToKhaiInfo.SheetInfos[idxPL03_3b]);
		LockUnlockMucBUuDaiCungTinh_PL03_3c(winFpSpread.Sheets[idxPL03_3c], _ToKhaiInfo.SheetInfos[idxPL03_3c]);
		LockUnlockMucBUuDaiCungTinh_PL03_3d(winFpSpread.Sheets[idxPL03_3d], _ToKhaiInfo.SheetInfos[idxPL03_3d]);
	}

	public void AutoBindData_22(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
		if (PeriodInfo.LoaiToKhai == "B")
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7_tenThueSuat");
			int[] array2 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7_thueSuat");
			int[] array3 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC9");
			string text = sheetView.Cells[array3[1], HTKKFile.ColumnLettersToIndex("C") - 1].Text;
			text = text.Replace("\n", " ");
			GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7_thueSuat", "Caption", sheetView.Cells[array[1], array[0]].Text);
			GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC9", "Caption", text);
		}
	}

	private string getDenNgayNC(string ngayBdNamTc, int iNam)
	{
		string result = "";
		try
		{
			if (ngayBdNamTc.Equals("01/01"))
			{
				result = "31/03/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
			else if (ngayBdNamTc.Equals("01/04"))
			{
				result = "30/06/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
			else if (ngayBdNamTc.Equals("01/07"))
			{
				result = "30/09/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
			else if (ngayBdNamTc.Equals("01/10"))
			{
				result = "30/12/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public void TongHopCtA1_TK(SheetView sView_TK)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("O_35");
			int[] array2 = GeneralFunctions.ArrayCellId("P_35");
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongHopCtB15_TK(SheetView sView_TK)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("O_51");
			int[] array2 = GeneralFunctions.ArrayCellId("P_51");
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongHopCtC15_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC15");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongHopCtC16_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC16");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongHopCtD2_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctD2");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
			LockCtD2_TK(sView_TK, sInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockCtD2_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongHopCtE6_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE6");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtE5_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE5");
			int[] array2 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE6");
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongHopCtG5_TK(SheetView sView_TK, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctG5");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sView_TK.Cells[array[1], array[0]].Value = sView_TK.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtC11_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			double num = 0.0;
			double num2 = 0.0;
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC11");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = _ToKhaiInfo.SheetInfos[idxPL03_3a];
			SheetView sheetView2 = winFpSpread.Sheets[idxPL03_3a];
			if (hTKKCMCSheetInfo != null)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong12"));
					if (!Convert.ToBoolean(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
					{
						num += Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					}
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = _ToKhaiInfo.SheetInfos[idxPL03_3b];
			SheetView sheetView3 = winFpSpread.Sheets[idxPL03_3b];
			if (hTKKCMCSheetInfo2 != null)
			{
				foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo2.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo3 = table2.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong12"));
					num2 += Convert.ToDouble(sheetView3.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				}
			}
			sheetView.Cells[array[1], array[0]].Value = num + num2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtC12_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			double num = 0.0;
			double num2 = 0.0;
			double num3 = 0.0;
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC12");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = _ToKhaiInfo.SheetInfos[idxPL03_3a];
			SheetView sheetView2 = winFpSpread.Sheets[idxPL03_3a];
			if (hTKKCMCSheetInfo != null)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong13"));
					if (!Convert.ToBoolean(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
					{
						num += Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					}
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = _ToKhaiInfo.SheetInfos[idxPL03_3b];
			SheetView sheetView3 = winFpSpread.Sheets[idxPL03_3b];
			if (hTKKCMCSheetInfo2 != null)
			{
				foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo2.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo3 = table2.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong13"));
					num2 += Convert.ToDouble(sheetView3.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo3 = _ToKhaiInfo.SheetInfos[idxPL03_3d];
			SheetView sheetView4 = winFpSpread.Sheets[idxPL03_3d];
			if (hTKKCMCSheetInfo3 != null)
			{
				foreach (HTKKCMCTableInfo table3 in hTKKCMCSheetInfo3.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo4 = table3.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong20"));
					num3 += Convert.ToDouble(sheetView4.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				}
			}
			sheetView.Cells[array[1], array[0]].Value = num + num2 + num3;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtC13_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			double num = 0.0;
			double num2 = 0.0;
			double num3 = 0.0;
			double num4 = 0.0;
			double num5 = 0.0;
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC13");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = _ToKhaiInfo.SheetInfos[idxPL03_3a];
			SheetView sheetView2 = winFpSpread.Sheets[idxPL03_3a];
			if (hTKKCMCSheetInfo != null)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong14"));
					if (!Convert.ToBoolean(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
					{
						num += Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					}
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = _ToKhaiInfo.SheetInfos[idxPL03_3b];
			SheetView sheetView3 = winFpSpread.Sheets[idxPL03_3b];
			if (hTKKCMCSheetInfo2 != null)
			{
				foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo2.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo3 = table2.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong14"));
					num2 += Convert.ToDouble(sheetView3.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo3 = _ToKhaiInfo.SheetInfos[idxPL03_3c];
			SheetView sheetView4 = winFpSpread.Sheets[idxPL03_3c];
			if (hTKKCMCSheetInfo3 != null)
			{
				foreach (HTKKCMCTableInfo table3 in hTKKCMCSheetInfo3.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo4 = table3.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct16"));
					num3 += Convert.ToDouble(sheetView4.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo4 = _ToKhaiInfo.SheetInfos[idxPL03_3d];
			SheetView sheetView5 = winFpSpread.Sheets[idxPL03_3d];
			if (hTKKCMCSheetInfo4 != null)
			{
				foreach (HTKKCMCTableInfo table4 in hTKKCMCSheetInfo4.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo5 = table4.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct14"));
					num5 += Convert.ToDouble(sheetView5.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value);
					HTKKCMCCellInfo hTKKCMCCellInfo6 = table4.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong21"));
					num4 += Convert.ToDouble(sheetView5.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value);
				}
			}
			sheetView.Cells[array[1], array[0]].Value = num + num2 + num5 + num4;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtD7_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			double num = 0.0;
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctD7");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = _ToKhaiInfo.SheetInfos[idxPL03_3a];
			SheetView sheetView2 = winFpSpread.Sheets[idxPL03_3a];
			if (hTKKCMCSheetInfo != null)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/tong12"));
					if (Convert.ToBoolean(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
					{
						num += Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					}
				}
			}
			sheetView.Cells[array[1], array[0]].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtG4_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctG4");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtG3_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctG3");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtG2_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctG2");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtG1_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[idxToKhai];
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctG1");
			int[] array2 = GeneralFunctions.ArrayCellId("P_" + (array[1] + 1));
			sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCtE1_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			double num = 0.0;
			double num2 = 0.0;
			double num3 = 0.0;
			double num4 = 0.0;
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC17");
			int[] array2 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE1");
			double num5 = Convert.ToDouble(sheetView.Cells[array[1], array[0]].Value);
			if (_ToKhaiInfo.SheetInfos[idxPL03_3a] != null)
			{
				foreach (HTKKCMCTableInfo table in _ToKhaiInfo.SheetInfos[idxPL03_3a].Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "duAn");
					HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct16");
					bool flag = Convert.ToBoolean(winFpSpread.Sheets[idxPL03_3a].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
					double num6 = Convert.ToDouble(winFpSpread.Sheets[idxPL03_3a].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					if (!flag)
					{
						num += num6;
					}
				}
			}
			if (_ToKhaiInfo.SheetInfos[idxPL03_3b] != null)
			{
				foreach (HTKKCMCTableInfo table2 in _ToKhaiInfo.SheetInfos[idxPL03_3b].Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo2 = table2.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct16");
					double num6 = Convert.ToDouble(winFpSpread.Sheets[idxPL03_3b].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					num2 += num6;
				}
			}
			if (_ToKhaiInfo.SheetInfos[idxPL03_3c] != null)
			{
				foreach (HTKKCMCTableInfo table3 in _ToKhaiInfo.SheetInfos[idxPL03_3c].Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo3 = table3.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct21");
					double num7 = Convert.ToDouble(winFpSpread.Sheets[idxPL03_3c].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
					num3 += num7;
				}
			}
			if (_ToKhaiInfo.SheetInfos[idxPL03_3d] != null)
			{
				foreach (HTKKCMCTableInfo table4 in _ToKhaiInfo.SheetInfos[idxPL03_3d].Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo4 = table4.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct23");
					double num8 = Convert.ToDouble(winFpSpread.Sheets[idxPL03_3d].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
					num4 += num8;
				}
			}
			double num9 = num5 - num - num2 - num3 - num4;
			sheetView.Cells[array2[1], array2[0]].Value = ((num9 < 0.0) ? 0.0 : num9);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockTinhCQT(bool isKhacTinh_TSC, SheetView sView, HTKKCMCCellInfo cellTenTinh, HTKKCMCCellInfo cellMaTinh, HTKKCMCCellInfo cellTenHuyen, HTKKCMCCellInfo cellMaHuyen, HTKKCMCCellInfo cellTenXa, HTKKCMCCellInfo cellMaXa, HTKKCMCCellInfo cellDiaChi, HTKKCMCCellInfo cellCQTCapCuc, HTKKCMCCellInfo cellCQTChiCuc)
	{
		try
		{
			if (isKhacTinh_TSC)
			{
				GeneralFunctions.LockUnLockCol(sView, cellCQTCapCuc.CellTemplate.RowIndex, cellCQTCapCuc.CellTemplate.ColIndex, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, cellCQTChiCuc.CellTemplate.RowIndex, cellCQTChiCuc.CellTemplate.ColIndex, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, cellTenTinh.CellTemplate.RowIndex, cellTenTinh.CellTemplate.ColIndex, locked: false, deleteData: false);
				if (cellTenXa != null)
				{
					GeneralFunctions.LockUnLockCol(sView, cellTenXa.CellTemplate.RowIndex, cellTenXa.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
				if (cellDiaChi != null)
				{
					GeneralFunctions.LockUnLockCol(sView, cellDiaChi.CellTemplate.RowIndex, cellDiaChi.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
			}
			else
			{
				GeneralFunctions.LockUnLockCol(sView, cellCQTCapCuc.CellTemplate.RowIndex, cellCQTCapCuc.CellTemplate.ColIndex, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sView, cellCQTChiCuc.CellTemplate.RowIndex, cellCQTChiCuc.CellTemplate.ColIndex, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sView, cellTenTinh.CellTemplate.RowIndex, cellTenTinh.CellTemplate.ColIndex, locked: true, deleteData: true);
				if (cellTenXa != null)
				{
					GeneralFunctions.LockUnLockCol(sView, cellTenXa.CellTemplate.RowIndex, cellTenXa.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
				if (cellDiaChi != null)
				{
					GeneralFunctions.LockUnLockCol(sView, cellDiaChi.CellTemplate.RowIndex, cellDiaChi.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiTinhTheoCQT(string maCQTCapCuc, string maCQTQuanLy, SheetView sView, HTKKCMCCellInfo cellTenTinh, HTKKCMCCellInfo cellMaTinh, HTKKCMCCellInfo cellTenHuyen, HTKKCMCCellInfo cellMaHuyen, HTKKCMCCellInfo cellTenXa, HTKKCMCCellInfo cellMaXa, HTKKCMCCellInfo cellDiaChi)
	{
		try
		{
			if (string.IsNullOrEmpty(maCQTCapCuc))
			{
				return;
			}
			if (maCQTCapCuc == "82500")
			{
				sView.Cells[cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1].Locked = false;
				return;
			}
			sView.Cells[cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1].Value = "";
			sView.Cells[cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1].Locked = false;
			sView.Cells[cellMaTinh.CellTemplate.RowIndex - 1, cellMaTinh.CellTemplate.ColIndex - 1].Value = "";
			if (cellTenXa != null)
			{
				sView.Cells[cellTenXa.CellTemplate.RowIndex - 1, cellTenXa.CellTemplate.ColIndex - 1].Value = "";
				sView.Cells[cellMaXa.CellTemplate.RowIndex - 1, cellMaXa.CellTemplate.ColIndex - 1].Value = "";
				sView.Cells[cellTenXa.CellTemplate.RowIndex - 1, cellTenXa.CellTemplate.ColIndex - 1].Locked = false;
			}
			string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(maCQTCapCuc.Substring(0, 3)) ? maCQTCapCuc.Substring(0, 3) : "");
			if (string.IsNullOrEmpty(maCQTCapCuc) || string.IsNullOrEmpty(maTinh))
			{
				return;
			}
			KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
			sView.Cells[cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1].Value = keyValuePair.Key;
			sView.Cells[cellMaTinh.CellTemplate.RowIndex - 1, cellMaTinh.CellTemplate.ColIndex - 1].Value = maTinh;
			sView.Cells[cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1].Locked = true;
			if (cellTenXa != null)
			{
				Dictionary<string, string> dic = HTKKWinDictionaryDiaDanh._diaBanXa.Where((DiaBanXa x) => x.MaTinh == maTinh).ToDictionary((DiaBanXa x) => x.MaXa, (DiaBanXa x) => x.TenXa);
				if (string.IsNullOrEmpty(sView.Cells[cellTenXa.CellTemplate.RowIndex - 1, cellTenXa.CellTemplate.ColIndex - 1].Text.Trim()))
				{
					sView.Cells[cellTenXa.CellTemplate.RowIndex - 1, cellTenXa.CellTemplate.ColIndex - 1].Value = "";
					sView.Cells[cellMaXa.CellTemplate.RowIndex - 1, cellMaXa.CellTemplate.ColIndex - 1].Value = "";
				}
				GeneralFunctions.InsertCombobox(sView, cellTenXa.CellTemplate.RowIndex, cellTenXa.CellTemplate.ColIndex, dic);
			}
			if (cellDiaChi != null && string.IsNullOrEmpty(sView.Cells[cellDiaChi.CellTemplate.RowIndex - 1, cellDiaChi.CellTemplate.ColIndex - 1].Text.Trim()))
			{
				sView.Cells[cellDiaChi.CellTemplate.RowIndex - 1, cellDiaChi.CellTemplate.ColIndex - 1].Value = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckBatNhapCQT(SheetView sView, HTKKCMCCellInfo cellCQTCapCuc, HTKKCMCCellInfo cellCQTCapCucMa, HTKKCMCCellInfo cellCQTChiCuc, HTKKCMCCellInfo cellCQTChiCucMa)
	{
		bool result = false;
		try
		{
			string text = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string value = sView.Cells[cellCQTCapCuc.CellTemplate.RowIndex - 1, cellCQTCapCuc.CellTemplate.ColIndex - 1].Text.Trim();
			string value2 = sView.Cells[cellCQTCapCucMa.CellTemplate.RowIndex - 1, cellCQTCapCucMa.CellTemplate.ColIndex - 1].Text.Trim();
			string value3 = sView.Cells[cellCQTChiCuc.CellTemplate.RowIndex - 1, cellCQTChiCuc.CellTemplate.ColIndex - 1].Text.Trim();
			string value4 = sView.Cells[cellCQTChiCucMa.CellTemplate.RowIndex - 1, cellCQTChiCucMa.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellCQTCapCuc.CellTemplate.RowIndex - 1, cellCQTCapCuc.CellTemplate.ColIndex - 1, text);
				result = true;
			}
			if (string.IsNullOrEmpty(value3) || string.IsNullOrEmpty(value4))
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellCQTChiCuc.CellTemplate.RowIndex - 1, cellCQTChiCuc.CellTemplate.ColIndex - 1, text);
				result = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return result;
		}
		return result;
	}

	public bool CheckCQTKhac_TSC(SheetView sView, string maCQTCapCuc_TSC, HTKKCMCCellInfo cellCQTCapCuc, HTKKCMCCellInfo cellCQTCapCucMa)
	{
		try
		{
			string text = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[172, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sView.Cells[cellCQTCapCucMa.CellTemplate.RowIndex - 1, cellCQTCapCucMa.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text2) && text2 == maCQTCapCuc_TSC)
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellCQTCapCuc.CellTemplate.RowIndex - 1, cellCQTCapCuc.CellTemplate.ColIndex - 1, text);
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public bool CheckBatNhapTinh(SheetView sView, HTKKCMCCellInfo cellTenTinh, HTKKCMCCellInfo cellMaTinh, HTKKCMCCellInfo cellTenHuyen, HTKKCMCCellInfo cellMaHuyen, HTKKCMCCellInfo cellTenXa, HTKKCMCCellInfo cellMaXa)
	{
		bool result = false;
		try
		{
			string text = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string value = sView.Cells[cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1].Text.Trim();
			string value2 = sView.Cells[cellMaTinh.CellTemplate.RowIndex - 1, cellMaTinh.CellTemplate.ColIndex - 1].Text.Trim();
			string text2 = sView.Cells[cellTenHuyen.CellTemplate.RowIndex - 1, cellTenHuyen.CellTemplate.ColIndex - 1].Text.Trim();
			string text3 = sView.Cells[cellMaHuyen.CellTemplate.RowIndex - 1, cellMaHuyen.CellTemplate.ColIndex - 1].Text.Trim();
			string value3 = sView.Cells[cellTenXa.CellTemplate.RowIndex - 1, cellTenXa.CellTemplate.ColIndex - 1].Text.Trim();
			string value4 = sView.Cells[cellMaXa.CellTemplate.RowIndex - 1, cellMaXa.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellTenTinh.CellTemplate.RowIndex - 1, cellTenTinh.CellTemplate.ColIndex - 1, text);
				result = true;
			}
			if (string.IsNullOrEmpty(value3) || string.IsNullOrEmpty(value4))
			{
				HTKKFile.SetNodeBackGroundPink(sView, cellTenXa.CellTemplate.RowIndex - 1, cellTenXa.CellTemplate.ColIndex - 1, text);
				result = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return result;
		}
		return result;
	}

	public bool CheckKhongDuocLonHonNamHienTai(SheetView sView, int row, int colNam)
	{
		try
		{
			string text = sView.Cells[row - 1, colNam - 1].Text.Trim();
			HTKKFile.RemoveValidate(sView, row - 1, colNam - 1);
			if (string.IsNullOrEmpty(text))
			{
				return false;
			}
			string text2 = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[152, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[151, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			if (!GeneralFunctions.IsNumeric(text))
			{
				HTKKFile.SetNodeBackGroundPink(sView, row - 1, colNam - 1, text2);
				return true;
			}
			if (Convert.ToInt32(text) > DateTime.Now.Year)
			{
				HTKKFile.SetNodeBackGroundPink(sView, row - 1, colNam - 1, text3);
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return false;
	}

	public void ThayDoiTinhTheoCQTDangBang(string maCQTCapCuc, string maCQTQuanLy, SheetView sView, HTKKCMCSheetInfo sInfo, int colTenTinh, int colMaTinh, int colTenHuyen, int colMaHuyen, int iRow, bool isLock)
	{
		try
		{
			if (string.IsNullOrEmpty(maCQTCapCuc))
			{
				sView.Cells[iRow - 1, colTenTinh - 1].Text = "";
				sView.Cells[iRow - 1, colMaTinh - 1].Text = "";
				GeneralFunctions.LockUnLockCol(sView, iRow, colTenTinh, locked: false, deleteData: false);
				sView.Cells[iRow - 1, colTenHuyen - 1].Text = "";
				sView.Cells[iRow - 1, colMaHuyen - 1].Text = "";
				return;
			}
			if (maCQTCapCuc == "82500")
			{
				GeneralFunctions.LockUnLockCol(sView, iRow, colTenTinh, locked: false, deleteData: false);
				return;
			}
			string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(maCQTCapCuc.Substring(0, 3)) ? maCQTCapCuc.Substring(0, 3) : "");
			KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
			sView.Cells[iRow - 1, colTenTinh - 1].Text = keyValuePair.Key;
			sView.Cells[iRow - 1, colMaTinh - 1].Text = maTinh;
			sView.Cells[iRow - 1, colTenHuyen - 1].Text = "";
			sView.Cells[iRow - 1, colMaHuyen - 1].Text = "";
			GeneralFunctions.LockUnLockCol(sView, iRow, colTenTinh, isLock, deleteData: false);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiCQTTheoTinhDynamic2(SheetView sView, string maTinh, string tenCQTCapCuc, string maCQTCapCuc, string tenCQTChiCuc, string maCQTChiCuc, int iRow, int colTenCQTCuc, int colMaCQTCuc, int colTenCQTChiCuc, int colMaCQTChiCuc, bool isLock)
	{
		try
		{
			if (string.IsNullOrEmpty(maTinh) && string.IsNullOrEmpty(maCQTCapCuc))
			{
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Value = "";
				sView.Cells[iRow - 1, colMaCQTCuc - 1].Value = "";
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Locked = false;
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Value = "";
				sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Value = "";
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Locked = false;
				return;
			}
			sView.Cells[iRow - 1, colTenCQTCuc - 1].Value = "";
			sView.Cells[iRow - 1, colMaCQTCuc - 1].Value = "";
			sView.Cells[iRow - 1, colTenCQTCuc - 1].Locked = false;
			sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Value = "";
			sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Value = "";
			sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Locked = false;
			tenCQTCapCuc = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(maCQTCapCuc, isGetFullName: false);
			if (string.IsNullOrEmpty(maCQTCapCuc) && !string.IsNullOrEmpty(maTinh) && HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(maTinh + "00"))
			{
				maCQTCapCuc = maTinh + "00";
				Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(maCQTCapCuc);
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCuc);
				maCQTChiCuc = maCQTCapCuc;
				tenCQTChiCuc = coQuanThueQuanLybyMaCoQuanThueCapCuc[maCQTChiCuc];
				int num = keyValuePair.Key.IndexOf(" - ");
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Text = tenCQTCapCuc;
				sView.Cells[iRow - 1, colMaCQTCuc - 1].Text = maCQTCapCuc;
				GeneralFunctions.InsertCombobox(sView, iRow, colTenCQTChiCuc, coQuanThueQuanLybyMaCoQuanThueCapCuc);
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Text = tenCQTChiCuc;
				sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Text = maCQTChiCuc;
				GeneralFunctions.LockUnLockCol(sView, iRow, colTenCQTCuc, isLock, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, iRow, colTenCQTChiCuc, isLock, deleteData: false);
			}
			if (!string.IsNullOrEmpty(maCQTCapCuc) && string.IsNullOrEmpty(maTinh))
			{
				Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(maCQTCapCuc);
				sView.Cells[iRow - 1, colTenCQTCuc - 1].Text = tenCQTCapCuc;
				sView.Cells[iRow - 1, colMaCQTCuc - 1].Text = maCQTCapCuc;
				GeneralFunctions.InsertCombobox(sView, iRow, colTenCQTChiCuc, coQuanThueQuanLybyMaCoQuanThueCapCuc);
				sView.Cells[iRow - 1, colTenCQTChiCuc - 1].Text = tenCQTChiCuc;
				sView.Cells[iRow - 1, colMaCQTChiCuc - 1].Text = maCQTChiCuc;
				GeneralFunctions.LockUnLockCol(sView, iRow, colTenCQTCuc, isLock, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, iRow, colTenCQTChiCuc, isLock, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetLockUnlockBang_P_03_8A(SheetView sView, HTKKCMCSheetInfo sInfo, string maCQTCapCuc, string maCQTChiCuc)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AU");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int colHide = HTKKFile.ColumnLettersToIndex("R");
			int num3 = HTKKFile.ColumnLettersToIndex("V");
			int colHide2 = HTKKFile.ColumnLettersToIndex("AA");
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				if (maCQTCapCuc == "82500")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
					continue;
				}
				if (!string.IsNullOrEmpty(maCQTCapCuc) && maCQTCapCuc == maCQTChiCuc)
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
					continue;
				}
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (text == "01")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: true, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide2, locked: true, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide2, locked: false, deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetLockUnlockBang_P_03_8(SheetView sView, HTKKCMCSheetInfo sInfo, string maCQTCapCuc, string maCQTChiCuc)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AR");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int colHide = HTKKFile.ColumnLettersToIndex("R");
			int num3 = HTKKFile.ColumnLettersToIndex("V");
			int colHide2 = HTKKFile.ColumnLettersToIndex("AA");
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (maCQTCapCuc == "82500")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
				else if (!string.IsNullOrEmpty(maCQTCapCuc) && maCQTCapCuc == maCQTChiCuc)
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
				else if (text == "01")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: true, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide2, locked: true, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide2, locked: false, deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetLockUnlockBang_P_03_8C(SheetView sView, HTKKCMCSheetInfo sInfo, string maCQTCapCuc, string maCQTChiCuc)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AP");
			int num2 = HTKKFile.ColumnLettersToIndex("P");
			int num3 = HTKKFile.ColumnLettersToIndex("AJ");
			int colHide = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("AI");
			int num5 = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("AK");
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int num8 = HTKKFile.ColumnLettersToIndex("AL");
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (maCQTCapCuc == "82500")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
				else if (!string.IsNullOrEmpty(maCQTCapCuc) && maCQTCapCuc == maCQTChiCuc)
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetLockUnlockBang_P_03_8B(SheetView sView, HTKKCMCSheetInfo sInfo, string maCQTCapCuc, string maCQTChiCuc)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AU");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int colHide = HTKKFile.ColumnLettersToIndex("R");
			int num3 = HTKKFile.ColumnLettersToIndex("AA");
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (maCQTCapCuc == "82500")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
				else if (!string.IsNullOrEmpty(maCQTCapCuc) && maCQTCapCuc == maCQTChiCuc)
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colHide, locked: false, deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3a];
			SheetView sheetView = winFpSpread.Sheets[idxPL03_3a];
			UpdateCacHangSauLuonGiongHangTruoc_PL03_3a(winFpSpread, ToKhaiInfo);
			TinhCt16PL03_3a(winFpSpread.Sheets[idxPL03_3a], _ToKhaiInfo.SheetInfos[idxPL03_3a]);
			((Control)winFpSpread).Refresh();
			TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			((Control)winFpSpread).Refresh();
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			((Control)winFpSpread).Refresh();
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			((Control)winFpSpread).Refresh();
			TinhCtD7_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			((Control)winFpSpread).Refresh();
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateCacHangSauLuonGiongHangTruoc_PL03_3a(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[idxPL03_3a];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3a];
			int item = HTKKFile.ColumnLettersToIndex("I");
			int item2 = HTKKFile.ColumnLettersToIndex("AX");
			int item3 = HTKKFile.ColumnLettersToIndex("K");
			int item4 = HTKKFile.ColumnLettersToIndex("AY");
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				GeneralFunctions.UpdateCacHangSauGiongHangTruocDo(sView, table, list, chiGiongHangLock: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			UpdateCacHangSauLuonGiongHangTruoc_PL03_3b(winFpSpread, ToKhaiInfo);
			TinhCt16PL03_3b(winFpSpread.Sheets[idxPL03_3b], _ToKhaiInfo.SheetInfos[idxPL03_3b]);
			TinhCtC11_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_7(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TinhCt21PL03_3c(winFpSpread.Sheets[idxPL03_3c], _ToKhaiInfo.SheetInfos[idxPL03_3c]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateCacHangSauLuonGiongHangTruoc_PL03_3b(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[idxPL03_3b];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3b];
			int item = HTKKFile.ColumnLettersToIndex("I");
			int item2 = HTKKFile.ColumnLettersToIndex("AY");
			int item3 = HTKKFile.ColumnLettersToIndex("K");
			int item4 = HTKKFile.ColumnLettersToIndex("AZ");
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				GeneralFunctions.UpdateCacHangSauGiongHangTruocDo(sView, table, list, chiGiongHangLock: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			UpdateCacHangSauLuonGiongHangTruoc_PL03_3d(winFpSpread, ToKhaiInfo);
			TinhCt16PL03_3d(winFpSpread.Sheets[idxPL03_3d], _ToKhaiInfo.SheetInfos[idxPL03_3d]);
			TinhCtC12_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtC13_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
			TinhCtE1_TK(winFpSpread, _ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateCacHangSauLuonGiongHangTruoc_PL03_3d(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[idxPL03_3d];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_3d];
			int item = HTKKFile.ColumnLettersToIndex("I");
			int item2 = HTKKFile.ColumnLettersToIndex("AW");
			int item3 = HTKKFile.ColumnLettersToIndex("K");
			int item4 = HTKKFile.ColumnLettersToIndex("AX");
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				GeneralFunctions.UpdateCacHangSauGiongHangTruocDo(sView, table, list, chiGiongHangLock: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_14(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[idxPL03_8b], ToKhaiInfo.SheetInfos[idxPL03_8b]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_15(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TinhTongTheoTinh_Pl03_8C(winFpSpread, ToKhaiInfo, -1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockMucBUuDaiCungTinh_PL03_3a(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			int[] array = GeneralFunctions.ArrayCellId("AW_3");
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo fromRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo toRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("duAn"));
				HTKKCMCCellInfo fromRow2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09c_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo toRow2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct16"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct17"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct18"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct19"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct20"));
				string text = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (text != UU_DAI_KHAC_TINH)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex, hTKKCMCCellInfo8.CellTemplate.ColIndex, locked: true, deleteData: true);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: true, isDelete: true);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow2, toRow2, null, isLock: true, isDelete: true);
				}
				else
				{
					if (Convert.ToInt32(sView.Cells[array[1], array[0]].Value) == 1)
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: true, deleteData: false);
					}
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.ColIndex, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex, hTKKCMCCellInfo8.CellTemplate.ColIndex, locked: false, deleteData: false);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: false, isDelete: false);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow2, toRow2, null, isLock: false, isDelete: false);
				}
				GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockMucBUuDaiCungTinh_PL03_3b(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			int[] array = GeneralFunctions.ArrayCellId("AW_3");
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo fromRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo toRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct16"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct17"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct18"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct19"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct20"));
				string text = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (text != UU_DAI_KHAC_TINH)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.ColIndex, locked: true, deleteData: true);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: true, isDelete: true);
				}
				else
				{
					if (Convert.ToInt32(sView.Cells[array[1], array[0]].Value) == 1)
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: true, deleteData: false);
					}
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.ColIndex, locked: false, deleteData: false);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: false, isDelete: false);
				}
				GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockMucBUuDaiCungTinh_PL03_3d(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			int[] array = GeneralFunctions.ArrayCellId("AW_3");
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo fromRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo toRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct23"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct24"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct25"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct26"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_phaiNop/ct27"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_chuyenGiao"));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_khoaHocCN"));
				HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("so_gcn"));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ngayCap"));
				HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("noiCap"));
				HTKKCMCCellInfo fromRow2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10"));
				HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct11"));
				HTKKCMCCellInfo toRow2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct12"));
				HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct14"));
				HTKKCMCCellInfo toRow3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct18"));
				bool flag = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value);
				bool flag2 = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value);
				if (flag)
				{
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow2, toRow2, null, isLock: false, isDelete: false);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo16.CellTemplate.RowIndex, hTKKCMCCellInfo16.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
				else
				{
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow2, toRow2, null, isLock: true, isDelete: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo16.CellTemplate.RowIndex, hTKKCMCCellInfo16.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
				if (flag2)
				{
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow2, toRow3, hTKKCMCCellInfo11, isLock: false, isDelete: false);
				}
				else
				{
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow2, toRow3, hTKKCMCCellInfo11, isLock: true, isDelete: true);
				}
				string text = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (text != UU_DAI_KHAC_TINH)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.ColIndex, locked: true, deleteData: true);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: true, isDelete: true);
				}
				else
				{
					if (Convert.ToInt32(sView.Cells[array[1], array[0]].Value) == 1)
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: true, deleteData: false);
					}
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.ColIndex, locked: false, deleteData: false);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: false, isDelete: false);
				}
				GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int CheckTrong1KhoangGiaTriBatNhap(SheetView sView, HTKKCMCTableInfo table, int fromRow, int toRow)
	{
		int num = 0;
		try
		{
			string text = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			IEnumerable<HTKKCMCCellInfo> enumerable = table.Header.StaticCells.Where((HTKKCMCCellInfo cell) => cell.CellTemplate.RowIndex >= fromRow && cell.CellTemplate.RowIndex <= toRow);
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			foreach (HTKKCMCCellInfo item in enumerable)
			{
				try
				{
					bool locked = sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].Locked;
					ICellType cellType = sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].CellType;
					object value = sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].Value;
					if (!locked && ((object)cellType.GetType() == numberCellType.GetType() || (object)cellType.GetType() == currencyCellType.GetType()) && (value == null || (value != null && Convert.ToDouble(value) == 0.0)))
					{
						HTKKFile.SetNodeBackGroundPink(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1, text);
						num++;
					}
					if (!locked && (object)cellType.GetType() != numberCellType.GetType() && (object)cellType.GetType() != currencyCellType.GetType() && (value == null || (value != null && value.ToString().Trim() == "")))
					{
						HTKKFile.SetNodeBackGroundPink(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1, text);
						num++;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return num;
		}
	}

	public void Reset1KhoangLienTiep(SheetView sView, HTKKCMCTableInfo table, int fromRow, int toRow)
	{
		try
		{
			IEnumerable<HTKKCMCCellInfo> enumerable = table.Header.StaticCells.Where((HTKKCMCCellInfo cell) => cell.CellTemplate.RowIndex >= fromRow && cell.CellTemplate.RowIndex <= toRow);
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			foreach (HTKKCMCCellInfo item in enumerable)
			{
				HTKKFile.RemoveValidate(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockTrong1Khoang(SheetView sView, HTKKCMCTableInfo table, HTKKCMCCellInfo fromRow, HTKKCMCCellInfo toRow, HTKKCMCCellInfo loaiCell, bool isLock, bool isDelete)
	{
		try
		{
			List<HTKKCMCCellInfo> list = null;
			list = ((loaiCell != null) ? table.Header.StaticCells.Where((HTKKCMCCellInfo cell) => cell.CellTemplate.RowIndex >= fromRow.CellTemplate.RowIndex && cell.CellTemplate.RowIndex <= toRow.CellTemplate.RowIndex && cell.CellTemplate.RowIndex != loaiCell.CellTemplate.RowIndex && cell.CellTemplate.RowIndex != loaiCell.CellTemplate.RowIndex).ToList() : table.Header.StaticCells.Where((HTKKCMCCellInfo cell) => cell.CellTemplate.RowIndex >= fromRow.CellTemplate.RowIndex && cell.CellTemplate.RowIndex <= toRow.CellTemplate.RowIndex).ToList());
			foreach (HTKKCMCCellInfo item in list)
			{
				GeneralFunctions.LockUnLockCol(sView, item.CellTemplate.RowIndex, item.CellTemplate.ColIndex, isLock, isDelete);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void KhoaCt17Ct19PL_03_3c(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct17"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("thueTNDN_uuDai/ct19"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_sdld_nu"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("dn_sdld_dtoc"));
				bool flag = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				bool flag2 = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				if (flag)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
				if (flag2)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockMucBUuDaiCungTinh_PL03_3c(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			int[] array = GeneralFunctions.ArrayCellId("S_3");
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
				HTKKCMCCellInfo fromRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09b_ten"));
				HTKKCMCCellInfo toRow = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct09a_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct21");
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct22");
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct23");
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct24");
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct25");
				string text = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (text != UU_DAI_KHAC_TINH)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: true, deleteData: true);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: true, isDelete: true);
				}
				else
				{
					if (Convert.ToInt32(sView.Cells[array[1], array[0]].Value) == 1)
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: false, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: true, deleteData: false);
					}
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.ColIndex, locked: false, deleteData: false);
					LockUnlockTrong1Khoang(sView, hTKKCMCTableInfo, fromRow, toRow, null, isLock: false, isDelete: false);
				}
				GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnLockC3bTK(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("O_51");
			int[] array2 = GeneralFunctions.ArrayCellId("O_53");
			int[] array3 = GeneralFunctions.ArrayCellId("O_54");
			int[] array4 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3a");
			int[] array5 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3b");
			double num = Convert.ToDouble(sView.Cells[array[1], array[0]].Value);
			double num2 = Convert.ToDouble(sView.Cells[array2[1], array2[0]].Value);
			double num3 = Convert.ToDouble(sView.Cells[array3[1], array3[0]].Value);
			double num4 = Convert.ToDouble(sView.Cells[array4[1], array4[0]].Value);
			if (num2 - num3 - num4 > 0.0)
			{
				GeneralFunctions.LockUnLockCol(sView, array5[1] + 1, array5[0] + 1, locked: false, deleteData: false);
			}
			else
			{
				GeneralFunctions.LockUnLockCol(sView, array5[1] + 1, array5[0] + 1, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetPhuLucND92(FpSpread winFpSpread)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("AW_3");
			int[] array2 = GeneralFunctions.ArrayCellId("AW_3");
			int[] array3 = GeneralFunctions.ArrayCellId("S_3");
			int[] array4 = GeneralFunctions.ArrayCellId("AW_3");
			int num = ((_ToKhaiInfo.SheetInfos[idxPL03_ND92] != null) ? 1 : 0);
			if (winFpSpread.Sheets[idxPL03_3a].Visible)
			{
				winFpSpread.Sheets[idxPL03_3a].Cells[array[1], array[0]].Value = num;
			}
			if (winFpSpread.Sheets[idxPL03_3b].Visible)
			{
				winFpSpread.Sheets[idxPL03_3b].Cells[array2[1], array2[0]].Value = num;
			}
			if (winFpSpread.Sheets[idxPL03_3c].Visible)
			{
				winFpSpread.Sheets[idxPL03_3c].Cells[array3[1], array3[0]].Value = num;
			}
			if (winFpSpread.Sheets[idxPL03_3d].Visible)
			{
				winFpSpread.Sheets[idxPL03_3d].Cells[array4[1], array4[0]].Value = num;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetAllPhuLuc(FpSpread winFpSpread)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("R_16");
			int[] array2 = GeneralFunctions.ArrayCellId("S_16");
			int[] array3 = GeneralFunctions.ArrayCellId("T_16");
			int[] array4 = GeneralFunctions.ArrayCellId("U_16");
			int[] array5 = GeneralFunctions.ArrayCellId("V_16");
			int[] array6 = GeneralFunctions.ArrayCellId("W_16");
			int[] array7 = GeneralFunctions.ArrayCellId("X_16");
			int[] array8 = GeneralFunctions.ArrayCellId("Y_16");
			int[] array9 = GeneralFunctions.ArrayCellId("Z_16");
			int[] array10 = GeneralFunctions.ArrayCellId("AA_16");
			int[] array11 = GeneralFunctions.ArrayCellId("AB_16");
			int[] array12 = GeneralFunctions.ArrayCellId("AC_16");
			int[] array13 = GeneralFunctions.ArrayCellId("AD_16");
			int[] array14 = GeneralFunctions.ArrayCellId("AE_16");
			int[] array15 = GeneralFunctions.ArrayCellId("AF_16");
			int[] array16 = GeneralFunctions.ArrayCellId("AG_16");
			int[] array17 = GeneralFunctions.ArrayCellId("AH_16");
			sheetView.Cells[array[1], array[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_1a] != null) ? 1 : 0);
			sheetView.Cells[array2[1], array2[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_1b] != null) ? 1 : 0);
			sheetView.Cells[array3[1], array3[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_1c] != null) ? 1 : 0);
			sheetView.Cells[array4[1], array4[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_2] != null) ? 1 : 0);
			sheetView.Cells[array5[1], array5[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_3a] != null) ? 1 : 0);
			sheetView.Cells[array6[1], array6[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_3b] != null) ? 1 : 0);
			sheetView.Cells[array7[1], array7[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_3c] != null) ? 1 : 0);
			sheetView.Cells[array8[1], array8[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_3d] != null) ? 1 : 0);
			sheetView.Cells[array9[1], array9[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_4] != null) ? 1 : 0);
			sheetView.Cells[array10[1], array10[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_5] != null) ? 1 : 0);
			sheetView.Cells[array11[1], array11[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_6] != null) ? 1 : 0);
			sheetView.Cells[array12[1], array12[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_ND92] != null) ? 1 : 0);
			sheetView.Cells[array13[1], array13[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_9] != null) ? 1 : 0);
			sheetView.Cells[array14[1], array14[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_8a] != null) ? 1 : 0);
			sheetView.Cells[array15[1], array15[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_8] != null) ? 1 : 0);
			sheetView.Cells[array16[1], array16[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_8b] != null) ? 1 : 0);
			sheetView.Cells[array17[1], array17[0]].Value = ((_ToKhaiInfo.SheetInfos[idxPL03_8c] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetTieuMucChoBang(SheetView sView, int row, int colMaTieuMuc, HTKKCMCTableInfo table, int colTenTieuMuc)
	{
		try
		{
			string maTieuMuc = sView.Cells[row - 1, colMaTieuMuc - 1].Text.Trim();
			DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(_TTinChung, _PeriodInfo);
			string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(maTieuMuc, ngayDauKyKeKhai);
			GeneralFunctions.SetAllRowSame(sView, table, colTenTieuMuc, tenTmByMaTm);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCt16PL03_3a(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = sInfo.Tables;
			foreach (HTKKCMCTableInfo item in tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "uuDai_ma");
				HTKKCMCCellInfo hTKKCMCCellInfo2 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct16");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/tong11");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/tong15");
				string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				double num = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				double num2 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				if (text == "01")
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else if (num - num2 < 0.0)
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num - num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCt16PL03_3b(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = sInfo.Tables;
			foreach (HTKKCMCTableInfo item in tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "uuDai_ma");
				HTKKCMCCellInfo hTKKCMCCellInfo2 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct16");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/tong11");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/tong15");
				string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				double num = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				double num2 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				if (text == "01")
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else if (num - num2 < 0.0)
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num - num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCt21PL03_3c(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = sInfo.Tables;
			foreach (HTKKCMCTableInfo item in tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "uuDai_ma");
				HTKKCMCCellInfo hTKKCMCCellInfo2 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct21");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/ct15");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/ct16");
				string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				double num = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				double num2 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				if (text == "01")
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else if (num - num2 < 0.0)
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num - num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCt16PL03_3d(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = sInfo.Tables;
			foreach (HTKKCMCTableInfo item in tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "uuDai_ma");
				HTKKCMCCellInfo hTKKCMCCellInfo2 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct23");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/tong19");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = item.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_uuDai/tong22");
				string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				double num = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
				double num2 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				if (text == "01")
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				}
				else
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num - num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckE2_TK(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			double num = 0.0;
			double num2 = 0.0;
			string pathCtD8 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctD8";
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtD8);
			string pathCtE2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE2";
			HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtE2);
			int[] array = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctD8");
			int[] array2 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctE2");
			int[] array3 = GeneralFunctions.ArrayCellId("E_29");
			double num3 = Convert.ToDouble(sheetView.Cells[array2[1], array2[0]].Value);
			double num4 = Convert.ToDouble(sheetView.Cells[array[1], array[0]].Value);
			num2 = Convert.ToDouble(winFpSpread.Sheets[idxPL03_ND92].Cells[array3[1], array3[0]].Value);
			if (_ToKhaiInfo.SheetInfos[idxPL03_3a] != null)
			{
				foreach (HTKKCMCTableInfo table in _ToKhaiInfo.SheetInfos[idxPL03_3a].Tables)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "duAn");
					HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "thueTNDN_phaiNop/ct16");
					bool flag = Convert.ToBoolean(winFpSpread.Sheets[idxPL03_3a].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
					double num5 = Convert.ToDouble(winFpSpread.Sheets[idxPL03_3a].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
					if (flag)
					{
						num += num5;
					}
				}
			}
			double num6 = num4 - num - num2;
			if (num3 != num6)
			{
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			if (ToKhaiInfo.SheetInfos[idxPL03_3a] != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("AK");
				int num2 = HTKKFile.ColumnLettersToIndex("AW");
				foreach (HTKKCMCTableInfo table in ToKhaiInfo.SheetInfos[idxPL03_3a].Tables)
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table.RowsValues.Count; i++)
					{
						string value = winFpSpread.Sheets[idxPL03_3a].Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
						object value2 = winFpSpread.Sheets[idxPL03_3a].Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value;
						if (string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(value2.ToString()))
						{
							winFpSpread.Sheets[idxPL03_3a].Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text = string.Concat(value2, " %");
						}
					}
				}
				foreach (HTKKCMCTableInfo table2 in ToKhaiInfo.SheetInfos[idxPL03_3a].Tables)
				{
					TinhLaiCQTTheoUuDai_PL03_3A(winFpSpread.Sheets[idxPL03_3a], _TTinChung, _PeriodInfo, table2);
				}
			}
			if (ToKhaiInfo.SheetInfos[idxPL03_3b] != null)
			{
				int num3 = HTKKFile.ColumnLettersToIndex("AL");
				int num4 = HTKKFile.ColumnLettersToIndex("AX");
				foreach (HTKKCMCTableInfo table3 in ToKhaiInfo.SheetInfos[idxPL03_3b].Tables)
				{
					int dataLocationTemplateRowStart = table3.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table3.RowsValues.Count; i++)
					{
						string value3 = winFpSpread.Sheets[idxPL03_3b].Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text.Trim();
						object value4 = winFpSpread.Sheets[idxPL03_3b].Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value;
						if (string.IsNullOrEmpty(value3) && !string.IsNullOrEmpty(value4.ToString()))
						{
							winFpSpread.Sheets[idxPL03_3b].Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text = string.Concat(value4, " %");
						}
					}
				}
				foreach (HTKKCMCTableInfo table4 in ToKhaiInfo.SheetInfos[idxPL03_3b].Tables)
				{
					TinhLaiCQTTheoUuDai_PL03_3B(winFpSpread.Sheets[idxPL03_3b], _TTinChung, _PeriodInfo, table4);
				}
			}
			if (ToKhaiInfo.SheetInfos[idxPL03_3d] != null)
			{
				foreach (HTKKCMCTableInfo table5 in ToKhaiInfo.SheetInfos[idxPL03_3d].Tables)
				{
					TinhLaiCQTTheoUuDai_PL03_3D(winFpSpread.Sheets[idxPL03_3d], _TTinChung, _PeriodInfo, table5);
				}
			}
			if (ToKhaiInfo.SheetInfos[idxPL03_8b] != null)
			{
				int colHeader = HTKKFile.ColumnLettersToIndex("F");
				int colHeader2 = HTKKFile.ColumnLettersToIndex("AF");
				int colHeader3 = HTKKFile.ColumnLettersToIndex("P");
				int colChilden = HTKKFile.ColumnLettersToIndex("BB");
				int colChilden2 = HTKKFile.ColumnLettersToIndex("BC");
				int colChilden3 = HTKKFile.ColumnLettersToIndex("BD");
				int num5 = 0;
				foreach (HTKKCMCTableInfo table6 in ToKhaiInfo.SheetInfos[idxPL03_8b].Tables)
				{
					for (int i = 0; i < table6.RowsValues.Count; i++)
					{
						int dataLocationTemplateRowStart = table6.Location.DataLocationTemplateRowStart;
						GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[idxPL03_8b], ToKhaiInfo.SheetInfos[idxPL03_8b], num5, dataLocationTemplateRowStart + i, colHeader2, colChilden);
						GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[idxPL03_8b], ToKhaiInfo.SheetInfos[idxPL03_8b], num5, dataLocationTemplateRowStart + i, colHeader3, colChilden2);
						GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[idxPL03_8b], ToKhaiInfo.SheetInfos[idxPL03_8b], num5, dataLocationTemplateRowStart + i, colHeader, colChilden3);
						GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[idxPL03_8b], dataLocationTemplateRowStart + i, -1, table6.RowsInfo, table6);
					}
					num5++;
				}
			}
			if (ToKhaiInfo.SheetInfos[idxPL03_8c] != null)
			{
				TinhTongTheoTinh_Pl03_8C(winFpSpread, ToKhaiInfo, -1);
				UpdateChoTongTheoTinhDuLieuCu_Pl03_8C(winFpSpread.Sheets[idxPL03_8c], ToKhaiInfo.SheetInfos[idxPL03_8c]);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct04_ma");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct04_ten");
			if (hTKKCMCCellInfo == null || hTKKCMCCellInfo2 == null)
			{
				return;
			}
			string text = winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(winFpSpread.Sheets[0], hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex);
				if (dMFromCombobox.ContainsKey(text))
				{
					winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = dMFromCombobox[text];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateChoTongTheoTinhDuLieuCu_Pl03_8C(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			List<int> list = new List<int>();
			int item = HTKKFile.ColumnLettersToIndex("M");
			int item2 = HTKKFile.ColumnLettersToIndex("AI");
			list.Add(item);
			list.Add(item2);
			_listRowAndKeyOld_Pl01_8C = GeneralFunctions.GetRowOnSViewCreateDictionary(sView, sInfo.Tables[0], list);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhTongTheoTinh_Pl03_8C(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[idxPL03_8c];
			SheetView sView = winFpSpread.Sheets[idxPL03_8c];
			int num = HTKKFile.ColumnLettersToIndex("M");
			int num2 = HTKKFile.ColumnLettersToIndex("AI");
			int value = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("AL");
			int num4 = HTKKFile.ColumnLettersToIndex("U");
			int num5 = HTKKFile.ColumnLettersToIndex("V");
			int num6 = HTKKFile.ColumnLettersToIndex("Y");
			int num7 = HTKKFile.ColumnLettersToIndex("Z");
			int num8 = HTKKFile.ColumnLettersToIndex("AA");
			int num9 = HTKKFile.ColumnLettersToIndex("AB");
			int num10 = HTKKFile.ColumnLettersToIndex("AC");
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(num, value);
			dictionary.Add(num2, num3);
			dictionary.Add(num6, num6);
			dictionary.Add(num7, num7);
			dictionary.Add(num8, num8);
			dictionary.Add(num9, num9);
			dictionary.Add(num10, num10);
			if (iCol < 0 || iCol == num)
			{
				Dictionary<string, int> checkExists = new Dictionary<string, int>();
				GeneralFunctions.CopyTableToTable2(sView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			}
			if (iCol == num4 || iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8 || iCol == num9 || iCol == num10)
			{
				List<int> list = new List<int>();
				list.Add(num6);
				list.Add(num7);
				list.Add(num8);
				list.Add(num9);
				list.Add(num10);
				GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(sView, hTKKCMCSheetInfo, "03_8C_TNDN_1", "03_8C_TNDN_2", num2, num3, list);
			}
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_15(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listKeyAndRowTo_Pl01_8C.Clear();
			_listRowAndKeyOld_Pl01_8C.Clear();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiCQTTheoUuDai_PL03_3A(SheetView sView, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKCMCTableInfo table)
	{
		int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("I");
		int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("AX");
		int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("K");
		int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AY");
		HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
		HTKKCMCCellInfo hTKKCMCCellInfo5 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
		string tenCQTCapCuc = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
		string maCQTCapCuc = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
		string tenCQTChiCuc = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
		string maCQTChiCuc = sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int num = 0; num < table.RowsValues.Count; num++)
		{
			if (text != UU_DAI_KHAC_TINH)
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, TTinChung.tenCQTCapCuc, TTinChung.maCQTCapCuc, TTinChung.tenCQTNoiNop, TTinChung.maCQTNoiNop, dataLocationTemplateRowStart + num, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc);
			}
			else
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, tenCQTCapCuc, maCQTCapCuc, tenCQTChiCuc, maCQTChiCuc, dataLocationTemplateRowStart + num, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc);
			}
		}
	}

	public void TinhLaiCQTTheoUuDai_PL03_3B(SheetView sView, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKCMCTableInfo table)
	{
		int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("I");
		int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("AY");
		int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("K");
		int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AZ");
		HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ten"));
		HTKKCMCCellInfo hTKKCMCCellInfo5 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10_ma"));
		string tenCQTCapCuc = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
		string maCQTCapCuc = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
		string tenCQTChiCuc = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
		string maCQTChiCuc = sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int num = 0; num < table.RowsValues.Count; num++)
		{
			if (text != UU_DAI_KHAC_TINH)
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, TTinChung.tenCQTCapCuc, TTinChung.maCQTCapCuc, TTinChung.tenCQTNoiNop, TTinChung.maCQTNoiNop, dataLocationTemplateRowStart + num, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc);
			}
			else
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, tenCQTCapCuc, maCQTCapCuc, tenCQTChiCuc, maCQTChiCuc, dataLocationTemplateRowStart + num, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc);
			}
		}
	}

	public void TinhLaiCQTTheoUuDai_PL03_3D(SheetView sView, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKCMCTableInfo table)
	{
		int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("I");
		int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("AW");
		int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("K");
		int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AX");
		HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("uuDai_ma"));
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathCucThue23"));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathMaCucThue23"));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10a_ten"));
		HTKKCMCCellInfo hTKKCMCCellInfo5 = table.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct10a_ma"));
		string tenCQTCapCuc = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
		string maCQTCapCuc = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
		string tenCQTChiCuc = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
		string maCQTChiCuc = sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int num = 0; num < table.RowsValues.Count; num++)
		{
			if (text != UU_DAI_KHAC_TINH)
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, TTinChung.tenCQTCapCuc, TTinChung.maCQTCapCuc, TTinChung.tenCQTNoiNop, TTinChung.maCQTNoiNop, dataLocationTemplateRowStart + num, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc);
			}
			else
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, tenCQTCapCuc, maCQTCapCuc, tenCQTChiCuc, maCQTChiCuc, dataLocationTemplateRowStart + num, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc);
			}
		}
	}

	public Dictionary<string, string> SetLaiGiaTriCombobox(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCTableInfo table, int colCombobox, int nguongSoSanh)
	{
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(sView, dataLocationTemplateRowStart, colCombobox);
		string nam = PeriodInfo.Nam;
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		if (Convert.ToInt32(nam) < 2025)
		{
			dictionary = dMFromCombobox.Where((KeyValuePair<string, string> x) => Convert.ToInt32(x.Key) < nguongSoSanh).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
		}
		else
		{
			dictionary = dMFromCombobox.Where((KeyValuePair<string, string> x) => Convert.ToInt32(x.Key) >= nguongSoSanh).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
			if (Convert.ToInt32(nam) != 2025)
			{
				dictionary = dictionary.Where((KeyValuePair<string, string> x) => x.Key != "0262").ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
			}
		}
		for (int num = 0; num < table.RowsValues.Count; num++)
		{
			GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart + num, colCombobox, dictionary, deleteText: false);
		}
		return dictionary;
	}

	public int[] ConvertPathToCellId(HTKKCMCSheetInfo sInfo, string path)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path);
		return new int[2]
		{
			hTKKCMCCellInfo.CellTemplate.ColIndex - 1,
			hTKKCMCCellInfo.CellTemplate.RowIndex - 1
		};
	}

	public void tinhCtC3a_tk(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
		SheetView sheetView2 = winFpSpread.Sheets[idxPL03_2];
		int[] array = GeneralFunctions.ArrayCellId("M_33");
		int[] array2 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3a");
		sheetView.Cells[array2[1], array2[0]].Value = sheetView2.Cells[array[1], array[0]].Value;
	}

	public void lockUnlockCtC3a(SheetView sViewTk, HTKKCMCSheetInfo sInfo)
	{
		int[] array = GeneralFunctions.ArrayCellId("O_53");
		int[] array2 = GeneralFunctions.ArrayCellId("O_54");
		int[] array3 = ConvertPathToCellId(sInfo, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC3a");
		double num = Convert.ToDouble(sViewTk.Cells[array[1], array[0]].Value);
		double num2 = Convert.ToDouble(sViewTk.Cells[array2[1], array2[0]].Value);
		if (num - num2 <= 0.0)
		{
			GeneralFunctions.LockUnLockCol(sViewTk, array3[1] + 1, array3[0] + 1, locked: true, deleteData: true);
		}
		else
		{
			GeneralFunctions.LockUnLockCol(sViewTk, array3[1] + 1, array3[0] + 1, locked: false, deleteData: false);
		}
	}
}
