using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn02TT21Nckt295
{
	private int[] cellCnbds = GeneralFunctions.ArrayCellId("C_18");

	private int[] cellBdgbds = GeneralFunctions.ArrayCellId("C_20");

	private int[] cellBds = GeneralFunctions.ArrayCellId("I_22");

	private int[] cellTsk = GeneralFunctions.ArrayCellId("AG_22");

	private int[] cellCt4 = GeneralFunctions.ArrayCellId("O_39");

	private int[] cellCt5 = GeneralFunctions.ArrayCellId("O_41");

	private int[] cellCt9cTen = GeneralFunctions.ArrayCellId("I_74");

	private int[] cellCt9cMa = GeneralFunctions.ArrayCellId("BM_74");

	private int[] cellCt9bTen = GeneralFunctions.ArrayCellId("I_78");

	private int[] cellCt9bMa = GeneralFunctions.ArrayCellId("BM_78");

	private int[] cellCt9a = GeneralFunctions.ArrayCellId("I_80");

	private int[] cellTenTinhHeader = GeneralFunctions.ArrayCellId("BN_74");

	private int[] cellMaTinhHeader = GeneralFunctions.ArrayCellId("BO_74");

	private int[] cellTenXaHeader = GeneralFunctions.ArrayCellId("BN_78");

	private int[] cellMaXaHeader = GeneralFunctions.ArrayCellId("BO_78");

	private int[] cellDiaChiHeader = GeneralFunctions.ArrayCellId("BN_80");

	private int colBN = HTKKFile.ColumnLettersToIndex("BN");

	private int colBO = HTKKFile.ColumnLettersToIndex("BO");

	private int[] cell_item_2_static_start = GeneralFunctions.ArrayCellId("T_51");

	private int[] cell_item_2_static_end = GeneralFunctions.ArrayCellId("I_80");

	private string pathCt10 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThongTinCoQuanThiHanhAn/ct10";

	private string pathCt11 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThongTinCoQuanThiHanhAn/ct11";

	private string path_item_3_static_start = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct16";

	private string path_item_3_static_end = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct23";

	private string pathCt25Ten = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct25_ten";

	private string pathCt25Ma = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct25";

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			sView.Cells[cellTenTinhHeader[1], cellTenTinhHeader[0]].Value = TTinChung.tenTinhNNT;
			sView.Cells[cellMaTinhHeader[1], cellMaTinhHeader[0]].Value = TTinChung.maTinhNNT;
			sView.Cells[cellTenXaHeader[1], cellTenXaHeader[0]].Value = TTinChung.tenXaPhuongNNT;
			sView.Cells[cellMaXaHeader[1], cellMaXaHeader[0]].Value = TTinChung.maXaPhuongNNT;
			sView.Cells[cellDiaChiHeader[1], cellDiaChiHeader[0]].Value = TTinChung.dchiNNT;
			sView.Cells[cellCnbds[1], cellCnbds[0]].Value = PeriodInfo.TieuMucHachToanVal == "1";
			sView.Cells[cellBdgbds[1], cellBdgbds[0]].Value = PeriodInfo.TieuMucHachToanVal == "2";
			sView.Cells[cellBds[1], cellBds[0]].Value = PeriodInfo.ToChucCNDNVonChuyenNhuong == 1;
			sView.Cells[cellTsk[1], cellTsk[0]].Value = PeriodInfo.ToChucCNDNVonChuyenNhuong == 2;
			LockUnlockItem_3(sView, sInfo);
			AutoBidingItem1(sView, sInfo);
			AutoBidingItem2(sView, sInfo, PeriodInfo);
			AutoBidingItem3(sView, sInfo);
			SetCt24(sView, sInfo);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt25Ten);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt25Ma);
			sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
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
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					continue;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
						continue;
					}
					break;
				}
				winFpSpread.TabStrip[key].BackColor = Color.White;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return false;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
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

	public void LockUnlockItem_2(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		bool flag = Convert.ToBoolean(sView.Cells[cellCnbds[1], cellCnbds[0]].Value);
		List<int> list = new List<int>();
		list.Add(HTKKFile.ColumnLettersToIndex("E"));
		list.Add(HTKKFile.ColumnLettersToIndex("I"));
		list.Add(HTKKFile.ColumnLettersToIndex("Q"));
		if (flag)
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cell_item_2_static_start[1] + 1, cell_item_2_static_end[1] + 1, isLock: false, isDeleteData: false);
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[0], list, locked: false, deleteData: false);
		}
		else
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cell_item_2_static_start[1] + 1, cell_item_2_static_end[1] + 1, isLock: true, isDeleteData: true);
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[0], list, locked: true, deleteData: true);
		}
	}

	public void LockUnlockItem_3(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		bool flag = Convert.ToBoolean(sView.Cells[cellCnbds[1], cellCnbds[0]].Value);
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_item_3_static_start);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_item_3_static_end);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt25Ten);
		if (flag)
		{
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: true, deleteData: true);
		}
		else
		{
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: false, deleteData: false);
		}
	}

	public void AutoBidingItem1(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		if (Convert.ToBoolean(sView.Cells[cellCnbds[1], cellCnbds[0]].Value))
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellCt4[1] + 1, cellCt5[1] + 1, isLock: true, isDeleteData: true);
			sView.Cells[cellCt4[1], cellCt4[0]].Value = sView.Cells[cellCt4[1], colBN - 1].Value;
			sView.Cells[cellCt5[1], cellCt5[0]].Value = sView.Cells[cellCt5[1], colBN - 1].Value;
		}
		else
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellCt4[1] + 1, cellCt5[1] + 1, isLock: false, isDeleteData: false);
		}
	}

	public void AutoBidingItem2(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo)
	{
		if (Convert.ToBoolean(sView.Cells[cellCnbds[1], cellCnbds[0]].Value))
		{
			GeneralFunctions.LockUnLockCol(sView, cellCt9cTen[1] + 1, cellCt9cTen[0] + 1, locked: true, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, cellCt9bTen[1] + 1, cellCt9bTen[0] + 1, locked: true, deleteData: false);
			if (!string.IsNullOrEmpty(PeriodInfo.TenTinh) && !string.IsNullOrEmpty(PeriodInfo.TenXa))
			{
				sView.Cells[cellCt9cTen[1], cellCt9cTen[0]].Value = PeriodInfo.TenTinh;
				sView.Cells[cellCt9bTen[1], cellCt9bTen[0]].Value = PeriodInfo.TenXa;
				sView.Cells[cellCt9cMa[1], cellCt9cMa[0]].Value = PeriodInfo.MaTinh;
				sView.Cells[cellCt9bMa[1], cellCt9bMa[0]].Value = PeriodInfo.MaXa;
			}
		}
		else
		{
			GeneralFunctions.LockUnLockCol(sView, cellCt9cTen[1] + 1, cellCt9cTen[0] + 1, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, cellCt9bTen[1] + 1, cellCt9bTen[0] + 1, locked: false, deleteData: false);
		}
	}

	public void AutoBidingItem3(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		bool flag = Convert.ToBoolean(sView.Cells[cellBdgbds[1], cellBdgbds[0]].Value);
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt10);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt11);
		if (flag)
		{
			sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, colBN - 1].Value;
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, colBN - 1].Value;
		}
	}

	public void SetCt24(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt25Ten);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt25Ma);
		bool flag = Convert.ToBoolean(sView.Cells[cellCnbds[1], cellCnbds[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellBdgbds[1], cellBdgbds[0]].Value);
		bool flag3 = Convert.ToBoolean(sView.Cells[cellBds[1], cellBds[0]].Value);
		bool flag4 = Convert.ToBoolean(sView.Cells[cellTsk[1], cellTsk[0]].Value);
		if (flag)
		{
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: true);
			sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = "20";
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = "20";
		}
		else if (flag2 && flag3)
		{
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: true);
			sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = "20";
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = "20";
		}
		else
		{
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: false, deleteData: false);
		}
	}
}
