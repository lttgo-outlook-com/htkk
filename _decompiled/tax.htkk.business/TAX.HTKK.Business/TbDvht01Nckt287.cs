using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TbDvht01Nckt287
{
	private string pathToKhaiLanDau = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/LoaiThongBao/nnt_ThongBaoLanDau";

	private string pathToKhaiLanDauChiDinh = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/LoaiThongBao/cqt_ThongBaoLanDau";

	private string pathTapDoanDaQuocGiaQDMTT = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DoiTuongApDung/doi_TuongApDungQDMTT";

	private string pathTapDoanDaQuocGiaQDMTTStart = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DoiTuongApDung/loai_ApDungQDMTT/dvht_TapDoanDaQuocGia";

	private string pathTapDoanDaQuocGiaQDMTTEnd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DoiTuongApDung/loai_ApDungQDMTT/dvht_KhongThanhVienTapDoanCon";

	private string pathDoiTuongApDungIIR = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DoiTuongApDung/doi_TuongApDungIIR";

	private string pathII_1_TenDonVi = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/ten_DVHTKeKhai";

	private string pathII_2_MaSoThue = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/mst_DVHTKeKhai";

	private string pathII_3_TenTinh = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/ten_TinhDVHTKeKhai";

	private string pathII_3_TenXa = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/ten_XaDVHTKeKhai";

	private string pathII_3_MaTinh = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/ma_TinhDVHTKeKhai";

	private string pathII_3_MaXa = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/ma_XaDVHTKeKhai";

	private string pathII_3_DiaChi = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TT_DVHTKeKhai/diaChi_DVHTKeKhai";

	private int[] ngayView = GeneralFunctions.ArrayCellId("S_23");

	private int[] cellTenKinhGui = GeneralFunctions.ArrayCellId("F_74");

	private int[] cellMaKinhGui = GeneralFunctions.ArrayCellId("BI_74");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			loadDefaultII(sheetView, sInfo, TTinChung);
			GeneralFunctions.setButton(sheetView, 2, 3, "Tải bảng kê");
			sheetView.Cells[ngayView[1], ngayView[0]].Text = PeriodInfo.kyKKhai;
			sheetView.Cells[cellTenKinhGui[1], cellTenKinhGui[0]].Text = "Chi cục Thuế doanh nghiệp lớn";
			sheetView.Cells[cellMaKinhGui[1], cellMaKinhGui[0]].Text = "82500";
			lockUnlockApDungQDMTT(sheetView, sInfo);
			xoaDataTable_III(sheetView, sInfo);
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
		{
			return false;
		}
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_38"));
			for (int i = 1; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				int colMst = 0;
				if (i == 2)
				{
					colMst = HTKKFile.ColumnLettersToIndex("J");
				}
				else
				{
					colMst = HTKKFile.ColumnLettersToIndex("E");
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				int indexMst = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMst);
				for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
				{
					string valueMst = hTKKCMCTableInfo.RowsValues[num].RowValue[indexMst];
					List<HTKKCMCRowValue> list = hTKKCMCTableInfo.RowsValues.FindAll((HTKKCMCRowValue x) => x.RowValue[indexMst] == valueMst);
					if (list.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + num - 1, colMst - 1, message);
						value = 2;
					}
				}
			}
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			xoaDataTable_III(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void xoaDataTable_III(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTT);
		bool flag = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
		List<int> list = new List<int>();
		foreach (HTKKCMCTableCellInfo cell in sInfo.Tables[1].RowsInfo.Cells)
		{
			list.Add(cell.CellTemplate.ColIndex);
		}
		if (!flag)
		{
			GeneralFunctions.RemoveDataOfTable(sView, 1, sInfo);
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[1], list, locked: true, deleteData: true);
		}
		else
		{
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[1], list, locked: false, deleteData: false);
		}
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathDoiTuongApDungIIR);
		bool flag2 = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
		List<int> list2 = new List<int>();
		foreach (HTKKCMCTableCellInfo cell2 in sInfo.Tables[2].RowsInfo.Cells)
		{
			list2.Add(cell2.CellTemplate.ColIndex);
		}
		List<int> list3 = new List<int>();
		foreach (HTKKCMCTableCellInfo cell3 in sInfo.Tables[3].RowsInfo.Cells)
		{
			list3.Add(cell3.CellTemplate.ColIndex);
		}
		if (!flag2)
		{
			GeneralFunctions.RemoveDataOfTable(sView, 2, sInfo);
			GeneralFunctions.RemoveDataOfTable(sView, 3, sInfo);
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[2], list2, locked: true, deleteData: true);
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[3], list3, locked: true, deleteData: true);
		}
		else
		{
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[2], list2, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockColsInTable(sView, sInfo.Tables[3], list3, locked: false, deleteData: false);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeRowIndex == 1 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_TB_DVHT.xlsx");
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathToKhaiLanDau);
			if (activeRowIndex == hTKKCMCCellInfo.CellTemplate.RowIndex - 1)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathToKhaiLanDauChiDinh);
				GeneralFunctions.resetOtherCheckboxStatic(sView, hTKKCMCSheetInfo, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.RowIndex, activeRowIndex + 1, activeColumnIndex + 1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTT);
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTTStart);
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTTEnd);
			HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathDoiTuongApDungIIR);
			if (activeRowIndex >= hTKKCMCCellInfo4.CellTemplate.RowIndex - 1 && activeRowIndex <= hTKKCMCCellInfo5.CellTemplate.RowIndex - 1)
			{
				GeneralFunctions.resetOtherCheckboxStatic(sView, hTKKCMCSheetInfo, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.RowIndex, activeRowIndex + 1, activeColumnIndex + 1);
			}
			lockUnlockApDungQDMTT(sView, hTKKCMCSheetInfo);
			if (activeRowIndex == hTKKCMCCellInfo3.CellTemplate.RowIndex - 1 || activeRowIndex == hTKKCMCCellInfo6.CellTemplate.RowIndex - 1)
			{
				xoaDataTable_III(sView, hTKKCMCSheetInfo);
			}
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void loadDefaultII(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTTinChungInfo TTinChung)
	{
		if (!string.IsNullOrEmpty(TTinChung.soLan))
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_1_TenDonVi);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_2_MaSoThue);
			HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_3_TenTinh);
			HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_3_TenXa);
			HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_3_MaTinh);
			HTKKCMCCellInfo hTKKCMCCellInfo6 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_3_MaXa);
			HTKKCMCCellInfo hTKKCMCCellInfo7 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathII_3_DiaChi);
			sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = TTinChung.tenNNT;
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = TTinChung.mst;
			sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = TTinChung.tenTinhNNT;
			sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text = TTinChung.tenXaPhuongNNT;
			sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text = TTinChung.maTinhNNT;
			sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = TTinChung.maXaPhuongNNT;
			sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text = TTinChung.dchiNNT;
		}
	}

	public void lockUnlockApDungQDMTT(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTT);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTTStart);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTapDoanDaQuocGiaQDMTTEnd);
		if (!Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
		}
		else
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.RowIndex, isLock: false, isDeleteData: false);
		}
	}
}
