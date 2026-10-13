using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnDtv04Nckt251
{
	private int[] cellCaNhanDuocMienGiamThue = GeneralFunctions.ArrayCellId("C_22");

	private int[] cellToChucKhaiThay = GeneralFunctions.ArrayCellId("C_24");

	private int[] cellCt12 = GeneralFunctions.ArrayCellId("L_46");

	private int[] cellCt13 = GeneralFunctions.ArrayCellId("H_48");

	private int[] cellCt14 = GeneralFunctions.ArrayCellId("H_50");

	private int[] cellCt15 = GeneralFunctions.ArrayCellId("H_52");

	private int[] cellCt16 = GeneralFunctions.ArrayCellId("H_54");

	private int[] cellCt17 = GeneralFunctions.ArrayCellId("F_56");

	private int[] cellCt18 = GeneralFunctions.ArrayCellId("L_56");

	private int[] cellCt19 = GeneralFunctions.ArrayCellId("S_56");

	private int[] cellTemPhuLuc = GeneralFunctions.ArrayCellId("AI_24");

	private HTKKHeaderInfo _HeaderInfo = null;

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_HeaderInfo = HeaderInfo;
			_ToKhaiInfo = ToKhaiInfo;
			bool isDefault = ((PeriodInfo.KieuKy != "D" && string.IsNullOrEmpty(sView.Cells[cellCt13[1], cellCt13[0]].Text.Trim())) ? true : false);
			ThayDoiHinhThucKhaiThay(sView, PeriodInfo, isDefault);
			sView.Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			LockBangTKNeuCoPhuLuc(sView);
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
			bool flag = Convert.ToBoolean(sheetView.Cells[cellCaNhanDuocMienGiamThue[1], cellCaNhanDuocMienGiamThue[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellToChucKhaiThay[1], cellToChucKhaiThay[0]].Value);
			if (iRow == cellCaNhanDuocMienGiamThue[1] + 1 && flag)
			{
				sheetView.Cells[cellToChucKhaiThay[1], cellToChucKhaiThay[0]].Value = false;
			}
			if (iRow == cellToChucKhaiThay[1] + 1 && flag2)
			{
				sheetView.Cells[cellCaNhanDuocMienGiamThue[1], cellCaNhanDuocMienGiamThue[0]].Value = false;
			}
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
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
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				ThayDoiHinhThucKhaiThay(winFpSpread.Sheets[0], PeriodInfo, isDefault: true);
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
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
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
		SheetView sheetView = winFpSpread.Sheets[0];
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		switch (value)
		{
		case 2:
			sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
			break;
		case 1:
			sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
			break;
		default:
			sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
			break;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("C_24");
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			if (hTKKCMCSheetInfo != null)
			{
				switch (value)
				{
				case 2:
					sheetView.FpSpread.TabStrip[1].BackColor = Color.Pink;
					break;
				case 1:
					sheetView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
					break;
				default:
					sheetView.FpSpread.TabStrip[1].BackColor = Color.White;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiHinhThucKhaiThay(SheetView sView, HTKKPeriodInfo PeriodInfo, bool isDefault)
	{
		if (Convert.ToBoolean(sView.Cells[cellToChucKhaiThay[1], cellToChucKhaiThay[0]].Value))
		{
			if (isDefault)
			{
				sView.Cells[cellCt12[1], cellCt12[0]].Value = _HeaderInfo.tenNNT;
				sView.Cells[cellCt13[1], cellCt13[0]].Value = _HeaderInfo.mst;
				sView.Cells[cellCt14[1], cellCt14[0]].Value = _HeaderInfo.dchiNNT;
				sView.Cells[cellCt15[1], cellCt15[0]].Value = _HeaderInfo.tenHuyenNNT;
				sView.Cells[cellCt16[1], cellCt16[0]].Value = _HeaderInfo.tenTinhNNT;
				sView.Cells[cellCt17[1], cellCt17[0]].Value = _HeaderInfo.dthoaiNNT;
				sView.Cells[cellCt18[1], cellCt18[0]].Value = _HeaderInfo.faxNNT;
				sView.Cells[cellCt19[1], cellCt19[0]].Value = _HeaderInfo.emailNNT;
			}
		}
		else if (isDefault)
		{
			sView.Cells[cellCt12[1], cellCt12[0]].Value = "";
			sView.Cells[cellCt13[1], cellCt13[0]].Value = "";
			sView.Cells[cellCt14[1], cellCt14[0]].Value = "";
			sView.Cells[cellCt15[1], cellCt15[0]].Value = "";
			sView.Cells[cellCt16[1], cellCt16[0]].Value = "";
			sView.Cells[cellCt17[1], cellCt17[0]].Value = "";
			sView.Cells[cellCt18[1], cellCt18[0]].Value = "";
			sView.Cells[cellCt19[1], cellCt19[0]].Value = "";
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			LockBangTKNeuCoPhuLuc(winFpSpread.Sheets[0]);
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
			LockBangTKNeuCoPhuLuc(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockBangTKNeuCoPhuLuc(SheetView sView)
	{
		if (_ToKhaiInfo.SheetInfos[1] != null)
		{
			for (int i = 0; i < _ToKhaiInfo.SheetInfos[0].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				if (hTKKCMCTableInfo.RowsValues.Count > 1)
				{
					GeneralFunctions.RemoveDataRows(sView, i, dataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, _ToKhaiInfo.SheetInfos[0]);
				}
			}
			for (int i = 0; i < _ToKhaiInfo.SheetInfos[0].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
					{
						sView.Cells[dataLocationTemplateRowStart + j - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
						GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + j, cell.CellTemplate.ColIndex, locked: true, deleteData: false);
					}
				}
			}
			return;
		}
		for (int i = 0; i < _ToKhaiInfo.SheetInfos[0].Tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[i];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + j, cell2.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
			}
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeRowIndex == 1 && activeColumnIndex == 3)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_04_1_DTV_TNCN_TT80.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
