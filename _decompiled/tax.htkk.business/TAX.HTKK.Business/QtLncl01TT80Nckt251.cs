using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class QtLncl01TT80Nckt251
{
	private int[] cellTenTruongHopQT = GeneralFunctions.ArrayCellId("I_30");

	private int[] cellMaTruongHopQT = GeneralFunctions.ArrayCellId("H_30");

	private int[] cellTenApDung = GeneralFunctions.ArrayCellId("AB_2");

	private int[] cellMaApDung = GeneralFunctions.ArrayCellId("AC_2");

	private int[] cellCt10 = GeneralFunctions.ArrayCellId("U_54");

	private int[] cellCt17 = GeneralFunctions.ArrayCellId("U_61");

	private int[] cellPhuLuc = GeneralFunctions.ArrayCellId("AC_5");

	private int[] cellCt04 = GeneralFunctions.ArrayCellId("AL_35");

	private int colMaLoai = HTKKFile.ColumnLettersToIndex("CB");

	private int col08 = HTKKFile.ColumnLettersToIndex("Q");

	private int col09 = HTKKFile.ColumnLettersToIndex("W");

	private int col09Ma = HTKKFile.ColumnLettersToIndex("CC");

	private int col11a = HTKKFile.ColumnLettersToIndex("AI");

	private int col11aMa = HTKKFile.ColumnLettersToIndex("CE");

	private int col11b = HTKKFile.ColumnLettersToIndex("AM");

	private int col11bMa = HTKKFile.ColumnLettersToIndex("CF");

	private int col12 = HTKKFile.ColumnLettersToIndex("AP");

	private int col14 = HTKKFile.ColumnLettersToIndex("AZ");

	private int col15 = HTKKFile.ColumnLettersToIndex("BE");

	private int col16 = HTKKFile.ColumnLettersToIndex("BJ");

	private int col17 = HTKKFile.ColumnLettersToIndex("BO");

	private int col18 = HTKKFile.ColumnLettersToIndex("BT");

	private int colTenTinh = HTKKFile.ColumnLettersToIndex("C");

	private int colMaTinh = HTKKFile.ColumnLettersToIndex("CC");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private List<string> _listKeyAndRowTo = new List<string>();

	private List<string> _listRowAndKeyOld = new List<string>();

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sheetView.Cells[cellTenTruongHopQT[1], cellTenTruongHopQT[0]].Text = PeriodInfo.TenQuyetToan;
				sheetView.Cells[cellMaTruongHopQT[1], cellMaTruongHopQT[0]].Text = PeriodInfo.MaQuyetToan;
			}
			if (!string.IsNullOrEmpty(PeriodInfo.TieuMucHachToanVal))
			{
				sheetView.Cells[cellTenApDung[1], cellTenApDung[0]].Text = PeriodInfo.NganhNgheKinhDoanh;
				sheetView.Cells[cellMaApDung[1], cellMaApDung[0]].Text = PeriodInfo.TieuMucHachToanVal;
			}
			switch (PeriodInfo.TieuMucHachToanVal)
			{
			case "01":
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "1153";
				break;
			case "02":
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "1151";
				break;
			}
			winFpSpread.Sheets[0].Cells[cellPhuLuc[1], cellPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		UpdateChoTongTheoTinhDuLieuCu_Pl01_1(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1]);
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iRow == cellCt10[0] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (ToKhaiInfo.SheetInfos[1] != null && iRow >= cellCt10[1] + 1 && iRow < cellCt17[1] + 1)
			{
				winFpSpread.Sheets[1].Cells[cellCt04[1], cellCt04[0]].Value = winFpSpread.Sheets[0].Cells[cellCt17[1], cellCt17[0]].Value;
				TongTheoTungTinh(winFpSpread, ToKhaiInfo, -1);
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
			SheetView sheetView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("AL_35");
			int[] array2 = GeneralFunctions.ArrayCellId("U_37");
			if (iCol == col09)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, col09Ma - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, col11a, col11aMa, col11b, col11bMa, isLock: false);
			}
			if (iRow == array[1] + 1 || iRow == array2[1] + 1)
			{
				TongTheoTungTinh(winFpSpread, ToKhaiInfo, -1);
			}
			else
			{
				TongTheoTungTinh(winFpSpread, ToKhaiInfo, (int)iCol);
			}
			GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col11a, locked: true, deleteData: false);
			string text = sheetView.Cells[(int)iRow - 1, colMaLoai - 1].Text.Trim();
			if (text == "03")
			{
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col08, locked: true, deleteData: true);
			}
			else
			{
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col08, locked: false, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongTheoTungTinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iCol)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		SheetView sheetView = winFpSpread.Sheets[1];
		if (hTKKCMCSheetInfo != null && sheetView != null)
		{
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(col09, colTenTinh);
			dictionary.Add(col09Ma, colMaTinh);
			dictionary.Add(col14, col14);
			dictionary.Add(col15, col15);
			dictionary.Add(col16, col16);
			dictionary.Add(col17, col17);
			dictionary.Add(col18, col18);
			if (iCol == col09 || iCol <= 0)
			{
				Dictionary<string, int> checkExists = new Dictionary<string, int>();
				GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			}
			if (iCol == col12 || iCol == col15 || iCol == col16)
			{
				List<int> list = new List<int>();
				list.Add(col14);
				list.Add(col15);
				list.Add(col16);
				list.Add(col17);
				list.Add(col18);
				GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(sheetView, hTKKCMCSheetInfo, "01_1_QT_LNCL_1", "01_1_QT_LNCL_2", col09Ma, colMaTinh, list);
			}
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[1]);
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
				else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else
				{
					winFpSpread.TabStrip[key].BackColor = Color.White;
				}
			}
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			switch (0)
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 0;
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, col08 - 1);
				string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, col08 - 1].Text.Trim();
				string text4 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, colMaLoai - 1].Text.Trim();
				if (string.IsNullOrEmpty(text3))
				{
					continue;
				}
				if (text4 == "01")
				{
					if ((text3.Length == 13 || text3.Length == 14) && !ValidateBase.CheckMaSoThue(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col08 - 1, msg);
						num = 2;
					}
					else if (text3.Length != 13 && text3.Length != 14)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col08 - 1, msg);
						num = 2;
					}
				}
				if (text4 == "02")
				{
					if (text3.Length == 5 && !ValidateBase.CheckMaSoThue5So(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col08 - 1, msg2);
						num = 2;
					}
					else if (text3.Length != 5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col08 - 1, msg2);
						num = 2;
					}
				}
			}
			switch (num)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
				break;
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc[1], cellPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			_listRowAndKeyOld.Clear();
			_listKeyAndRowTo.Clear();
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc[1], cellPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			_listRowAndKeyOld.Clear();
			_listKeyAndRowTo.Clear();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TongTheoTungTinh(winFpSpread, ToKhaiInfo, -1);
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
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				UpdateChoTongTheoTinhDuLieuCu_Pl01_1(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateChoTongTheoTinhDuLieuCu_Pl01_1(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(col09, colTenTinh);
			dictionary.Add(col09Ma, colMaTinh);
			_listRowAndKeyOld = GeneralFunctions.GetRowOnSViewCreateDictionary(sView, sInfo.Tables[0], dictionary.Keys.ToList());
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
