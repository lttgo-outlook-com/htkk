using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnQtt05TT80Nckt279
{
	private string _valueMstOld_PL01_3 = "";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("H_5");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("AG_5");

	private int[] cellCt17_TK = GeneralFunctions.ArrayCellId("U_49");

	private int[] cellCt17Formualar_TK = GeneralFunctions.ArrayCellId("AV_3");

	private int colCt07_PL_05_1 = HTKKFile.ColumnLettersToIndex("D");

	private int colCt08_PL_05_1 = HTKKFile.ColumnLettersToIndex("E");

	private int colCt09a_PL_05_1 = HTKKFile.ColumnLettersToIndex("G");

	private int colCt09aMa_PL_05_1 = HTKKFile.ColumnLettersToIndex("BE");

	private int colCt09_PL_05_1 = HTKKFile.ColumnLettersToIndex("H");

	private int colCt10_PL_05_1 = HTKKFile.ColumnLettersToIndex("I");

	private int colCt11_PL_05_1 = HTKKFile.ColumnLettersToIndex("K");

	private int colCt12_PL_05_1 = HTKKFile.ColumnLettersToIndex("M");

	private int colCt13_PL_05_1 = HTKKFile.ColumnLettersToIndex("O");

	private int colCt14_PL_05_1 = HTKKFile.ColumnLettersToIndex("Q");

	private int colCt15_PL_05_1 = HTKKFile.ColumnLettersToIndex("S");

	private int colCt16_PL_05_1 = HTKKFile.ColumnLettersToIndex("U");

	private int colCt17_PL_05_1 = HTKKFile.ColumnLettersToIndex("W");

	private int colCt17Formualar_PL_05_1 = HTKKFile.ColumnLettersToIndex("BI");

	private int colCt20_PL_05_1 = HTKKFile.ColumnLettersToIndex("AC");

	private int colCt22_PL_05_1 = HTKKFile.ColumnLettersToIndex("AG");

	private int colCt23_PL_05_1 = HTKKFile.ColumnLettersToIndex("AI");

	private int colCt24_PL_05_1 = HTKKFile.ColumnLettersToIndex("AK");

	private int colCt24Formualar_PL_05_1 = HTKKFile.ColumnLettersToIndex("BJ");

	private int colCt25_PL_05_1 = HTKKFile.ColumnLettersToIndex("AM");

	private int colCt25Formualar_PL_05_1 = HTKKFile.ColumnLettersToIndex("BK");

	private int colTongGiaTriNPT_PL_05_1 = HTKKFile.ColumnLettersToIndex("BF");

	private int colBieuThucLuyTien_PL_05_1 = HTKKFile.ColumnLettersToIndex("BH");

	private int colCt07_PL_05_2 = HTKKFile.ColumnLettersToIndex("D");

	private int colCt08_PL_05_2 = HTKKFile.ColumnLettersToIndex("E");

	private int colCt09a_PL_05_2 = HTKKFile.ColumnLettersToIndex("G");

	private int colCt09_PL_05_2 = HTKKFile.ColumnLettersToIndex("H");

	private int colCt12_PL_05_2 = HTKKFile.ColumnLettersToIndex("M");

	private int colCt13_PL_05_2 = HTKKFile.ColumnLettersToIndex("P");

	private int colCt14_PL_05_2 = HTKKFile.ColumnLettersToIndex("R");

	private int colCt15_PL_05_2 = HTKKFile.ColumnLettersToIndex("T");

	private int colCt16_PL_05_2 = HTKKFile.ColumnLettersToIndex("V");

	private int colCt7_PL_05_3 = HTKKFile.ColumnLettersToIndex("C");

	private int colCt08_PL_05_3 = HTKKFile.ColumnLettersToIndex("E");

	private int colCt9_PL_05_3 = HTKKFile.ColumnLettersToIndex("G");

	private int colCt10_PL_05_3 = HTKKFile.ColumnLettersToIndex("I");

	private int colCt11_PL_05_3 = HTKKFile.ColumnLettersToIndex("K");

	private int colCt12_PL_05_3 = HTKKFile.ColumnLettersToIndex("M");

	private int colCt12Ma_PL_05_3 = HTKKFile.ColumnLettersToIndex("AJ");

	private int colCt13_PL_05_3 = HTKKFile.ColumnLettersToIndex("O");

	private int colCt14_PL_05_3 = HTKKFile.ColumnLettersToIndex("Q");

	private int colCt14Ma_PL_05_3 = HTKKFile.ColumnLettersToIndex("AI");

	private int colCt15_PL_05_3 = HTKKFile.ColumnLettersToIndex("S");

	private int colCt16_PL_05_3 = HTKKFile.ColumnLettersToIndex("U");

	private int colCt15ThoaMan_PL_05_3 = HTKKFile.ColumnLettersToIndex("AG");

	private int colCt16ThoaMan_PL_05_3 = HTKKFile.ColumnLettersToIndex("AH");

	private string mst_PL05_3 = "";

	private HTKKPeriodInfo _PeriodInfo = null;

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_PeriodInfo = PeriodInfo;
			_ToKhaiInfo = ToKhaiInfo;
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[cellTenTHQT_TK[1], cellTenTHQT_TK[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[cellMaTHQT_TK[1], cellMaTHQT_TK[0]].Value = PeriodInfo.MaQuyetToan;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("AX_5");
			int num = TinhSoThang(PeriodInfo.TuThang, PeriodInfo.DenThang);
			sheetView.Cells[array[1], array[0]].Value = num;
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo, dataLocationTemplateRowStart + i);
				LockUnlockCol24Col25_PL05_1(sheetView, dataLocationTemplateRowStart + i);
				TinhChiTieu26_27_PL05_1(sheetView, dataLocationTemplateRowStart + i);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[3].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo2, dataLocationTemplateRowStart2 + i);
				LockUnlockCol24Col25_PL05_1(winFpSpread.Sheets[1], dataLocationTemplateRowStart2 + i);
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
			int[] array = GeneralFunctions.ArrayCellId("Y_7");
			if (iRow == array[1] + 1)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[1].Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo, dataLocationTemplateRowStart + i);
					TinhChiTieu26_27_PL05_1(winFpSpread.Sheets[1], dataLocationTemplateRowStart + i);
				}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
			int num = HTKKFile.ColumnLettersToIndex("I");
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("W");
			int num7 = HTKKFile.ColumnLettersToIndex("AA");
			int num8 = HTKKFile.ColumnLettersToIndex("AC");
			int num9 = HTKKFile.ColumnLettersToIndex("AK");
			int num10 = HTKKFile.ColumnLettersToIndex("BJ");
			int num11 = HTKKFile.ColumnLettersToIndex("AG");
			string text = sheetView.Cells[(int)iRow - 1, colCt08_PL_05_1 - 1].Text.Trim();
			int tableRowIndex = GetTableRowIndex(dataLocationTemplateRowStart, (int)iRow);
			if (iCol == colCt07_PL_05_1)
			{
				SetValueCt17_TK(winFpSpread);
			}
			if (iCol == colCt08_PL_05_1)
			{
				int soNguoiPhuThuoc = 0;
				double num12 = Cot16VaTongSoTienNPT_PL_05_1(text, ToKhaiInfo.SheetInfos[3].Tables[0], ref soNguoiPhuThuoc);
				sheetView.Cells[(int)iRow - 1, colCt16_PL_05_1 - 1].Value = ((!string.IsNullOrEmpty(text)) ? soNguoiPhuThuoc : 0);
				sheetView.Cells[(int)iRow - 1, colTongGiaTriNPT_PL_05_1 - 1].Value = (string.IsNullOrEmpty(text) ? 0.0 : num12);
			}
			if (iCol == colCt08_PL_05_1)
			{
				SetValueCt17_PL_05_1(winFpSpread, ToKhaiInfo, (int)iRow);
			}
			if (iCol == colCt08_PL_05_1 || iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8)
			{
				sheetView.Cells[(int)iRow - 1, num9 - 1].Value = sheetView.Cells[(int)iRow - 1, num10 - 1].Value;
			}
			if (iCol == colCt08_PL_05_1 || iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num6 || iCol == num7 || iCol == num8 || iCol == num11 || iCol == num9)
			{
				sheetView.Cells[(int)iRow - 1, colCt25_PL_05_1 - 1].Value = sheetView.Cells[(int)iRow - 1, colCt25Formualar_PL_05_1 - 1].Value;
				LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo, (int)iRow);
				LockUnlockCol24Col25_PL05_1(sheetView, (int)iRow);
				TinhChiTieu26_27_PL05_1(sheetView, (int)iRow);
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int tableRowIndex = GetTableRowIndex(dataLocationTemplateRowStart, activeRowIndex + 1);
			GeneralFunctions.UpdateFromSViewToSInfo(sheetView, activeRowIndex + 1, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
			if (activeRowIndex == 2 && activeColumnIndex == 3)
			{
				ShowBangKe("Bang_Ke_05_1_QTT_TT80.xls");
			}
			int num = HTKKFile.ColumnLettersToIndex("I");
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("W");
			int num7 = HTKKFile.ColumnLettersToIndex("AA");
			int num8 = HTKKFile.ColumnLettersToIndex("AK");
			int num9 = HTKKFile.ColumnLettersToIndex("BJ");
			int num10 = HTKKFile.ColumnLettersToIndex("AG");
			if (activeColumnIndex == num - 1 || activeColumnIndex == num2 - 1)
			{
				sheetView.Cells[activeRowIndex, num8 - 1].Value = sheetView.Cells[activeRowIndex, num9 - 1].Value;
				sheetView.Cells[activeRowIndex, colCt25_PL_05_1 - 1].Value = sheetView.Cells[activeRowIndex, colCt25Formualar_PL_05_1 - 1].Value;
				LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo, activeRowIndex + 1);
				LockUnlockCol24Col25_PL05_1(sheetView, activeRowIndex + 1);
				TinhChiTieu26_27_PL05_1(sheetView, activeRowIndex + 1);
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
			int num = HTKKFile.ColumnLettersToIndex("M");
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int num3 = HTKKFile.ColumnLettersToIndex("AR");
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[2].Tables[0];
			GeneralFunctions.UpdateFromSViewToSInfo(sView, (int)iRow, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
			if (iCol == num)
			{
				winFpSpread.Sheets[2].Cells[(int)iRow - 1, num2 - 1].Value = winFpSpread.Sheets[2].Cells[(int)iRow - 1, num3 - 1].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			if (activeRowIndex == 2 && activeColumnIndex == 2)
			{
				ShowBangKe("Bang_Ke_05_3_QTT_TT80.xls");
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
			int colCheck = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int num3 = HTKKFile.ColumnLettersToIndex("U");
			int num4 = HTKKFile.ColumnLettersToIndex("BF");
			SheetView sheetView = winFpSpread.Sheets[3];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[3].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			string text = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
			int num5 = TinhSoThang(PeriodInfo.TuThang, PeriodInfo.DenThang);
			GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
			if (iCol == num2 || iCol == num3)
			{
				TinhCot15VaCot16VaSoTienNPTTuPL_05_3(winFpSpread, ToKhaiInfo, _PeriodInfo, (int)iRow);
			}
			if ((iCol == colCt08_PL_05_3 || iCol == num2 || iCol == num3) && !string.IsNullOrEmpty(text))
			{
				int dataLocationTemplateRowStart2 = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart;
				for (int i = 0; i < 2; i++)
				{
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[0], colCheck, text);
					if (i == 1 && !string.IsNullOrEmpty(_valueMstOld_PL01_3))
					{
						text = _valueMstOld_PL01_3;
						rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[0], colCheck, text);
					}
					if (i != 0 && (i != 1 || string.IsNullOrEmpty(_valueMstOld_PL01_3)))
					{
						continue;
					}
					foreach (int item in rowIndexInSViewByCompare)
					{
						int tableRowIndex = GetTableRowIndex(dataLocationTemplateRowStart2, item);
						int soNguoiPhuThuoc = 0;
						double num6 = Cot16VaTongSoTienNPT_PL_05_1(text, ToKhaiInfo.SheetInfos[3].Tables[0], ref soNguoiPhuThuoc);
						sheetView2.Cells[item - 1, colCt16_PL_05_1 - 1].Value = ((num6 != 0.0) ? soNguoiPhuThuoc : 0);
						sheetView2.Cells[item - 1, colTongGiaTriNPT_PL_05_1 - 1].Value = ((num6 == 0.0) ? 0.0 : num6);
						SetValueCt17_PL_05_1(winFpSpread, ToKhaiInfo, item);
						double num7 = Convert.ToDouble(sheetView2.Cells[item - 1, colBieuThucLuyTien_PL_05_1 - 1].Value);
						sheetView2.Cells[item - 1, colCt24_PL_05_1 - 1].Value = sheetView2.Cells[item - 1, colCt24Formualar_PL_05_1 - 1].Value;
						sheetView2.Cells[item - 1, colCt25_PL_05_1 - 1].Value = sheetView2.Cells[item - 1, colCt25Formualar_PL_05_1 - 1].Value;
						LockUnLockCt13_23_05_1(winFpSpread, ToKhaiInfo.SheetInfos[1].Tables[0], item);
						LockUnlockCol24Col25_PL05_1(winFpSpread.Sheets[1], item);
						TinhChiTieu26_27_PL05_1(winFpSpread.Sheets[1], item);
					}
				}
				SetValueCt17_TK(winFpSpread);
			}
			if (iCol == colCt08_PL_05_3 && _valueMstOld_PL01_3 != sheetView.Cells[(int)iRow - 1, colCt08_PL_05_3 - 1].Text.Trim())
			{
				_valueMstOld_PL01_3 = sheetView.Cells[(int)iRow - 1, colCt08_PL_05_3 - 1].Text.Trim();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("E");
			if (iCol == num)
			{
				_valueMstOld_PL01_3 = winFpSpread.Sheets[3].Cells[(int)iRow - 1, num - 1].Text.Trim();
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
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
			{
				result = false;
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			StringBuilder stringBuilder = new StringBuilder();
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_39"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52"));
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt07_PL_05_1);
			int num = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt08_PL_05_1);
			int index2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt09a_PL_05_1);
			int num2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt09a_PL_05_1);
			int index3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt09_PL_05_1);
			int index4 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt10_PL_05_1);
			int index5 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt11_PL_05_1);
			int index6 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt12_PL_05_1);
			int index7 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt13_PL_05_1);
			int index8 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt14_PL_05_1);
			int index9 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt15_PL_05_1);
			int index10 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt17_PL_05_1);
			int index11 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt20_PL_05_1);
			int index12 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt22_PL_05_1);
			int index13 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt23_PL_05_1);
			int index14 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "nguoiVNSongNN_NguoiNN");
			for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt07_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt09a_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt09_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt10_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt11_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt12_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt13_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt14_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt15_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt17_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt20_PL_05_1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt23_PL_05_1 - 1);
				string value2 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index];
				string text = hTKKCMCTableInfo.RowsValues[num3].RowValue[num];
				bool flag = hTKKCMCTableInfo.RowsValues[num3].RowValue[index14].ToBoolOrDefault();
				if (num3 == 0 && string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text))
				{
					continue;
				}
				batBuocNhapHoChieuCccd(hTKKCMCTableInfo, sheetView, dataLocationTemplateRowStart, num3, colCt09_PL_05_1, 1, ref value);
				string value3 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index2];
				string value4 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index3];
				bool flag2 = ((hTKKCMCTableInfo.RowsValues[num3].RowValue[index4] == "1" || hTKKCMCTableInfo.RowsValues[num3].RowValue[index4] == "True") ? true : false);
				bool flag3 = ((hTKKCMCTableInfo.RowsValues[num3].RowValue[index5] == "1" || hTKKCMCTableInfo.RowsValues[num3].RowValue[index5] == "True") ? true : false);
				if (string.IsNullOrEmpty(text) && (flag2 || flag3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_1 - 1, message2, num3 + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text))
				{
					List<HTKKCMCRowValue> rowIndexInTableByCompare = GeneralFunctions.GetRowIndexInTableByCompare(hTKKCMCTableInfo, num, text);
					if (rowIndexInTableByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_1 - 1, message3, num3 + 1);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text) && (!ValidateBase.CheckMaSoThueCCCD(text) || (flag && text.Length != 10) || (!flag && text.Length != 12)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_1 - 1, message, num3 + 1);
					value = 2;
				}
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt07_PL_05_1 - 1, message2, num3 + 1);
					value = 2;
				}
				if ((string.IsNullOrEmpty(text) && string.IsNullOrEmpty(value3)) || (!string.IsNullOrEmpty(value4) && string.IsNullOrEmpty(value3)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt09a_PL_05_1 - 1, message4, num3 + 1);
					value = 2;
				}
				if ((string.IsNullOrEmpty(text) && string.IsNullOrEmpty(value4)) || (!string.IsNullOrEmpty(value3) && string.IsNullOrEmpty(value4)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt09_PL_05_1 - 1, message2, num3 + 1);
					value = 2;
				}
				double num4 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index6])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index6]) : 0.0);
				double num5 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index7])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index7]) : 0.0);
				if (num5 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt13_PL_05_1 - 1, message5, num3 + 1);
					value = 2;
				}
				double num6 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index8])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index8]) : 0.0);
				if (num6 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt14_PL_05_1 - 1, message6, num3 + 1);
					value = 2;
				}
				double num7 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index9])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index9]) : 0.0);
				if (num7 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt15_PL_05_1 - 1, message7, num3 + 1);
					value = 2;
				}
				double num8 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index11])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index11]) : 0.0);
				if (num8 > 12000000.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt20_PL_05_1 - 1, message8, num3 + 1);
					value = 2;
				}
				double num9 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index12])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index12]) : 0.0);
				double num10 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index13])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index13]) : 0.0);
				if (num10 > num9)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt23_PL_05_1 - 1, message9, num3 + 1);
					value = 2;
				}
				double num11 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index10])) ? Convert.ToDouble(hTKKCMCTableInfo.RowsValues[num3].RowValue[index10]) : 0.0);
				double num12 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + num3 - 1, colCt17Formualar_PL_05_1 - 1].Value);
				if (num11 > num12)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt17_PL_05_1 - 1, message10, num3 + 1);
					value = 2;
				}
			}
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
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_46"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51"));
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt07_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt08_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt09_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt09a_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt12_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt13_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt14_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt15_PL_05_2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, colCt16_PL_05_2 - 1);
				int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct07");
				int num = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct08");
				int index2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct09");
				int index3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct09a_ten");
				int index4 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct09a_ma");
				int index5 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct11");
				int index6 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct12");
				int index7 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct13");
				int index8 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct14");
				int index9 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct15");
				int index10 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct16");
				int index11 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "nguoiVNSongNN_NguoiNN");
				string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[index];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[num];
				string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[index2];
				string value4 = hTKKCMCTableInfo.RowsValues[i].RowValue[index3];
				string value5 = hTKKCMCTableInfo.RowsValues[i].RowValue[index4];
				bool flag = hTKKCMCTableInfo.RowsValues[i].RowValue[index11].ToBoolOrDefault();
				double num2 = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[i].RowValue[index5]) ? 0.0 : Convert.ToDouble(hTKKCMCTableInfo.RowsValues[i].RowValue[index5]));
				double num3 = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[i].RowValue[index6]) ? 0.0 : Convert.ToDouble(hTKKCMCTableInfo.RowsValues[i].RowValue[index6]));
				double num4 = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[i].RowValue[index7]) ? 0.0 : Convert.ToDouble(hTKKCMCTableInfo.RowsValues[i].RowValue[index7]));
				double num5 = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[i].RowValue[index8]) ? 0.0 : Convert.ToDouble(hTKKCMCTableInfo.RowsValues[i].RowValue[index8]));
				double num6 = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[i].RowValue[index9]) ? 0.0 : Convert.ToDouble(hTKKCMCTableInfo.RowsValues[i].RowValue[index9]));
				double num7 = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[i].RowValue[index10]) ? 0.0 : Convert.ToDouble(hTKKCMCTableInfo.RowsValues[i].RowValue[index10]));
				if (i == 0 && string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text))
				{
					continue;
				}
				batBuocNhapHoChieuCccd(hTKKCMCTableInfo, sheetView, dataLocationTemplateRowStart, i, colCt09_PL_05_2, 2, ref value);
				if (string.IsNullOrEmpty(value2) && (!string.IsNullOrEmpty(text) || !string.IsNullOrEmpty(value3) || !string.IsNullOrEmpty(value4)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt07_PL_05_2 - 1, message2, i + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text))
				{
					List<HTKKCMCRowValue> rowIndexInTableByCompare = GeneralFunctions.GetRowIndexInTableByCompare(hTKKCMCTableInfo, num, text);
					if (rowIndexInTableByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt08_PL_05_2 - 1, message3, i + 1);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text) && (!ValidateBase.CheckMaSoThueCCCD(text) || (flag && text.Length != 10) || (!flag && text.Length != 12)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt08_PL_05_2 - 1, message, i + 1);
					value = 2;
				}
				if ((!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text) && string.IsNullOrEmpty(value5)) || (!string.IsNullOrEmpty(value2) && !string.IsNullOrEmpty(value3) && string.IsNullOrEmpty(value5)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt09a_PL_05_2 - 1, message4, i + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text) && string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt09_PL_05_2 - 1, message2, i + 1);
					value = 2;
				}
				if (num3 > num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt12_PL_05_2 - 1, message5, i + 1);
					value = 2;
				}
				if (num4 > num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt13_PL_05_2 - 1, message6, i + 1);
					value = 2;
				}
				if (num5 > num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt14_PL_05_2 - 1, message7, i + 1);
					value = 2;
				}
				if (num7 > num6)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, colCt16_PL_05_2 - 1, message8, i + 1);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(3, out value);
			SheetView sheetView = winFpSpread.Sheets[3];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55"));
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt7_PL_05_3);
			int index2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt08_PL_05_3);
			int index3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt9_PL_05_3);
			int index4 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt10_PL_05_3);
			int index5 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt11_PL_05_3);
			int num = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt12_PL_05_3);
			int index6 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt12Ma_PL_05_3);
			int index7 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt13_PL_05_3);
			int num2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt14_PL_05_3);
			int index8 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt14Ma_PL_05_3);
			int index9 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt15_PL_05_3);
			int index10 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt15ThoaMan_PL_05_3);
			int index11 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt16_PL_05_3);
			int index12 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt16ThoaMan_PL_05_3);
			int index13 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "nguoiVNSongNN_NguoiNN");
			for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
			{
				string[] array = hTKKCMCTableInfo.RowsValues[num3].RowValue.ToArray();
				array[0] = "";
				string value2 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index];
				string text = hTKKCMCTableInfo.RowsValues[num3].RowValue[index2];
				bool flag = hTKKCMCTableInfo.RowsValues[num3].RowValue[index13].ToBoolOrDefault();
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt7_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt9_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt10_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt11_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt12_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt13_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt14_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt15_PL_05_3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt16_PL_05_3 - 1);
				if (num3 == 0 && string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text))
				{
					continue;
				}
				batBuocNhapHoChieuCccd(hTKKCMCTableInfo, sheetView, dataLocationTemplateRowStart, num3, colCt13_PL_05_3, 3, ref value);
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThueCCCD(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt08_PL_05_3 - 1, message, num3 + 1);
					value = 2;
				}
				string text2 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index5];
				if (!string.IsNullOrEmpty(text2) && (!ValidateBase.CheckMaSoThueCCCD(text2) || (flag && text2.Length != 10) || (!flag && text2.Length != 12)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt11_PL_05_3 - 1, message, num3 + 1);
					value = 2;
				}
				string text3 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index9];
				int num4 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index10])) ? Convert.ToInt32(hTKKCMCTableInfo.RowsValues[num3].RowValue[index10]) : 0);
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt15_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				else if (num4 > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt15_PL_05_3 - 1, message4, num3 + 1);
					value = 2;
				}
				string text4 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index11];
				int num5 = ((!string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num3].RowValue[index12])) ? Convert.ToInt32(hTKKCMCTableInfo.RowsValues[num3].RowValue[index12]) : 0);
				if (string.IsNullOrEmpty(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt16_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				else if (num5 > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt16_PL_05_3 - 1, message4, num3 + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text4))
				{
					DateTime t = ConvertHelpers.ConvertToDateTime2(text3, "MM/yyyy");
					DateTime t2 = ConvertHelpers.ConvertToDateTime2(text4, "MM/yyyy");
					if (DateTime.Compare(t, t2) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt16_PL_05_3 - 1, message5, num3 + 1);
						value = 2;
					}
				}
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt7_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				string value3 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index3];
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt9_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				string text5 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index4];
				DateTime dtDate;
				if (string.IsNullOrEmpty(text5))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt10_PL_05_3 - 1, message6);
					if (value != 2)
					{
						value = 1;
					}
				}
				else if (!HTKKFormater.IsDateTime(GeneralFunctions.FormartDateStr(text5), "dd/MM/yyyy", out dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt10_PL_05_3 - 1, message7, num3 + 1);
					value = 2;
				}
				string value4 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index8];
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt14_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				string value5 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index7];
				string value6 = hTKKCMCTableInfo.RowsValues[num3].RowValue[index6];
				if ((string.IsNullOrEmpty(value6) && string.IsNullOrEmpty(text2)) || (!string.IsNullOrEmpty(value5) && string.IsNullOrEmpty(value6)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt12_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num3 - 1, colCt13_PL_05_3 - 1, message2, num3 + 1);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[3].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[3].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[3].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void batBuocNhapHoChieuCccd(HTKKCMCTableInfo table, SheetView sView, int rowStart, int row, int colSoGiayTo, int sheetIndex, ref int iwarningColor)
	{
		string text = "Hộ chiếu";
		string text2 = "Thẻ CCCD/Số định danh cá nhân";
		string text3 = "Giấy khai sinh";
		string text4 = "Khác";
		string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_56"));
		string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57"));
		string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23"));
		string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_58"));
		int num = HTKKFile.ColumnLettersToIndex("E");
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "nguoiVNSongNN_NguoiNN");
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "DC_128");
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo3 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "DC_129");
		if (hTKKCMCTableCellInfo == null)
		{
			return;
		}
		int num2 = rowStart + row;
		HTKKFile.RemoveValidate(sView, num2 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1);
		if (hTKKCMCTableCellInfo2 != null)
		{
			HTKKFile.RemoveValidate(sView, num2 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1);
		}
		if (hTKKCMCTableCellInfo3 != null)
		{
			HTKKFile.RemoveValidate(sView, num2 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1);
		}
		HTKKFile.RemoveValidate(sView, num2 - 1, colSoGiayTo - 1);
		object obj = null;
		object value = sView.Cells[num2 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value;
		if (hTKKCMCTableCellInfo2 != null)
		{
			obj = sView.Cells[num2 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Value;
		}
		if (hTKKCMCTableCellInfo3 != null)
		{
			obj = sView.Cells[num2 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1].Value;
		}
		object value2 = sView.Cells[num2 - 1, colSoGiayTo - 1].Value;
		object value3 = sView.Cells[num2 - 1, num - 1].Value;
		if ((sheetIndex == 1 || sheetIndex == 2 || sheetIndex == 3) && (obj == null || obj.ToString().Trim() == "") && (value2 == null || value2.ToString().Trim() == "") && value3 != null && value3.ToString() != "")
		{
			if (sheetIndex == 3)
			{
				HTKKFile.SetNodeBackGroundYellow(sView, num2 - 1, colSoGiayTo - 1, "Nhập số giấy tờ");
				if (hTKKCMCTableCellInfo2 != null)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num2 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1, "Nhập Loại giấy tờ");
				}
				if (hTKKCMCTableCellInfo3 != null)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num2 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1, "Nhập Loại giấy tờ");
				}
				if (iwarningColor != 2)
				{
					iwarningColor = 1;
				}
			}
			return;
		}
		if (Convert.ToBoolean(value))
		{
			if (obj == null || (obj.ToString().Trim().ToLower() != text.ToLower() && obj.ToString().Trim().ToLower() != text3.ToLower() && obj.ToString().Trim().ToLower() != text4.ToLower()))
			{
				if (hTKKCMCTableCellInfo2 != null)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num2 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1, message, row + 1);
				}
				if (hTKKCMCTableCellInfo3 != null)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num2 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1, message4, row + 1);
				}
				iwarningColor = 2;
			}
			if (obj != null && (obj.ToString().Trim().ToLower() == text.ToLower() || obj.ToString().Trim().ToLower() == text3.ToLower() || obj.ToString().Trim().ToLower() == text4.ToLower()) && (value2 == null || string.IsNullOrEmpty(value2.ToString())))
			{
				HTKKFile.SetNodeBackGroundPink(sView, num2 - 1, colSoGiayTo - 1, message3, row + 1);
				iwarningColor = 2;
			}
			return;
		}
		if (obj == null || obj.ToString().ToLower() != text2.ToLower())
		{
			if (hTKKCMCTableCellInfo2 != null)
			{
				HTKKFile.SetNodeBackGroundPink(sView, num2 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1, message2, row + 1);
			}
			if (hTKKCMCTableCellInfo3 != null)
			{
				HTKKFile.SetNodeBackGroundPink(sView, num2 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1, message2, row + 1);
			}
			iwarningColor = 2;
		}
		if (obj != null && obj.ToString().Trim().ToLower() == text2.ToLower() && (value2 == null || string.IsNullOrEmpty(value2.ToString())))
		{
			HTKKFile.SetNodeBackGroundPink(sView, num2 - 1, colSoGiayTo - 1, message3, row + 1);
			iwarningColor = 2;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			if (activeRowIndex == 2 && activeColumnIndex == 3)
			{
				ShowBangKe("Bang_Ke_05_2_QTT_TT80.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.ActiveSheet.AutoCalculation = false;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView activeSheet = winFpSpread.ActiveSheet;
			if ((activeSheetIndex == 1 || activeSheetIndex == 3) && ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count > 1)
			{
				int dataLocationTemplateRowStart2 = ToKhaiInfo.SheetInfos[3].Tables[0].Location.DataLocationTemplateRowStart;
				for (int i = 0; i < ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count; i++)
				{
					TinhCot15VaCot16VaSoTienNPTTuPL_05_3(winFpSpread, ToKhaiInfo, PeriodInfo, dataLocationTemplateRowStart2 + i);
				}
			}
			if ((activeSheetIndex == 1 || activeSheetIndex == 3) && ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count > 1)
			{
				TinhLaiCot16VaTongSoTienNPTPhuLuc05_1TuPhuLuc05_3(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[0], ToKhaiInfo.SheetInfos[3].Tables[0], isShowOnSview: true);
			}
			if (activeSheetIndex == 1)
			{
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				int i = 0;
				foreach (HTKKCMCRowValue rowsValue in hTKKCMCTableInfo.RowsValues)
				{
					try
					{
						bool flag = Convert.ToBoolean(activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt10_PL_05_1 - 1].Value);
						bool flag2 = Convert.ToBoolean(activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt11_PL_05_1 - 1].Value);
						double num = Convert.ToDouble(activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt22_PL_05_1 - 1].Value);
						double num2 = Convert.ToDouble(activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt24Formualar_PL_05_1 - 1].Value);
						double num3 = 0.0;
						if ((flag || flag2) && num2 - num < 0.0)
						{
							num3 = Math.Round((num2 - num) * -1.0, 0);
						}
						activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt24_PL_05_1 - 1].Value = num2;
						activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt25_PL_05_1 - 1].Value = num3;
						LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo, dataLocationTemplateRowStart + i);
						LockUnlockCol24Col25_PL05_1(activeSheet, dataLocationTemplateRowStart + i);
						TinhChiTieu26_27_PL05_1(activeSheet, dataLocationTemplateRowStart + i);
					}
					catch (Exception ex)
					{
						string fnName = "rowStart" + dataLocationTemplateRowStart + "colCt25Formualar_PL_05_1" + colCt25Formualar_PL_05_1 + "_r_" + i + "Formula" + activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colCt25Formualar_PL_05_1 - 1].Formula;
						HTKKFile.Log(fnName, ex);
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
					i++;
				}
			}
			if (activeSheetIndex == 2)
			{
			}
			SetValueCt17_TK(winFpSpread);
			winFpSpread.ActiveSheet.AutoCalculation = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetValueCt17_TK(FpSpread winFpSpread)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCt17_TK[1], cellCt17_TK[0]].Value = winFpSpread.Sheets[1].Cells[cellCt17Formualar_TK[1], cellCt17Formualar_TK[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetValueCt17_PL_05_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int row)
	{
		try
		{
			if (row <= 0)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					winFpSpread.Sheets[1].Cells[dataLocationTemplateRowStart + i - 1, colCt17_PL_05_1 - 1].Value = winFpSpread.Sheets[1].Cells[dataLocationTemplateRowStart + i - 1, colCt17Formualar_PL_05_1 - 1].Value;
					LockUnLockCot27_PL05_1(winFpSpread.Sheets[1], dataLocationTemplateRowStart + i);
				}
			}
			else
			{
				winFpSpread.Sheets[1].Cells[row - 1, colCt17_PL_05_1 - 1].Value = winFpSpread.Sheets[1].Cells[row - 1, colCt17Formualar_PL_05_1 - 1].Value;
				LockUnLockCot27_PL05_1(winFpSpread.Sheets[1], row);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiCot16VaTongSoTienNPTPhuLuc05_1TuPhuLuc05_3(SheetView sView_PL_05_1, HTKKCMCTableInfo table_PL05_1, HTKKCMCTableInfo table_PL05_3, bool isShowOnSview)
	{
		try
		{
			HTKKCMCRowInfo rowsInfo = table_PL05_1.RowsInfo;
			HTKKCMCRowInfo rowsInfo2 = table_PL05_3.RowsInfo;
			int index = table_PL05_1.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct08");
			int index2 = table_PL05_1.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct16");
			int index3 = table_PL05_1.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_26");
			int dataLocationTemplateRowStart = table_PL05_1.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("BF");
			int num2 = 0;
			foreach (HTKKCMCRowValue rowsValue in table_PL05_1.RowsValues)
			{
				string text = rowsValue.RowValue[index];
				int num3 = dataLocationTemplateRowStart + num2;
				if (!string.IsNullOrEmpty(text))
				{
					int soNguoiPhuThuoc = 0;
					double num4 = Cot16VaTongSoTienNPT_PL_05_1(text, table_PL05_3, ref soNguoiPhuThuoc);
					rowsValue.RowValue[index2] = soNguoiPhuThuoc.ToString();
					rowsValue.RowValue[index3] = num4.ToString();
					if (isShowOnSview)
					{
						sView_PL_05_1.Cells[num3 - 1, colCt16_PL_05_1 - 1].Value = soNguoiPhuThuoc;
						sView_PL_05_1.Cells[num3 - 1, num - 1].Value = num4;
					}
				}
				num2++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public double Cot16VaTongSoTienNPT_PL_05_1(string mst_PL05_1, HTKKCMCTableInfo table_PL05_3, ref int soNguoiPhuThuoc)
	{
		try
		{
			int indexCt08_PL05_3 = table_PL05_3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct08");
			int index = table_PL05_3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_16");
			IEnumerable<HTKKCMCRowValue> enumerable = table_PL05_3.RowsValues.Where((HTKKCMCRowValue x) => x.RowValue[indexCt08_PL05_3] == mst_PL05_1);
			double num = 0.0;
			foreach (HTKKCMCRowValue item in enumerable)
			{
				int num2 = ((!string.IsNullOrEmpty(item.RowValue[index])) ? Convert.ToInt32(item.RowValue[index]) : 0);
				num += (double)(num2 * 4400000);
			}
			soNguoiPhuThuoc = enumerable.Count();
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return 0.0;
	}

	public int TinhSoThang(string tuThang, string denThang)
	{
		try
		{
			if (string.IsNullOrEmpty(tuThang) || string.IsNullOrEmpty(denThang))
			{
				return 0;
			}
			return (ConvertHelpers.ConvertToDateTime2(denThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime2(tuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime2(denThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime2(tuThang, "MM/yyyy").Month + 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public void TinhCot15VaCot16VaSoTienNPTTuPL_05_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, int row)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[3].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int tableRowIndex = GetTableRowIndex(dataLocationTemplateRowStart, row);
			HTKKCMCRowValue hTKKCMCRowValue = hTKKCMCTableInfo.RowsValues[tableRowIndex];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int num3 = HTKKFile.ColumnLettersToIndex("U");
			int num4 = HTKKFile.ColumnLettersToIndex("AK");
			int num5 = HTKKFile.ColumnLettersToIndex("AG");
			int num6 = HTKKFile.ColumnLettersToIndex("AH");
			int num7 = HTKKFile.ColumnLettersToIndex("AL");
			DateTime t = ConvertHelpers.ConvertToDateTime2("01/" + PeriodInfo.TuThang, "dd/MM/yyyy");
			DateTime t2 = ConvertHelpers.ConvertToDateTime2("01/" + PeriodInfo.DenThang, "dd/MM/yyyy");
			string text = sheetView.Cells[row - 1, num2 - 1].Text.Trim();
			string text2 = sheetView.Cells[row - 1, num3 - 1].Text.Trim();
			DateTime dateTime = ConvertHelpers.ConvertToDateTime2("01/" + text, "dd/MM/yyyy");
			DateTime dateTime2 = ConvertHelpers.ConvertToDateTime2("01/" + text2, "dd/MM/yyyy");
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
			{
				int num8 = TinhSoThang(text, text2);
				sheetView.Cells[row - 1, num4 - 1].Value = num8;
				sheetView.Cells[row - 1, num7 - 1].Value = num8 * 4400000;
				int num9 = ((DateTime.Compare(dateTime, t) < 0 || DateTime.Compare(t2, dateTime) < 0) ? 1 : 0);
				sheetView.Cells[row - 1, num5 - 1].Value = num9;
				int num10 = ((DateTime.Compare(dateTime2, t) < 0 || DateTime.Compare(t2, dateTime2) < 0) ? 1 : 0);
				sheetView.Cells[row - 1, num6 - 1].Value = num10;
				int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_14");
				int index2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_15");
				int index3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_16");
				hTKKCMCRowValue.RowValue[index2] = num9.ToString();
				hTKKCMCRowValue.RowValue[index] = num10.ToString();
				hTKKCMCRowValue.RowValue[index3] = num8.ToString();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnLockCot27_PL05_1(SheetView sView, int row)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AO");
			int colHide = HTKKFile.ColumnLettersToIndex("AQ");
			double num2 = Convert.ToDouble(sView.Cells[row - 1, num - 1].Value);
			if (num2 < 50000.0 && num2 > 0.0)
			{
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, locked: false, deleteData: false, false);
			}
			else
			{
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, locked: true, deleteData: true, false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockCol24Col25_PL05_1(SheetView sView, int row)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("I");
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int colHide = HTKKFile.ColumnLettersToIndex("AK");
			int colHide2 = HTKKFile.ColumnLettersToIndex("AM");
			bool flag = Convert.ToBoolean(sView.Cells[row - 1, num - 1].Value);
			bool flag2 = Convert.ToBoolean(sView.Cells[row - 1, num2 - 1].Value);
			if (flag && !flag2)
			{
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, locked: true, deleteData: false, 0);
				GeneralFunctions.LockUnLockCol2(sView, row, colHide2, locked: true, deleteData: false, 0);
			}
			else if ((!flag && flag2) || (flag && flag2))
			{
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, locked: false, deleteData: false, 0);
				GeneralFunctions.LockUnLockCol2(sView, row, colHide2, locked: false, deleteData: false, 0);
			}
			else if (!flag && !flag2)
			{
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, locked: true, deleteData: true, 0);
				GeneralFunctions.LockUnLockCol2(sView, row, colHide2, locked: true, deleteData: true, 0);
			}
			else
			{
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, locked: true, deleteData: true, 0);
				GeneralFunctions.LockUnLockCol2(sView, row, colHide2, locked: true, deleteData: true, 0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		UpdateToKhaiVaPL05_1(winFpSpread, ToKhaiInfo);
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TinhLaiCot16VaTongSoTienNPTPhuLuc05_1TuPhuLuc05_3(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[0], ToKhaiInfo.SheetInfos[3].Tables[0], isShowOnSview: true);
		UpdateToKhaiVaPL05_1(winFpSpread, ToKhaiInfo);
		return true;
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		UpdateToKhaiVaPL05_1(winFpSpread, _ToKhaiInfo);
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
	}

	public bool UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		return true;
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType)
	{
	}

	public void ClearData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		UpdateToKhaiVaPL05_1(winFpSpread, ToKhaiInfo);
	}

	public bool UpdateFormula_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TinhLaiCot16VaTongSoTienNPTPhuLuc05_1TuPhuLuc05_3(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[0], ToKhaiInfo.SheetInfos[3].Tables[0], isShowOnSview: true);
		UpdateToKhaiVaPL05_1(winFpSpread, ToKhaiInfo);
		SheetView sView = winFpSpread.Sheets[1];
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			LockUnLockCt13_23_05_1(winFpSpread, hTKKCMCTableInfo, dataLocationTemplateRowStart + i);
			LockUnlockCol24Col25_PL05_1(sView, dataLocationTemplateRowStart + i);
			TinhChiTieu26_27_PL05_1(sView, dataLocationTemplateRowStart + i);
		}
		return true;
	}

	public void Refresh_3(FpSpread winFpSpread, int actionType)
	{
		UpdateToKhaiVaPL05_1(winFpSpread, _ToKhaiInfo);
	}

	public void UpdateToKhaiVaPL05_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[3].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				TinhCot15VaCot16VaSoTienNPTTuPL_05_3(winFpSpread, ToKhaiInfo, _PeriodInfo, dataLocationTemplateRowStart + i);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[1].Tables[0];
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = TinhSoThang(_PeriodInfo.TuThang, _PeriodInfo.DenThang);
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				int soNguoiPhuThuoc = 0;
				string text = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colCt08_PL_05_1 - 1].Text.Trim();
				double num2 = Cot16VaTongSoTienNPT_PL_05_1(text, ToKhaiInfo.SheetInfos[3].Tables[0], ref soNguoiPhuThuoc);
				sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colCt16_PL_05_1 - 1].Value = ((!string.IsNullOrEmpty(text)) ? soNguoiPhuThuoc : 0);
				sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colTongGiaTriNPT_PL_05_1 - 1].Value = (string.IsNullOrEmpty(text) ? 0.0 : num2);
				SetValueCt17_PL_05_1(winFpSpread, ToKhaiInfo, dataLocationTemplateRowStart2 + i);
				double num3 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colBieuThucLuyTien_PL_05_1 - 1].Value);
				sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colCt24_PL_05_1 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colCt24Formualar_PL_05_1 - 1].Value;
				sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colCt25_PL_05_1 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, colCt25Formualar_PL_05_1 - 1].Value;
				LockUnlockCol24Col25_PL05_1(winFpSpread.Sheets[1], dataLocationTemplateRowStart2 + i);
			}
			SetValueCt17_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int GetTableRowIndex(int startRow, int iRow)
	{
		return (iRow - startRow >= 0) ? (iRow - startRow) : (-1);
	}

	public void LockUnLockCt13_23_05_1(FpSpread winFpSpread, HTKKCMCTableInfo table, int row)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("Y_7");
			int colHide = HTKKFile.ColumnLettersToIndex("O");
			int colHide2 = HTKKFile.ColumnLettersToIndex("AI");
			int num = HTKKFile.ColumnLettersToIndex("AO");
			int num2 = HTKKFile.ColumnLettersToIndex("AQ");
			int num3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct26");
			bool flag = !Convert.ToBoolean(sheetView.Cells[array[1], array[0]].Value);
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			if (row < 0)
			{
				for (int num4 = 0; num4 < table.RowsValues.Count; num4++)
				{
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num4, colHide, flag, flag, 0);
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num4, colHide2, flag, flag, 0);
				}
			}
			else
			{
				int tableRowIndex = GetTableRowIndex(dataLocationTemplateRowStart, row);
				HTKKCMCRowValue hTKKCMCRowValue = table.RowsValues[tableRowIndex];
				GeneralFunctions.LockUnLockCol2(sView, row, colHide, flag, flag, 0);
				GeneralFunctions.LockUnLockCol2(sView, row, colHide2, flag, flag, 0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.ActiveSheet.AutoCalculation = false;
			int num = 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView sheetView = winFpSpread.Sheets[num];
			int dataLocationTemplateRowStart2 = ToKhaiInfo.SheetInfos[3].Tables[0].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count; i++)
			{
				TinhCot15VaCot16VaSoTienNPTTuPL_05_3(winFpSpread, ToKhaiInfo, PeriodInfo, dataLocationTemplateRowStart2 + i);
			}
			if (num == 1 && ToKhaiInfo.SheetInfos[3] != null && ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count > 1)
			{
				TinhLaiCot16VaTongSoTienNPTPhuLuc05_1TuPhuLuc05_3(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[0], ToKhaiInfo.SheetInfos[3].Tables[0], isShowOnSview: true);
			}
			for (int j = 0; j < ToKhaiInfo.SheetInfos.Count - 1; j++)
			{
				GeneralFunctions.fpSpReindexingSheetTemplate2(winFpSpread.Sheets[j], ToKhaiInfo.SheetInfos[j]);
			}
			winFpSpread.ActiveSheet.AutoCalculation = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhChiTieu26_27_PL05_1(SheetView sView, int row)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AG");
			int num2 = HTKKFile.ColumnLettersToIndex("AK");
			int num3 = HTKKFile.ColumnLettersToIndex("AO");
			int num4 = HTKKFile.ColumnLettersToIndex("AQ");
			double num5 = Convert.ToDouble(sView.Cells[row - 1, num - 1].Value);
			double num6 = Convert.ToDouble(sView.Cells[row - 1, num2 - 1].Value);
			long num7 = 0L;
			if (num6 - num5 > 0.0)
			{
				num7 = Convert.ToInt64(num6 - num5);
			}
			sView.Cells[row - 1, num3 - 1].Value = num7;
			bool flag = ((num7 <= 0 || num7 > 50000) ? true : false);
			bool flag2 = num7 > 0 && num7 <= 50000 && Convert.ToBoolean(sView.Cells[row - 1, num4 - 1].Value);
			GeneralFunctions.LockUnLockCol2(sView, row, num4, flag, flag, flag2);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
