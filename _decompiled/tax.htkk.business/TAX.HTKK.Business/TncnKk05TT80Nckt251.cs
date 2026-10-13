using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnKk05TT80Nckt251
{
	private int[] cellTemPhuLuc = GeneralFunctions.ArrayCellId("AF_5");

	private int[] cellTemTruSoChinh = GeneralFunctions.ArrayCellId("AG_5");

	private string maTSC = "01";

	private string maTinhTSC = "";

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			sView.Cells[cellTemTruSoChinh[1], cellTemTruSoChinh[0]].Value = ((PeriodInfo.maCQTCapCuc == headerInfo.maCQTCapCuc) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		SheetView sheetView = winFpSpread.Sheets[1];
		int num = HTKKFile.ColumnLettersToIndex("C");
		int num2 = HTKKFile.ColumnLettersToIndex("G");
		int num3 = HTKKFile.ColumnLettersToIndex("R");
		int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("I");
		int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("T");
		int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("J");
		int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("U");
		if (iCol == num2)
		{
			string maTinh = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
			GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc, isLock: false);
		}
		if (iCol == num)
		{
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("F");
		int num2 = HTKKFile.ColumnLettersToIndex("J");
		int num3 = HTKKFile.ColumnLettersToIndex("C");
		int num4 = HTKKFile.ColumnLettersToIndex("V");
		int num5 = HTKKFile.ColumnLettersToIndex("R");
		int num6 = HTKKFile.ColumnLettersToIndex("G");
		SheetView sheetView = winFpSpread.Sheets[1];
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
		string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
		string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40"));
		string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41"));
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
			string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				if (text.Length == 5 && !ValidateBase.CheckMaSoThue5So(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
					value = 2;
				}
				else if (text.Length != 5 && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
					value = 2;
				}
			}
			string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text2))
			{
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num2, text2);
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, message2);
					value = 2;
				}
			}
			string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
			if (text3.Equals(maTSC))
			{
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num4, text3);
				maTinhTSC = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Text.Trim();
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, message3);
					value = 2;
				}
			}
		}
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string text4 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Text.Trim();
			string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
			if (!text3.Equals(maTSC) && text4.Equals(maTinhTSC))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, message4);
				value = 2;
			}
		}
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

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc[1], cellTemPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			GeneralFunctions.checkLockUnLockCells(_ToKhaiInfo, winFpSpread);
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
			GeneralFunctions.checkLockUnLockCells(_ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
