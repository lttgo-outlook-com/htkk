using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnQtt02TT80Nckt251
{
	private int[] cellChenhLechSoThang = GeneralFunctions.ArrayCellId("W_17");

	private int[] cellTenQuyetToan = GeneralFunctions.ArrayCellId("F_11");

	private int[] cellMaQuyetToan = GeneralFunctions.ArrayCellId("AA_12");

	private int[] cellChuyenKhoan = GeneralFunctions.ArrayCellId("D_149");

	private int[] cellTienMat = GeneralFunctions.ArrayCellId("D_153");

	private int[] cell34 = GeneralFunctions.ArrayCellId("M_109");

	private int[] cell41 = GeneralFunctions.ArrayCellId("M_116");

	private int[] cell44 = GeneralFunctions.ArrayCellId("M_124");

	private int[] cell46 = GeneralFunctions.ArrayCellId("M_126");

	private int[] cell47 = GeneralFunctions.ArrayCellId("M_127");

	private int col5 = HTKKFile.ColumnLettersToIndex("J");

	private int col7 = HTKKFile.ColumnLettersToIndex("M");

	private int col8_1 = HTKKFile.ColumnLettersToIndex("O");

	private int col8_1Ma = HTKKFile.ColumnLettersToIndex("AC");

	private int col9 = HTKKFile.ColumnLettersToIndex("Q");

	private int col9Ma = HTKKFile.ColumnLettersToIndex("AE");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[cellTenQuyetToan[1], cellTenQuyetToan[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[cellMaQuyetToan[1], cellMaQuyetToan[0]].Value = PeriodInfo.MaQuyetToan;
			}
			LuonKhoaCot5(sView, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = TinhSoThang(PeriodInfo.TuThang, PeriodInfo.DenThang);
			sheetView.Cells[cellChenhLechSoThang[1], cellChenhLechSoThang[0]].Value = num;
			TinhCot16VaCot17TuPL_02_1(winFpSpread, ToKhaiInfo, PeriodInfo, -1);
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			if ((int)iRow == cellChuyenKhoan[1] + 1 && (int)iCol == cellChuyenKhoan[0] + 1 && Convert.ToBoolean(sheetView.Cells[cellChuyenKhoan[1], cellChuyenKhoan[0]].Value))
			{
				sheetView.Cells[cellTienMat[1], cellTienMat[0]].Value = false;
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if ((int)iRow == cellTienMat[1] + 1 && (int)iCol == cellTienMat[0] + 1 && Convert.ToBoolean(sheetView.Cells[cellTienMat[1], cellTienMat[0]].Value))
			{
				sheetView.Cells[cellChuyenKhoan[1], cellChuyenKhoan[0]].Value = false;
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iCol == cell44[0] + 1 && Convert.ToDouble(sheetView.Cells[cell44[1], cell44[0]].Value) == 0.0)
			{
				sheetView.Cells[cell46[1], cell46[0]].Value = 0;
				sheetView.Cells[cell47[1], cell47[0]].Value = 0;
			}
			if (iRow == cell46[1] + 1 || iRow == cell47[1] + 1 || (iRow >= cell34[1] + 1 && iRow <= cell41[1] + 1))
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				LuonKhoaCot5(sheetView, ToKhaiInfo);
			}
			int num = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
			if (iCol == col7 && num >= 0)
			{
				sheetView.Cells[(int)iRow - 1, col5 - 1].Value = sheetView.Cells[(int)iRow - 1, col7 - 1].Value;
			}
			if (iCol == col8_1)
			{
				string maCQTCapCuc = sheetView.Cells[(int)iRow - 1, col8_1Ma - 1].Text.Trim();
				ThayDoiTinhTheoCQT(maCQTCapCuc, sheetView, (int)iRow, col9, col9Ma);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			if (iCol == num || iCol == num2 || iCol == num3)
			{
				TinhCot16VaCot17TuPL_02_1(winFpSpread, ToKhaiInfo, PeriodInfo, (int)iRow);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL2_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
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

	private void checkPL2_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
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
			int num = HTKKFile.ColumnLettersToIndex("E");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49"));
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num, text);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message3);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiTinhTheoCQT(string maCQTCapCuc, SheetView sView, int row, int colTenTinh, int colMaTinh)
	{
		try
		{
			if (maCQTCapCuc == "82500")
			{
				sView.Cells[row - 1, colTenTinh - 1].Locked = false;
			}
			else
			{
				if (string.IsNullOrEmpty(maCQTCapCuc))
				{
					return;
				}
				sView.Cells[row - 1, colTenTinh - 1].Value = "";
				sView.Cells[row - 1, colTenTinh - 1].Locked = false;
				sView.Cells[row - 1, colMaTinh - 1].Value = "";
				string maTinh = (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsValue(maCQTCapCuc.Substring(0, 3)) ? maCQTCapCuc.Substring(0, 3) : "");
				if (!string.IsNullOrEmpty(maCQTCapCuc))
				{
					KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
					sView.Cells[row - 1, colTenTinh - 1].Value = keyValuePair.Key;
					sView.Cells[row - 1, colMaTinh - 1].Value = maTinh;
					sView.Cells[row - 1, colTenTinh - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LuonKhoaCot5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[ToKhaiInfo.SheetInfos[0].Tables.Count - 1];
			int num = HTKKFile.ColumnLettersToIndex("J");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
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

	public void TinhCot16VaCot17TuPL_02_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, int row)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("O");
			int num2 = HTKKFile.ColumnLettersToIndex("Q");
			int num3 = HTKKFile.ColumnLettersToIndex("AB");
			int num4 = HTKKFile.ColumnLettersToIndex("X");
			int num5 = HTKKFile.ColumnLettersToIndex("Y");
			DateTime t = ConvertHelpers.ConvertToDateTime2("01/" + PeriodInfo.TuThang, "dd/MM/yyyy");
			DateTime t2 = ConvertHelpers.ConvertToDateTime2("01/" + PeriodInfo.DenThang, "dd/MM/yyyy");
			if (row < 0)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
					DateTime dateTime = ConvertHelpers.ConvertToDateTime2("01/" + text, "dd/MM/yyyy");
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime2("01/" + text2, "dd/MM/yyyy");
					if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
					{
						int num6 = TinhSoThang(text, text2);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value = num6;
						int num7 = ((DateTime.Compare(dateTime, t) < 0 || DateTime.Compare(t2, dateTime) < 0) ? 1 : 0);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value = num7;
						int num8 = ((DateTime.Compare(dateTime2, t) < 0 || DateTime.Compare(t2, dateTime2) < 0) ? 1 : 0);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Value = num8;
					}
				}
			}
			else
			{
				string text3 = sheetView.Cells[row - 1, num - 1].Text.Trim();
				string text = sheetView.Cells[row - 1, num2 - 1].Text.Trim();
				DateTime dateTime = ConvertHelpers.ConvertToDateTime2("01/" + text3, "dd/MM/yyyy");
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime2("01/" + text, "dd/MM/yyyy");
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text))
				{
					int num6 = TinhSoThang(text3, text);
					sheetView.Cells[row - 1, num3 - 1].Value = num6;
					int num7 = ((DateTime.Compare(dateTime, t) < 0 || DateTime.Compare(t2, dateTime) < 0) ? 1 : 0);
					sheetView.Cells[row - 1, num4 - 1].Value = num7;
					int num8 = ((DateTime.Compare(dateTime2, t) < 0 || DateTime.Compare(t2, dateTime2) < 0) ? 1 : 0);
					sheetView.Cells[row - 1, num5 - 1].Value = num8;
				}
				else
				{
					sheetView.Cells[row - 1, num3 - 1].Value = 0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count == 1)
			{
				TinhCot16VaCot17TuPL_02_1(winFpSpread, ToKhaiInfo, PeriodInfo, 38);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AB");
			SheetView sheetView = winFpSpread.Sheets[1];
			sheetView.Cells[37, num - 1].Value = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
