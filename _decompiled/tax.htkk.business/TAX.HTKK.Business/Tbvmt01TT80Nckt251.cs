using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tbvmt01TT80Nckt251
{
	public List<CacheFormula> formulas = new List<CacheFormula>();

	private Dictionary<string, string[]> dicDMThueBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueBVMT4 = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMThueBVMTTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMXangDau = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMXangDauTmp = new Dictionary<string, string>();

	private Dictionary<string, double> dicToKhaiHangHoaCol4 = new Dictionary<string, double>();

	private Dictionary<string, double> dicToKhaiHangHoaCol7 = new Dictionary<string, double>();

	private Common common = new Common();

	private string MA_XANG_RON92 = "020206";

	private int[] cellTemPhuLuc1 = GeneralFunctions.ArrayCellId("AE_9");

	private int[] cellTemPhuLuc2 = GeneralFunctions.ArrayCellId("AF_9");

	private int[] cellCt11d = GeneralFunctions.ArrayCellId("H_56");

	private int[] cellCt11dMa = GeneralFunctions.ArrayCellId("AC_56");

	private int[] cellCt11c = GeneralFunctions.ArrayCellId("H_58");

	private int[] cellCt11cMa = GeneralFunctions.ArrayCellId("AC_58");

	private int[] cellCt11b = GeneralFunctions.ArrayCellId("H_60");

	private int[] cellCt11bMa = GeneralFunctions.ArrayCellId("AC_60");

	private int[] cellCt11a = GeneralFunctions.ArrayCellId("H_62");

	private int[] cellCt09 = GeneralFunctions.ArrayCellId("H_45");

	private int[] cellCt10 = GeneralFunctions.ArrayCellId("H_47");

	private string LOAI_MA_CHINH = "01";

	private string LOAI_MA_CHI_NHANH = "02";

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private HTKKCMCTTinChungInfo _TTinChung = null;

	public Dictionary<string, List<double>> _dicMucThue = new Dictionary<string, List<double>>();

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_O => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_R => HTKKFile.ColumnLettersToIndex("R") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_AC => HTKKFile.ColumnLettersToIndex("AC") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			_TTinChung = TTinChung;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			string text = "";
			string text2 = "#010108#0106#020109#020120#020121#020119#020112#020205#010102#020206#";
			if (PeriodInfo.Thang == "07" && PeriodInfo.Nam == "2022" && _dicMucThue.Count == 0)
			{
				_dicMucThue.Add("020201", new List<double> { 2000.0, 1000.0 });
				_dicMucThue.Add("020202", new List<double> { 1000.0, 500.0 });
				_dicMucThue.Add("020203", new List<double> { 300.0, 300.0 });
				_dicMucThue.Add("020205", new List<double> { 1500.0, 1000.0 });
				_dicMucThue.Add("020206", new List<double> { 2000.0, 1000.0 });
				_dicMucThue.Add("020207", new List<double> { 1000.0, 300.0 });
				_dicMucThue.Add("020208", new List<double> { 1000.0, 300.0 });
				_dicMucThue.Add("020209", new List<double> { 1000.0, 300.0 });
			}
			int num = HTKKFile.ColumnLettersToIndex("AD");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sView.Cells[dataLocationTemplateRowStart + i - 1, idxCol_AC].Text;
				if (!string.IsNullOrEmpty(text) && string.IsNullOrEmpty(sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text))
				{
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(text, "BVMT");
					sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text = maTieuMucByMaHH;
				}
			}
			LuuToanBoHangHoaToKhai(sView, hTKKCMCTableInfo);
			sView.Cells[cellTemPhuLuc1[1], cellTemPhuLuc1[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			sView.Cells[cellTemPhuLuc2[1], cellTemPhuLuc2[0]].Value = ((ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, cellCt11d, cellCt11dMa, cellCt11c, cellCt11cMa, cellCt11b, cellCt11bMa, cellCt11a, cellCt09, cellCt10);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			GeneralFunctions.fpSpReindexingSheetTemplate(sView, hTKKCMCSheetInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			LockUnLockPl01_2(sView, hTKKCMCSheetInfo, TTinChung, -1);
			GeneralFunctions.fpSpReindexingSheetTemplate(sView, hTKKCMCSheetInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
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
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("L");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("AD");
			int num9 = HTKKFile.ColumnLettersToIndex("AC");
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + rowsValues.Count && (int)iCol == num)
			{
				string text = sheetView.Cells[(int)iRow - 1, num9 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text) && text != MA_XANG_RON92)
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = 100;
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num5, locked: true, deleteData: false);
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = 0;
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num5, locked: false, deleteData: false);
				}
			}
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + rowsValues.Count && (int)iCol == num)
			{
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim()];
				if (array.Length == 4)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[2];
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = array[3];
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[0], "BVMT");
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = maTieuMucByMaHH;
				}
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iCol == num || iCol == num4 || iCol == num5 || iCol == num6)
			{
				LuuToanBoHangHoaToKhai(sheetView, hTKKCMCTableInfo);
			}
			if (ToKhaiInfo.SheetInfos[2] == null)
			{
				return;
			}
			TinhCot10Pl_01_2(winFpSpread, ToKhaiInfo.SheetInfos[2], -1);
			TinhCot11Pl_01_2(winFpSpread, ToKhaiInfo.SheetInfos[2]);
			int fromColHeaderKey = HTKKFile.ColumnLettersToIndex("P");
			int num10 = HTKKFile.ColumnLettersToIndex("AB");
			int toColHeaderKey = HTKKFile.ColumnLettersToIndex("G");
			int num11 = HTKKFile.ColumnLettersToIndex("G");
			List<int> list = new List<int>();
			int item = HTKKFile.ColumnLettersToIndex("AD");
			int item2 = HTKKFile.ColumnLettersToIndex("AE");
			int num12 = HTKKFile.ColumnLettersToIndex("AF");
			list.Add(item);
			list.Add(item2);
			list.Add(num12);
			List<int> list2 = new List<int>();
			list2.Add(num12);
			GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangF7SangBangF7(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], "01_2_TBVMT_1", "01_2_TBVMT_2", fromColHeaderKey, num10, toColHeaderKey, num11, list2);
			List<HTKKCMCTableInfo> list3 = new List<HTKKCMCTableInfo>();
			foreach (HTKKCMCTableInfo table in ToKhaiInfo.SheetInfos[2].Tables)
			{
				if (table.TableName == "01_2_TBVMT_2")
				{
					list3.Add(table);
				}
			}
			GeneralFunctions.CalculateSumByGroup(winFpSpread.Sheets[2], list3, num12);
			GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], "01_2_TBVMT_1", "01_2_TBVMT_3", num10, num11, list);
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
			int[] array = GeneralFunctions.ArrayCellId("S_50");
			int[] array2 = GeneralFunctions.ArrayCellId("X_50");
			int[] cellTenCQTCuc = GeneralFunctions.ArrayCellId("Z_50");
			int[] cellMaCQTCuc = GeneralFunctions.ArrayCellId("AB_50");
			int[] cellTenCQTChiCuc = GeneralFunctions.ArrayCellId("AA_50");
			int[] cellMaCQTChiCuc = GeneralFunctions.ArrayCellId("AC_50");
			int[] array3 = GeneralFunctions.ArrayCellId("AE_50");
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			int num4 = HTKKFile.ColumnLettersToIndex("U");
			int colHeader = HTKKFile.ColumnLettersToIndex("W");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("X");
			int num7 = HTKKFile.ColumnLettersToIndex("Z");
			int num8 = HTKKFile.ColumnLettersToIndex("AB");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AC");
			int num11 = HTKKFile.ColumnLettersToIndex("AE");
			int colCheck = HTKKFile.ColumnLettersToIndex("AD");
			int num12 = HTKKFile.ColumnLettersToIndex("AF");
			int num13 = HTKKFile.ColumnLettersToIndex("AG");
			int num14 = HTKKFile.ColumnLettersToIndex("AH");
			int num15 = HTKKFile.ColumnLettersToIndex("AI");
			int num16 = HTKKFile.ColumnLettersToIndex("AJ");
			int num17 = HTKKFile.ColumnLettersToIndex("AK");
			int num18 = HTKKFile.ColumnLettersToIndex("AL");
			int num19 = HTKKFile.ColumnLettersToIndex("AN");
			int num20 = HTKKFile.ColumnLettersToIndex("AO");
			int num21 = HTKKFile.ColumnLettersToIndex("AP");
			int num22 = HTKKFile.ColumnLettersToIndex("AR");
			int num23 = HTKKFile.ColumnLettersToIndex("AQ");
			int colChilden = HTKKFile.ColumnLettersToIndex("AS");
			int colChilden2 = HTKKFile.ColumnLettersToIndex("AT");
			int colChilden3 = HTKKFile.ColumnLettersToIndex("AU");
			int colChilden4 = HTKKFile.ColumnLettersToIndex("AV");
			int colChilden5 = HTKKFile.ColumnLettersToIndex("AW");
			int colChilden6 = HTKKFile.ColumnLettersToIndex("AX");
			int colChilden7 = HTKKFile.ColumnLettersToIndex("AY");
			int colChilden8 = HTKKFile.ColumnLettersToIndex("AZ");
			int colChilden9 = HTKKFile.ColumnLettersToIndex("BA");
			int colChilden10 = HTKKFile.ColumnLettersToIndex("BB");
			int num24 = HTKKFile.ColumnLettersToIndex("E");
			int num25 = HTKKFile.ColumnLettersToIndex("G");
			int num26 = HTKKFile.ColumnLettersToIndex("E");
			int num27 = HTKKFile.ColumnLettersToIndex("G");
			int num28 = HTKKFile.ColumnLettersToIndex("F");
			int num29 = HTKKFile.ColumnLettersToIndex("E");
			int num30 = HTKKFile.ColumnLettersToIndex("G");
			int num31 = HTKKFile.ColumnLettersToIndex("AJ");
			int num32 = HTKKFile.ColumnLettersToIndex("AL");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num33 = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (iCol == num11)
			{
				object value = sheetView.Cells[(int)iRow - 1, num11 - 1].Value;
				if (value != null)
				{
					string[] array4 = dicDMThueBVMT[value.ToString().Trim()];
					if (array4.Length == 4)
					{
						sheetView.Cells[(int)iRow - 1, num17 - 1].Text = array4[3];
						sheetView.Cells[(int)iRow - 1, num19 - 1].Text = array4[3];
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array4[0], "BVMT");
						sheetView.Cells[(int)iRow - 1, num22 - 1].Text = maTieuMucByMaHH;
						DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
						string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(maTieuMucByMaHH, ngayDauKyKeKhai);
						sheetView.Cells[(int)iRow - 1, num23 - 1].Text = tenTmByMaTm;
					}
					GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				}
			}
			if (iCol == num11)
			{
				object value = sheetView.Cells[(int)iRow - 1, num11 - 1].Value;
				if (value != null)
				{
					string[] array4 = dicDMThueBVMT[value.ToString().Trim()];
					if (array4.Length == 4)
					{
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array4[0], "BVMT");
						sheetView.Cells[(int)iRow - 1, num22 - 1].Text = maTieuMucByMaHH;
						DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
						string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(maTieuMucByMaHH, ngayDauKyKeKhai);
						sheetView.Cells[(int)iRow - 1, num23 - 1].Text = tenTmByMaTm;
					}
				}
			}
			if (iCol == num11)
			{
				int index = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[index], colCheck, "");
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num11 - 1, text);
				}
			}
			if (iCol == num11 || iCol == num12 || iCol == num13)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, num11 - 1].Text.Trim();
				List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num11, text2);
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					if (table.TableName != "01_1_TBVMT_2")
					{
						continue;
					}
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table.RowsValues.Count; i++)
					{
						string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Text.Trim();
						if (text3 == text2)
						{
							if (rowIndexInSViewByCompare2.Count > 0)
							{
								sheetView.Cells[dataLocationTemplateRowStart + i - 1, num20 - 1].Value = sheetView.Cells[rowIndexInSViewByCompare2[0] - 1, num12 - 1].Value;
							}
							else
							{
								sheetView.Cells[(int)iRow - 1, num20 - 1].Value = 0;
							}
						}
					}
				}
			}
			if (iCol == num11 || iCol == num12 || iCol == num13)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, num11 - 1].Text.Trim();
				List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num11, text2);
				foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo.Tables)
				{
					if (table2.TableName != "01_1_TBVMT_2")
					{
						continue;
					}
					int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table2.RowsValues.Count; i++)
					{
						string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Text.Trim();
						if (text3 == text2)
						{
							if (rowIndexInSViewByCompare2.Count > 0)
							{
								sheetView.Cells[dataLocationTemplateRowStart + i - 1, num21 - 1].Value = sheetView.Cells[rowIndexInSViewByCompare2[0] - 1, num13 - 1].Value;
							}
							else
							{
								sheetView.Cells[(int)iRow - 1, num21 - 1].Value = 0;
							}
						}
					}
				}
			}
			if (iCol == num5 && iRow == array[1] + 1)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhStatic(sheetView, maTinh, cellTenCQTCuc, cellMaCQTCuc, cellTenCQTChiCuc, cellMaCQTChiCuc, isLock: false);
			}
			if (iCol == num5 && num33 >= 0 && hTKKCMCSheetInfo.Tables[num33].TableName == "01_1_TBVMT_2")
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, num7, num8, num9, num10, isLock: false);
			}
			if (num33 >= 0 && (iCol == num2 || iCol == num3 || iCol == num4 || iCol == num5 || iCol == num7 || iCol == num9))
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num3, colChilden);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num5, colChilden2);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num6, colChilden3);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num4, colChilden4);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, colHeader, colChilden5);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num7, colChilden6);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num8, colChilden7);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num9, colChilden8);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num10, colChilden9);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num33, -1, num2, colChilden10);
			}
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			foreach (HTKKCMCTableInfo table3 in hTKKCMCSheetInfo.Tables)
			{
				if (table3.TableName == "01_1_TBVMT_2")
				{
					list.Add(table3);
				}
			}
			GeneralFunctions.CalculateSumByGroup(sheetView, list, num18);
			if (iCol == num5 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num15)
			{
				TinhTongTheoTinhPl01_1(winFpSpread, ToKhaiInfo);
			}
			if (iCol == num11 || iCol == num12 || iCol == num13 || iCol == num15 || iCol == num18)
			{
				TinhBangTongCongPl01_1(winFpSpread, ToKhaiInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ComboCloseUp_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
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
			SheetView sheetView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("AB_49");
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("AI");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("N");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int colHeader = HTKKFile.ColumnLettersToIndex("T");
			int num7 = HTKKFile.ColumnLettersToIndex("P");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("W");
			int num10 = HTKKFile.ColumnLettersToIndex("Y");
			int num11 = HTKKFile.ColumnLettersToIndex("X");
			int num12 = HTKKFile.ColumnLettersToIndex("Z");
			int num13 = HTKKFile.ColumnLettersToIndex("AB");
			int num14 = HTKKFile.ColumnLettersToIndex("AA");
			int num15 = HTKKFile.ColumnLettersToIndex("AC");
			int num16 = HTKKFile.ColumnLettersToIndex("AD");
			int num17 = HTKKFile.ColumnLettersToIndex("AE");
			int num18 = HTKKFile.ColumnLettersToIndex("AF");
			int num19 = HTKKFile.ColumnLettersToIndex("AH");
			int colChilden = HTKKFile.ColumnLettersToIndex("AL");
			int colChilden2 = HTKKFile.ColumnLettersToIndex("AM");
			int colChilden3 = HTKKFile.ColumnLettersToIndex("AN");
			int colChilden4 = HTKKFile.ColumnLettersToIndex("AO");
			int colChilden5 = HTKKFile.ColumnLettersToIndex("AP");
			int colChilden6 = HTKKFile.ColumnLettersToIndex("AQ");
			int colChilden7 = HTKKFile.ColumnLettersToIndex("AR");
			int colChilden8 = HTKKFile.ColumnLettersToIndex("AS");
			int colChilden9 = HTKKFile.ColumnLettersToIndex("AT");
			int colChilden10 = HTKKFile.ColumnLettersToIndex("AU");
			int num20 = HTKKFile.ColumnLettersToIndex("AK");
			int num21 = HTKKFile.ColumnLettersToIndex("AJ");
			int num22 = HTKKFile.ColumnLettersToIndex("E");
			int toColHeaderKey = HTKKFile.ColumnLettersToIndex("G");
			int num23 = HTKKFile.ColumnLettersToIndex("E");
			int toColBodyKey = HTKKFile.ColumnLettersToIndex("G");
			int num24 = HTKKFile.ColumnLettersToIndex("F");
			int num25 = HTKKFile.ColumnLettersToIndex("E");
			int toColKey = HTKKFile.ColumnLettersToIndex("G");
			int num26 = HTKKFile.ColumnLettersToIndex("AD");
			int num27 = HTKKFile.ColumnLettersToIndex("AF");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int count = hTKKCMCSheetInfo.Tables.Count;
			int num28 = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (iCol == array[0] + 1)
			{
				string[] array2 = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, num13 - 1].Text.Trim()];
				if (array2.Length == 4)
				{
					sheetView.Cells[(int)iRow - 1, num15 - 1].Text = array2[2];
					sheetView.Cells[(int)iRow - 1, num17 - 1].Text = array2[3];
					sheetView.Cells[(int)iRow - 1, num19 - 1].Text = array2[3];
				}
			}
			if (iCol == num13)
			{
				object value = sheetView.Cells[(int)iRow - 1, num13 - 1].Value;
				if (value != null)
				{
					string[] array2 = dicDMThueBVMT[value.ToString().Trim()];
					if (array2.Length == 4)
					{
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array2[0], "BVMT");
						sheetView.Cells[(int)iRow - 1, num20 - 1].Text = maTieuMucByMaHH;
						DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
						string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(maTieuMucByMaHH, ngayDauKyKeKhai);
						sheetView.Cells[(int)iRow - 1, num21 - 1].Text = tenTmByMaTm;
					}
				}
			}
			if (iCol == num13)
			{
				int index = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[index], num14, "");
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num13 - 1, text);
				}
			}
			if (iCol == num7)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num8 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, num9, num10, num11, num12, isLock: false);
			}
			if (iCol == num2)
			{
				LockUnLockPl01_2(sheetView, hTKKCMCSheetInfo, TTinChung, (int)iRow);
			}
			if (num28 >= 0 && hTKKCMCSheetInfo.Tables[num28].TableName == "01_2_TBVMT_1" && (iCol == num2 || iCol == num4 || iCol == num5 || iCol == num6 || iCol == num7 || iCol == num9 || iCol == num11 || iCol == num13))
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num5, colChilden);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num7, colChilden2);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num8, colChilden3);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num6, colChilden4);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, colHeader, colChilden5);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num9, colChilden6);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num10, colChilden7);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num11, colChilden8);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num12, colChilden9);
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, num28, -1, num4, colChilden10);
			}
			if (iCol == num13 || iCol == num16)
			{
				TinhCot10Pl_01_2(winFpSpread, hTKKCMCSheetInfo, (int)iRow);
			}
			if (iCol == num13 || iCol == num16 || iCol == num17)
			{
				TinhCot11Pl_01_2(winFpSpread, hTKKCMCSheetInfo);
			}
			if (iCol == num7 || iCol == num13)
			{
				TinhTongTheoTinhPl01_2(winFpSpread, ToKhaiInfo);
			}
			if (iCol == num13)
			{
				TinhBangTongCongPl01_2(winFpSpread, ToKhaiInfo);
			}
			if (iCol == num16 || iCol == num17)
			{
				List<int> list = new List<int>();
				list.Add(num16);
				list.Add(num17);
				list.Add(num18);
				List<int> list2 = new List<int>();
				list2.Add(num18);
				GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangF7SangBangF7(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], "01_2_TBVMT_1", "01_2_TBVMT_2", num7, num13, toColHeaderKey, toColBodyKey, list2);
				GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(sheetView, hTKKCMCSheetInfo, "01_2_TBVMT_1", "01_2_TBVMT_3", num13, toColKey, list);
			}
			if (iCol == num2 || iCol == num7 || iCol == num13 || iCol == num16 || iCol == num17)
			{
				List<HTKKCMCTableInfo> list3 = new List<HTKKCMCTableInfo>();
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					if (table.TableName == "01_2_TBVMT_2")
					{
						list3.Add(table);
					}
				}
				GeneralFunctions.CalculateSumByGroup(sheetView, list3, num18);
			}
			Dictionary<string, double> dictionary = GeneralFunctions.SumValueInTableByKey(sheetView, hTKKCMCSheetInfo, "01_2_TBVMT_1", num14, num17);
			foreach (KeyValuePair<string, double> item in dictionary)
			{
				foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo.Tables)
				{
					if (table2.TableName == "01_2_TBVMT_3")
					{
						int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
						for (int i = 0; i < table2.RowsValues.Count; i++)
						{
							string key = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num25 - 1].Text.Trim();
							sheetView.Cells[dataLocationTemplateRowStart + i - 1, num17 - 1].Value = (dictionary.ContainsKey(key) ? dictionary[key] : 0.0);
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

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, ref levelWarning);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, ref levelWarning);
		Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, ref levelWarning);
		if (winFpSpread.Sheets[0].FpSpread.TabStrip[1].BackColor == Color.Pink)
		{
			return false;
		}
		if (winFpSpread.Sheets[1].FpSpread.TabStrip[1].BackColor == Color.Pink)
		{
			return false;
		}
		if (winFpSpread.Sheets[2].FpSpread.TabStrip[1].BackColor == Color.Pink)
		{
			return false;
		}
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, ref Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AC");
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51"));
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cellCt11d, cellCt11dMa, cellCt11c, cellCt11cMa, cellCt11b, cellCt11bMa, cellCt11a, cellCt09, cellCt10))
			{
				value = 2;
			}
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_C);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_L);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_O);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_R);
				array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, idxCol_C].Value);
				array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, idxCol_J].Value);
				array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, idxCol_L].Value);
				array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, idxCol_O].Value);
				array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, idxCol_R].Value);
				array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, idxCol_V].Value);
				string key = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + i, num - 1].Value);
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					continue;
				}
				if (!string.IsNullOrEmpty(array[0]))
				{
					List<int> list = new List<int>();
					list.Add(idxCol_C + 1);
					list.Add(idxCol_R + 1);
					List<string> list2 = new List<string>();
					list2.Add(array[0]);
					list2.Add(array[4]);
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
					if (rowIndexInSViewByCompares.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_C, message5);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message5);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(array[0]) && ConvertHelpers.ConvertToDouble(array[2]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, num2 + i, idxCol_L, message);
					value = 1;
				}
				if (!string.IsNullOrEmpty(array[0]) && ConvertHelpers.ConvertToDouble(array[3]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, num2 + i, idxCol_O, message);
					value = 1;
				}
				if (string.IsNullOrEmpty(array[0]) || !(PeriodInfo.Thang == "07") || !(PeriodInfo.Nam == "2022") || _dicMucThue.Count <= 0 || !_dicMucThue.ContainsKey(key) || _dicMucThue[key].Contains(ConvertHelpers.ConvertToDouble(array[4])))
				{
					continue;
				}
				string text = "";
				foreach (double item in _dicMucThue[key])
				{
					double num3 = item;
					text = text + num3 + " hoặc ";
				}
				text = text.Substring(0, text.Length - " hoặc ".Length);
				HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message6 + " " + text);
				value = 2;
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
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, ref Dictionary<int, int> levelWarning)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		SheetView sheetView = winFpSpread.Sheets[1];
		int num = HTKKFile.ColumnLettersToIndex("C");
		int num2 = HTKKFile.ColumnLettersToIndex("G");
		int num3 = HTKKFile.ColumnLettersToIndex("Q");
		int num4 = HTKKFile.ColumnLettersToIndex("U");
		int num5 = HTKKFile.ColumnLettersToIndex("W");
		int num6 = HTKKFile.ColumnLettersToIndex("S");
		int num7 = HTKKFile.ColumnLettersToIndex("X");
		int num8 = HTKKFile.ColumnLettersToIndex("Z");
		int num9 = HTKKFile.ColumnLettersToIndex("AB");
		int num10 = HTKKFile.ColumnLettersToIndex("AA");
		int num11 = HTKKFile.ColumnLettersToIndex("AC");
		int num12 = HTKKFile.ColumnLettersToIndex("AE");
		int num13 = HTKKFile.ColumnLettersToIndex("AD");
		int num14 = HTKKFile.ColumnLettersToIndex("AF");
		int num15 = HTKKFile.ColumnLettersToIndex("AG");
		int num16 = HTKKFile.ColumnLettersToIndex("AH");
		int num17 = HTKKFile.ColumnLettersToIndex("AI");
		int num18 = HTKKFile.ColumnLettersToIndex("AJ");
		int num19 = HTKKFile.ColumnLettersToIndex("AK");
		int num20 = HTKKFile.ColumnLettersToIndex("AL");
		int num21 = HTKKFile.ColumnLettersToIndex("AN");
		int num22 = HTKKFile.ColumnLettersToIndex("D");
		int num23 = HTKKFile.ColumnLettersToIndex("G");
		int num24 = HTKKFile.ColumnLettersToIndex("E");
		int num25 = HTKKFile.ColumnLettersToIndex("G");
		int num26 = HTKKFile.ColumnLettersToIndex("F");
		int num27 = HTKKFile.ColumnLettersToIndex("E");
		int num28 = HTKKFile.ColumnLettersToIndex("G");
		int num29 = HTKKFile.ColumnLettersToIndex("AJ");
		int num30 = HTKKFile.ColumnLettersToIndex("AL");
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0627");
			string message2 = xMLMessage.GetMessage("0628");
			string message3 = xMLMessage.GetMessage("0629");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1);
				string valueCompare = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num13 - 1].Text.Trim();
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0].Tables[0], HTKKFile.ColumnLettersToIndex("AC"), valueCompare);
				List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num13, valueCompare);
				if (rowIndexInSViewByCompare2.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1, text6);
					value = 2;
				}
				else if (rowIndexInSViewByCompare.Count == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1, text13);
					value = 2;
				}
			}
			List<HTKKCMCTableInfo> list = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == "01_1_TBVMT_2").ToList();
			int index = 2;
			foreach (HTKKCMCTableInfo item in list)
			{
				int rowIndex = item.Header.StaticCells[index].CellTemplate.RowIndex;
				int colIndex = item.Header.StaticCells[index].CellTemplate.ColIndex;
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, colIndex - 1);
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num3 - 1);
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, num10 - 1);
				string text15 = sheetView.Cells[rowIndex - 1, colIndex - 1].Text.Trim();
				if (string.IsNullOrEmpty(text15))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, colIndex - 1, text4);
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(text15))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, colIndex - 1, text2);
					value = 2;
				}
				string value2 = sheetView.Cells[rowIndex - 1, num2 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num2 - 1, text4);
					value = 2;
				}
				string value3 = sheetView.Cells[rowIndex - 1, num6 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num6 - 1, text4);
					value = 2;
				}
				string value4 = sheetView.Cells[rowIndex - 1, num4 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num4 - 1, text4);
					value = 2;
				}
				string value5 = sheetView.Cells[rowIndex - 1, num8 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num8 - 1, text4);
					value = 2;
				}
				string value6 = sheetView.Cells[rowIndex - 1, num10 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, num10 - 1, text4);
					value = 2;
				}
				int dataLocationTemplateRowStart2 = item.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < item.RowsValues.Count; i++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + i - 1, num12 - 1);
					string valueCompare = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num13 - 1].Text.Trim();
					string text16 = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num12 - 1].Text.Trim();
					List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num13, valueCompare);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + i - 1, num12 - 1);
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num12, text16);
					if (string.IsNullOrEmpty(valueCompare) || string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + i - 1, num12 - 1, text4);
						value = 2;
					}
					else if (rowIndexInSViewByCompare2.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + i - 1, num12 - 1, text6);
						value = 2;
					}
					else if (rowIndexInSViewByCompare.Count == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + i - 1, num12 - 1, text14);
						value = 2;
					}
				}
			}
			Dictionary<string, double> dictionary = GeneralFunctions.SumValueInTableByKey(sheetView, hTKKCMCSheetInfo, "01_1_TBVMT_1", num13, num17);
			Dictionary<string, double> dictionary2 = GeneralFunctions.SumValueInTableByKey(sheetView, hTKKCMCSheetInfo, "01_1_TBVMT_2", num13, num17);
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				if (!(table.TableName == "01_1_TBVMT_2"))
				{
					continue;
				}
				int dataLocationTemplateRowStart2 = table.Location.DataLocationTemplateRowStart;
				for (int num31 = 0; num31 < table.RowsValues.Count; num31++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + num31 - 1, num17 - 1);
					string key = sheetView.Cells[dataLocationTemplateRowStart2 + num31 - 1, num13 - 1].Text.Trim();
					double value7 = (dictionary.ContainsKey(key) ? dictionary[key] : 0.0);
					double value8 = (dictionary2.ContainsKey(key) ? dictionary2[key] : 0.0);
					if (Math.Round(value7, 6) != Math.Round(value8, 6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num31 - 1, num17 - 1, text10);
						value = 2;
					}
				}
			}
			List<string> list2 = new List<string>();
			list2.Add("ct5_ten");
			list2.Add("ct4_ten");
			list2.Add("path_7");
			list2.Add("ct6_ten");
			List<string> list3 = new List<string>();
			list3.Add("ct5_ma");
			list3.Add("ct4_ma");
			list3.Add("path_8");
			list3.Add("ct6_ma");
			foreach (HTKKCMCTableInfo item2 in list)
			{
				int num32 = GeneralFunctions.ValidateComboboxTableHeader(sheetView, item2, list2, list3, message2, message, message3);
				if (num32 > 0)
				{
					value = 2;
				}
			}
			List<int> list4 = new List<int>();
			list4.Add(num12);
			List<int> list5 = new List<int>();
			list5.Add(num13);
			foreach (HTKKCMCTableInfo item3 in list)
			{
				int num32 = GeneralFunctions.ValidateCombobox(sheetView, item3, list4, list5, message2, message, message3);
				if (num32 > 0)
				{
					value = 2;
				}
			}
			List<string> list6 = new List<string>();
			list6.Add("PLuc_01_1_TBVMT/DonViTieuThu/ct5_ten");
			list6.Add("PLuc_01_1_TBVMT/DonViTieuThu/ct4_ten");
			list6.Add("path_7");
			list6.Add("PLuc_01_1_TBVMT/DonViTieuThu/ct6_ten");
			List<string> list7 = new List<string>();
			list7.Add("PLuc_01_1_TBVMT/DonViTieuThu/ct5_ma");
			list7.Add("PLuc_01_1_TBVMT/DonViTieuThu/ct4_ma");
			list7.Add("path_8");
			list7.Add("PLuc_01_1_TBVMT/DonViTieuThu/ct6_ma");
			int num33 = GeneralFunctions.ValidateComboboxStatic(sheetView, hTKKCMCSheetInfo, list6, list7, message2, message, message3);
			if (num33 > 0)
			{
				value = 2;
			}
			List<int> list8 = new List<int>();
			list8.Add(num12);
			List<int> list9 = new List<int>();
			list9.Add(num13);
			foreach (HTKKCMCTableInfo item4 in list)
			{
				int num32 = GeneralFunctions.ValidateCombobox(sheetView, item4, list8, list9, message2, message, message3);
				if (num32 > 0)
				{
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, ref Dictionary<int, int> levelWarning)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		SheetView sheetView = winFpSpread.Sheets[2];
		int[] array = GeneralFunctions.ArrayCellId("AI_42");
		int[] array2 = GeneralFunctions.ArrayCellId("AI_43");
		int num = HTKKFile.ColumnLettersToIndex("D");
		int num2 = HTKKFile.ColumnLettersToIndex("AI");
		int num3 = HTKKFile.ColumnLettersToIndex("C");
		int num4 = HTKKFile.ColumnLettersToIndex("G");
		int num5 = HTKKFile.ColumnLettersToIndex("N");
		int num6 = HTKKFile.ColumnLettersToIndex("R");
		int num7 = HTKKFile.ColumnLettersToIndex("T");
		int num8 = HTKKFile.ColumnLettersToIndex("P");
		int num9 = HTKKFile.ColumnLettersToIndex("U");
		int num10 = HTKKFile.ColumnLettersToIndex("W");
		int num11 = HTKKFile.ColumnLettersToIndex("Y");
		int num12 = HTKKFile.ColumnLettersToIndex("X");
		int num13 = HTKKFile.ColumnLettersToIndex("Z");
		int num14 = HTKKFile.ColumnLettersToIndex("AB");
		int num15 = HTKKFile.ColumnLettersToIndex("AA");
		int num16 = HTKKFile.ColumnLettersToIndex("AC");
		int num17 = HTKKFile.ColumnLettersToIndex("AD");
		int num18 = HTKKFile.ColumnLettersToIndex("AE");
		int num19 = HTKKFile.ColumnLettersToIndex("AF");
		int num20 = HTKKFile.ColumnLettersToIndex("AH");
		int num21 = HTKKFile.ColumnLettersToIndex("E");
		int value = 0;
		levelWarning.TryGetValue(2, out value);
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75"));
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[45, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message4 = xMLMessage.GetMessage("0627");
			string message5 = xMLMessage.GetMessage("0628");
			string message6 = xMLMessage.GetMessage("0629");
			List<HTKKCMCTableInfo> list = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == "01_2_TBVMT_1").ToList();
			Dictionary<string, double> dictionary = new Dictionary<string, double>();
			Dictionary<string, double> dictionary2 = new Dictionary<string, double>();
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TableName == "01_2_TBVMT_3");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num22 = 0; num22 < hTKKCMCTableInfo.RowsValues.Count; num22++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num22 - 1, num18 - 1);
			}
			foreach (HTKKCMCTableInfo item in list)
			{
				int dataLocationTemplateRowStart2 = item.Location.DataLocationTemplateRowStart;
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 - 2, num - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 - 2, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 - 2, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 - 2, num18 - 1);
				int num23 = Convert.ToInt32(sheetView.Cells[array[1], array[0]].Value);
				int num24 = Convert.ToInt32(sheetView.Cells[array2[1], array2[0]].Value);
				string text16 = ((sheetView.Cells[dataLocationTemplateRowStart2 - 2, num2 - 1].Value != null) ? sheetView.Cells[dataLocationTemplateRowStart2 - 2, num2 - 1].Value.ToString().Trim() : "");
				if (num23 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 2, num - 1, text8);
					value = 2;
				}
				else if (num23 > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 2, num - 1, text9);
					value = 2;
				}
				else if (num24 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 2, num - 1, text7);
					value = 2;
				}
				string text17 = sheetView.Cells[dataLocationTemplateRowStart2 - 2, num5 - 1].Text.Trim();
				if (string.IsNullOrEmpty(text17))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 2, num5 - 1, text4);
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(text17))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 2, num5 - 1, text2);
					value = 2;
				}
				string text18 = sheetView.Cells[dataLocationTemplateRowStart2 - 2, num9 - 1].Text.Trim();
				string text19 = text18 + "00";
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num9, text18);
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 2, num8 - 1, message3);
					value = 2;
				}
				for (int num25 = 0; num25 < item.RowsValues.Count; num25++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num17 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num19 - 1);
					string text20 = ((sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num15 - 1].Value != null) ? sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num15 - 1].Value.ToString().Trim() : "");
					List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0].Tables[0], HTKKFile.ColumnLettersToIndex("AC"), text20);
					List<int> rowIndexInSViewByCompare3 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, item, num15, text20);
					if (rowIndexInSViewByCompare3.Count > 1)
					{
						if (string.IsNullOrEmpty(text20) || !(PeriodInfo.Thang == "07") || !(PeriodInfo.Nam == "2022"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1, text6);
							value = 2;
						}
					}
					else if (rowIndexInSViewByCompare2.Count == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1, text13);
						value = 2;
					}
					double num26 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num17 - 1].Value);
					if (num26 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num17 - 1, text4);
						value = 2;
					}
					double num27 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num19 - 1].Value);
					if (text16 == LOAI_MA_CHINH && !string.IsNullOrEmpty(text16) && !string.IsNullOrEmpty(text20))
					{
						string key = text20;
						if (!dictionary.ContainsKey(key))
						{
							dictionary.Add(key, num27);
						}
						else
						{
							dictionary[key] += num27;
						}
					}
					if (text16 == LOAI_MA_CHI_NHANH && !string.IsNullOrEmpty(text16) && !string.IsNullOrEmpty(text20))
					{
						string key = text20;
						if (!dictionary2.ContainsKey(key))
						{
							dictionary2.Add(key, num27);
						}
						else
						{
							dictionary2[key] += num27;
						}
					}
				}
			}
			foreach (HTKKCMCTableInfo item2 in list)
			{
				int dataLocationTemplateRowStart2 = item2.Location.DataLocationTemplateRowStart;
				string text16 = ((sheetView.Cells[dataLocationTemplateRowStart2 - 2, num2 - 1].Value != null) ? sheetView.Cells[dataLocationTemplateRowStart2 - 2, num2 - 1].Value.ToString().Trim() : "");
				for (int num25 = 0; num25 < item2.RowsValues.Count; num25++)
				{
					string text20 = ((sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num15 - 1].Value != null) ? sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num15 - 1].Value.ToString().Trim() : "");
					string text21 = ((sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num14 - 1].Value != null) ? sheetView.Cells[dataLocationTemplateRowStart2 + num25 - 1, num14 - 1].Value.ToString().Trim() : "");
					List<int> rowIndexInSViewByCompare3 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, item2, num15, text20);
					if (string.IsNullOrEmpty(text20))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1, text4);
						value = 2;
					}
					else if (rowIndexInSViewByCompare3.Count > 1)
					{
						if (string.IsNullOrEmpty(text21) || !(PeriodInfo.Thang == "07") || !(PeriodInfo.Nam == "2022"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1, text6);
							value = 2;
						}
					}
					else if (!string.IsNullOrEmpty(text16) && !string.IsNullOrEmpty(text20) && !(text16 != LOAI_MA_CHI_NHANH))
					{
						string key = text20;
						if (dictionary.ContainsKey(key) && !dictionary2.ContainsKey(key))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1, text11 + " " + text21);
							value = 2;
						}
						else if (!dictionary.ContainsKey(key) && dictionary2.ContainsKey(key))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num25 - 1, num14 - 1, text12 + " " + text21);
							value = 2;
						}
					}
				}
			}
			Dictionary<string, double> dictionary3 = GeneralFunctions.SumValueInTableByKey(sheetView, hTKKCMCSheetInfo, "01_2_TBVMT_1", num15, num18);
			foreach (KeyValuePair<string, double> item3 in dictionary3)
			{
				if (Math.Round(item3.Value, 6) == 100.0)
				{
					continue;
				}
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					if (!(table.TableName == "01_2_TBVMT_3"))
					{
						continue;
					}
					int dataLocationTemplateRowStart2 = table.Location.DataLocationTemplateRowStart;
					for (int num22 = 0; num22 < table.RowsValues.Count; num22++)
					{
						string text22 = sheetView.Cells[dataLocationTemplateRowStart2 + num22 - 1, num21 - 1].Text.Trim();
						if (text22 == item3.Key)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 + num22 - 1, num18 - 1, text15);
							value = 2;
						}
					}
				}
			}
			List<string> list2 = new List<string>();
			list2.Add("ct5_ten");
			list2.Add("ct4_ten");
			list2.Add("path_9");
			list2.Add("ct6_ten");
			List<string> list3 = new List<string>();
			list3.Add("ct5_ma");
			list3.Add("ct4_ma");
			list3.Add("path_10");
			list3.Add("ct6_ma");
			foreach (HTKKCMCTableInfo item4 in list)
			{
				int num28 = GeneralFunctions.ValidateComboboxTableHeader(sheetView, item4, list2, list3, message5, message4, message6);
				if (num28 > 0)
				{
					value = 2;
				}
			}
			List<int> list4 = new List<int>();
			list4.Add(num14);
			List<int> list5 = new List<int>();
			list5.Add(num15);
			foreach (HTKKCMCTableInfo item5 in list)
			{
				int num28 = GeneralFunctions.ValidateCombobox(sheetView, item5, list4, list5, message5, message4, message6);
				if (num28 > 0)
				{
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

	public Dictionary<string, string> GetDMThueBVMT(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Expected O, but got Unknown
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012027#102026#052026#042026#032026#012026#072025#012025#012024#012023#072022#042022#112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMThueBVMTTmp.Count > 0)
			{
				return dicDMThueBVMTTmp;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
				DateTime t2 = new DateTime(year, month, 1).AddMonths(1).AddDays(-1.0);
				for (int i = 0; i < array.Length; i++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = array[i];
					}
				}
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string key = array2[0];
				string value2 = array2[1];
				dictionary2.Add(key, value2);
				dictionary.Add(array2[1], array2);
			}
			dicDMThueBVMT = dictionary;
			dicDMThueBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThan(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Expected O, but got Unknown
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012027#102026#052026#042026#032026#012026#072025#012025#012024#012023#072022#042022#112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMThueBVMT4.Count > 0)
			{
				return dicDMThueBVMT4;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
				DateTime t2 = new DateTime(year, month, 1).AddMonths(1).AddDays(-1.0);
				for (int i = 0; i < array.Length; i++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = array[i];
					}
				}
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string text4 = array2[0];
				string value2 = array2[1];
				if (text4.Substring(0, 4).Equals("0102") || text4.Equals("020118"))
				{
					dictionary2.Add(text4, value2);
				}
			}
			dicDMThueBVMT4 = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMXangDau(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Expected O, but got Unknown
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012027#102026#052026#042026#032026#012026#072025#012025#012024#012023#072022#042022#112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int year = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMXangDau.Count > 0)
			{
				return dicDMXangDau;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
				DateTime t2 = new DateTime(year, month, 1).AddMonths(1).AddDays(-1.0);
				for (int i = 0; i < array.Length; i++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = array[i];
					}
				}
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string text4 = array2[0];
				string value2 = array2[1];
				if (text4.Substring(0, 4).Equals("0202"))
				{
					dictionary2.Add(text4, value2);
				}
			}
			dicDMXangDau = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			string text2 = "#010108#0106#020109#020120#020121#020119#020112#020205#010102#020206#";
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num2 - 1].Text;
				if (text2.Contains("#" + text + "#"))
				{
					sheetView.Cells[num + i, idxCol_R].Locked = false;
				}
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				TinhCot11Pl_01_2(winFpSpread, ToKhaiInfo.SheetInfos[2]);
			}
			TinhTongTheoTinhPl01_1(winFpSpread, ToKhaiInfo);
			TinhBangTongCongPl01_1(winFpSpread, ToKhaiInfo);
			TinhTongTheoTinhPl01_2(winFpSpread, ToKhaiInfo);
			TinhBangTongCongPl01_2(winFpSpread, ToKhaiInfo);
			int colHeader;
			int colHeader2;
			int colHeader3;
			int colHeader4;
			int colHeader5;
			int colHeader6;
			int colHeader7;
			int colHeader8;
			int colHeader9;
			int colHeader10;
			int colChilden;
			int colChilden2;
			int colChilden3;
			int colChilden4;
			int colChilden5;
			int colChilden6;
			int colChilden7;
			int colChilden8;
			int colChilden9;
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1]);
				GeneralFunctions.fpSpReindexingTableTemplate(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[ToKhaiInfo.SheetInfos[1].Tables.Count - 1]);
				colHeader = HTKKFile.ColumnLettersToIndex("G");
				colHeader2 = HTKKFile.ColumnLettersToIndex("Q");
				colHeader3 = HTKKFile.ColumnLettersToIndex("U");
				colHeader4 = HTKKFile.ColumnLettersToIndex("W");
				colHeader5 = HTKKFile.ColumnLettersToIndex("S");
				colHeader6 = HTKKFile.ColumnLettersToIndex("X");
				colHeader7 = HTKKFile.ColumnLettersToIndex("Z");
				colHeader8 = HTKKFile.ColumnLettersToIndex("AB");
				colHeader9 = HTKKFile.ColumnLettersToIndex("AA");
				colHeader10 = HTKKFile.ColumnLettersToIndex("AC");
				colChilden = HTKKFile.ColumnLettersToIndex("AS");
				colChilden2 = HTKKFile.ColumnLettersToIndex("AT");
				colChilden3 = HTKKFile.ColumnLettersToIndex("AU");
				colChilden4 = HTKKFile.ColumnLettersToIndex("AV");
				colChilden5 = HTKKFile.ColumnLettersToIndex("AW");
				colChilden6 = HTKKFile.ColumnLettersToIndex("AX");
				colChilden7 = HTKKFile.ColumnLettersToIndex("AY");
				colChilden8 = HTKKFile.ColumnLettersToIndex("AZ");
				colChilden9 = HTKKFile.ColumnLettersToIndex("BA");
				int colChilden10 = HTKKFile.ColumnLettersToIndex("BB");
				for (int j = 0; j < ToKhaiInfo.SheetInfos[1].Tables.Count; j++)
				{
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader2, colChilden);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader5, colChilden2);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader6, colChilden3);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader3, colChilden4);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader4, colChilden5);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader7, colChilden6);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader8, colChilden7);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader9, colChilden8);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader10, colChilden9);
					GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], j, -1, colHeader, colChilden10);
				}
				foreach (HTKKCMCTableInfo table in ToKhaiInfo.SheetInfos[1].Tables)
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int k = 0; k < table.RowsValues.Count; k++)
					{
						GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[1], dataLocationTemplateRowStart + k, -1, table.RowsInfo, table);
					}
				}
			}
			if (ToKhaiInfo.SheetInfos[2] == null)
			{
				return;
			}
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2]);
			GeneralFunctions.fpSpReindexingTableTemplate(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2].Tables[ToKhaiInfo.SheetInfos[2].Tables.Count - 1]);
			LockUnLockPl01_2(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], _TTinChung, -1);
			colHeader = HTKKFile.ColumnLettersToIndex("G");
			colHeader2 = HTKKFile.ColumnLettersToIndex("N");
			colHeader3 = HTKKFile.ColumnLettersToIndex("R");
			colHeader4 = HTKKFile.ColumnLettersToIndex("T");
			colHeader5 = HTKKFile.ColumnLettersToIndex("P");
			colHeader6 = HTKKFile.ColumnLettersToIndex("U");
			colHeader7 = HTKKFile.ColumnLettersToIndex("W");
			colHeader8 = HTKKFile.ColumnLettersToIndex("Y");
			colHeader9 = HTKKFile.ColumnLettersToIndex("X");
			colHeader10 = HTKKFile.ColumnLettersToIndex("Z");
			colChilden = HTKKFile.ColumnLettersToIndex("AL");
			colChilden2 = HTKKFile.ColumnLettersToIndex("AM");
			colChilden3 = HTKKFile.ColumnLettersToIndex("AN");
			colChilden4 = HTKKFile.ColumnLettersToIndex("AO");
			colChilden5 = HTKKFile.ColumnLettersToIndex("AP");
			colChilden6 = HTKKFile.ColumnLettersToIndex("AQ");
			colChilden7 = HTKKFile.ColumnLettersToIndex("AR");
			colChilden8 = HTKKFile.ColumnLettersToIndex("AS");
			colChilden9 = HTKKFile.ColumnLettersToIndex("AT");
			int colChilden11 = HTKKFile.ColumnLettersToIndex("AU");
			for (int j = 0; j < ToKhaiInfo.SheetInfos[2].Tables.Count; j++)
			{
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader2, colChilden);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader5, colChilden2);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader6, colChilden3);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader3, colChilden4);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader4, colChilden5);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader7, colChilden6);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader8, colChilden7);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader9, colChilden8);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader10, colChilden9);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2], j, -1, colHeader, colChilden11);
			}
			foreach (HTKKCMCTableInfo table2 in ToKhaiInfo.SheetInfos[2].Tables)
			{
				int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
				for (int k = 0; k < table2.RowsValues.Count; k++)
				{
					GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[2], dataLocationTemplateRowStart + k, -1, table2.RowsInfo, table2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public double TinhTongTheoIsDonViHH(SheetView sView, HTKKCMCTableInfo dataTable, int colIsDonViOrHH, int colSum)
	{
		try
		{
			int dataLocationTemplateRowStart = dataTable.Location.DataLocationTemplateRowStart;
			double num = 0.0;
			for (int i = 0; i < dataTable.RowsValues.Count; i++)
			{
				bool flag = Convert.ToBoolean(sView.Cells[dataLocationTemplateRowStart + i - 1, colIsDonViOrHH - 1].Value);
				double num2 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, colSum - 1].Value);
				if (!flag)
				{
					num += num2;
				}
			}
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0.0;
		}
	}

	public Dictionary<string, int> DetectChildenRowByValue(SheetView sView, HTKKCMCTableInfo table, int startRow, int colParentCheckTo, int colChildenCheckTo, string valueParentCompareFrom, string valueChildenCompareFrom)
	{
		try
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				string text = sView.Cells[startRow + i - 1, colParentCheckTo - 1].Text.Trim();
				string text2 = sView.Cells[startRow + i - 1, colChildenCheckTo - 1].Text.Trim();
				if (string.IsNullOrEmpty(text2))
				{
					return dictionary;
				}
				if (!dictionary.ContainsKey(text))
				{
					dictionary.Add(text, startRow + i);
				}
				else
				{
					dictionary[text] = startRow + i;
				}
				if (text == valueParentCompareFrom && text2 == valueChildenCompareFrom)
				{
					return dictionary;
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc1[1], cellTemPhuLuc1[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc2[1], cellTemPhuLuc2[0]].Value = ((ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		SheetView sView = winFpSpread.Sheets[1];
		int colSum = HTKKFile.ColumnLettersToIndex("AL");
		TinhTongTheoTinhPl01_1(winFpSpread, _ToKhaiInfo);
		TinhBangTongCongPl01_1(winFpSpread, _ToKhaiInfo);
		List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
		foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
		{
			if (table.TableName == "01_1_TBVMT_2")
			{
				list.Add(table);
			}
		}
		GeneralFunctions.CalculateSumByGroup(sView, list, colSum);
		GeneralFunctions.fpSpReindexingTableTemplate(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1].Tables[ToKhaiInfo.SheetInfos[1].Tables.Count - 1]);
		GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1]);
		return true;
	}

	public bool UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sView = winFpSpread.Sheets[2];
			int colSum = HTKKFile.ColumnLettersToIndex("AF");
			TinhTongTheoTinhPl01_2(winFpSpread, ToKhaiInfo);
			TinhBangTongCongPl01_2(winFpSpread, ToKhaiInfo);
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				if (table.TableName == "01_2_TBVMT_2")
				{
					list.Add(table);
				}
			}
			GeneralFunctions.CalculateSumByGroup(sView, list, colSum);
			GeneralFunctions.fpSpReindexingTableTemplate(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2].Tables[ToKhaiInfo.SheetInfos[2].Tables.Count - 1]);
			GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2]);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LuuToanBoHangHoaToKhai(SheetView sView, HTKKCMCTableInfo table)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("L");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			dicToKhaiHangHoaCol4.Clear();
			dicToKhaiHangHoaCol7.Clear();
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				string key = sView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
				double num5 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
				double num6 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value);
				if (!dicToKhaiHangHoaCol4.ContainsKey(key))
				{
					dicToKhaiHangHoaCol4.Add(key, num5);
				}
				else
				{
					dicToKhaiHangHoaCol4[key] += num5;
				}
				if (!dicToKhaiHangHoaCol7.ContainsKey(key))
				{
					dicToKhaiHangHoaCol7.Add(key, num6);
				}
				else
				{
					dicToKhaiHangHoaCol7[key] += num6;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCot11Pl_01_2(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			if (sheetView == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("AB");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			int num3 = HTKKFile.ColumnLettersToIndex("AE");
			int num4 = HTKKFile.ColumnLettersToIndex("AF");
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (!(table.TableName != "01_2_TBVMT_1"))
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table.RowsValues.Count; i++)
					{
						string key = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
						double num5 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value = ((dicToKhaiHangHoaCol7.ContainsKey(key) && dicToKhaiHangHoaCol7[key] != 0.0) ? Math.Round(num5 / 100.0 * dicToKhaiHangHoaCol7[key], 0) : 0.0);
					}
				}
			}
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
			{
				if (table2.TableName == "01_2_TBVMT_1")
				{
					list.Add(table2);
				}
			}
			GeneralFunctions.CalculateSumByGroup(sheetView, list, num4);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCot10Pl_01_2(FpSpread winFpSpread, HTKKCMCSheetInfo sInfo, int iRow)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			if (sheetView == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("AB");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			int num3 = HTKKFile.ColumnLettersToIndex("AD");
			int num4 = HTKKFile.ColumnLettersToIndex("AE");
			string key;
			double num5;
			if (iRow < 0)
			{
				foreach (HTKKCMCTableInfo table in sInfo.Tables)
				{
					if (!(table.TableName != "01_2_TBVMT_1"))
					{
						int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
						for (int i = 0; i < table.RowsValues.Count; i++)
						{
							key = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
							num5 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
							sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value = ((dicToKhaiHangHoaCol4.ContainsKey(key) && dicToKhaiHangHoaCol4[key] != 0.0) ? Math.Round(num5 / dicToKhaiHangHoaCol4[key] * 100.0, 6) : 0.0);
						}
					}
				}
				return;
			}
			key = sheetView.Cells[iRow - 1, num2 - 1].Text.Trim();
			num5 = Convert.ToDouble(sheetView.Cells[iRow - 1, num3 - 1].Value);
			sheetView.Cells[iRow - 1, num4 - 1].Value = ((dicToKhaiHangHoaCol4.ContainsKey(key) && dicToKhaiHangHoaCol4[key] != 0.0) ? Math.Round(num5 / dicToKhaiHangHoaCol4[key] * 100.0, 6) : 0.0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhTongTheoTinhPl01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (sheetView == null || hTKKCMCSheetInfo == null)
			{
				return;
			}
			int key = HTKKFile.ColumnLettersToIndex("S");
			int key2 = HTKKFile.ColumnLettersToIndex("X");
			int key3 = HTKKFile.ColumnLettersToIndex("AE");
			int key4 = HTKKFile.ColumnLettersToIndex("AD");
			int num = HTKKFile.ColumnLettersToIndex("AL");
			int value = HTKKFile.ColumnLettersToIndex("E");
			int value2 = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(key2, value);
			Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
			dictionary2.Add(key, value2);
			Dictionary<int, int> dictionary3 = new Dictionary<int, int>();
			dictionary3.Add(key3, value2);
			dictionary3.Add(key4, value);
			dictionary3.Add(num, num);
			GeneralFunctions.CopyTableToTableGroup_F7_F7(sheetView, hTKKCMCSheetInfo, "01_1_TBVMT_2", "01_1_TBVMT_3", dictionary, dictionary2, dictionary3, usingCellType: true);
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				if (table.TableName == "01_1_TBVMT_3")
				{
					list.Add(table);
				}
			}
			GeneralFunctions.CalculateSumByGroup(sheetView, list, num);
			GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, hTKKCMCSheetInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhBangTongCongPl01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (sheetView != null && hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("S");
				int num2 = HTKKFile.ColumnLettersToIndex("X");
				int key = HTKKFile.ColumnLettersToIndex("AE");
				int key2 = HTKKFile.ColumnLettersToIndex("AD");
				int num3 = HTKKFile.ColumnLettersToIndex("AJ");
				int num4 = HTKKFile.ColumnLettersToIndex("AL");
				int value = HTKKFile.ColumnLettersToIndex("E");
				int num5 = HTKKFile.ColumnLettersToIndex("G");
				int value2 = HTKKFile.ColumnLettersToIndex("E");
				int value3 = HTKKFile.ColumnLettersToIndex("G");
				int num6 = HTKKFile.ColumnLettersToIndex("F");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(key2, value);
				Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
				dictionary2.Add(key, value3);
				dictionary2.Add(key2, value2);
				dictionary2.Add(num3, num3);
				dictionary2.Add(num4, num4);
				GeneralFunctions.CopyTableToTableGroup_F7_F5(sheetView, hTKKCMCSheetInfo, "01_1_TBVMT_2", "01_1_TBVMT_4", dictionary, dictionary2, usingCellType: true);
				HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1];
				double num7 = GeneralFunctions.CalculateColSumTable(sheetView, table, num4);
				string pathTong = "PLuc_01_1_TBVMT/TongCong/tong";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTong));
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num7;
				GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, hTKKCMCSheetInfo);
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhTongTheoTinhPl01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (sheetView != null && hTKKCMCSheetInfo != null)
			{
				int key = HTKKFile.ColumnLettersToIndex("P");
				int key2 = HTKKFile.ColumnLettersToIndex("U");
				int key3 = HTKKFile.ColumnLettersToIndex("AB");
				int key4 = HTKKFile.ColumnLettersToIndex("AA");
				int num = HTKKFile.ColumnLettersToIndex("AF");
				int value = HTKKFile.ColumnLettersToIndex("E");
				int value2 = HTKKFile.ColumnLettersToIndex("G");
				int num2 = HTKKFile.ColumnLettersToIndex("E");
				int num3 = HTKKFile.ColumnLettersToIndex("G");
				int num4 = HTKKFile.ColumnLettersToIndex("F");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(key2, value);
				Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
				dictionary2.Add(key, value2);
				Dictionary<int, int> dictionary3 = new Dictionary<int, int>();
				dictionary3.Add(key3, value2);
				dictionary3.Add(key4, value);
				dictionary3.Add(num, num);
				GeneralFunctions.CopyTableToTableGroup_F7_F7(sheetView, hTKKCMCSheetInfo, "01_2_TBVMT_1", "01_2_TBVMT_2", dictionary, dictionary2, dictionary3, usingCellType: true);
				GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, hTKKCMCSheetInfo);
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhBangTongCongPl01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (sheetView != null && hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("P");
				int num2 = HTKKFile.ColumnLettersToIndex("U");
				int key = HTKKFile.ColumnLettersToIndex("AB");
				int key2 = HTKKFile.ColumnLettersToIndex("AA");
				int num3 = HTKKFile.ColumnLettersToIndex("AC");
				int num4 = HTKKFile.ColumnLettersToIndex("AD");
				int num5 = HTKKFile.ColumnLettersToIndex("AE");
				int num6 = HTKKFile.ColumnLettersToIndex("AF");
				int value = HTKKFile.ColumnLettersToIndex("E");
				int value2 = HTKKFile.ColumnLettersToIndex("G");
				int num7 = HTKKFile.ColumnLettersToIndex("E");
				int num8 = HTKKFile.ColumnLettersToIndex("G");
				int num9 = HTKKFile.ColumnLettersToIndex("F");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(key2, value);
				Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
				dictionary2.Add(key, value2);
				dictionary2.Add(key2, value);
				dictionary2.Add(num3, num3);
				dictionary2.Add(num4, num4);
				dictionary2.Add(num5, num5);
				dictionary2.Add(num6, num6);
				GeneralFunctions.CopyTableToTableGroup_F7_F5(sheetView, hTKKCMCSheetInfo, "01_2_TBVMT_1", "01_2_TBVMT_3", dictionary, dictionary2, usingCellType: true);
				HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1];
				double num10 = GeneralFunctions.CalculateColSumTable(sheetView, table, num6);
				string pathTong = "PLuc_01_2_TBVMT/TongCong/tong";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTong));
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num10;
				GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, hTKKCMCSheetInfo);
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnLockPl01_2(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTTinChungInfo TTinChung, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("AI");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("N");
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (!(table.TableName != "01_2_TBVMT_1"))
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					string text = sView.Cells[dataLocationTemplateRowStart - 2, num2 - 1].Text.Trim();
					if (text == "01")
					{
						sView.Cells[dataLocationTemplateRowStart - 2, num3 - 1].Value = TTinChung.tenNNT;
						sView.Cells[dataLocationTemplateRowStart - 2, num4 - 1].Value = TTinChung.mst;
						GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart - 1, num3, locked: true, deleteData: false);
						GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart - 1, num4, locked: true, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart - 1, num3, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart - 1, num4, locked: false, deleteData: false);
					}
				}
			}
			if (iRow >= 0)
			{
				string text = sView.Cells[iRow - 1, num2 - 1].Text.Trim();
				if (text != "01")
				{
					sView.Cells[iRow - 1, num3 - 1].Value = "";
					sView.Cells[iRow - 1, num4 - 1].Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
