using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ttdb01TT80Nckt251
{
	private Dictionary<string, string[]> dicDMThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDichVuThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDichVuThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDanhMucSoXoDienToan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDanhMucSoXoDienToanTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private string FOLDER_DM = "..\\InterfaceIni";

	private string nameTable_01_TTDB_I_1 = "01_TTDB_I_1";

	private string nameTable_01_TTDB_I_2 = "01_TTDB_I_2";

	private string nameTable_01_TTDB_II = "01_TTDB_II";

	private string nameTable_01_TTDB_III_A = "01_TTDB_III_A";

	private string nameTable_01_TTDB_III_B = "01_TTDB_III_B";

	private string nameTable_01_TTDB_III_C = "01_TTDB_III_C";

	private int[] cellNopTruSoChinh = GeneralFunctions.ArrayCellId("AT_13");

	private int[] cellPhuLuc1 = GeneralFunctions.ArrayCellId("AT_15");

	private int[] cellPhuLuc2 = GeneralFunctions.ArrayCellId("AT_16");

	private int[] cellCt11c = GeneralFunctions.ArrayCellId("M_33");

	private int[] cellCt11cMa = GeneralFunctions.ArrayCellId("AT_33");

	private int[] cellCt11b = GeneralFunctions.ArrayCellId("M_35");

	private int[] cellCt11bMa = GeneralFunctions.ArrayCellId("AT_35");

	private int[] cellCt11a = GeneralFunctions.ArrayCellId("M_37");

	private int[] cellCt11aMa = GeneralFunctions.ArrayCellId("AT_37");

	private int[] cellCt09 = GeneralFunctions.ArrayCellId("AC_27");

	private int[] cellCt10 = GeneralFunctions.ArrayCellId("S_29");

	private string cellId_TieuMuc = "L_14";

	private string cellId_N47 = "N_47";

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private int[] CellId_TieuMuc => GeneralFunctions.ArrayCellId(cellId_TieuMuc);

	private int[] CellId_N47 => GeneralFunctions.ArrayCellId(cellId_N47);

	private int idxCol_F => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_G => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_P => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_S => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_AA => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_AD => HTKKFile.ColumnLettersToIndex("AD") - 1;

	private int idxCol_AG => HTKKFile.ColumnLettersToIndex("AG") - 1;

	private int idxCol_AN => HTKKFile.ColumnLettersToIndex("AN") - 1;

	private int idxCol_C_PL1 => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_D_PL1 => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_E_PL1 => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_F_PL1 => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_G_PL1 => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_H_PL1 => HTKKFile.ColumnLettersToIndex("H") - 1;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("M_41");
			int[] array2 = GeneralFunctions.ArrayCellId("N_41");
			if (string.IsNullOrEmpty(sheetView.Cells[array[1], array[0]].Text.Trim()))
			{
				sheetView.Cells[array[1], array[0]].Text = PeriodInfo.TieuMucHachToanVal;
				sheetView.Cells[array2[1], array2[0]].Text = PeriodInfo.NganhNgheKinhDoanh;
			}
			sheetView.Cells[cellPhuLuc1[1], cellPhuLuc1[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			sheetView.Cells[cellPhuLuc2[1], cellPhuLuc2[0]].Value = ((_ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			int num = ((PeriodInfo.maCQTCapCuc == headerInfo.maCQTCapCuc) ? 1 : 0);
			if (num == 1)
			{
				GeneralFunctions.LockUnLockCol(sheetView, cellCt09[1] + 1, cellCt09[0] + 1, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt10[1] + 1, cellCt10[0] + 1, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt11a[1] + 1, cellCt11a[0] + 1, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt11b[1] + 1, cellCt11b[0] + 1, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt11c[1] + 1, cellCt11c[0] + 1, locked: true, deleteData: true);
			}
			else
			{
				GeneralFunctions.LockUnLockCol(sheetView, cellCt09[1] + 1, cellCt09[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt10[1] + 1, cellCt10[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt11a[1] + 1, cellCt11a[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt11b[1] + 1, cellCt11b[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt11c[1] + 1, cellCt11c[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cellCt11c, cellCt11cMa, cellCt11b, cellCt11bMa, cellCt11a, cellCt11aMa, null, cellCt09, cellCt10);
			}
			NeuChonXoSoDienToan(sheetView, ToKhaiInfo, PeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			if (activeRowIndex == 2 && activeColumnIndex == 1)
			{
				ShowBangKe("Bangkekhautru_01TTDB_TT80.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		ChangeCombobox_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
		int num = HTKKFile.ColumnLettersToIndex("G");
		int num2 = HTKKFile.ColumnLettersToIndex("P");
		int num3 = HTKKFile.ColumnLettersToIndex("S");
		int num4 = HTKKFile.ColumnLettersToIndex("V");
		int num5 = HTKKFile.ColumnLettersToIndex("AA");
		int num6 = HTKKFile.ColumnLettersToIndex("AW");
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_II);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		int num7 = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
		if (iRow <= num7 && (iCol == num || iCol == num2 || iCol == num3 || iCol == num4))
		{
			winFpSpread.Sheets[0].Cells[(int)iRow - 1, num5 - 1].Value = winFpSpread.Sheets[0].Cells[(int)iRow - 1, num6 - 1].Value;
		}
		TinhTongTheoTungTinh(winFpSpread, ToKhaiInfo);
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			try
			{
				SheetView sheetView = winFpSpread.Sheets[1];
				HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
				int num = HTKKFile.ColumnLettersToIndex("H");
				int num2 = HTKKFile.ColumnLettersToIndex("I");
				int num3 = HTKKFile.ColumnLettersToIndex("J");
				int num4 = HTKKFile.ColumnLettersToIndex("K");
				int num5 = HTKKFile.ColumnLettersToIndex("Q");
				int num6 = HTKKFile.ColumnLettersToIndex("L");
				int num7 = HTKKFile.ColumnLettersToIndex("Q");
				int num8 = HTKKFile.ColumnLettersToIndex("N");
				int num9 = HTKKFile.ColumnLettersToIndex("R");
				int num10 = HTKKFile.ColumnLettersToIndex("O");
				int num11 = HTKKFile.ColumnLettersToIndex("S");
				int num12 = HTKKFile.ColumnLettersToIndex("M");
				int num13 = HTKKFile.ColumnLettersToIndex("R");
				int num14 = HTKKFile.ColumnLettersToIndex("N");
				int num15 = HTKKFile.ColumnLettersToIndex("S");
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[1];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				int num16 = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
				if (iRow < dataLocationTemplateRowStart && (iCol == num2 || iCol == num))
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = sheetView.Cells[(int)iRow - 1, num7 - 1].Value;
				}
				if (num16 == 0 && (iCol == num3 || iCol == num4 || iCol == num6))
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = sheetView.Cells[(int)iRow - 1, num9 - 1].Value;
				}
				if (num16 == 0 && (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num6 || iCol == num12 || iCol == num8))
				{
					sheetView.Cells[(int)iRow - 1, num10 - 1].Value = sheetView.Cells[(int)iRow - 1, num11 - 1].Value;
				}
				if (num16 == 1 && (iCol == num2 || iCol == num))
				{
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = sheetView.Cells[(int)iRow - 1, num5 - 1].Value;
				}
				if (num16 == 1 && (iCol == num || iCol == num2 || iCol == num3 || iCol == num4))
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Value = sheetView.Cells[(int)iRow - 1, num13 - 1].Value;
				}
				if (num16 == 1 && (iCol == num || iCol == num2 || iCol == num3 || iCol == num4))
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Value = sheetView.Cells[(int)iRow - 1, num13 - 1].Value;
				}
				if (num16 == 1 && (iCol == num || iCol == num2 || iCol == num3 || iCol == num4 || iCol == num6 || iCol == num12))
				{
					sheetView.Cells[(int)iRow - 1, num14 - 1].Value = sheetView.Cells[(int)iRow - 1, num15 - 1].Value;
				}
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				throw ex;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("M");
			int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("U");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("N");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("V");
			int num6 = HTKKFile.ColumnLettersToIndex("Q");
			int num7 = HTKKFile.ColumnLettersToIndex("O");
			int num8 = HTKKFile.ColumnLettersToIndex("G");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[2];
			if (iCol == num)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			if (iCol == num4)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc, isLock: false);
			}
			if (num4 == iCol || iCol == num7 || iRow == 21)
			{
				TinhTongTheoTungTinh(winFpSpread, ToKhaiInfo);
			}
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[1]);
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
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("G");
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			bool flag = false;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cellCt11c, cellCt11cMa, cellCt11b, cellCt11bMa, cellCt11a, cellCt11aMa, null, cellCt09, cellCt10))
			{
				value = 2;
			}
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text))
					{
						List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, table, num, text);
						if (rowIndexInSViewByCompare.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message8);
							value = 2;
						}
					}
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
			throw ex;
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
			string message = HTKKFile.GetMessage("0071");
			string message2 = HTKKFile.GetMessage("0112");
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			HTKKCMCTableInfo hTKKCMCTableInfo = null;
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				if (hTKKCMCTableInfo == null)
				{
					continue;
				}
				int num = table.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[table.RowsInfo.Cells.Count];
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					array[3] = sheetView.Cells[num + i, idxCol_E_PL1].Text;
					if (!GeneralFunctions.HaveToValidateRow(table.RowsInfo, array))
					{
						HTKKFile.RemoveValidate(sheetView, num + i, idxCol_E_PL1);
					}
					else if (!string.IsNullOrEmpty(array[3]))
					{
						if (DateTime.TryParseExact(array[3], "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
						{
							int num2 = 12 * ConvertHelpers.ConvertToInt(PeriodInfo.Nam) + ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
							int num3 = 12 * result.Year + result.Month;
							if (num3 > num2)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_E_PL1, message2);
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num + i, idxCol_E_PL1);
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_E_PL1, message);
							value = 2;
						}
					}
					else if (i > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_E_PL1, message3);
						value = 2;
					}
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
			throw ex;
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
			int num = HTKKFile.ColumnLettersToIndex("W");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("U");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("V");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_84"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_78"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_76"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_79"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_82"));
			int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
			{
			}
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text.Trim();
				if (text == "1")
				{
					if (!string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, message4);
						value = 2;
					}
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num3, text2);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, message5);
						value = 2;
					}
				}
				if (text == "2" && !string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue5So(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, message6);
					value = 2;
				}
				string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Text.Trim();
				string text4 = text3 + "00";
				if (!string.IsNullOrEmpty(PeriodInfo.MaTinh))
				{
					if (text3 == PeriodInfo.MaTinh)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, message);
						value = 2;
					}
				}
				else if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && text4 == PeriodInfo.maCQTCapCuc)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, message2);
					value = 2;
				}
				List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[0], num4, text3);
				if (rowIndexInSViewByCompare2.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, message3);
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
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
				{
					string key = array[1];
					string text = array[2].Trim();
					string text2 = array[4].Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						text = text + " (" + text2 + " %)";
					}
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDMThueDB = dictionary;
			dicDMThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_04a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_04aa: Expected O, but got Unknown
		//IL_04d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_04df: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("1"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								if (!string.IsNullOrEmpty(item.thuesuat))
								{
									text = text + " (" + item.thuesuat + " %)";
								}
								if (!dictionary2.ContainsKey(key))
								{
									dictionary2.Add(key, text);
								}
								if (!dictionary.ContainsKey(text))
								{
									dictionary.Add(text, array);
								}
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("1") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							if (!string.IsNullOrEmpty(item.thuesuat))
							{
								text = text + " (" + item.thuesuat + " %)";
							}
							if (!dictionary2.ContainsKey(key))
							{
								dictionary2.Add(key, text);
							}
							if (!dictionary.ContainsKey(text))
							{
								dictionary.Add(text, array);
							}
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						string text2 = array[4].Trim();
						if (!string.IsNullOrEmpty(text2))
						{
							text = text + " (" + text2 + " %)";
						}
						if (!dictionary2.ContainsKey(key))
						{
							dictionary2.Add(key, text);
						}
						if (!dictionary.ContainsKey(text))
						{
							dictionary.Add(text, array);
						}
					}
				}
			}
			dicDMThueDB = dictionary;
			dicDMThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDichVuThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDichVuThueDBTmp.Count > 0)
			{
				return dicDichVuThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[1].StartsWith("2") && array[^1] == "1")
				{
					string key = array[1];
					string text = array[2].Trim();
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDichVuThueDB = dictionary;
			dicDichVuThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDichVuThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_0447: Unknown result type (might be due to invalid IL or missing references)
		//IL_044e: Expected O, but got Unknown
		//IL_047c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0483: Expected O, but got Unknown
		try
		{
			if (dicDichVuThueDBTmp.Count > 0)
			{
				return dicDichVuThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("2"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								if (!dictionary2.ContainsKey(key))
								{
									dictionary2.Add(key, text);
								}
								if (!dictionary.ContainsKey(text))
								{
									dictionary.Add(text, array);
								}
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("2") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							if (!dictionary2.ContainsKey(key))
							{
								dictionary2.Add(key, text);
							}
							if (!dictionary.ContainsKey(text))
							{
								dictionary.Add(text, array);
							}
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("2") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						if (!dictionary2.ContainsKey(key))
						{
							dictionary2.Add(key, text);
						}
						if (!dictionary.ContainsKey(text))
						{
							dictionary.Add(text, array);
						}
					}
				}
			}
			dicDichVuThueDB = dictionary;
			dicDichVuThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMTTDBDonVi(string pathFile)
	{
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		try
		{
			if (dicDMDonVi.Count > 0)
			{
				return dicDMDonVi;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			int num = 0;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["value"]).Value;
				num++;
				dictionary.Add(num.ToString(), value);
			}
			dicDMDonVi = dictionary;
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMSoXoDienToan()
	{
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_0066: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Expected O, but got Unknown
		try
		{
			if (dicDanhMucSoXoDienToanTmp.Count > 0)
			{
				return dicDanhMucSoXoDienToanTmp;
			}
			string text = FOLDER_DM + "\\Catalogue_DM_XoSoDienToan.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2)
				{
					string key = array[0];
					string value2 = array[1].Trim();
					List<string> list = new List<string>();
					for (int i = 0; i < array.Count(); i++)
					{
						list.Add(array[i]);
					}
					dicDanhMucSoXoDienToan.Add(key, list.ToArray());
					if (!dicDanhMucSoXoDienToanTmp.ContainsKey(key))
					{
						dicDanhMucSoXoDienToanTmp.Add(key, value2);
					}
				}
			}
			return dicDanhMucSoXoDienToanTmp;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("AY");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string text = sheetView.Cells[num3 + j, idxCol_G].Text;
					if (!string.IsNullOrEmpty(text) && dicDMThueDB.ContainsKey(text))
					{
						string[] array2 = dicDMThueDB[text];
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array2[1], "TTDB");
						sheetView.Cells[num3 + j, num2 - 1].Text = maTieuMucByMaHH;
					}
					else
					{
						sheetView.Cells[num3 + j, num2 - 1].Text = "";
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		try
		{
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int num5 = HTKKFile.ColumnLettersToIndex("W");
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo2 == null)
			{
				return;
			}
			SheetView sheetView2 = winFpSpread.Sheets[2];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(sheetView2, dataLocationTemplateRowStart, num4);
			for (int j = 0; j < hTKKCMCTableInfo2.RowsValues.Count; j++)
			{
				string key = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text.Trim();
				if (dMFromCombobox.ContainsKey(key))
				{
					sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text = dMFromCombobox[key];
				}
				else
				{
					sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text = "";
				}
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc1[1], cellPhuLuc1[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			winFpSpread.Sheets[0].Cells[cellPhuLuc1[2], cellPhuLuc1[0]].Value = ((_ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc1[1], cellPhuLuc1[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			winFpSpread.Sheets[0].Cells[cellPhuLuc1[2], cellPhuLuc1[0]].Value = ((_ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
			GeneralFunctions.checkLockUnLockCells(_ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		int num = HTKKFile.ColumnLettersToIndex("F");
		int num2 = HTKKFile.ColumnLettersToIndex("G");
		int num3 = HTKKFile.ColumnLettersToIndex("J");
		int formColTenTinh = HTKKFile.ColumnLettersToIndex("K");
		int num4 = HTKKFile.ColumnLettersToIndex("S");
		int num5 = HTKKFile.ColumnLettersToIndex("M");
		int num6 = HTKKFile.ColumnLettersToIndex("U");
		int num7 = HTKKFile.ColumnLettersToIndex("N");
		int num8 = HTKKFile.ColumnLettersToIndex("V");
		int num9 = HTKKFile.ColumnLettersToIndex("Q");
		int num10 = HTKKFile.ColumnLettersToIndex("O");
		int toColTenTinh = HTKKFile.ColumnLettersToIndex("G");
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
		SheetView sheetView = winFpSpread.Sheets[2];
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		dictionary.Add(num9, num9);
		GeneralFunctions.TongCongTheoTungTinh(hTKKCMCSheetInfo, sheetView, hTKKCMCSheetInfo.Tables[0], num4, formColTenTinh, sheetView, hTKKCMCSheetInfo.Tables[1], 1, num4, toColTenTinh, dictionary);
		return true;
	}

	public void NeuChonXoSoDienToan(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		int col = HTKKFile.ColumnLettersToIndex("G");
		List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
		string tieuMucHachToanVal = PeriodInfo.TieuMucHachToanVal;
		if (!(tieuMucHachToanVal == "03"))
		{
			return;
		}
		GetDMSoXoDienToan();
		HTKKCMCTableInfo hTKKCMCTableInfo = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TableName == "01_TTDB_II");
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
		{
			GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart + num, col, dicDanhMucSoXoDienToanTmp, deleteText: false);
		}
		foreach (HTKKCMCTableInfo item in tables)
		{
			if (item.TableName == "01_TTDB_II")
			{
				continue;
			}
			int dataLocationTemplateRowStart2 = item.Location.DataLocationTemplateRowStart;
			for (int num2 = 0; num2 < item.RowsValues.Count; num2++)
			{
				foreach (HTKKCMCTableCellInfo cell in item.RowsInfo.Cells)
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart2 + num2, cell.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
			}
		}
	}

	public void TinhTongTheoTungTinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
		if (hTKKCMCSheetInfo != null)
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			int formColTenTinh = HTKKFile.ColumnLettersToIndex("K");
			int num = HTKKFile.ColumnLettersToIndex("S");
			int num2 = HTKKFile.ColumnLettersToIndex("Q");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int toColTenTinh = HTKKFile.ColumnLettersToIndex("G");
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(num2, num2);
			GeneralFunctions.TongCongTheoTungTinh(hTKKCMCSheetInfo, sheetView, hTKKCMCSheetInfo.Tables[0], num, formColTenTinh, sheetView, hTKKCMCSheetInfo.Tables[1], 1, num, toColTenTinh, dictionary);
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

	public void ChangeCombobox_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("G");
			if (iCol != num)
			{
				return;
			}
			int num2 = HTKKFile.ColumnLettersToIndex("AT");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("AD");
			int num5 = HTKKFile.ColumnLettersToIndex("AY");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_I_1);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_I_2);
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_II);
			HTKKCMCTableInfo hTKKCMCTableInfo4 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_III_A);
			HTKKCMCTableInfo hTKKCMCTableInfo5 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_III_B);
			HTKKCMCTableInfo hTKKCMCTableInfo6 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_III_C);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num6 = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num7 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			int num8 = dataLocationTemplateRowStart3 + hTKKCMCTableInfo3.RowsValues.Count - 1;
			int dataLocationTemplateRowStart4 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart;
			int num9 = dataLocationTemplateRowStart4 + hTKKCMCTableInfo4.RowsValues.Count - 1;
			int dataLocationTemplateRowStart5 = hTKKCMCTableInfo5.Location.DataLocationTemplateRowStart;
			int num10 = dataLocationTemplateRowStart5 + hTKKCMCTableInfo5.RowsValues.Count - 1;
			int dataLocationTemplateRowStart6 = hTKKCMCTableInfo6.Location.DataLocationTemplateRowStart;
			int num11 = dataLocationTemplateRowStart6 + hTKKCMCTableInfo6.RowsValues.Count - 1;
			if (iRow >= dataLocationTemplateRowStart && iRow <= num7)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].ResetValue();
					sheetView.Cells[(int)iRow - 1, num3 - 1].ResetValue();
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
				}
				else if (!dicDMThueDB.ContainsKey(text))
				{
					sheetView.Cells[(int)iRow - 1, num - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "";
				}
				else
				{
					string[] array = dicDMThueDB[text];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = array[3];
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(array[4]);
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "TTDB");
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = maTieuMucByMaHH;
				}
			}
			else if (iRow >= dataLocationTemplateRowStart3 && iRow <= num8)
			{
				string tenDV = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = "";
				if (HTKKFormater.isNullString(tenDV))
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].ResetValue();
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
					return;
				}
				string tieuMucHachToanVal = PeriodInfo.TieuMucHachToanVal;
				string[] array;
				string maTieuMucByMaHH;
				if (tieuMucHachToanVal != "03")
				{
					if (!dicDichVuThueDB.ContainsKey(tenDV))
					{
						sheetView.Cells[(int)iRow - 1, num - 1].Text = "";
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
						sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "";
					}
					else
					{
						array = dicDichVuThueDB[tenDV];
						sheetView.Cells[(int)iRow - 1, num2 - 1].Value = array[1];
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(array[4]);
						maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "TTDB");
						sheetView.Cells[(int)iRow - 1, num5 - 1].Value = maTieuMucByMaHH;
					}
					return;
				}
				KeyValuePair<string, string[]> keyValuePair = dicDanhMucSoXoDienToan.FirstOrDefault((KeyValuePair<string, string[]> x) => x.Value[1] == tenDV);
				if (keyValuePair.Value != null && keyValuePair.Value.Count() == 0)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "";
					return;
				}
				array = keyValuePair.Value;
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = array[0];
				sheetView.Cells[(int)iRow - 1, num3 - 1].Value = array[3];
				sheetView.Cells[(int)iRow - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(array[2]);
				maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[0], "TTDB");
				sheetView.Cells[(int)iRow - 1, num5 - 1].Value = maTieuMucByMaHH;
			}
			else if (iRow >= dataLocationTemplateRowStart4 && iRow <= num11)
			{
				string text = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].ResetValue();
				}
				else if (!dicDMThueDB.ContainsKey(text))
				{
					sheetView.Cells[(int)iRow - 1, num - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "";
				}
				else
				{
					string[] array = dicDMThueDB[text];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = array[3];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
