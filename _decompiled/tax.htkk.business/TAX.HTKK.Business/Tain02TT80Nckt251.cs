using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tain02TT80Nckt251
{
	private Dictionary<string, string> dicDMMo = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMKhoangSanChiTiet = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMKhoangSanChiTietTem = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private List<string> pathDonGiaTinh = new List<string>();

	private Common common = new Common();

	private string FOLDER_DM = "..\\InterfaceIni";

	private long rowLastClickSheet0 = 0L;

	private long rowLastClickSheet1 = 0L;

	private XmlNodeList nodeListDmMoKhoangSan = null;

	private string IS_MO = "mo";

	private int[] cellCoPhuLuc = GeneralFunctions.ArrayCellId("AM_11");

	private int[] cellCoMaV5 = GeneralFunctions.ArrayCellId("AM_12");

	private int[] cell09c = GeneralFunctions.ArrayCellId("M_14");

	private int[] cell09cMa = GeneralFunctions.ArrayCellId("AO_14");

	private int[] cell09b = GeneralFunctions.ArrayCellId("M_16");

	private int[] cell09bMa = GeneralFunctions.ArrayCellId("AO_16");

	private int[] cell09a = GeneralFunctions.ArrayCellId("M_18");

	private int[] cell09aMa = GeneralFunctions.ArrayCellId("AO_18");

	private int colMaLoai_PL = HTKKFile.ColumnLettersToIndex("AL");

	private int colLoai_PL = HTKKFile.ColumnLettersToIndex("F");

	private int[] cellCountNoiCoVanPhongDh_PL = GeneralFunctions.ArrayCellId("AL_30");

	private int[] cellCountNoiCoLongHoThuyDien_PL = GeneralFunctions.ArrayCellId("AL_32");

	private int col8_PL = HTKKFile.ColumnLettersToIndex("S");

	private int col8Ma_PL = HTKKFile.ColumnLettersToIndex("X");

	private int col7_PL = HTKKFile.ColumnLettersToIndex("U");

	private int col7Ma_PL = HTKKFile.ColumnLettersToIndex("W");

	private int col9Cuc_PL = HTKKFile.ColumnLettersToIndex("Z");

	private int col9MaCuc_PL = HTKKFile.ColumnLettersToIndex("AB");

	private int col9QuanLy_PL = HTKKFile.ColumnLettersToIndex("AA");

	private int col9MaQuanLy_PL = HTKKFile.ColumnLettersToIndex("AC");

	private int col05Header_PL = HTKKFile.ColumnLettersToIndex("G");

	private int col05DongCon_PL = HTKKFile.ColumnLettersToIndex("AP");

	private int col06Header_PL = HTKKFile.ColumnLettersToIndex("Q");

	private int col10Header_PL = HTKKFile.ColumnLettersToIndex("AE");

	private int col06DongCon_PL = HTKKFile.ColumnLettersToIndex("AM");

	private int col10DongCon_PL = HTKKFile.ColumnLettersToIndex("AN");

	private int colCt10_PL = HTKKFile.ColumnLettersToIndex("AE");

	private int colCt11_PL = HTKKFile.ColumnLettersToIndex("AF");

	private int colCt12_PL = HTKKFile.ColumnLettersToIndex("AG");

	private int[] cellCt10Sum = GeneralFunctions.ArrayCellId("AE_48");

	private int[] cellCt12Sum = GeneralFunctions.ArrayCellId("AG_48");

	private int[] cellCt13Sum = GeneralFunctions.ArrayCellId("AH_48");

	private int[] cellCt14Sum = GeneralFunctions.ArrayCellId("AI_48");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, cell09c, cell09cMa, cell09b, cell09bMa, cell09a, cell09aMa, null, null, null);
			sView.Cells[cellCoPhuLuc[1], cellCoPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			GeneralFunctions.fpSpReindexingSheetTemplate(sView, ToKhaiInfo.SheetInfos[0]);
			int[] array = GeneralFunctions.ArrayCellId("H_36");
			int[] array2 = GeneralFunctions.ArrayCellId("G_36");
			int[] array3 = GeneralFunctions.ArrayCellId("I_36");
			int[] array4 = GeneralFunctions.ArrayCellId("Q_36");
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[array[1], array[0]].Text = PeriodInfo.TenQuyetToan;
				sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaQuyetToan;
			}
			if (!string.IsNullOrEmpty(HeaderInfo.NgayBatDauNamTaiChinh))
			{
				sView.Cells[array3[1], array3[0]].Text = HeaderInfo.NgayBatDauNamTaiChinh;
				sView.Cells[array4[1], array4[0]].Text = GeneralFunctions.GetNgayTaiChinhDen(HeaderInfo.NgayBatDauNamTaiChinh);
			}
			List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			KeyValuePair<string, string[]> keyValuePair = default(KeyValuePair<string, string[]>);
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string valueCol2Ma = sView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					keyValuePair = dicDMKhoangSanChiTiet.FirstOrDefault((KeyValuePair<string, string[]> x) => x.Value.Length > 1 && x.Value[1].ToString() == valueCol2Ma);
					if (keyValuePair.Value != null && keyValuePair.Value.Length > 5 && (keyValuePair.Value[5] == "1" || keyValuePair.Value[1] == "1101"))
					{
						sView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Locked = false;
						HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, num2 - 1);
					}
					else
					{
						sView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Locked = true;
						HTKKFile.RemoveValidateBackGround(sView, dataLocationTemplateRowStart + j - 1, num2 - 1);
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

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("U");
			int num6 = HTKKFile.ColumnLettersToIndex("W");
			int num7 = HTKKFile.ColumnLettersToIndex("S");
			int num8 = HTKKFile.ColumnLettersToIndex("X");
			int num9 = HTKKFile.ColumnLettersToIndex("Z");
			int num10 = HTKKFile.ColumnLettersToIndex("AB");
			int num11 = HTKKFile.ColumnLettersToIndex("AA");
			int num12 = HTKKFile.ColumnLettersToIndex("AC");
			int num13 = HTKKFile.ColumnLettersToIndex("AE");
			int num14 = HTKKFile.ColumnLettersToIndex("AF");
			int num15 = HTKKFile.ColumnLettersToIndex("AG");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("AM");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("S");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("AA");
			int num9 = HTKKFile.ColumnLettersToIndex("AD");
			int num10 = HTKKFile.ColumnLettersToIndex("AN");
			int num11 = HTKKFile.ColumnLettersToIndex("AO");
			int num12 = HTKKFile.ColumnLettersToIndex("AL");
			int colChilden = HTKKFile.ColumnLettersToIndex("AR");
			int num13 = -1;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string text = sheetView.Cells[(int)iRow - 1, num2 - 1].Text.Trim();
			if (text == IS_MO)
			{
				return;
			}
			if (text == IS_MO)
			{
				num13 = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow + 1);
			}
			else
			{
				num13 = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
			}
			object value = sheetView.Cells[(int)iRow - 1, num9 - 1].Value;
			double num14 = ((value == null) ? 0.0 : ConvertHelpers.ConvertToDouble(value));
			object value2 = sheetView.Cells[(int)iRow - 1, num7 - 1].Value;
			double num15 = ((value2 == null) ? 0.0 : ConvertHelpers.ConvertToDouble(value2));
			string valueCol2 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
			KeyValuePair<string, string[]> keyValuePair = default(KeyValuePair<string, string[]>);
			ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
			ICellType cellType = sheetView.Cells[(int)iRow - 1, num3 - 1].CellType;
			if ((object)comboBoxCellType.GetType() == cellType.GetType())
			{
				string[] items = ((ComboBoxCellType)cellType).Items;
				string[] itemData = ((ComboBoxCellType)cellType).ItemData;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int i = 0; i < itemData.Length; i++)
				{
					if (!dictionary.ContainsKey(itemData[i]))
					{
						dictionary.Add(itemData[i], items[i]);
					}
				}
			}
			keyValuePair = dicDMKhoangSanChiTiet.FirstOrDefault((KeyValuePair<string, string[]> x) => x.Value.Length > 1 && x.Value[2].ToString() == valueCol2);
			if (!string.IsNullOrEmpty(keyValuePair.Key) && keyValuePair.Value.Length > 1)
			{
				sheetView.Cells[(int)iRow - 1, num4 - 1].Text = keyValuePair.Value[1];
				sheetView.Cells[(int)iRow - 1, num10 - 1].Text = ((keyValuePair.Value.Length > 7) ? keyValuePair.Value[7] : "");
			}
			if (iCol == num3)
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Text = "";
				sheetView.Cells[(int)iRow - 1, num7 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num8 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num11 - 1].Value = 0;
			}
			if (iCol == num3 && keyValuePair.Key != null && keyValuePair.Value.Length > 3)
			{
				if (ConvertHelpers.ConvertToInt(keyValuePair.Value[4]) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num3 - 1, message);
					return;
				}
				HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num3 - 1);
				sheetView.Cells[(int)iRow - 1, num5 - 1].Text = keyValuePair.Value[3].ToString();
				double donGia = getDonGia(keyValuePair.Key);
				sheetView.Cells[(int)iRow - 1, num7 - 1].Value = donGia;
				sheetView.Cells[(int)iRow - 1, num8 - 1].Text = keyValuePair.Value[4];
				sheetView.Cells[(int)iRow - 1, num11 - 1].Text = keyValuePair.Value[4];
				if (keyValuePair.Value[5] == "1" || keyValuePair.Value[1] == "1101")
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = true;
				}
				sheetView.Cells[(int)iRow - 1, num12 - 1].Text = keyValuePair.Value[0];
			}
			GeneralFunctions.SetLaiDongCon(sheetView, ToKhaiInfo.SheetInfos[0], 0, (int)iRow, num3, colChilden);
			double num16 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, num9 - 1].Value);
			double num17 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, num7 - 1].Value);
			if (iCol == num9 && num14 > 0.0)
			{
				sheetView.Cells[(int)iRow - 1, num8 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num7 - 1].Value = 0;
			}
			if (iCol == num7 && num15 > 0.0)
			{
				sheetView.Cells[(int)iRow - 1, num9 - 1].Value = 0;
			}
			object value3 = sheetView.Cells[(int)iRow - 1, num11 - 1].Value;
			if (iCol == num7 && num15 > 0.0)
			{
				sheetView.Cells[(int)iRow - 1, num8 - 1].Value = value3;
			}
			TinhTongCot_4_8_9_10(sheetView, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void SetTenMoDeTongHopKHBS(SheetView sView, HTKKCMCTableInfo table)
	{
		if (table.Header.StaticCells.Count != 0)
		{
			int num = HTKKFile.ColumnLettersToIndex("AR");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int rowIndex = table.Header.StaticCells[1].CellTemplate.RowIndex;
			string text = sView.Cells[rowIndex - 1, num2 - 1].Text.Trim();
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				string text2 = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				text2 = (text2.Contains(text) ? text2 : GeneralFunctions.FormatKey2(text2, text, " - "));
			}
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			int index = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (iCol == colLoai_PL || iCol == col05Header_PL)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, col05Header_PL, col05DongCon_PL);
			}
			if (iCol == colLoai_PL || iCol == col06Header_PL)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, col06Header_PL, col06DongCon_PL);
			}
			if (iCol == colLoai_PL || iCol == col10Header_PL || iCol == colCt11_PL)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, col10Header_PL, col10DongCon_PL);
			}
			if (iCol == colCt10_PL || iCol == colCt11_PL)
			{
				int rowIndex = hTKKCMCSheetInfo.Tables[index].Header.StaticCells[0].CellTemplate.RowIndex;
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[index];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				double num = GeneralFunctions.CalculateColSumTable(sheetView, hTKKCMCTableInfo, colCt11_PL);
				sheetView.Cells[dataLocationTemplateRowStart - 2, colCt11_PL - 1].Value = num;
				double num2 = GeneralFunctions.CalculateColSumTable(sheetView, hTKKCMCTableInfo, colCt12_PL);
				sheetView.Cells[dataLocationTemplateRowStart - 2, colCt12_PL - 1].Value = num2;
			}
			object value = sheetView.Cells[(int)iRow - 1, colMaLoai_PL - 1].Value;
			int num3 = Convert.ToInt32(sheetView.Cells[cellCountNoiCoVanPhongDh_PL[1], cellCountNoiCoVanPhongDh_PL[0]].Value);
			int num4 = Convert.ToInt32(sheetView.Cells[cellCountNoiCoLongHoThuyDien_PL[1], cellCountNoiCoLongHoThuyDien_PL[0]].Value);
			if (iCol == colLoai_PL)
			{
				if (value != null && value.ToString().Trim() == "01")
				{
					if (num3 > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, colLoai_PL - 1, msg);
					}
					GeneralFunctions.ThayDoiTinhTheoCQTDangBang(sheetView, hTKKCMCSheetInfo, PeriodInfo, -1, col8_PL, col8Ma_PL, col7_PL, col7Ma_PL, -1, -1, (int)iRow, isLock: false);
					GeneralFunctions.ThayDoiCQTTheoCQTOKyTinhThueDangBang(sheetView, hTKKCMCSheetInfo, PeriodInfo, -1, col9Cuc_PL, col9MaCuc_PL, col9QuanLy_PL, col9MaQuanLy_PL, (int)iRow, isLock: false);
					if (PeriodInfo.maCQTCapCuc == "82500")
					{
						GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col8_PL, locked: false, deleteData: true);
					}
				}
				if (value != null && value.ToString().Trim() == "02")
				{
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col8_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col8Ma_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col7_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col7Ma_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col9Cuc_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col9MaCuc_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col9QuanLy_PL, locked: false, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, col9MaQuanLy_PL, locked: false, deleteData: true);
				}
			}
			if (iCol == col8_PL)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, col8Ma_PL - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, col9Cuc_PL, col9MaCuc_PL, col9QuanLy_PL, col9MaQuanLy_PL, isLock: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			rowLastClickSheet0 = iRow;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			rowLastClickSheet1 = iRow;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("AL");
			int num6 = HTKKFile.ColumnLettersToIndex("P");
			int num7 = HTKKFile.ColumnLettersToIndex("S");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AD");
			int num11 = HTKKFile.ColumnLettersToIndex("AO");
			int num12 = HTKKFile.ColumnLettersToIndex("AP");
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cell09c, cell09cMa, cell09b, cell09bMa, cell09a, cell09aMa, null))
			{
				value = 2;
			}
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				bool flag = false;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart - 2, num3 - 1);
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text.Trim();
					string value2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text.Trim();
					double num13 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num7 - 1].Value);
					double num14 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Value);
					double num15 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Value);
					double num16 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Value);
					object value3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Value;
					string item = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Text.Trim();
					if (!flag && tables[i].TableName == "01_TAIN_1" && (!string.IsNullOrEmpty(text) || !string.IsNullOrEmpty(value2) || num13 != 0.0 || num14 != 0.0 || num15 != 0.0 || num16 != 0.0))
					{
						flag = true;
					}
					int num17 = Convert.ToInt32(sheetView.Cells[cellCoPhuLuc[1], cellCoPhuLuc[0]].Value);
					int num18 = Convert.ToInt32(sheetView.Cells[cellCoMaV5[1], cellCoMaV5[0]].Value);
					if (string.IsNullOrEmpty(text) && (num13 > 0.0 || num14 > 0.0 || num16 > 0.0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, msg);
						value = 2;
					}
					if (num17 == 1 && num18 == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, msg2);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text) && (value3 == null || (value3 != null && Convert.ToDouble(value3) == 0.0)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, msg4);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text))
					{
						List<int> list = new List<int>();
						list.Add(num12);
						list.Add(num3);
						list.Add(num8);
						List<string> list2 = new List<string>();
						list2.Add(item);
						list2.Add(text);
						list2.Add(num14.ToString());
						List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, tables[i], list, list2);
						if (rowIndexInSViewByCompares.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, msg3);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, msg3);
							value = 2;
						}
					}
					if (num13 == 0.0 && !string.IsNullOrEmpty(text))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, msg);
						if (value != 2)
						{
							value = 1;
						}
					}
					string valueCol2ID = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text.Trim();
					if (HTKKFormater.isNullString(valueCol2ID) || _dmKSData == null || _dmKSData.Count <= 0 || !(num16.ToNumber() == 0m) || !(num14.ToNumber() != 0m))
					{
						continue;
					}
					bool flag2 = true;
					item item2 = _dmKSData.Find((item x) => x.id.Contains(valueCol2ID));
					if (item2 == null || !flag2 || string.IsNullOrEmpty(item2.giatoithieu) || string.IsNullOrEmpty(item2.giatoida))
					{
						continue;
					}
					if (num14.ToNumber() < item2.giatoithieu.ToNumber() - item2.giatoithieu.ToNumber() * 20m / 100m)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
						value = 2;
					}
					if (num14.ToNumber() > item2.giatoida.ToNumber())
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
						if (value != 2)
						{
							value = 1;
						}
					}
				}
				if (tables[i].TableName == "01_TAIN_1")
				{
					string value4 = sheetView.Cells[dataLocationTemplateRowStart - 2, num3 - 1].Text.Trim();
					if (string.IsNullOrEmpty(value4) && flag)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart - 2, num3 - 1, msg);
						value = 2;
					}
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
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("AL");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("Q");
			int num6 = HTKKFile.ColumnLettersToIndex("U");
			int num7 = HTKKFile.ColumnLettersToIndex("W");
			int num8 = HTKKFile.ColumnLettersToIndex("S");
			int num9 = HTKKFile.ColumnLettersToIndex("X");
			int num10 = HTKKFile.ColumnLettersToIndex("Z");
			int num11 = HTKKFile.ColumnLettersToIndex("AB");
			int num12 = HTKKFile.ColumnLettersToIndex("AA");
			int num13 = HTKKFile.ColumnLettersToIndex("AC");
			int num14 = HTKKFile.ColumnLettersToIndex("AE");
			int num15 = HTKKFile.ColumnLettersToIndex("AF");
			int num16 = HTKKFile.ColumnLettersToIndex("AG");
			int num17 = HTKKFile.ColumnLettersToIndex("AP");
			int num18 = HTKKFile.ColumnLettersToIndex("AO");
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[34, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			foreach (HTKKCMCTableInfo table in ToKhaiInfo.SheetInfos[1].Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1, testColor: true);
					}
				}
				int rowIndex = table.Header.StaticCells[1].CellTemplate.RowIndex;
				foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
				{
					HTKKFile.RemoveValidate(sheetView, rowIndex - 1, staticCell.CellTemplate.ColIndex - 1, testColor: true);
				}
			}
			foreach (HTKKCMCTableInfo table2 in ToKhaiInfo.SheetInfos[1].Tables)
			{
				int rowIndex2 = table2.Header.StaticCells[0].CellTemplate.RowIndex;
				string text3 = sheetView.Cells[rowIndex2 - 1, num4 - 1].Text.Trim();
				List<int> rowIndexInSViewByCompareHeader = GeneralFunctions.GetRowIndexInSViewByCompareHeader(sheetView, ToKhaiInfo.SheetInfos[1].Tables, num4, text3);
				if (string.IsNullOrEmpty(text3.ToString().Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num4 - 1, msg);
					value = 2;
				}
				else if (rowIndexInSViewByCompareHeader.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num4 - 1, msg6);
					value = 2;
				}
				else if (table2.RowsValues.Count < 2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num4 - 1, msg2);
					value = 2;
				}
				else if (table2.RowsValues.Count >= 2)
				{
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, table2, colMaLoai_PL, "01");
					List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, table2, colMaLoai_PL, "02");
					if (rowIndexInSViewByCompare.Count < 1 || rowIndexInSViewByCompare2.Count < 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num4 - 1, msg2);
						value = 2;
					}
				}
				object value2 = sheetView.Cells[rowIndex2 - 1, num5 - 1].Value;
				if (value2 == null || string.IsNullOrEmpty(value2.ToString().Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num5 - 1, msg);
					value = 2;
				}
				else if (value2 != null && !string.IsNullOrEmpty(value2.ToString().Trim()) && value2.ToString().Trim().Length == 5 && !ValidateBase.CheckMaSoThue5So(value2.ToString().Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num5 - 1, msg5);
					value = 2;
				}
				else if (value2 != null && !string.IsNullOrEmpty(value2.ToString().Trim()) && (value2.ToString().Trim().Length == 10 || value2.ToString().Trim().Length == 13 || value2.ToString().Trim().Length == 14) && !ValidateBase.CheckMaSoThue(value2.ToString().Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num5 - 1, msg5);
					value = 2;
				}
				else if (value2 != null && !string.IsNullOrEmpty(value2.ToString().Trim()) && value2.ToString().Trim().Length != 5 && value2.ToString().Trim().Length != 10 && value2.ToString().Trim().Length != 13 && value2.ToString().Trim().Length != 14)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num5 - 1, msg5);
					value = 2;
				}
				double num19 = Convert.ToDouble(sheetView.Cells[rowIndex2 - 1, num15 - 1].Value);
				if (num19 != 100.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num15 - 1, msg3);
					value = 2;
				}
				int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = table2.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string value3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, colMaLoai_PL - 1].Text.Trim();
					string value4 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, colLoai_PL - 1].Text.Trim();
					string text4 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text.Trim();
					string text5 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text.Trim();
					string text6 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text.Trim();
					string text7 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Text.Trim();
					string text8 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Text.Trim();
					string text9 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Text.Trim();
					string text10 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Text.Trim();
					double num20 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num14 - 1].Value);
					double num21 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num15 - 1].Value);
					string text11 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num17 - 1].Text.Trim();
					string item = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num18 - 1].Text.Trim();
					if (string.IsNullOrEmpty(value4) || string.IsNullOrEmpty(value3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, colLoai_PL - 1, msg);
						value = 2;
					}
					if (string.IsNullOrEmpty(text6) && !sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Locked)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, msg);
						value = 2;
					}
					if (string.IsNullOrEmpty(text7) && !sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Locked)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, msg);
						value = 2;
					}
					List<int> rowIndexInSViewByCompare3 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, table2, num2, "01");
					if (rowIndexInSViewByCompare3.Count > 0)
					{
						string text12 = sheetView.Cells[rowIndexInSViewByCompare3[0] - 1, num9 - 1].Text.Trim();
						if (!string.IsNullOrEmpty(text12))
						{
							List<int> list = new List<int>();
							list.Add(num2);
							list.Add(num9);
							List<string> list2 = new List<string>();
							list2.Add("02");
							list2.Add(text12);
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, table2, list, list2);
							if (rowIndexInSViewByCompares.Count > 0)
							{
								foreach (int item2 in rowIndexInSViewByCompares)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num8 - 1, msg4);
									value = 2;
								}
							}
						}
					}
					if (string.IsNullOrEmpty(text9) && !sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Locked)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, msg);
						value = 2;
					}
					if (string.IsNullOrEmpty(text10) && !sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Locked)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, msg);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text11))
					{
						List<int> list = new List<int>();
						list.Add(num17);
						list.Add(num18);
						list.Add(num6);
						list.Add(num8);
						list.Add(num10);
						list.Add(num12);
						List<string> list2 = new List<string>();
						list2.Add(text11);
						list2.Add(item);
						list2.Add(text6);
						list2.Add(text7);
						list2.Add(text9);
						list2.Add(text10);
						List<int> rowIndexInSViewByCompares2 = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, table2, list, list2);
						if (rowIndexInSViewByCompares2.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num4 - 1, msg7);
							HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex2 - 1, num5 - 1, msg7);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, msg7);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, msg7);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, msg7);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, msg7);
							value = 2;
						}
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

	public Dictionary<string, string> GetDMKhoangSanTuMaMo(string maMo, string tenMo)
	{
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Expected O, but got Unknown
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cb: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			string text = FOLDER_DM + "\\DM_MO_KHOANG_SAN.xml";
			if (nodeListDmMoKhoangSan == null)
			{
				XmlDocument val = new XmlDocument();
				if (!File.Exists(text))
				{
					return null;
				}
				val.Load(text);
				nodeListDmMoKhoangSan = val.GetElementsByTagName("Item");
			}
			string valueMo = maMo + "###" + tenMo;
			XmlNode val2 = ((IEnumerable)nodeListDmMoKhoangSan).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode n) => ((XmlNode)n.Attributes["Value"]).InnerText == valueMo));
			if (val2 == null)
			{
				return dictionary;
			}
			foreach (XmlNode childNode in val2.ChildNodes)
			{
				XmlNode val3 = childNode;
				string input = ((XmlNode)val3.Attributes["Value"]).Value.Trim();
				string[] array = Regex.Split(input, "###");
				string key = array[0];
				string value = array[1];
				if (!dictionary.ContainsKey(key))
				{
					dictionary.Add(key, value);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return dictionary;
		}
	}

	public string GetMaKhoangSan(string maMo, string tenKhoangSan)
	{
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Expected O, but got Unknown
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Expected O, but got Unknown
		try
		{
			string text = FOLDER_DM + "\\DM_MO_KHOANG_SAN.xml";
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("dm_mo_khoang_san");
			IEnumerator enumerator = elementsByTagName.GetEnumerator();
			try
			{
				if (enumerator.MoveNext())
				{
					XmlNode val2 = (XmlNode)enumerator.Current;
					XmlNode val3 = ((IEnumerable)val2.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode n) => ((XmlNode)n.Attributes["ma"]).InnerText == maMo));
					if (val3 == null)
					{
						return null;
					}
					XmlNode val4 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode n) => ((XmlNode)n.Attributes["ten"]).InnerText == tenKhoangSan));
					if (val4 == null)
					{
						return null;
					}
					return ((XmlNode)val4.Attributes["ma"]).Value.Trim();
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public Dictionary<string, string> GetDMKhoangSan(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanChiTietTem.Count > 0)
			{
				return dicDMKhoangSanChiTietTem;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[2];
				if (array[^1] == "1")
				{
					dictionary2.Add(key, value2);
					dictionary.Add(key, array);
				}
			}
			dicDMKhoangSanChiTiet = dictionary;
			dicDMKhoangSanChiTietTem = dictionary2;
			return dicDMKhoangSanChiTietTem;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKhoangSanNew(List<PathDm> pathDms, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0375: Unknown result type (might be due to invalid IL or missing references)
		//IL_037c: Expected O, but got Unknown
		//IL_03aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b1: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanChiTietTem.Count > 0)
			{
				return dicDMKhoangSanChiTietTem;
			}
			_dmKSData = new List<item>();
			List<item> list = new List<item>();
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
					if (pathDm.pathDonGiaTinh != null && pathDm.pathDonGiaTinh.Count > 0)
					{
						foreach (string p in pathDm.pathDonGiaTinh)
						{
							List<string> list2 = pathDonGiaTinh;
							Predicate<string> match = (string o) => o.Equals(p);
							if (HTKKFormater.isNullString(list2.Find(match)))
							{
								pathDonGiaTinh.Add(p);
							}
						}
					}
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
					for (int num = 0; num < items.Count; num++)
					{
						item item = items[num];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							list.Add(item);
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							list.Add(item);
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
				foreach (XmlNode item7 in elementsByTagName)
				{
					XmlNode val2 = item7;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array[^1] == "1")
					{
						item item3 = new item();
						item3.id = array[0];
						item3.ma = array[1];
						item3.ten = array[2];
						item3.dvt = array[3];
						item3.thuesuat = array[4];
						list.Add(item3);
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				List<item> dmTaiNguyenForTinh = GeneralFunctions.getDmTaiNguyenForTinh(pathDonGiaTinh, PeriodInfo);
				List<item> list3 = Serializer.DeepCopy(list);
				if (dmTaiNguyenForTinh != null && dmTaiNguyenForTinh.Count > 0)
				{
					foreach (item iT in list)
					{
						Predicate<item> match2 = (item obj) => obj.id.Equals(iT.id);
						List<item> list4 = dmTaiNguyenForTinh.FindAll(match2);
						if (list4 != null && list4.Count > 0)
						{
							item item4 = list3.Find((item obj) => obj.id == iT.id);
							int num2 = list3.FindIndex((item obj) => obj.id == iT.id);
							if (item4 != null)
							{
								list3.Remove(item4);
							}
							for (int num3 = 0; num3 < list4.Count; num3++)
							{
								item item5 = Serializer.DeepCopy(iT);
								item5.cqtgia = list4[num3].cqtgia;
								list3.Insert(num2 + num3, item5);
							}
						}
					}
					list = Serializer.DeepCopy(list3);
				}
				list3 = new List<item>();
				foreach (item item8 in list)
				{
					item item6 = Serializer.DeepCopy(item8);
					string ten = item6.ten;
					List<item> list5 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list5 != null && list5.Count > 1)
					{
						item6 = GeneralFunctions.getTaiNguyenForCombobox(item6, list5);
						if (item6 == null && item8.id == list5[0].id)
						{
							list3.Add(Serializer.DeepCopy(item8));
						}
					}
					if (item6 != null)
					{
						if (!HTKKFormater.isNullString(item6.level) && HTKKFormater.IsNumber(item6.level))
						{
							item6.ten = HTKKFile.addSpace(ConvertHelpers.ConvertToInt(item6.level)) + item6.ten;
						}
						list3.Add(item6);
					}
				}
				for (int num4 = 0; num4 < list3.Count; num4++)
				{
					item item6 = Serializer.DeepCopy(list3[num4]);
					string text = item6.id + "_" + item6.ma + "_" + item6.thuesuat + "_" + item6.giatoithieu + "_" + item6.cqtgia;
					string[] array = new string[8] { text, item6.ma, item6.ten, item6.dvt, item6.thuesuat, "0", "1", item6.matieumuc };
					item6.id = text;
					if (!dictionary.ContainsKey(text))
					{
						string value2 = array[2];
						dictionary2.Add(text, value2);
						dictionary.Add(text, array);
						_dmKSData.Add(item6);
					}
				}
			}
			dicDMKhoangSanChiTiet = dictionary;
			dicDMKhoangSanChiTietTem = dictionary2;
			return dicDMKhoangSanChiTietTem;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<item> GetItemDm_GetDMKhoangSan()
	{
		return _dmKSData;
	}

	public Dictionary<string, string> GetDMDonVi(string pathFile)
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

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("AN");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("AO");
			int colChilden = HTKKFile.ColumnLettersToIndex("AR");
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			int dataLocationTemplateRowStart = tables[0].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < tables[0].RowsValues.Count; i++)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, ToKhaiInfo.SheetInfos[0], 0, dataLocationTemplateRowStart + i, num, colChilden);
			}
			foreach (HTKKCMCTableInfo item in tables)
			{
				int dataLocationTemplateRowStart2 = item.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < item.RowsValues.Count; i++)
				{
					object valueCol2 = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num - 1].Value;
					if (valueCol2 == null)
					{
						continue;
					}
					KeyValuePair<string, string[]> keyValuePair = dicDMKhoangSanChiTiet.FirstOrDefault((KeyValuePair<string, string[]> x) => x.Value.Length > 1 && x.Value[2].ToString() == valueCol2.ToString().Trim());
					if (!string.IsNullOrEmpty(keyValuePair.Key) && keyValuePair.Value.Length > 1)
					{
						sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num2 - 1].Text = keyValuePair.Value[1];
						sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num3 - 1].Text = ((keyValuePair.Value.Length > 7) ? keyValuePair.Value[7] : "");
						if (keyValuePair.Value[5] == "1" || keyValuePair.Value[1] == "1101")
						{
							sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num4 - 1].Locked = false;
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + i - 1, num4 - 1);
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num4 - 1].Locked = true;
							HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart2 + i - 1, num4 - 1);
						}
					}
				}
			}
			for (int num6 = 0; num6 < tables.Count; num6++)
			{
				int dataLocationTemplateRowStart2 = tables[num6].Location.DataLocationTemplateRowStart;
				for (int i = 0; i < tables[num6].RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num5 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num4 - 1].Value;
				}
			}
			if (hTKKCMCSheetInfo2 != null)
			{
				GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo2, -1, -1, col05Header_PL, col05DongCon_PL);
				GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo2, -1, -1, col06Header_PL, col06DongCon_PL);
				GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo2, -1, -1, col10Header_PL, col10DongCon_PL);
			}
			foreach (HTKKCMCTableInfo item2 in tables)
			{
				int dataLocationTemplateRowStart2 = item2.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < item2.RowsValues.Count; i++)
				{
					GeneralFunctions.UpdateFromSViewToSInfo(sheetView, dataLocationTemplateRowStart2 + i, -1, item2.RowsInfo, item2);
				}
			}
			if (hTKKCMCSheetInfo2 == null)
			{
				return;
			}
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo2.Tables)
			{
				int dataLocationTemplateRowStart2 = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					GeneralFunctions.UpdateFromSViewToSInfo(sView, dataLocationTemplateRowStart2 + i, -1, table.RowsInfo, table);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public double getDonGia(string ma)
	{
		double result = 0.0;
		try
		{
			if (!HTKKFormater.isNullString(ma) && _dmKSData != null && _dmKSData.Count > 0)
			{
				item item = _dmKSData.Find((item x) => x.id.Contains(ma));
				if (item != null && item.cqtgia.ToNumber() > 0m)
				{
					result = ConvertHelpers.ConvertToDouble1(item.cqtgia.ToNumber());
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}

	public void TinhTongCot_4_8_9_10(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		int num = HTKKFile.ColumnLettersToIndex("S");
		int num2 = HTKKFile.ColumnLettersToIndex("AG");
		int num3 = HTKKFile.ColumnLettersToIndex("AJ");
		int num4 = HTKKFile.ColumnLettersToIndex("AK");
		string path4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/tongSanLuong";
		string path8 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/tongThuePhaiNop";
		string path9 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/tongThuePhaiNop1";
		string path10 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietTNguyen/tongThuePhaiNop2";
		HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path4));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path8));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path9));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path10));
		double num5 = 0.0;
		double num6 = 0.0;
		double num7 = 0.0;
		double num8 = 0.0;
		List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
		foreach (HTKKCMCTableInfo item in tables)
		{
			int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
			for (int num9 = 0; num9 < item.RowsValues.Count; num9++)
			{
				double num10 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num9 - 1, num - 1].Value);
				num5 += num10;
				double num11 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num9 - 1, num2 - 1].Value);
				num6 += num11;
				double num12 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Value);
				num7 += num12;
				double num13 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num9 - 1, num4 - 1].Value);
				num8 += num13;
			}
		}
		sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num5;
		sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num6;
		sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = num7;
		sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = num8;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc[1], cellCoPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
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
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc[1], cellCoPhuLuc[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockCell(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
		{
			if (!(table.TableName == "01_1_TAIN_1"))
			{
				continue;
			}
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				if (sView.Cells[dataLocationTemplateRowStart + i - 1, colMaLoai_PL - 1].Text.Trim() == "01")
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, col8_PL, locked: true, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, col8_PL, locked: false, deleteData: false);
				}
			}
		}
	}
}
