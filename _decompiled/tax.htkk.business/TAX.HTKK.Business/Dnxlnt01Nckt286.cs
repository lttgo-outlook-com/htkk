using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

internal class Dnxlnt01Nckt286
{
	private int colD = HTKKFile.ColumnLettersToIndex("D");

	private int[] cellPhuLuc = GeneralFunctions.ArrayCellId("BJ_15");

	private int[] cellHoanThue = GeneralFunctions.ArrayCellId("D_19");

	private int[] cellNopThua = GeneralFunctions.ArrayCellId("D_21");

	private int[] cellTraTienNopThua = GeneralFunctions.ArrayCellId("D_30");

	private int[] cellNopThua2_1 = GeneralFunctions.ArrayCellId("F_24");

	private int[] cellNopThua2_2 = GeneralFunctions.ArrayCellId("F_26");

	private int[] cellNopThua2_3 = GeneralFunctions.ArrayCellId("F_28");

	private int[] cellKGTenXa = GeneralFunctions.ArrayCellId("M_10");

	private int[] cellKGMaXa = GeneralFunctions.ArrayCellId("BE_10");

	private int[] cellCQTCuc = GeneralFunctions.ArrayCellId("L_8");

	private int[] cellMaCQTCuc = GeneralFunctions.ArrayCellId("BJ_8");

	private int[] cellCQTChiCuc = GeneralFunctions.ArrayCellId("L_9");

	private int[] cellCQTMaChiCuc = GeneralFunctions.ArrayCellId("BJ_9");

	private int colCt05_Tb1_TK = HTKKFile.ColumnLettersToIndex("S");

	private int colCt05Ma_Tb1_TK = HTKKFile.ColumnLettersToIndex("BJ");

	private int colCt06_1_Tb1_TK = HTKKFile.ColumnLettersToIndex("Y");

	private int colCt06_1_Ma_Tb1_TK = HTKKFile.ColumnLettersToIndex("BK");

	private int colCt06_2_Tb1_TK = HTKKFile.ColumnLettersToIndex("AB");

	private int colCt06_2_Ma_Tb1_TK = HTKKFile.ColumnLettersToIndex("BM");

	private int colCt08_Tb1_TK = HTKKFile.ColumnLettersToIndex("AI");

	private int colCt09_Tb1_TK = HTKKFile.ColumnLettersToIndex("AN");

	private int colCt10_Tb1_TK = HTKKFile.ColumnLettersToIndex("AS");

	private int colCt07_Tb2_TK = HTKKFile.ColumnLettersToIndex("AB");

	private int colCt07Ma_Tb2_TK = HTKKFile.ColumnLettersToIndex("BJ");

	private int colCt12_Tb2_TK = HTKKFile.ColumnLettersToIndex("AY");

	private int colCt13_Tb2_TK = HTKKFile.ColumnLettersToIndex("AZ");

	private int colCt14_Tb2_TK = HTKKFile.ColumnLettersToIndex("BA");

	private string pathNopThuaBangIII = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDeNghiXuLyHoan/TTinNNT/CTietTTinNNT";

	private int colCt07_TbIII_TK = HTKKFile.ColumnLettersToIndex("AB");

	private int colCt08_1_TbIII_TK = HTKKFile.ColumnLettersToIndex("AF");

	private int colCt08_1_Ma_TbIII_TK = HTKKFile.ColumnLettersToIndex("BK");

	private int colCt08_2_TbIII_TK = HTKKFile.ColumnLettersToIndex("AM");

	private int colCt08_2_Ma_TbIII_TK = HTKKFile.ColumnLettersToIndex("BL");

	private int colCt09_1_TbIII_TK = HTKKFile.ColumnLettersToIndex("AQ");

	private int colCt09_1_Ma_TbIII_TK = HTKKFile.ColumnLettersToIndex("BM");

	private int colCt09_2_TbIII_TK = HTKKFile.ColumnLettersToIndex("AV");

	private int colCt09_2_Ma_TbIII_TK = HTKKFile.ColumnLettersToIndex("BN");

	private int[] cellTuKy = GeneralFunctions.ArrayCellId("G_4");

	private int[] cellTuKyIsNumber = GeneralFunctions.ArrayCellId("N_4");

	private int[] cellDenKy = GeneralFunctions.ArrayCellId("I_4");

	private int[] cellDenKyIsNumber = GeneralFunctions.ArrayCellId("O_4");

	private int colKieuKy = HTKKFile.ColumnLettersToIndex("I");

	private int colKyKK = HTKKFile.ColumnLettersToIndex("J");

	private int colKyKKTemp = HTKKFile.ColumnLettersToIndex("T");

	private int colMaKieuKy = HTKKFile.ColumnLettersToIndex("O");

	private ComboBoxCellType comboQuy = new ComboBoxCellType();

	private ComboBoxCellType comboThang = new ComboBoxCellType();

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.Cells[cellPhuLuc[1], cellPhuLuc[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			DefaultLockUnLockNopThua_2_1(sheetView);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (!string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy) && !string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
			{
				sheetView.Cells[cellCQTCuc[1], cellCQTCuc[0]].Value = PeriodInfo.tenCQTCapCuc;
				sheetView.Cells[cellMaCQTCuc[1], cellMaCQTCuc[0]].Value = PeriodInfo.maCQTCapCuc;
				sheetView.Cells[cellCQTChiCuc[1], cellCQTChiCuc[0]].Value = PeriodInfo.tenCQTQuanLy;
				sheetView.Cells[cellCQTMaChiCuc[1], cellCQTMaChiCuc[0]].Value = PeriodInfo.maCQTQuanLy;
			}
			sheetView.Cells[cellKGTenXa[1], cellKGTenXa[0]].Value = ((!string.IsNullOrEmpty(PeriodInfo.TenXa)) ? PeriodInfo.TenXa : "");
			sheetView.Cells[cellKGMaXa[1], cellKGMaXa[0]].Value = ((!string.IsNullOrEmpty(PeriodInfo.MaXa)) ? PeriodInfo.MaXa : "");
			ThayDoiTheoBuTruHoan_BangII_TK(sheetView, deleteText: false);
			bool flag = !(PeriodInfo.maCQTQuanLy == headerInfo.maCQTQuanLy);
			int[] array = GeneralFunctions.ArrayCellId("BL_9");
			sheetView.Cells[array[1], array[0]].Value = (flag ? 1 : 0);
			LockUnLockCQT_II_TK(sheetView, sInfo, PeriodInfo, flag, TTinChung.mst, isSet: false);
			ThayDoiTheoBuTruHoan_BangIII_TK(sheetView, deleteText: false);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			string input = sheetView.Cells[cellTuKy[1], cellTuKy[0]].Text.Trim();
			sheetView.Cells[cellTuKyIsNumber[1], cellTuKyIsNumber[0]].Value = (GeneralFunctions.IsNumeric(input) ? 1 : 0);
			string input2 = sheetView.Cells[cellDenKy[1], cellDenKy[0]].Text.Trim();
			sheetView.Cells[cellDenKyIsNumber[1], cellDenKyIsNumber[0]].Value = (GeneralFunctions.IsNumeric(input2) ? 1 : 0);
			int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, colKyKK - 1].Text.Trim();
				string typeKeKhai = sheetView.Cells[dataLocationTemplateRowStart + i - 1, colMaKieuKy - 1].Text.Trim();
				GeneralFunctions.SetCellTypeTheoKieuKy(sheetView, typeKeKhai, dataLocationTemplateRowStart + i, colKyKK);
				sheetView.Cells[dataLocationTemplateRowStart + i - 1, colMaKieuKy - 1].Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			iCol = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			iRow = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			int[] array = GeneralFunctions.ArrayCellId("D_19");
			bool flag = Convert.ToBoolean(sheetView.Cells[cellHoanThue[1], cellHoanThue[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellNopThua[1], cellNopThua[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[cellTraTienNopThua[1], cellTraTienNopThua[0]].Value);
			if (iRow == cellHoanThue[1] + 1 && flag)
			{
				sheetView.Cells[cellNopThua[1], cellNopThua[0]].Value = false;
				sheetView.Cells[cellTraTienNopThua[1], cellTraTienNopThua[0]].Value = false;
			}
			if (iRow == cellNopThua[1] + 1 && flag2)
			{
				sheetView.Cells[cellHoanThue[1], cellHoanThue[0]].Value = false;
				sheetView.Cells[cellTraTienNopThua[1], cellTraTienNopThua[0]].Value = false;
			}
			if (iRow == cellTraTienNopThua[1] + 1 && flag3)
			{
				sheetView.Cells[cellHoanThue[1], cellHoanThue[0]].Value = false;
				sheetView.Cells[cellNopThua[1], cellNopThua[0]].Value = false;
			}
			DefaultLockUnLockNopThua_2_1(sheetView);
			if (iRow >= cellNopThua[1] + 1 && iRow <= cellNopThua2_3[1] + 1)
			{
				LockUnLockHoanNopThua(sheetView, iRow);
			}
			if (iRow >= cellHoanThue[1] + 1 && iRow <= cellTraTienNopThua[1] + 1)
			{
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
				bool isVangLai = !(PeriodInfo.maCQTQuanLy == headerInfo.maCQTQuanLy);
				LockUnLockCQT_II_TK(sheetView, sInfo, PeriodInfo, isVangLai, TTinChung.mst, isSet: true);
			}
			int num = GeneralFunctions.DetectIndexOfTable(sInfo, iRow);
			if (iRow >= cellHoanThue[1] + 1 && iRow <= cellTraTienNopThua[1] + 1)
			{
				ThayDoiTheoBuTruHoan_BangII_TK(sheetView, deleteText: true);
				ThayDoiTheoBuTruHoan_BangIII_TK(sheetView, deleteText: true);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			string value = sheetView.Cells[(int)iRow - 1, colCt07Ma_Tb2_TK - 1].Text.Trim();
			string text = sheetView.Cells[(int)iRow - 1, colCt08_1_Ma_TbIII_TK - 1].Text.Trim();
			if (string.IsNullOrEmpty(value) || iCol != colCt07_Tb2_TK || hTKKCMCSheetInfo.Tables[num].TableName == "01_DNXLNT_3")
			{
			}
			if ((iCol != colCt07_TbIII_TK && iCol != colCt08_1_TbIII_TK && iCol != colCt08_2_TbIII_TK) || num == 2)
			{
			}
			if (num == 0 && (iCol == colCt05_Tb1_TK || iCol == colCt08_Tb1_TK || iCol == colCt09_Tb1_TK))
			{
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(colCt05_Tb1_TK, colD);
				dictionary.Add(colCt05Ma_Tb1_TK, colCt05Ma_Tb1_TK);
				dictionary.Add(colCt08_Tb1_TK, colCt08_Tb1_TK);
				dictionary.Add(colCt09_Tb1_TK, colCt09_Tb1_TK);
				dictionary.Add(colCt10_Tb1_TK, colCt10_Tb1_TK);
				Dictionary<string, int> checkExists = new Dictionary<string, int>();
				GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			}
			if (num == 2 && (iCol == colCt07_Tb2_TK || iCol == colCt12_Tb2_TK || iCol == colCt13_Tb2_TK))
			{
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(colCt07_Tb2_TK, colD);
				dictionary.Add(colCt07Ma_Tb2_TK, colCt07Ma_Tb2_TK);
				dictionary.Add(colCt12_Tb2_TK, colCt12_Tb2_TK);
				dictionary.Add(colCt13_Tb2_TK, colCt13_Tb2_TK);
				dictionary.Add(colCt14_Tb2_TK, colCt14_Tb2_TK);
				Dictionary<string, int> checkExists = new Dictionary<string, int>();
				GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[2], 3, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDeNghiXuLyHoan/TTinDeNghi/HinhThucHoanTra/chuyenKhoan");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDeNghiXuLyHoan/TTinDeNghi/HinhThucHoanTra/tienMat");
			bool flag = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo.CellTemplate.ColIndex && flag)
			{
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = false;
			}
			if (iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo2.CellTemplate.ColIndex && flag2)
			{
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = false;
			}
			if ((iRow == hTKKCMCCellInfo.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo.CellTemplate.ColIndex) || (iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo2.CellTemplate.ColIndex))
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
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
			if (iCol == cellTuKy[0] + 1 && iRow == cellTuKy[1] + 1)
			{
				string input = sheetView.Cells[cellTuKy[1], cellTuKy[0]].Text.Trim();
				sheetView.Cells[cellTuKyIsNumber[1], cellTuKyIsNumber[0]].Value = (GeneralFunctions.IsNumeric(input) ? 1 : 0);
			}
			if (iCol == cellDenKy[0] + 1 && iRow == cellDenKy[1] + 1)
			{
				string input2 = sheetView.Cells[cellDenKy[1], cellDenKy[0]].Text.Trim();
				sheetView.Cells[cellDenKyIsNumber[1], cellDenKyIsNumber[0]].Value = (GeneralFunctions.IsNumeric(input2) ? 1 : 0);
			}
			if (iCol == colKieuKy)
			{
				sheetView.Cells[(int)iRow - 1, colKyKK - 1].Text = "";
				string text = sheetView.Cells[(int)iRow - 1, colMaKieuKy - 1].Text;
				GeneralFunctions.SetCellTypeTheoKieuKy(sheetView, text, (int)iRow, colKyKK);
			}
			if (iCol == colKieuKy || iCol == colKyKK)
			{
				TaoKyThueNgayThang(sheetView, (int)iRow);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void initComboboxQuy()
	{
		comboQuy = new ComboBoxCellType();
		string[] array = new string[4];
		for (int i = 1; i <= 4; i++)
		{
			array[i - 1] = "Q" + i;
		}
		comboQuy.ItemData = array;
		comboQuy.Items = array;
	}

	private void initComboboxThang()
	{
		comboThang = new ComboBoxCellType();
		string[] array = new string[12];
		for (int i = 1; i <= 12; i++)
		{
			array[i - 1] = i.ToString();
		}
		comboThang.ItemData = array;
		comboThang.Items = array;
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/tenXa");
			XmlNode val2 = xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maXa");
			val.InnerText = ((val != null) ? PeriodInfo.TenXa : "");
			val2.InnerText = ((val2 != null) ? PeriodInfo.MaXa : "");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int num = HTKKFile.ColumnLettersToIndex("AI");
			int num2 = HTKKFile.ColumnLettersToIndex("AN");
			int num3 = HTKKFile.ColumnLettersToIndex("D");
			int num4 = HTKKFile.ColumnLettersToIndex("AY");
			int num5 = HTKKFile.ColumnLettersToIndex("AZ");
			int num6 = HTKKFile.ColumnLettersToIndex("AW");
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num7 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				double num8 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num7 + i, num - 1].Value);
				double num9 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num7 + i, num2 - 1].Value);
				HTKKFile.RemoveValidate(sheetView, num7 + i, num2 - 1);
				if (num9 < num8)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, num2 - 1, text);
					value = 2;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
			int num10 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text5 = sheetView.Cells[num10 + i, num3 - 1].Text;
				double num11 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num10 + i, num4 - 1].Value);
				double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num10 + i, num5 - 1].Value);
				string text6 = sheetView.Cells[num10 + i, num6 - 1].Text;
				HTKKFile.RemoveValidate(sheetView, num10 + i, num3 - 1);
				HTKKFile.RemoveValidate(sheetView, num10 + i, num5 - 1);
				if (!string.IsNullOrEmpty(text5) && !ValidateBase.CheckMaSoThue(text5.Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, num3 - 1, text2);
					value = 2;
				}
				if (num12 > num11)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, num3 - 1, text3);
					value = 2;
				}
				if (string.IsNullOrEmpty(text6) && (!HTKKFormater.IsDateTime(text6, "dd/MM/yyyy", out var dtDate) || !(dtDate >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num10 + i, num6 - 1, text4);
					value = 2;
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
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			string message = HTKKFile.GetMessage("0071");
			if (hTKKCMCSheetInfo != null)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				if (hTKKCMCTableInfo != null)
				{
					int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
					{
						string text = sheetView.Cells[num2 + i, num].Text;
						HTKKFile.RemoveValidate(sheetView, num2 + i, num);
						if (!string.IsNullOrEmpty(text) && !HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out var _))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, num, message);
							value = 2;
						}
					}
				}
			}
			switch (value)
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiTheoBuTruHoan_BangII_TK(SheetView sView, bool deleteText)
	{
		TextCellType cellType = new TextCellType();
		int[] array = GeneralFunctions.ArrayCellId("D_19");
		int num = HTKKFile.ColumnLettersToIndex("N");
		bool flag = Convert.ToBoolean(sView.Cells[array[1], array[0]].Value);
		HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (flag)
		{
			Dictionary<string, string> dic = DicCellTypeChuong();
			string text = "1001";
			Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(sView, dataLocationTemplateRowStart, colCt05_Tb1_TK);
			string text2 = (dMFromCombobox.ContainsKey(text) ? dMFromCombobox[text] : "");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart + i, num, dic, deleteText);
				sView.Cells[dataLocationTemplateRowStart + i - 1, colCt05_Tb1_TK - 1].Text = text2;
				sView.Cells[dataLocationTemplateRowStart + i - 1, colCt05Ma_Tb1_TK - 1].Text = text;
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colCt05_Tb1_TK, locked: true, deleteData: false);
			}
			return;
		}
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].CellType = cellType;
			if (deleteText)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text = "";
			}
			GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colCt05_Tb1_TK, locked: false, deleteData: false);
		}
	}

	public Dictionary<string, string> DicCellTypeChuong()
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		dictionary.Add("557", "557");
		dictionary.Add("757", "757");
		return dictionary;
	}

	public void ThayDoiTheoBuTruHoan_BangIII_TK(SheetView sView, bool deleteText)
	{
		TextCellType cellType = new TextCellType();
		int[] array = GeneralFunctions.ArrayCellId("D_19");
		int num = HTKKFile.ColumnLettersToIndex("X");
		bool flag = Convert.ToBoolean(sView.Cells[array[1], array[0]].Value);
		int index = _ToKhaiInfo.SheetInfos[0].Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == pathNopThuaBangIII);
		HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[index];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (flag)
		{
			Dictionary<string, string> dic = DicCellTypeChuong();
			for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
			{
				GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart + num2, num, dic, deleteText);
			}
			return;
		}
		for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
		{
			sView.Cells[dataLocationTemplateRowStart + num2 - 1, num - 1].CellType = cellType;
			if (deleteText)
			{
				sView.Cells[dataLocationTemplateRowStart + num2 - 1, num - 1].Text = "";
			}
			GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + num2, colCt05_Tb1_TK, locked: false, deleteData: false);
		}
	}

	public void ThayDoiCQTTheoTTinChung_BangIII_TK(SheetView sView, HTKKCMCTTinChungInfo TTinChung, int iRow)
	{
		int index = _ToKhaiInfo.SheetInfos[0].Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == pathNopThuaBangIII);
		HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[index];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
		if (iRow < 0)
		{
			for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
			{
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, headerInfo.tenCQTCapCuc, headerInfo.maCQTCapCuc, headerInfo.tenCQTQuanLy, headerInfo.maCQTQuanLy, dataLocationTemplateRowStart + num, colCt08_1_TbIII_TK, colCt08_1_Ma_TbIII_TK, colCt08_2_TbIII_TK, colCt08_2_Ma_TbIII_TK);
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + num, colCt08_1_TbIII_TK, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + num, colCt08_2_TbIII_TK, locked: false, deleteData: false);
			}
		}
		else
		{
			GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, headerInfo.tenCQTCapCuc, headerInfo.maCQTCapCuc, headerInfo.tenCQTQuanLy, headerInfo.maCQTQuanLy, iRow, colCt08_1_TbIII_TK, colCt08_1_Ma_TbIII_TK, colCt08_2_TbIII_TK, colCt08_2_Ma_TbIII_TK);
			GeneralFunctions.LockUnLockCol(sView, iRow, colCt08_1_TbIII_TK, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, iRow, colCt08_2_TbIII_TK, locked: false, deleteData: false);
		}
	}

	public void ThayDoiTinhTheoCQT_BangIII_TK(SheetView sView, bool awaySet, int iRow)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = _ToKhaiInfo.SheetInfos[0];
		int index = hTKKCMCSheetInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == pathNopThuaBangIII);
		HTKKCMCTableInfo hTKKCMCTableInfo = _ToKhaiInfo.SheetInfos[0].Tables[index];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow < 0)
		{
			for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
			{
				string maCQTCapCuc_TSC = sView.Cells[dataLocationTemplateRowStart + num - 1, colCt08_1_Ma_TbIII_TK - 1].Text.Trim();
				string maCQTNoiNop_TSC = sView.Cells[dataLocationTemplateRowStart + num - 1, colCt08_2_Ma_TbIII_TK - 1].Text.Trim();
				GeneralFunctions.ThayDoiTinhTheoCQTDangBang2(maCQTCapCuc_TSC, maCQTNoiNop_TSC, sView, colCt09_1_TbIII_TK, colCt09_1_Ma_TbIII_TK, colCt09_2_TbIII_TK, colCt09_2_Ma_TbIII_TK, dataLocationTemplateRowStart + num, awaySet);
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + num, colCt09_1_TbIII_TK, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + num, colCt09_2_TbIII_TK, locked: false, deleteData: false);
			}
		}
		else
		{
			string maCQTCapCuc_TSC = sView.Cells[iRow - 1, colCt08_1_Ma_TbIII_TK - 1].Text.Trim();
			string maCQTNoiNop_TSC = sView.Cells[iRow - 1, colCt08_2_Ma_TbIII_TK - 1].Text.Trim();
			GeneralFunctions.ThayDoiTinhTheoCQTDangBang2(maCQTCapCuc_TSC, maCQTNoiNop_TSC, sView, colCt09_1_TbIII_TK, colCt09_1_Ma_TbIII_TK, colCt09_2_TbIII_TK, colCt09_2_Ma_TbIII_TK, iRow, awaySet);
			GeneralFunctions.LockUnLockCol(sView, iRow, colCt09_1_TbIII_TK, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, iRow, colCt09_2_TbIII_TK, locked: false, deleteData: false);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			if (activeRowIndex == 4)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_DNXLNT_TT80.xls");
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView activeSheet = winFpSpread.ActiveSheet;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
			{
				string text = activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colKyKKTemp - 1].Text.Trim();
				string typeKeKhai = activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colMaKieuKy - 1].Text.Trim();
				GeneralFunctions.SetCellTypeTheoKieuKy(activeSheet, typeKeKhai, dataLocationTemplateRowStart + i, colKyKK);
				activeSheet.Cells[dataLocationTemplateRowStart + i - 1, colKyKK - 1].Text = text;
			}
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				TaoKyThueNgayThang(activeSheet, dataLocationTemplateRowStart + i);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TaoKyThueNgayThang(SheetView sViewPL, int iRow)
	{
		string text = sViewPL.Cells[iRow - 1, colMaKieuKy - 1].Text.Trim();
		string text2 = sViewPL.Cells[iRow - 1, colKyKK - 1].Text.Trim();
		if (string.IsNullOrEmpty(text2))
		{
			return;
		}
		switch (text)
		{
		case "M":
			if (!(ConvertHelpers.ConvertToDateTime("01/" + text2, "dd/MM/yyyy") == DateTime.MinValue))
			{
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("P") - 1].Text = DateTime.ParseExact("01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text = DateTime.ParseExact("01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
					.ToString("dd/MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = text2;
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("S") - 1].Text = text2;
			}
			break;
		case "D":
			if (!(ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") == DateTime.MinValue))
			{
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("P") - 1].Text = text2;
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text = text2;
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy").ToString("MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("S") - 1].Text = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy").ToString("MM/yyyy");
			}
			break;
		case "Y":
			if (!(ConvertHelpers.ConvertToDateTime("01/01/" + text2, "dd/MM/yyyy") == DateTime.MinValue))
			{
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("P") - 1].Text = DateTime.ParseExact("01/01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text = DateTime.ParseExact("01/12/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
					.ToString("dd/MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = "01/" + text2;
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("S") - 1].Text = "12/" + text2;
			}
			break;
		case "Q":
			if (!(ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text2.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text2.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue))
			{
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("P") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text2.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text2.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text2.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text2.Split(new char[1] { '/' })[1]), 1, 1).AddMonths(3).AddDays(-1.0)
					.ToString("dd/MM/yyyy");
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = text2;
				sViewPL.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("S") - 1].Text = text2;
			}
			break;
		}
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			List<int> list = new List<int>();
			list.Add(colCt05_Tb1_TK);
			list.Add(colCt05Ma_Tb1_TK);
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			bool locked = sheetView.Cells[dataLocationTemplateRowStart - 1, colCt05_Tb1_TK - 1].Locked;
			if (locked)
			{
				GeneralFunctions.SetAllValueInRowSame_Table(sheetView, hTKKCMCTableInfo, list, locked, isDelete: false);
			}
			List<int> list2 = new List<int>();
			list2.Add(colCt06_1_Tb1_TK);
			list2.Add(colCt06_1_Ma_Tb1_TK);
			list2.Add(colCt06_2_Tb1_TK);
			list2.Add(colCt06_2_Ma_Tb1_TK);
			bool locked2 = sheetView.Cells[dataLocationTemplateRowStart - 1, colCt06_1_Tb1_TK - 1].Locked;
			if (locked2)
			{
				GeneralFunctions.SetAllValueInRowSame_Table(sheetView, hTKKCMCTableInfo, list2, locked2, isDelete: false);
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void DefaultLockUnLockNopThua_2_1(SheetView sView)
	{
		try
		{
			if (Convert.ToBoolean(sView.Cells[cellNopThua[1], cellNopThua[0]].Value))
			{
				GeneralFunctions.LockUnLockCol(sView, cellNopThua2_1[1] + 1, cellNopThua2_1[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, cellNopThua2_2[1] + 1, cellNopThua2_2[0] + 1, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, cellNopThua2_3[1] + 1, cellNopThua2_3[0] + 1, locked: false, deleteData: false);
			}
			else
			{
				GeneralFunctions.LockUnLockCol(sView, cellNopThua2_1[1] + 1, cellNopThua2_1[0] + 1, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sView, cellNopThua2_2[1] + 1, cellNopThua2_2[0] + 1, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sView, cellNopThua2_3[1] + 1, cellNopThua2_3[0] + 1, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnLockHoanNopThua(SheetView sView, int iRow)
	{
		try
		{
			bool flag = Convert.ToBoolean(sView.Cells[cellNopThua[1], cellNopThua[0]].Value);
			bool flag2 = Convert.ToBoolean(sView.Cells[cellNopThua2_1[1], cellNopThua2_1[0]].Value);
			bool flag3 = Convert.ToBoolean(sView.Cells[cellNopThua2_2[1], cellNopThua2_2[0]].Value);
			bool flag4 = Convert.ToBoolean(sView.Cells[cellNopThua2_3[1], cellNopThua2_3[0]].Value);
			if (iRow == cellNopThua[1] + 1 && flag)
			{
				sView.Cells[cellNopThua2_1[1], cellNopThua2_1[0]].Value = true;
				sView.Cells[cellNopThua2_2[1], cellNopThua2_2[0]].Value = false;
				sView.Cells[cellNopThua2_3[1], cellNopThua2_3[0]].Value = false;
			}
			if (iRow == cellNopThua2_1[1] + 1 && flag2)
			{
				sView.Cells[cellNopThua2_2[1], cellNopThua2_2[0]].Value = false;
				sView.Cells[cellNopThua2_3[1], cellNopThua2_3[0]].Value = false;
			}
			if (iRow == cellNopThua2_2[1] + 1 && flag3)
			{
				sView.Cells[cellNopThua2_1[1], cellNopThua2_1[0]].Value = false;
				sView.Cells[cellNopThua2_3[1], cellNopThua2_3[0]].Value = false;
			}
			if (iRow == cellNopThua2_3[1] + 1 && flag4)
			{
				sView.Cells[cellNopThua2_1[1], cellNopThua2_1[0]].Value = false;
				sView.Cells[cellNopThua2_2[1], cellNopThua2_2[0]].Value = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnLockCQT_II_TK(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo, bool isVangLai, string mst, bool isSet)
	{
		try
		{
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(mst);
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			int dataLocationTemplateRowStart = sInfo.Tables[0].Location.DataLocationTemplateRowStart;
			bool flag = Convert.ToBoolean(sView.Cells[cellHoanThue[1], cellHoanThue[0]].Value);
			bool flag2 = Convert.ToBoolean(sView.Cells[cellNopThua[1], cellNopThua[0]].Value);
			bool flag3 = Convert.ToBoolean(sView.Cells[cellTraTienNopThua[1], cellTraTienNopThua[0]].Value);
			bool flag4 = Convert.ToBoolean(sView.Cells[cellNopThua2_1[1], cellNopThua2_1[0]].Value);
			bool flag5 = Convert.ToBoolean(sView.Cells[cellNopThua2_2[1], cellNopThua2_2[0]].Value);
			bool flag6 = Convert.ToBoolean(sView.Cells[cellNopThua2_3[1], cellNopThua2_3[0]].Value);
			bool locked;
			if (isVangLai)
			{
				locked = !flag5;
				string tenCQTCapCuc = "";
				if (!string.IsNullOrEmpty(PeriodInfo.maCQTCapCuc))
				{
					tenCQTCapCuc = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(PeriodInfo.maCQTCapCuc, isGetFullName: false);
				}
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string value = sView.Cells[dataLocationTemplateRowStart + i - 1, colCt06_1_Tb1_TK - 1].Text.Trim();
					if (isSet || (!isSet && string.IsNullOrEmpty(value)))
					{
						GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, tenCQTCapCuc, PeriodInfo.maCQTCapCuc, PeriodInfo.tenCQTQuanLy, PeriodInfo.maCQTQuanLy, dataLocationTemplateRowStart + i, colCt06_1_Tb1_TK, colCt06_1_Ma_Tb1_TK, colCt06_2_Tb1_TK, colCt06_2_Ma_Tb1_TK);
					}
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colCt06_1_Tb1_TK, locked, deleteData: false);
					GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colCt06_2_Tb1_TK, locked, deleteData: false);
				}
				return;
			}
			locked = ((!flag && !flag3 && !flag5) ? true : false);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart + i - 1, colCt06_1_Tb1_TK - 1].Text.Trim();
				if (isSet || (!isSet && string.IsNullOrEmpty(value)))
				{
					GeneralFunctions.ThayDoiCQTTheoTinhDynamic2(sView, headerInfo.tenCQTCapCuc, headerInfo.maCQTCapCuc, headerInfo.tenCQTQuanLy, headerInfo.maCQTQuanLy, dataLocationTemplateRowStart + i, colCt06_1_Tb1_TK, colCt06_1_Ma_Tb1_TK, colCt06_2_Tb1_TK, colCt06_2_Ma_Tb1_TK);
				}
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colCt06_1_Tb1_TK, locked, deleteData: false);
				GeneralFunctions.LockUnLockCol(sView, dataLocationTemplateRowStart + i, colCt06_2_Tb1_TK, locked, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void kyThueUpsert00(HTKKCMCSheetInfo sInfo)
	{
		for (int i = 0; i < sInfo.Tables[0].RowsValues.Count; i++)
		{
		}
	}
}
