using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class CcttTmdtTT80Nckt260
{
	private int colCt07_TK = HTKKFile.ColumnLettersToIndex("D");

	private int colCt08_TK = HTKKFile.ColumnLettersToIndex("G");

	private int colCt09_TK = HTKKFile.ColumnLettersToIndex("L");

	private int colCt10_TK = HTKKFile.ColumnLettersToIndex("O");

	private int colCt11_TK = HTKKFile.ColumnLettersToIndex("R");

	private int colCt12_TK = HTKKFile.ColumnLettersToIndex("U");

	private int colCt13_TK = HTKKFile.ColumnLettersToIndex("X");

	private int colCt14_TK = HTKKFile.ColumnLettersToIndex("AA");

	private int colCt15_TK = HTKKFile.ColumnLettersToIndex("AB");

	private int colCt16_TK = HTKKFile.ColumnLettersToIndex("AE");

	private int colCt17_TK = HTKKFile.ColumnLettersToIndex("AH");

	private int colCt18_TK = HTKKFile.ColumnLettersToIndex("AK");

	private int colCt19_TK = HTKKFile.ColumnLettersToIndex("AN");

	private int colCt20_TK = HTKKFile.ColumnLettersToIndex("AO");

	private int colMaLoaiHinh_TK = HTKKFile.ColumnLettersToIndex("AX");

	private int colMaLoaiGiayTo_TK = HTKKFile.ColumnLettersToIndex("AY");

	private int colMaTrangThai_TK = HTKKFile.ColumnLettersToIndex("AZ");

	private int colMaQLSan_PL01_1 = HTKKFile.ColumnLettersToIndex("D");

	private int colTenLoaiHinh_PL01_1 = HTKKFile.ColumnLettersToIndex("G");

	private int colEmail_PL01_1 = HTKKFile.ColumnLettersToIndex("J");

	private int colPhone_PL01_1 = HTKKFile.ColumnLettersToIndex("M");

	private int colQuocGia1_PL01_1 = HTKKFile.ColumnLettersToIndex("P");

	private int colTinh1_PL01_1 = HTKKFile.ColumnLettersToIndex("R");

	private int colHuyen1_PL01_1 = HTKKFile.ColumnLettersToIndex("U");

	private int colPhuong1_PL01_1 = HTKKFile.ColumnLettersToIndex("X");

	private int colSoNha1_PL01_1 = HTKKFile.ColumnLettersToIndex("AA");

	private int colQuocGia2_PL01_1 = HTKKFile.ColumnLettersToIndex("AC");

	private int colTinh2_PL01_1 = HTKKFile.ColumnLettersToIndex("AE");

	private int colHuyen2_PL01_1 = HTKKFile.ColumnLettersToIndex("AH");

	private int colPhuong2_PL01_1 = HTKKFile.ColumnLettersToIndex("AK");

	private int colSoNha2_PL01_1 = HTKKFile.ColumnLettersToIndex("AN");

	private int colTrangThai_PL01_1 = HTKKFile.ColumnLettersToIndex("AP");

	private int colMaLoaiHinh_PL01_1 = HTKKFile.ColumnLettersToIndex("AZ");

	private int colMaQuocGia1_PL01_1 = HTKKFile.ColumnLettersToIndex("BA");

	private int colMaTinh1_PL01_1 = HTKKFile.ColumnLettersToIndex("BB");

	private int colMaHuyen1_PL01_1 = HTKKFile.ColumnLettersToIndex("BC");

	private int colMaPhuong1_PL01_1 = HTKKFile.ColumnLettersToIndex("BD");

	private int colMaQuocGia2_PL01_1 = HTKKFile.ColumnLettersToIndex("BE");

	private int colMaTinh2_PL01_1 = HTKKFile.ColumnLettersToIndex("BF");

	private int colMaHuyen2_PL01_1 = HTKKFile.ColumnLettersToIndex("BG");

	private int colMaPhuong2_PL01_1 = HTKKFile.ColumnLettersToIndex("BH");

	private int colMaTrangThai_PL01_1 = HTKKFile.ColumnLettersToIndex("BI");

	private int colMaQLSan_PL01_2 = HTKKFile.ColumnLettersToIndex("C");

	private int colMaNganhHang_PL01_2 = HTKKFile.ColumnLettersToIndex("D");

	private int colLoaiHangHoa_PL01_2 = HTKKFile.ColumnLettersToIndex("E");

	private int colTrangThai_PL01_2 = HTKKFile.ColumnLettersToIndex("F");

	private int colMaTrangThai_PL01_2 = HTKKFile.ColumnLettersToIndex("T");

	private int colMaQLSan_PL01_3 = HTKKFile.ColumnLettersToIndex("C");

	private int colSoTaiKhoan_PL01_3 = HTKKFile.ColumnLettersToIndex("D");

	private int colChuTaiKhoan_PL01_3 = HTKKFile.ColumnLettersToIndex("E");

	private int colMaNganHang_PL01_3 = HTKKFile.ColumnLettersToIndex("F");

	private int colTenNganHang_PL01_3 = HTKKFile.ColumnLettersToIndex("G");

	private int colTrangThai_PL01_3 = HTKKFile.ColumnLettersToIndex("H");

	private int colMaTrangThai_PL01_3 = HTKKFile.ColumnLettersToIndex("T");

	private string maLoaiHinh_ToChuc_VN = "01";

	private string maLoaiHinh_CaNhan_VN = "02";

	private string maLoaiHinh_ToChuc_NN = "03";

	private string maLoaiHinh_CaNhan_NN = "04";

	private string maLoaiGiayTo_CCCD = "01";

	private string maLoaiGiayTo_CMDN = "02";

	private string maLoaiGiayTo_HoChieu = "03";

	private string maTrangThaiThayThe = "04";

	private string tenTrangThaiThayThe = "Thay thế";

	private string maTrangThaiLanDau = "01";

	private string tenTrangThaiLanDau = "Thêm mới";

	private DateTime ngayDauQuy = DateTime.Now;

	private DateTime ngayCuoiQuy = DateTime.Now;

	private string tableName1_PL01_1 = "01_1_CCTT_TMDT_1";

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

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			ngayDauQuy = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(PeriodInfo.Quy), ConvertHelpers.ConvertToInt(PeriodInfo.Nam), 1, 1);
			ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(PeriodInfo.Quy), ConvertHelpers.ConvertToInt(PeriodInfo.Nam), 1, 1);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart - 1 + i, colCt20_TK - 1].Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					sView.Cells[dataLocationTemplateRowStart - 1 + i, colTrangThai_PL01_1 - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart - 1 + i, colTrangThai_PL01_2 - 1].Locked = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					sView.Cells[dataLocationTemplateRowStart - 1 + i, colTrangThai_PL01_3 - 1].Locked = true;
				}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("AK");
			int num2 = HTKKFile.ColumnLettersToIndex("AN");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num3 = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			string text = sheetView.Cells[(int)iRow - 1, colMaLoaiHinh_TK - 1].Text.Trim();
			if (num3 >= 0 && iCol == colCt09_TK && string.IsNullOrEmpty(sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim()))
			{
				SetDefaultQuy(sheetView, PeriodInfo, (int)iRow);
			}
			if (iCol == colCt07_TK && string.IsNullOrEmpty(sheetView.Cells[(int)iRow - 1, colCt20_TK - 1].Text.Trim()))
			{
				SetDefaultThayThe(sheetView, PeriodInfo, (int)iRow, colCt20_TK, colMaTrangThai_TK);
				SetDefaultLanDau(sheetView, PeriodInfo, (int)iRow, colCt20_TK, colMaTrangThai_TK);
			}
			if (num3 >= 0)
			{
				GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
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
			int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (num >= 0)
			{
				string text = sheetView.Cells[(int)iRow - 1, colMaLoaiHinh_PL01_1 - 1].Text.Trim();
				if (iCol == colMaQLSan_PL01_1 && string.IsNullOrEmpty(sheetView.Cells[(int)iRow - 1, colTrangThai_PL01_1 - 1].Text.Trim()))
				{
					SetDefaultThayThe(sheetView, PeriodInfo, (int)iRow, colTrangThai_PL01_1, colMaTrangThai_PL01_1);
					SetDefaultLanDau(sheetView, PeriodInfo, (int)iRow, colTrangThai_PL01_1, colMaTrangThai_PL01_1);
				}
				if (num >= 0)
				{
					GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[0].RowsInfo, hTKKCMCSheetInfo.Tables[0]);
					GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[1].RowsInfo, hTKKCMCSheetInfo.Tables[1]);
				}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[2];
			if (iCol == colMaQLSan_PL01_2 && string.IsNullOrEmpty(sheetView.Cells[(int)iRow - 1, colTrangThai_PL01_2 - 1].Text.Trim()))
			{
				SetDefaultThayThe(sheetView, PeriodInfo, (int)iRow, colTrangThai_PL01_2, colMaTrangThai_PL01_2);
				SetDefaultLanDau(sheetView, PeriodInfo, (int)iRow, colTrangThai_PL01_2, colMaTrangThai_PL01_2);
			}
			int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (num >= 0)
			{
				GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[0].RowsInfo, hTKKCMCSheetInfo.Tables[0]);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			SheetView sheetView = winFpSpread.Sheets[3];
			if (iCol == colMaQLSan_PL01_3 && string.IsNullOrEmpty(sheetView.Cells[(int)iRow - 1, colTrangThai_PL01_3 - 1].Text.Trim()))
			{
				SetDefaultThayThe(sheetView, PeriodInfo, (int)iRow, colTrangThai_PL01_3, colMaTrangThai_PL01_3);
				SetDefaultLanDau(sheetView, PeriodInfo, (int)iRow, colTrangThai_PL01_3, colMaTrangThai_PL01_3);
			}
			int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (num >= 0)
			{
				GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[0].RowsInfo, hTKKCMCSheetInfo.Tables[0]);
				GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[1].RowsInfo, hTKKCMCSheetInfo.Tables[1]);
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
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_CCTT_TMDT.xls");
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
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_CCTT_TMDT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_2_CCTT_TMDT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_3_CCTT_TMDT.xls");
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29"));
			string message12 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			string message13 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31"));
			string message14 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int indexColCt7 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt07_TK);
			int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt08_TK);
			int index2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt09_TK);
			int index3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt10_TK);
			int index4 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt11_TK);
			int index5 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt12_TK);
			int index6 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt13_TK);
			int index7 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt14_TK);
			int index8 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt15_TK);
			int index9 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt16_TK);
			int index10 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt17_TK);
			int index11 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt18_TK);
			int index12 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt19_TK);
			int index13 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt20_TK);
			int index14 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaLoaiHinh_TK);
			int index15 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaLoaiGiayTo_TK);
			int index16 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaTrangThai_TK);
			for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					if (sheetView.Cells[dataLocationTemplateRowStart + num - 1, cell.CellTemplate.ColIndex - 1].Locked && HTKKFile.ColumnLettersToIndex("B") != cell.CellTemplate.ColIndex)
					{
						HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + num - 1, cell.CellTemplate.ColIndex - 1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num - 1, cell.CellTemplate.ColIndex - 1);
					}
				}
				string valueColCt7 = hTKKCMCTableInfo.RowsValues[num].RowValue[indexColCt7];
				string value2 = hTKKCMCTableInfo.RowsValues[num].RowValue[index];
				string value3 = hTKKCMCTableInfo.RowsValues[num].RowValue[index2];
				string value4 = hTKKCMCTableInfo.RowsValues[num].RowValue[index3];
				string text = hTKKCMCTableInfo.RowsValues[num].RowValue[index4];
				string value5 = hTKKCMCTableInfo.RowsValues[num].RowValue[index5];
				string text2 = hTKKCMCTableInfo.RowsValues[num].RowValue[index6];
				string value6 = hTKKCMCTableInfo.RowsValues[num].RowValue[index7];
				string value7 = hTKKCMCTableInfo.RowsValues[num].RowValue[index8];
				string value8 = hTKKCMCTableInfo.RowsValues[num].RowValue[index9];
				string value9 = hTKKCMCTableInfo.RowsValues[num].RowValue[index10];
				string text3 = hTKKCMCTableInfo.RowsValues[num].RowValue[index11];
				string text4 = hTKKCMCTableInfo.RowsValues[num].RowValue[index12];
				string value10 = hTKKCMCTableInfo.RowsValues[num].RowValue[index13];
				string text5 = hTKKCMCTableInfo.RowsValues[num].RowValue[index14];
				string text6 = hTKKCMCTableInfo.RowsValues[num].RowValue[index15];
				string value11 = hTKKCMCTableInfo.RowsValues[num].RowValue[index16];
				if (PeriodInfo.LoaiToKhai == "B" && !GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo.RowsValues[num].RowValue.ToArray()))
				{
					continue;
				}
				if (string.IsNullOrEmpty(valueColCt7))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt07_TK - 1, GeneralFunctions.FomartTenChiTieu("7", "", "") + message, num + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(valueColCt7))
				{
					int num2 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue x) => x.RowValue[indexColCt7] == valueColCt7);
					if (num2 > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt07_TK - 1, message2, num + 1);
						value = 2;
					}
				}
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt08_TK - 1, GeneralFunctions.FomartTenChiTieu("8", "", "") + message, num + 1);
					value = 2;
				}
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt09_TK - 1, GeneralFunctions.FomartTenChiTieu("9", "", "") + message, num + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(value3) && string.IsNullOrEmpty(text5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt09_TK - 1, message3, num + 1);
					value = 2;
				}
				if ((text5 == maLoaiHinh_CaNhan_VN || text5 == maLoaiHinh_CaNhan_NN) && string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt10_TK - 1, GeneralFunctions.FomartTenChiTieu("10", "", "") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				else if ((text5 == maLoaiHinh_CaNhan_VN || text5 == maLoaiHinh_CaNhan_NN) && !string.IsNullOrEmpty(value4) && string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt10_TK - 1, message3);
					if (value != 2)
					{
						value = 1;
					}
				}
				else if (text5 == maLoaiHinh_CaNhan_VN && !string.IsNullOrEmpty(value4) && !string.IsNullOrEmpty(text6) && text6 == maLoaiGiayTo_HoChieu)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt10_TK - 1, message7);
					if (value != 2)
					{
						value = 1;
					}
				}
				else if ((text5 == maLoaiHinh_CaNhan_NN || text5 == maLoaiHinh_ToChuc_NN) && !string.IsNullOrEmpty(value4) && !string.IsNullOrEmpty(text6) && text6 != maLoaiGiayTo_HoChieu)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt10_TK - 1, message9);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value4) && string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt11_TK - 1, GeneralFunctions.FomartTenChiTieu("11", "", "") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value4) && !string.IsNullOrEmpty(text) && text6 == maLoaiGiayTo_CMDN && text.Length != 9 && text.Length != 12)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt11_TK - 1, message10);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value4) && !string.IsNullOrEmpty(text) && text6 == maLoaiGiayTo_CCCD && text.Length != 12)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt11_TK - 1, message11);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (text5 == maLoaiHinh_ToChuc_VN && string.IsNullOrEmpty(value5) && string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt12_TK - 1, message4);
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt13_TK - 1, message4);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt13_TK - 1, message5);
					value = 2;
				}
				if (!string.IsNullOrEmpty(value6) && Convert.ToDouble(value6) < 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt14_TK - 1, message6);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (text5 == maLoaiHinh_ToChuc_VN && Convert.ToDouble(value7) < 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt15_TK - 1, message6);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (Convert.ToDouble(value7) > Convert.ToDouble(value6))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt15_TK - 1, message8);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value8) && Convert.ToDouble(value8) < 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt16_TK - 1, message6);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value9) && Convert.ToDouble(value9) < 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt17_TK - 1, message6);
					if (value != 2)
					{
						value = 1;
					}
				}
				DateTime dateTime = default(DateTime);
				DateTime dateTime2 = default(DateTime);
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt18_TK - 1, GeneralFunctions.FomartTenChiTieu("18", "", "") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(text3))
				{
					try
					{
						dateTime = ((text3.IndexOf("/") > 0) ? DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture) : DateTime.ParseExact(text3, "ddMMyyyy", CultureInfo.CurrentCulture));
					}
					catch
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt18_TK - 1, message14, num + 1);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text3) && (DateTime.Compare(dateTime, ngayDauQuy) < 0 || DateTime.Compare(ngayCuoiQuy, dateTime) < 0))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt18_TK - 1, message12, num + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text4))
				{
					try
					{
						dateTime2 = ((text4.IndexOf("/") > 0) ? DateTime.ParseExact(text4, "dd/MM/yyyy", CultureInfo.CurrentCulture) : DateTime.ParseExact(text4, "ddMMyyyy", CultureInfo.CurrentCulture));
					}
					catch
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt19_TK - 1, message14, num + 1);
						value = 2;
					}
				}
				if (string.IsNullOrEmpty(text4))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt19_TK - 1, GeneralFunctions.FomartTenChiTieu("19", "", "") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(text4) && (DateTime.Compare(dateTime2, ngayDauQuy) < 0 || DateTime.Compare(ngayCuoiQuy, dateTime2) < 0))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt19_TK - 1, message12, num + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text4) && DateTime.Compare(dateTime, dateTime2) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt18_TK - 1, message13, num + 1);
					value = 2;
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colCt19_TK - 1, message13, num + 1);
					value = 2;
				}
				if (string.IsNullOrEmpty(value10))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt20_TK - 1, GeneralFunctions.FomartTenChiTieu("20", "", "") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value10) && string.IsNullOrEmpty(value11))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colCt20_TK - 1, message3);
					if (value != 2)
					{
						value = 1;
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
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35"));
			bool flag = true;
			int num = 0;
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				int indexColCt7_TK = ToKhaiInfo.SheetInfos[0].Tables[0].RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt07_TK);
				int indexColMaLoaiHinh_TK = ToKhaiInfo.SheetInfos[0].Tables[0].RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaLoaiHinh_TK);
				int indexMaQLSan_PL01_1 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaQLSan_PL01_1);
				int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTenLoaiHinh_PL01_1);
				int index2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colEmail_PL01_1);
				int index3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colPhone_PL01_1);
				int index4 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colQuocGia1_PL01_1);
				int index5 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTinh1_PL01_1);
				int index6 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colHuyen1_PL01_1);
				int index7 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colPhuong1_PL01_1);
				int index8 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colSoNha1_PL01_1);
				int index9 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colQuocGia2_PL01_1);
				int index10 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTinh2_PL01_1);
				int index11 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colHuyen2_PL01_1);
				int index12 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colPhuong2_PL01_1);
				int index13 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colSoNha2_PL01_1);
				int index14 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTrangThai_PL01_1);
				int index15 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaLoaiHinh_PL01_1);
				int index16 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaQuocGia1_PL01_1);
				int index17 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaTinh1_PL01_1);
				int index18 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaHuyen1_PL01_1);
				int index19 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaPhuong1_PL01_1);
				int index20 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaQuocGia2_PL01_1);
				int index21 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaTinh2_PL01_1);
				int index22 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaHuyen2_PL01_1);
				int index23 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaPhuong2_PL01_1);
				int index24 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaTrangThai_PL01_1);
				for (int num2 = 0; num2 < table.RowsValues.Count; num2++)
				{
					foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
					{
						if (sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1].Locked && HTKKFile.ColumnLettersToIndex("B") != cell.CellTemplate.ColIndex)
						{
							HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1);
						}
					}
					bool flag2 = false;
					flag2 = GeneralFunctions.HaveToValidateRow(table.RowsInfo, table.RowsValues[num2].RowValue.ToArray());
					if (PeriodInfo.LoaiToKhai == "B" && !flag2)
					{
						continue;
					}
					if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count == 1 && hTKKCMCSheetInfo.Tables[1].RowsValues.Count == 1)
					{
						bool flag3 = GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[0].RowsInfo, hTKKCMCSheetInfo.Tables[0].RowsValues[0].RowValue.ToArray());
						bool flag4 = GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[1].RowsInfo, hTKKCMCSheetInfo.Tables[1].RowsValues[0].RowValue.ToArray());
						if (!flag3 && !flag4)
						{
							int dataLocationTemplateRowStart2 = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
							int dataLocationTemplateRowStart3 = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart;
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 1, colMaQLSan_PL01_1 - 1, message11, num2 + 1);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart3 - 1, colMaQLSan_PL01_1 - 1, message11, num2 + 1);
							value = 2;
							flag = false;
						}
					}
					else if (!GeneralFunctions.HaveToValidateRow(table.RowsInfo, table.RowsValues[0].RowValue.ToArray()))
					{
						continue;
					}
					if (!flag2 && flag)
					{
						continue;
					}
					string valueMaQLSan_PL01_1 = table.RowsValues[num2].RowValue[indexMaQLSan_PL01_1];
					string value2 = table.RowsValues[num2].RowValue[index];
					string value3 = table.RowsValues[num2].RowValue[index2];
					string value4 = table.RowsValues[num2].RowValue[index3];
					string value5 = table.RowsValues[num2].RowValue[index4];
					string value6 = table.RowsValues[num2].RowValue[index5];
					string value7 = table.RowsValues[num2].RowValue[index6];
					string value8 = table.RowsValues[num2].RowValue[index7];
					string value9 = table.RowsValues[num2].RowValue[index8];
					string value10 = table.RowsValues[num2].RowValue[index9];
					string value11 = table.RowsValues[num2].RowValue[index10];
					string value12 = table.RowsValues[num2].RowValue[index11];
					string value13 = table.RowsValues[num2].RowValue[index12];
					string value14 = table.RowsValues[num2].RowValue[index13];
					string value15 = table.RowsValues[num2].RowValue[index14];
					string valueMaLoaiHinh_PL01_1 = table.RowsValues[num2].RowValue[index15];
					string value16 = table.RowsValues[num2].RowValue[index16];
					string text = table.RowsValues[num2].RowValue[index17];
					string text2 = table.RowsValues[num2].RowValue[index18];
					string text3 = table.RowsValues[num2].RowValue[index19];
					string value17 = table.RowsValues[num2].RowValue[index20];
					string text4 = table.RowsValues[num2].RowValue[index21];
					string text5 = table.RowsValues[num2].RowValue[index22];
					string text6 = table.RowsValues[num2].RowValue[index23];
					string value18 = table.RowsValues[num2].RowValue[index24];
					if (!string.IsNullOrEmpty(value5) && string.IsNullOrEmpty(value16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colQuocGia1_PL01_1 - 1, message2, num2 + 1);
						value = 2;
					}
					if (table.TableName == tableName1_PL01_1)
					{
						if (!string.IsNullOrEmpty(value6) && string.IsNullOrEmpty(text))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh1_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value7) && string.IsNullOrEmpty(text2))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen1_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value8) && string.IsNullOrEmpty(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong1_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value10) && string.IsNullOrEmpty(value17))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colQuocGia2_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value11) && string.IsNullOrEmpty(text4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh2_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value12) && string.IsNullOrEmpty(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen2_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value13) && string.IsNullOrEmpty(text6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong2_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(value15) && string.IsNullOrEmpty(value18))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTrangThai_PL01_1 - 1, message2, num2 + 1);
							value = 2;
						}
					}
					if (string.IsNullOrEmpty(valueMaQLSan_PL01_1))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colMaQLSan_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("5", "", "") + message, num2 + 1);
						value = 2;
					}
					if (!string.IsNullOrEmpty(valueMaQLSan_PL01_1))
					{
						int num3 = table.RowsValues.Count((HTKKCMCRowValue x) => x.RowValue[indexMaQLSan_PL01_1] == valueMaQLSan_PL01_1);
						if (num3 > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colMaQLSan_PL01_1 - 1, message8, num2 + 1);
							value = 2;
						}
					}
					if (string.IsNullOrEmpty(value2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTenLoaiHinh_PL01_1 - 1, GeneralFunctions.FomartTenChiTieu("6", "", "") + message, num2 + 1);
						value = 2;
					}
					if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(valueMaLoaiHinh_PL01_1))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTenLoaiHinh_PL01_1 - 1, message2, num2 + 1);
						value = 2;
					}
					if (!string.IsNullOrEmpty(valueMaLoaiHinh_PL01_1) && !string.IsNullOrEmpty(valueMaQLSan_PL01_1))
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
						int num3 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue x) => x.RowValue[indexColCt7_TK].Trim() == valueMaQLSan_PL01_1 && x.RowValue[indexColMaLoaiHinh_TK].Trim() == valueMaLoaiHinh_PL01_1);
						if (num3 <= 0)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTenLoaiHinh_PL01_1 - 1, message10);
							if (value != 2)
							{
								value = 1;
							}
						}
					}
					if (string.IsNullOrEmpty(value3) && string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colEmail_PL01_1 - 1, message3);
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhone_PL01_1 - 1, message3);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (valueMaLoaiHinh_PL01_1 == maLoaiHinh_ToChuc_VN || valueMaLoaiHinh_PL01_1 == maLoaiHinh_ToChuc_NN)
					{
						if (string.IsNullOrEmpty(value5) && table.TableName != tableName1_PL01_1)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colQuocGia1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("9", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value6))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("10", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value7))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("11", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value8))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("12", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value9))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoNha1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("13", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1 && (!string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value13) || !string.IsNullOrEmpty(value14)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colQuocGia2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("14", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value11) && (!string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value13) || !string.IsNullOrEmpty(value14) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("15", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value12) && (!string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value13) || !string.IsNullOrEmpty(value14) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("16", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value13) && (!string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value14) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("17", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value14) && (!string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value13) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoNha2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("18", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
					}
					else
					{
						if (string.IsNullOrEmpty(value6) && string.IsNullOrEmpty(value7) && string.IsNullOrEmpty(value8) && string.IsNullOrEmpty(value9) && string.IsNullOrEmpty(value11) && string.IsNullOrEmpty(value12) && string.IsNullOrEmpty(value13) && string.IsNullOrEmpty(value14))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh1_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh2_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen1_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen2_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong1_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong2_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoNha1_PL01_1 - 1, message7);
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoNha2_PL01_1 - 1, message7);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value5) && table.TableName != tableName1_PL01_1 && (!string.IsNullOrEmpty(value6) || !string.IsNullOrEmpty(value7) || !string.IsNullOrEmpty(value8) || !string.IsNullOrEmpty(value9)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colQuocGia1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("9", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value6) && (!string.IsNullOrEmpty(value7) || !string.IsNullOrEmpty(value8) || !string.IsNullOrEmpty(value9) || (!string.IsNullOrEmpty(value5) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("10", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value7) && (!string.IsNullOrEmpty(value6) || !string.IsNullOrEmpty(value8) || !string.IsNullOrEmpty(value9) || (!string.IsNullOrEmpty(value5) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("11", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value8) && (!string.IsNullOrEmpty(value7) || !string.IsNullOrEmpty(value6) || !string.IsNullOrEmpty(value9) || (!string.IsNullOrEmpty(value5) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("12", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value9) && (!string.IsNullOrEmpty(value7) || !string.IsNullOrEmpty(value6) || !string.IsNullOrEmpty(value8) || (!string.IsNullOrEmpty(value5) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoNha1_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("13", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1 && (!string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value13) || !string.IsNullOrEmpty(value14)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colQuocGia2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("14", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value11) && (!string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value13) || !string.IsNullOrEmpty(value14) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("14", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value12) && (!string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value13) || !string.IsNullOrEmpty(value14) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("16", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value13) && (!string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value14) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("17", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
						if (string.IsNullOrEmpty(value14) && (!string.IsNullOrEmpty(value12) || !string.IsNullOrEmpty(value11) || !string.IsNullOrEmpty(value13) || (!string.IsNullOrEmpty(value10) && table.TableName != tableName1_PL01_1)))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoNha2_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("18", "", "") + message);
							if (value != 2)
							{
								value = 1;
							}
						}
					}
					if (string.IsNullOrEmpty(value18))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTrangThai_PL01_1 - 1, TenBang_PL01_1(num) + GeneralFunctions.FomartTenChiTieu("19", "", "") + message);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (table.TableName == tableName1_PL01_1 && !string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && !GeneralFunctions.ValidateDiaBan(text, text2, text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh1_PL01_1 - 1, message9, num2 + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen1_PL01_1 - 1, message9, num2 + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong1_PL01_1 - 1, message9, num2 + 1);
						value = 2;
					}
					if (table.TableName == tableName1_PL01_1 && !string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text5) && !string.IsNullOrEmpty(text6) && !GeneralFunctions.ValidateDiaBan(text4, text5, text6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colTinh2_PL01_1 - 1, message9, num2 + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colHuyen2_PL01_1 - 1, message9, num2 + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colPhuong2_PL01_1 - 1, message9, num2 + 1);
						value = 2;
					}
				}
				num++;
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
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaQLSan_PL01_2);
			int index2 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaNganhHang_PL01_2);
			int index3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colLoaiHangHoa_PL01_2);
			int index4 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTrangThai_PL01_2);
			int index5 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaTrangThai_PL01_2);
			for (int num = 0; num < hTKKCMCTableInfo.RowsValues.Count; num++)
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num - 1, cell.CellTemplate.ColIndex - 1);
				}
				string value2 = hTKKCMCTableInfo.RowsValues[num].RowValue[index];
				string value3 = hTKKCMCTableInfo.RowsValues[num].RowValue[index2];
				string value4 = hTKKCMCTableInfo.RowsValues[num].RowValue[index3];
				string value5 = hTKKCMCTableInfo.RowsValues[num].RowValue[index4];
				string value6 = hTKKCMCTableInfo.RowsValues[num].RowValue[index5];
				if (PeriodInfo.LoaiToKhai == "B" && !GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo.RowsValues[num].RowValue.ToArray()))
				{
					continue;
				}
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num - 1, colMaQLSan_PL01_2 - 1, GeneralFunctions.FomartTenChiTieu("", "", "Mã quản lý trên sàn giao dịch TMĐT - ") + message, num + 1);
					value = 2;
				}
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colMaNganhHang_PL01_2 - 1, GeneralFunctions.FomartTenChiTieu("", "", "Mã ngành hàng, dịch vụ kinh doanh theo Sàn - ") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colLoaiHangHoa_PL01_2 - 1, GeneralFunctions.FomartTenChiTieu("", "", "Loại hàng hóa tương ứng với mã ngành hàng, dịch vụ của Sàn - ") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colTrangThai_PL01_2 - 1, GeneralFunctions.FomartTenChiTieu("", "", "Trạng thái - ") + message);
					if (value != 2)
					{
						value = 1;
					}
				}
				if (!string.IsNullOrEmpty(value5) && string.IsNullOrEmpty(value6))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num - 1, colTrangThai_PL01_2 - 1, message2);
					if (value != 2)
					{
						value = 1;
					}
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
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35"));
			bool flag = true;
			int num = 0;
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				int indexColMaQLSan_PL01_3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaQLSan_PL01_3);
				int indexColSoTaiKhoan_PL01_3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colSoTaiKhoan_PL01_3);
				int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colChuTaiKhoan_PL01_3);
				int index2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaNganHang_PL01_3);
				int index3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTenNganHang_PL01_3);
				int index4 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colTrangThai_PL01_3);
				int index5 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaTrangThai_PL01_3);
				for (int num2 = 0; num2 < table.RowsValues.Count; num2++)
				{
					foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
					{
						if (sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1].Locked && HTKKFile.ColumnLettersToIndex("B") != cell.CellTemplate.ColIndex)
						{
							HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1);
						}
					}
					bool flag2 = false;
					flag2 = GeneralFunctions.HaveToValidateRow(table.RowsInfo, table.RowsValues[num2].RowValue.ToArray());
					if (PeriodInfo.LoaiToKhai == "B" && !flag2)
					{
						continue;
					}
					if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count == 1 && hTKKCMCSheetInfo.Tables[1].RowsValues.Count == 1)
					{
						bool flag3 = GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[0].RowsInfo, hTKKCMCSheetInfo.Tables[0].RowsValues[0].RowValue.ToArray());
						bool flag4 = GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[1].RowsInfo, hTKKCMCSheetInfo.Tables[1].RowsValues[0].RowValue.ToArray());
						if (!flag3 && !flag4)
						{
							int dataLocationTemplateRowStart2 = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
							int dataLocationTemplateRowStart3 = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart;
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart2 - 1, colMaQLSan_PL01_3 - 1, message4, num2 + 1);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart3 - 1, colMaQLSan_PL01_3 - 1, message4, num2 + 1);
							value = 2;
							flag = false;
						}
					}
					else if (!GeneralFunctions.HaveToValidateRow(table.RowsInfo, table.RowsValues[0].RowValue.ToArray()))
					{
						continue;
					}
					if (!flag2 && flag)
					{
						continue;
					}
					string valueColMaQLSan_PL01_3 = table.RowsValues[num2].RowValue[indexColMaQLSan_PL01_3];
					string valueColSoTaiKhoan_PL01_3 = table.RowsValues[num2].RowValue[indexColSoTaiKhoan_PL01_3];
					string value2 = table.RowsValues[num2].RowValue[index];
					string value3 = table.RowsValues[num2].RowValue[index2];
					string value4 = table.RowsValues[num2].RowValue[index3];
					string value5 = table.RowsValues[num2].RowValue[index4];
					string value6 = table.RowsValues[num2].RowValue[index5];
					if (string.IsNullOrEmpty(valueColMaQLSan_PL01_3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colMaQLSan_PL01_3 - 1, TenBang_PL01_3(num) + GeneralFunctions.FomartTenChiTieu("", "", "Mã quản lý trên sàn giao dịch TMĐT - ") + message, num2 + 1);
						value = 2;
					}
					if (string.IsNullOrEmpty(valueColSoTaiKhoan_PL01_3))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoTaiKhoan_PL01_3 - 1, TenBang_PL01_3(num) + GeneralFunctions.FomartTenChiTieu("", "", "Số tài khoản - ") + message);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (!string.IsNullOrEmpty(valueColMaQLSan_PL01_3) && !string.IsNullOrEmpty(valueColSoTaiKhoan_PL01_3))
					{
						IEnumerable<HTKKCMCRowValue> enumerable = table.RowsValues.Where((HTKKCMCRowValue x) => x.RowValue[indexColMaQLSan_PL01_3] == valueColMaQLSan_PL01_3 && x.RowValue[indexColSoTaiKhoan_PL01_3] == valueColSoTaiKhoan_PL01_3);
						if (enumerable != null && enumerable.Count() > 1)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colSoTaiKhoan_PL01_3 - 1, message3);
							if (value != 2)
							{
								value = 1;
							}
						}
					}
					if (string.IsNullOrEmpty(value2))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colChuTaiKhoan_PL01_3 - 1, TenBang_PL01_3(num) + GeneralFunctions.FomartTenChiTieu("", "", "Chủ tài khoản - ") + message);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (string.IsNullOrEmpty(value3))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colMaNganHang_PL01_3 - 1, TenBang_PL01_3(num) + GeneralFunctions.FomartTenChiTieu("", "", "Mã ngân hàng - ") + message);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTenNganHang_PL01_3 - 1, TenBang_PL01_3(num) + GeneralFunctions.FomartTenChiTieu("", "", "Tên ngân hàng - ") + message);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTrangThai_PL01_3 - 1, TenBang_PL01_3(num) + GeneralFunctions.FomartTenChiTieu("", "", "Trạng thái - ") + message);
						if (value != 2)
						{
							value = 1;
						}
					}
					if (!string.IsNullOrEmpty(value5) && string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + num2 - 1, colTrangThai_PL01_3 - 1, message2);
						if (value != 2)
						{
							value = 1;
						}
					}
				}
				num++;
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

	public string TenBang_PL01_1(int tableIndex)
	{
		return tableIndex switch
		{
			0 => "Mục I. Thông tin chung nhà cung cấp trong nước - ", 
			1 => "Mục II. Thông tin chung nhà cung cấp nước ngoài - ", 
			_ => "", 
		};
	}

	public string TenBang_PL01_3(int tableIndex)
	{
		return tableIndex switch
		{
			0 => "Mục I. Thông tin tài khoản nhà cung cấp trong nước - ", 
			1 => "Mục II. Thông tin tài khoản nhà cung cấp nước ngoài - ", 
			_ => "", 
		};
	}

	public void SetDefaultQuy(SheetView sView, HTKKPeriodInfo PeriodInfo, int iRow)
	{
		sView.Cells[iRow - 1, colCt18_TK - 1].Text = ngayDauQuy.ToString("dd/MM/yyyy");
		sView.Cells[iRow - 1, colCt19_TK - 1].Text = ngayCuoiQuy.ToString("dd/MM/yyyy");
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		string text = "";
		string text2 = "";
		int activeSheetIndex = winFpSpread.ActiveSheetIndex;
		if (PeriodInfo.LoaiToKhai == "C")
		{
			text = tenTrangThaiLanDau;
			text2 = maTrangThaiLanDau;
		}
		else
		{
			text = tenTrangThaiThayThe;
			text2 = maTrangThaiThayThe;
		}
		HTKKCMCSheetInfo hTKKCMCSheetInfo;
		if (winFpSpread.ActiveSheetIndex == 0)
		{
			hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[activeSheetIndex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colCt20_TK - 1].Text = text;
				winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colMaTrangThai_TK - 1].Text = text2;
				GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colCt20_TK, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
				GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colMaTrangThai_TK, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
			}
		}
		if (winFpSpread.ActiveSheetIndex == 1)
		{
			hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			if (hTKKCMCSheetInfo != null)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table.RowsValues.Count; i++)
					{
						string value = table.RowsValues[i].RowValue[1];
						if (!string.IsNullOrEmpty(value))
						{
							winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colTrangThai_PL01_1 - 1].Text = text;
							winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colMaTrangThai_PL01_1 - 1].Text = text2;
							GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colTrangThai_PL01_1, table.RowsInfo, table);
							GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colMaTrangThai_PL01_1, table.RowsInfo, table);
						}
					}
				}
			}
		}
		if (winFpSpread.ActiveSheetIndex == 2)
		{
			hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colTrangThai_PL01_2 - 1].Text = text;
				winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colMaTrangThai_PL01_2 - 1].Text = text2;
				GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colTrangThai_PL01_2, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
				GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colMaTrangThai_PL01_2, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
			}
		}
		if (winFpSpread.ActiveSheetIndex != 3)
		{
			return;
		}
		hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
		foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo.Tables)
		{
			int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table2.RowsValues.Count; i++)
			{
				string value = table2.RowsValues[i].RowValue[1];
				if (!string.IsNullOrEmpty(value))
				{
					winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colTrangThai_PL01_3 - 1].Text = text;
					winFpSpread.Sheets[activeSheetIndex].Cells[dataLocationTemplateRowStart - 1 + i, colMaTrangThai_PL01_3 - 1].Text = text2;
					GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colTrangThai_PL01_3, table2.RowsInfo, table2);
					GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[activeSheetIndex], dataLocationTemplateRowStart + i, colMaTrangThai_PL01_3, table2.RowsInfo, table2);
				}
			}
		}
	}

	public void LockUnLockCols_TK(SheetView sView, HTKKCMCTableInfo table)
	{
		if (table == null)
		{
			return;
		}
		int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaLoaiHinh_TK);
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int num = 0; num < table.RowsValues.Count; num++)
		{
			string text = table.RowsValues[num].RowValue[index];
			if (text == maLoaiHinh_ToChuc_VN)
			{
				GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colCt10_TK, locked: true, deleteData: true, "");
				GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colCt11_TK, locked: true, deleteData: true, "");
				sView.Cells[dataLocationTemplateRowStart + num - 1, colMaLoaiGiayTo_TK - 1].Text = "";
			}
			else
			{
				GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colCt10_TK, locked: false, deleteData: false, "");
				GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colCt11_TK, locked: false, deleteData: false, "");
			}
		}
	}

	public void LockUnLockCols_PL01_1(SheetView sView, List<HTKKCMCTableInfo> tables)
	{
		foreach (HTKKCMCTableInfo table in tables)
		{
			if (table == null)
			{
				continue;
			}
			int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colMaLoaiHinh_PL01_1);
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int num = 0; num < table.RowsValues.Count; num++)
			{
				string text = table.RowsValues[num].RowValue[index];
				if (text == maLoaiHinh_ToChuc_VN || text == maLoaiHinh_ToChuc_NN)
				{
					if (table.TableName != tableName1_PL01_1 && text == maLoaiHinh_ToChuc_NN)
					{
						GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colQuocGia2_PL01_1, locked: true, deleteData: true, "");
					}
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colTinh2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colHuyen2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colPhuong2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colSoNha2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colMaTinh2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colMaHuyen2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colMaPhuong2_PL01_1, locked: true, deleteData: true, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colSoNha2_PL01_1, locked: true, deleteData: true, "");
				}
				else
				{
					if (table.TableName != tableName1_PL01_1 && text == maLoaiHinh_ToChuc_NN)
					{
						GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colQuocGia2_PL01_1, locked: false, deleteData: false, "");
					}
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colTinh2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colHuyen2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colPhuong2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colSoNha2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colMaTinh2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colMaHuyen2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colMaPhuong2_PL01_1, locked: false, deleteData: false, "");
					GeneralFunctions.LockUnLockCol2(sView, dataLocationTemplateRowStart + num, colSoNha2_PL01_1, locked: false, deleteData: false, "");
				}
			}
		}
	}

	public void SetDefaultThayThe(SheetView sView, HTKKPeriodInfo PeriodInfo, int iRow, int colTen, int colMa)
	{
		if (PeriodInfo.LoaiToKhai == "T")
		{
			sView.Cells[iRow - 1, colTen - 1].Text = tenTrangThaiThayThe;
			sView.Cells[iRow - 1, colMa - 1].Text = maTrangThaiThayThe;
		}
	}

	public void SetDefaultLanDau(SheetView sView, HTKKPeriodInfo PeriodInfo, int iRow, int colTen, int colMa)
	{
		if (PeriodInfo.LoaiToKhai == "C")
		{
			sView.Cells[iRow - 1, colTen - 1].Text = tenTrangThaiLanDau;
			sView.Cells[iRow - 1, colMa - 1].Text = maTrangThaiLanDau;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			for (int i = 0; i < ToKhaiInfo.SheetInfos.Count - 1; i++)
			{
				GeneralFunctions.fpSpReindexingSheetTemplate2(winFpSpread.Sheets[i], ToKhaiInfo.SheetInfos[i]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
