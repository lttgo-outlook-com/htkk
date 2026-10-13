using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class CnkdTmdt01Nckt282
{
	private string tk_pathCt04_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThu/ct09";

	private string tk_pathCt04_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThu/ct10";

	private string tk_pathCt04_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThu/ct11";

	private string tk_pathCt04_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThu/ct12";

	private string tk_pathCt04_5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongCong/ct13_1";

	private string tk_pathCt05_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGT/ct09";

	private string tk_pathCt05_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGT/ct10";

	private string tk_pathCt05_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGT/ct11";

	private string tk_pathCt05_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGT/ct12";

	private string tk_pathCt05_5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongCong/ct13_2";

	private string tk_pathCt06_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCN/ct09";

	private string tk_pathCt06_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCN/ct10";

	private string tk_pathCt06_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCN/ct11";

	private string tk_pathCt06_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCN/ct12";

	private string tk_pathCt06_5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongCong/ct13_3";

	private string tk_pathCt07_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTru/ct09";

	private string tk_pathCt07_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTru/ct10";

	private string tk_pathCt07_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTru/ct11";

	private string tk_pathCt07_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTru/ct12";

	private string tk_pathCt07_5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongCong/ct13_4";

	private string pathTable0 = "PLuc_01_1_BK_CNKD_TMDT/BKeCNKD-NguoiVN/CTietCNKD";

	private string pl01_1_pathCt01 = "path_stt";

	private string pl01_1_pathCt02 = "ct02";

	private string pl01_1_pathCt04 = "ct04";

	private string pl01_1_pathCt05 = "ct05";

	private string pl01_1_pathCt06 = "ct06";

	private string pl01_1_pathCt07 = "ct07";

	private string pl01_1_pathCt08Ma = "ct08_ma";

	private string pl01_1_pathCt08Ten = "ct08_ten";

	private string pl01_1_pathMaGiayTo = "ma_giayto";

	private string pl01_1_pathTenGiayTo = "ten_giayto";

	private string pl01_1_pathCt09 = "ct09";

	private string pl01_1_pathTotal09_HangHoa = "PLuc_01_1_BK_CNKD_TMDT/TongCongHangHoa/ct13";

	private string pl01_1_pathTotal09_DichVu = "PLuc_01_1_BK_CNKD_TMDT/TongCongDichVu/ct13";

	private string pl01_1_pathTotal09_VanTai = "PLuc_01_1_BK_CNKD_TMDT/TongCongVanTai/ct13";

	private string pl01_1_pathTotal09_Khac = "PLuc_01_1_BK_CNKD_TMDT/TongCongKinhDoanhKhac/ct13";

	private string pl01_1_pathCt10 = "ct10";

	private string pl01_1_pathTotal10_HangHoa = "PLuc_01_1_BK_CNKD_TMDT/TongCongHangHoa/ct14";

	private string pl01_1_pathTotal10_DichVu = "PLuc_01_1_BK_CNKD_TMDT/TongCongDichVu/ct14";

	private string pl01_1_pathTotal10_VanTai = "PLuc_01_1_BK_CNKD_TMDT/TongCongVanTai/ct14";

	private string pl01_1_pathTotal10_Khac = "PLuc_01_1_BK_CNKD_TMDT/TongCongKinhDoanhKhac/ct14";

	private string pl01_1_pathCt11 = "ct11";

	private string pl01_1_pathTotal11_HangHoa = "PLuc_01_1_BK_CNKD_TMDT/TongCongHangHoa/ct15";

	private string pl01_1_pathTotal11_DichVu = "PLuc_01_1_BK_CNKD_TMDT/TongCongDichVu/ct15";

	private string pl01_1_pathTotal11_VanTai = "PLuc_01_1_BK_CNKD_TMDT/TongCongVanTai/ct15";

	private string pl01_1_pathTotal11_Khac = "PLuc_01_1_BK_CNKD_TMDT/TongCongKinhDoanhKhac/ct15";

	private string pl01_1_pathCt12 = "ct12";

	private string pl01_1_pathTotal12_HangHoa = "PLuc_01_1_BK_CNKD_TMDT/TongCongHangHoa/ct16";

	private string pl01_1_pathTotal12_DichVu = "PLuc_01_1_BK_CNKD_TMDT/TongCongDichVu/ct16";

	private string pl01_1_pathTotal12_VanTai = "PLuc_01_1_BK_CNKD_TMDT/TongCongVanTai/ct16";

	private string pl01_1_pathTotal12_Khac = "PLuc_01_1_BK_CNKD_TMDT/TongCongKinhDoanhKhac/ct16";

	private string pl01_1_nameColEnd = "AC";

	private string MA_NHOM_NGANH_MA_HANG_HOA = "01";

	private string MA_NHOM_NGANH_MA_DICH_VU = "02";

	private string MA_NHOM_NGANH_MA_VAN_TAI = "03";

	private string MA_NHOM_NGANH_MA_KHAC = "04";

	private string MA_LOAI_GIAY_TO_MST = "1000";

	private string TEN_LOAI_GIAY_TO_MST = "Mã sô thuế";

	private string MA_LOAI_GIAY_TO_CCCD = "2080";

	private string TEN_LOAI_GIAY_TO_CCCD = "CCCD";

	private bool dataChange = false;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
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
			sheetView.OperationMode = OperationMode.Normal;
			GeneralFunctions.setButton(sheetView, 3, 2, "Tải bảng kê");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int sheetIndex = 1;
			GeneralFunctions._page = 0;
			GeneralFunctions.dataInfoOld = null;
			GeneralFunctions.showPageLargeData1(winFpSpread, sheetIndex, hTKKCMCSheetInfo, GeneralFunctions._page, pl01_1_nameColEnd, "P2");
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Invalid comparison between Unknown and I4
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Invalid comparison between Unknown and I4
		try
		{
			dataChange = true;
			if ((int)GeneralFunctions.KEY_PRESS == 116 || (int)GeneralFunctions.KEY_PRESS == 117)
			{
				return;
			}
			int num = 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = null;
			if (GeneralFunctions._pageSize == 0)
			{
				GeneralFunctions._pageSize = GeneralFunctions.tinhPageSize_Large(hTKKCMCSheetInfo, "P2");
			}
			int startRowSView = GeneralFunctions._page * GeneralFunctions._pageSize + (int)iRow;
			int num2 = GeneralFunctions.timBangBatDauSView(hTKKCMCSheetInfo, startRowSView);
			if (num2 >= 0)
			{
				hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[num2];
			}
			if (hTKKCMCTableInfo == null)
			{
				return;
			}
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("Y");
			int num5 = HTKKFile.ColumnLettersToIndex("L");
			int num6 = HTKKFile.ColumnLettersToIndex("P");
			int num7 = HTKKFile.ColumnLettersToIndex("Q");
			int num8 = HTKKFile.ColumnLettersToIndex("R");
			int num9 = HTKKFile.ColumnLettersToIndex("U");
			string text = "";
			object value = sheetView.Cells[(int)iRow - 1, num4 - 1].Value;
			if (value != null && !string.IsNullOrEmpty(value.ToString()))
			{
				text = value.ToString().Trim();
				bool flag = Convert.ToBoolean(sheetView.Cells[(int)iRow - 1, num5 - 1].Value);
				double num10 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, num6 - 1].Value);
				Cell cell = sheetView.Cells[(int)iRow - 1, num7 - 1];
				if (iCol == num3 || iCol == num6)
				{
					if (text == MA_NHOM_NGANH_MA_HANG_HOA)
					{
						cell.Value = Math.Round(num10 * 0.01, 0);
					}
					if (text == MA_NHOM_NGANH_MA_DICH_VU)
					{
						cell.Value = Math.Round(num10 * 0.05, 0);
					}
					if (text == MA_NHOM_NGANH_MA_VAN_TAI)
					{
						cell.Value = Math.Round(num10 * 0.03, 0);
					}
					if (text == MA_NHOM_NGANH_MA_KHAC)
					{
						cell.Value = Math.Round(num10 * 0.05, 0);
					}
				}
				Cell cell2 = sheetView.Cells[(int)iRow - 1, num8 - 1];
				if (iCol == num3 || iCol == num5 || iCol == num6)
				{
					if (!flag)
					{
						if (text == MA_NHOM_NGANH_MA_HANG_HOA)
						{
							cell2.Value = Math.Round(num10 * 0.005, 0);
						}
						if (text == MA_NHOM_NGANH_MA_DICH_VU)
						{
							cell2.Value = Math.Round(num10 * 0.02, 0);
						}
						if (text == MA_NHOM_NGANH_MA_VAN_TAI)
						{
							cell2.Value = Math.Round(num10 * 0.015, 0);
						}
						if (text == MA_NHOM_NGANH_MA_KHAC)
						{
							cell2.Value = Math.Round(num10 * 0.02, 0);
						}
					}
					if (flag)
					{
						if (text == MA_NHOM_NGANH_MA_HANG_HOA)
						{
							cell2.Value = Math.Round(num10 * 0.01, 0);
						}
						if (text == MA_NHOM_NGANH_MA_DICH_VU)
						{
							cell2.Value = Math.Round(num10 * 0.05, 0);
						}
						if (text == MA_NHOM_NGANH_MA_VAN_TAI)
						{
							cell2.Value = Math.Round(num10 * 0.02, 0);
						}
						if (text == MA_NHOM_NGANH_MA_KHAC)
						{
							cell2.Value = Math.Round(num10 * 0.05, 0);
						}
					}
				}
				if (iCol >= num3)
				{
					sheetView.Cells[(int)iRow - 1, num9 - 1].Value = Math.Round(Convert.ToDouble(cell.Value) + Convert.ToDouble(cell2.Value), 0);
				}
				GeneralFunctions.UpdateFromSViewToSInfoLarge(sheetView, (int)iRow, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCSheetInfo, num2);
				if (GeneralFunctions.dataInfoOld.showFooter)
				{
					tinhTong_Pl01_1(ToKhaiInfo);
				}
				GeneralFunctions.showPageLargeData1(winFpSpread, num, hTKKCMCSheetInfo, GeneralFunctions._page, pl01_1_nameColEnd, "P2");
				GeneralFunctions.showRowsErrorLarge(sheetView);
			}
			else
			{
				GeneralFunctions.UpdateFromSViewToSInfoLarge(sheetView, (int)iRow, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCSheetInfo, num2);
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
			if (activeRowIndex == 2 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_CNKD_TMDT.xlsx");
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
			dataChange = false;
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			tinhTk(winFpSpread, ToKhaiInfo);
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
			return false;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			if (GeneralFunctions._pageSize == 0)
			{
				GeneralFunctions._pageSize = GeneralFunctions.tinhPageSize_Large(hTKKCMCSheetInfo, "P2");
			}
			int num2 = -1;
			int num3 = 0;
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				string tableName = ((num3 == 0) ? "Mục I" : "Mục II");
				List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
				int index = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt01);
				int index2 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt02);
				int index3 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt04);
				int index4 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt05);
				int index5 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt06);
				int index6 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt07);
				int index7 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt08Ten);
				int index8 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt08Ma);
				int index9 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt09);
				int index10 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt10);
				int index11 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt11);
				int index12 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathMaGiayTo);
				int index13 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathTenGiayTo);
				StringBuilder stringBuilder = new StringBuilder();
				string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
				string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
				string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
				string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num4 = 0; num4 < table.RowsValues.Count; num4++)
				{
					List<HTKKCMCRowValue> rowsValues = table.RowsValues;
					HTKKCMCRowValue hTKKCMCRowValue = rowsValues[num4];
					hTKKCMCRowValue.RowValue = hTKKCMCRowValue.RowValue.GetRange(0, cells.Count);
					hTKKCMCRowValue.RowValue.Take(table.RowsInfo.Cells.Count).ToList();
					if (num4 == 0 && table.RowsValues.Count == 1)
					{
						hTKKCMCRowValue.RowValue[index6] = ((hTKKCMCRowValue.RowValue[index6] == "1" || hTKKCMCRowValue.RowValue[index6] == "true") ? "true" : "false");
						if (!GeneralFunctions.checkHaveNewDataOfTable(table.RowsInfo, hTKKCMCRowValue.RowValue))
						{
							continue;
						}
					}
					string value2 = rowsValues[num4].RowValue[index];
					if (string.IsNullOrEmpty(value2))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index].CellTemplate.ColIndex.ToString());
						string text = "ct01 " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string value3 = rowsValues[num4].RowValue[index2];
					if (string.IsNullOrEmpty(value3))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index2].CellTemplate.ColIndex.ToString());
						string text = "ct02 " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string value4 = rowsValues[num4].RowValue[index3];
					if (string.IsNullOrEmpty(value4))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index3].CellTemplate.ColIndex.ToString());
						string text = "ct04 " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string value5 = rowsValues[num4].RowValue[index4];
					if (string.IsNullOrEmpty(value5))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index4].CellTemplate.ColIndex.ToString());
						string text = "ct05 " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string text2 = rowsValues[num4].RowValue[index12];
					if (string.IsNullOrEmpty(text2))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index13].CellTemplate.ColIndex.ToString());
						string text = "Loại giấy tờ " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string text3 = rowsValues[num4].RowValue[index5];
					if (string.IsNullOrEmpty(text3))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index5].CellTemplate.ColIndex.ToString());
						string text = "ct06 " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					if (text2 == MA_LOAI_GIAY_TO_MST && !ValidateBase.CheckMaSoThueCCCD(text3) && !string.IsNullOrEmpty(text3))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index5].CellTemplate.ColIndex.ToString());
						string text = message2;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					if (text2 == MA_LOAI_GIAY_TO_CCCD && text3.Length != 12 && !string.IsNullOrEmpty(text3))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index5].CellTemplate.ColIndex.ToString());
						string text = message3;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string value6 = rowsValues[num4].RowValue[index8];
					if (string.IsNullOrEmpty(value6))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index7].CellTemplate.ColIndex.ToString());
						string text = "ct08 " + message;
						hTKKCMCRowValue.RowValue.Add(text);
						value = 2;
						hTKKCMCRowValue.RowValue.Add(value.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
					string value7 = rowsValues[num4].RowValue[index9];
					string value8 = rowsValues[num4].RowValue[index10];
					string value9 = rowsValues[num4].RowValue[index11];
					if (Convert.ToDouble(value7) < Convert.ToDouble(value8) || Convert.ToDouble(value7) < Convert.ToDouble(value9))
					{
						hTKKCMCRowValue.RowValue.Add("error");
						hTKKCMCRowValue.RowValue.Add(cells[index9].CellTemplate.ColIndex.ToString());
						string text = "ct09 " + message4;
						hTKKCMCRowValue.RowValue.Add(text);
						if (value != 2)
						{
							value = 1;
						}
						hTKKCMCRowValue.RowValue.Add(1.ToString());
						GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
					}
				}
				if (value == 2)
				{
					table.RowsValues = table.RowsValues.OrderByDescending((HTKKCMCRowValue row) => row.RowValue.Count > cells.Count && row.RowValue[cells.Count] == "error" && row.RowValue[cells.Count + 3] == 2.ToString()).ToList();
					if (num2 < 0)
					{
						num2 = num3;
					}
				}
				else
				{
					table.RowsValues = table.RowsValues.OrderBy((HTKKCMCRowValue x) => int.Parse(x.RowValue[0])).ToList();
				}
				num3++;
			}
			if (num2 >= 0)
			{
				int dataLocationTemplateRowStart2 = hTKKCMCSheetInfo.Tables[num2].Location.DataLocationTemplateRowStart;
				GeneralFunctions._page = (int)Math.Round((decimal)dataLocationTemplateRowStart2 / (decimal)GeneralFunctions._pageSize, 0) - 1;
			}
			GeneralFunctions.showPageLargeData1(winFpSpread, num, hTKKCMCSheetInfo, GeneralFunctions._page, pl01_1_nameColEnd, "P2");
			if (num2 >= 0)
			{
				GeneralFunctions.showRowsErrorLarge(sheetView);
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			dataChange = true;
			winFpSpread.SuspendLayout();
			winFpSpread.ActiveSheet.AutoCalculation = false;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				IEnumerable<HTKKCMCRowValue> enumerable = table.RowsValues.Where((HTKKCMCRowValue x) => x.RowValue[0] == "");
				int num = 1;
				foreach (HTKKCMCRowValue item in enumerable)
				{
					item.RowValue[0] = num.ToString();
					num++;
				}
			}
			SheetView activeSheet = winFpSpread.ActiveSheet;
			GeneralFunctions._page = 0;
			tinhTk(winFpSpread, ToKhaiInfo);
			tinhTong_Pl01_1(ToKhaiInfo);
			GeneralFunctions.showPageLargeData1(winFpSpread, activeSheetIndex, hTKKCMCSheetInfo, GeneralFunctions._page, pl01_1_nameColEnd, "P2");
			winFpSpread.ResumeLayout();
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
			dataChange = true;
			winFpSpread.SuspendLayout();
			winFpSpread.ActiveSheet.AutoCalculation = false;
			int num = 1;
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			GeneralFunctions._page = 0;
			tinhTk(winFpSpread, ToKhaiInfo);
			tinhTong_Pl01_1(ToKhaiInfo);
			GeneralFunctions.showPageLargeData1(winFpSpread, num, sInfo, GeneralFunctions._page, pl01_1_nameColEnd, "P2");
			winFpSpread.ActiveSheet.AutoCalculation = true;
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpadatePage_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int page)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = 1;
			if (page + 1 >= GeneralFunctions._totalPage)
			{
				tinhTong_Pl01_1(ToKhaiInfo);
			}
			GeneralFunctions.showPageLargeData1(winFpSpread, num, hTKKCMCSheetInfo, page, pl01_1_nameColEnd, "P2");
			GeneralFunctions.showRowsErrorLarge(winFpSpread.Sheets[num]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		if (GeneralFunctions.dataInfoOld.showFooter)
		{
			tinhTong_Pl01_1(ToKhaiInfo);
		}
		GeneralFunctions.showPageLargeData1(winFpSpread, 1, ToKhaiInfo.SheetInfos[1], GeneralFunctions._page, pl01_1_nameColEnd, "P2");
		return true;
	}

	public void tinhTong_Pl01_1(HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int index = 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_HangHoa);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_DichVu);
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_VanTai);
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_Khac);
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_HangHoa);
			HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_DichVu);
			HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_VanTai);
			HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_Khac);
			HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_HangHoa);
			HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_DichVu);
			HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_VanTai);
			HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_Khac);
			HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_HangHoa);
			HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_DichVu);
			HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_VanTai);
			HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_Khac);
			hTKKCMCCellInfo.Value = "0";
			hTKKCMCCellInfo2.Value = "0";
			hTKKCMCCellInfo3.Value = "0";
			hTKKCMCCellInfo4.Value = "0";
			hTKKCMCCellInfo5.Value = "0";
			hTKKCMCCellInfo6.Value = "0";
			hTKKCMCCellInfo7.Value = "0";
			hTKKCMCCellInfo8.Value = "0";
			hTKKCMCCellInfo9.Value = "0";
			hTKKCMCCellInfo10.Value = "0";
			hTKKCMCCellInfo11.Value = "0";
			hTKKCMCCellInfo12.Value = "0";
			hTKKCMCCellInfo13.Value = "0";
			hTKKCMCCellInfo14.Value = "0";
			hTKKCMCCellInfo15.Value = "0";
			hTKKCMCCellInfo16.Value = "0";
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				for (int num = 0; num < table.RowsValues.Count; num++)
				{
					int index2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt08Ma);
					int index3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt10);
					int index4 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt11);
					int index5 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_pathCt12);
					string text = table.RowsValues[num].RowValue[index2];
					table.RowsValues[num].RowValue[index5] = Convert.ToString(Math.Round(Convert.ToDouble(table.RowsValues[num].RowValue[index3]) + Convert.ToDouble(table.RowsValues[num].RowValue[index4]), 0));
					for (int num2 = 0; num2 < table.RowsInfo.Cells.Count; num2++)
					{
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells[num2];
						if (!(hTKKCMCTableCellInfo.Path != pl01_1_pathCt09) || !(hTKKCMCTableCellInfo.Path != pl01_1_pathCt10) || !(hTKKCMCTableCellInfo.Path != pl01_1_pathCt11) || !(hTKKCMCTableCellInfo.Path != pl01_1_pathCt12))
						{
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt09 && text == MA_NHOM_NGANH_MA_HANG_HOA)
							{
								hTKKCMCCellInfo.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt09 && text == MA_NHOM_NGANH_MA_DICH_VU)
							{
								hTKKCMCCellInfo2.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo2.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt09 && text == MA_NHOM_NGANH_MA_VAN_TAI)
							{
								hTKKCMCCellInfo3.Value = Convert.ToString(Convert.ToDouble(hTKKCMCCellInfo3.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt09 && text == MA_NHOM_NGANH_MA_KHAC)
							{
								hTKKCMCCellInfo4.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo4.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt10 && text == MA_NHOM_NGANH_MA_HANG_HOA)
							{
								hTKKCMCCellInfo5.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo5.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt10 && text == MA_NHOM_NGANH_MA_DICH_VU)
							{
								hTKKCMCCellInfo6.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo6.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt10 && text == MA_NHOM_NGANH_MA_VAN_TAI)
							{
								hTKKCMCCellInfo7.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo7.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt10 && text == MA_NHOM_NGANH_MA_KHAC)
							{
								hTKKCMCCellInfo8.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo8.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt11 && text == MA_NHOM_NGANH_MA_HANG_HOA)
							{
								hTKKCMCCellInfo9.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo9.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt11 && text == MA_NHOM_NGANH_MA_DICH_VU)
							{
								hTKKCMCCellInfo10.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo10.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt11 && text == MA_NHOM_NGANH_MA_VAN_TAI)
							{
								hTKKCMCCellInfo11.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo11.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt11 && text == MA_NHOM_NGANH_MA_KHAC)
							{
								hTKKCMCCellInfo12.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo12.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt12 && text == MA_NHOM_NGANH_MA_HANG_HOA)
							{
								hTKKCMCCellInfo13.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo13.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt12 && text == MA_NHOM_NGANH_MA_DICH_VU)
							{
								hTKKCMCCellInfo14.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo14.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt12 && text == MA_NHOM_NGANH_MA_VAN_TAI)
							{
								hTKKCMCCellInfo15.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo15.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
							if (hTKKCMCTableCellInfo.Path == pl01_1_pathCt12 && text == MA_NHOM_NGANH_MA_KHAC)
							{
								hTKKCMCCellInfo16.Value = Convert.ToString(Math.Round(Convert.ToDouble(hTKKCMCCellInfo16.Value) + Convert.ToDouble(table.RowsValues[num].RowValue[num2]), 0));
							}
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

	public void SheetTabChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex)
	{
		try
		{
			if (sheetIndex == 0 && dataChange)
			{
				tinhTk(winFpSpread, ToKhaiInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void tinhTk(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[0];
		SheetView sheetView = winFpSpread.Sheets[0];
		tinhTong_Pl01_1(ToKhaiInfo);
		HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_HangHoa);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt04_1);
		hTKKCMCCellInfo2.Value = hTKKCMCCellInfo.Value;
		sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo2.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_DichVu);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt04_2);
		hTKKCMCCellInfo4.Value = hTKKCMCCellInfo3.Value;
		sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo4.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_VanTai);
		HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt04_3);
		hTKKCMCCellInfo6.Value = hTKKCMCCellInfo5.Value;
		sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo6.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal09_Khac);
		HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt04_4);
		hTKKCMCCellInfo8.Value = hTKKCMCCellInfo7.Value;
		sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo8.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt04_5);
		hTKKCMCCellInfo9.Value = Convert.ToString(Convert.ToDouble(hTKKCMCCellInfo2.Value) + Convert.ToDouble(hTKKCMCCellInfo4.Value) + Convert.ToDouble(hTKKCMCCellInfo6.Value) + Convert.ToDouble(hTKKCMCCellInfo8.Value));
		sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo9.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_HangHoa);
		HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt05_1);
		hTKKCMCCellInfo11.Value = hTKKCMCCellInfo10.Value;
		sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo11.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_DichVu);
		HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt05_2);
		hTKKCMCCellInfo13.Value = hTKKCMCCellInfo12.Value;
		sheetView.Cells[hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo13.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_VanTai);
		HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt05_3);
		hTKKCMCCellInfo15.Value = hTKKCMCCellInfo14.Value;
		sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo15.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal10_Khac);
		HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt05_4);
		hTKKCMCCellInfo17.Value = hTKKCMCCellInfo16.Value;
		sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo17.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt05_5);
		hTKKCMCCellInfo18.Value = Convert.ToString(Convert.ToDouble(hTKKCMCCellInfo11.Value) + Convert.ToDouble(hTKKCMCCellInfo13.Value) + Convert.ToDouble(hTKKCMCCellInfo15.Value) + Convert.ToDouble(hTKKCMCCellInfo17.Value));
		sheetView.Cells[hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo18.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_HangHoa);
		HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt06_1);
		hTKKCMCCellInfo20.Value = hTKKCMCCellInfo19.Value;
		sheetView.Cells[hTKKCMCCellInfo20.CellTemplate.RowIndex - 1, hTKKCMCCellInfo20.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo20.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo21 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_DichVu);
		HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt06_2);
		hTKKCMCCellInfo22.Value = hTKKCMCCellInfo21.Value;
		sheetView.Cells[hTKKCMCCellInfo22.CellTemplate.RowIndex - 1, hTKKCMCCellInfo22.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo22.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo23 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_VanTai);
		HTKKCMCCellInfo hTKKCMCCellInfo24 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt06_3);
		hTKKCMCCellInfo24.Value = hTKKCMCCellInfo23.Value;
		sheetView.Cells[hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, hTKKCMCCellInfo24.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo24.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo25 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal11_Khac);
		HTKKCMCCellInfo hTKKCMCCellInfo26 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt06_4);
		hTKKCMCCellInfo26.Value = hTKKCMCCellInfo25.Value;
		sheetView.Cells[hTKKCMCCellInfo26.CellTemplate.RowIndex - 1, hTKKCMCCellInfo26.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo26.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo27 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt06_5);
		hTKKCMCCellInfo27.Value = Convert.ToString(Convert.ToDouble(hTKKCMCCellInfo20.Value) + Convert.ToDouble(hTKKCMCCellInfo22.Value) + Convert.ToDouble(hTKKCMCCellInfo24.Value) + Convert.ToDouble(hTKKCMCCellInfo26.Value));
		sheetView.Cells[hTKKCMCCellInfo27.CellTemplate.RowIndex - 1, hTKKCMCCellInfo27.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo27.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo28 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_HangHoa);
		HTKKCMCCellInfo hTKKCMCCellInfo29 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt07_1);
		hTKKCMCCellInfo29.Value = hTKKCMCCellInfo28.Value;
		sheetView.Cells[hTKKCMCCellInfo29.CellTemplate.RowIndex - 1, hTKKCMCCellInfo29.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo29.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo30 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_DichVu);
		HTKKCMCCellInfo hTKKCMCCellInfo31 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt07_2);
		hTKKCMCCellInfo31.Value = hTKKCMCCellInfo30.Value;
		sheetView.Cells[hTKKCMCCellInfo31.CellTemplate.RowIndex - 1, hTKKCMCCellInfo31.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo31.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo32 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_VanTai);
		HTKKCMCCellInfo hTKKCMCCellInfo33 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt07_3);
		hTKKCMCCellInfo33.Value = hTKKCMCCellInfo32.Value;
		sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, hTKKCMCCellInfo33.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo33.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo34 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pl01_1_pathTotal12_Khac);
		HTKKCMCCellInfo hTKKCMCCellInfo35 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt07_4);
		hTKKCMCCellInfo35.Value = hTKKCMCCellInfo34.Value;
		sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, hTKKCMCCellInfo35.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo35.Value;
		HTKKCMCCellInfo hTKKCMCCellInfo36 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == tk_pathCt07_5);
		hTKKCMCCellInfo36.Value = Convert.ToString(Convert.ToDouble(hTKKCMCCellInfo29.Value) + Convert.ToDouble(hTKKCMCCellInfo31.Value) + Convert.ToDouble(hTKKCMCCellInfo33.Value) + Convert.ToDouble(hTKKCMCCellInfo35.Value));
		sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, hTKKCMCCellInfo36.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo36.Value;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.SuspendLayout();
			GeneralFunctions.showPageLargeData1(winFpSpread, 1, ToKhaiInfo.SheetInfos[1], GeneralFunctions._page, pl01_1_nameColEnd, "P2");
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
