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

public class CnkdTmdtNccnn01Nckt308
{
	private string path_I_sum_HH_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuNCCNN/ct09";

	private string path_I_sum_HH_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTNCCNN/ct09";

	private string path_I_sum_HH_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNNCCNN/ct09";

	private string path_I_sum_HH_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruNCCNN/ct09";

	private string path_I_sum_DV_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuNCCNN/ct10";

	private string path_I_sum_DV_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTNCCNN/ct10";

	private string path_I_sum_DV_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNNCCNN/ct10";

	private string path_I_sum_DV_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruNCCNN/ct10";

	private string path_I_sum_VT_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuNCCNN/ct11";

	private string path_I_sum_VT_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTNCCNN/ct11";

	private string path_I_sum_VT_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNNCCNN/ct11";

	private string path_I_sum_VT_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruNCCNN/ct11";

	private string path_I_sum_KHAC_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuNCCNN/ct12";

	private string path_I_sum_KHAC_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTNCCNN/ct12";

	private string path_I_sum_KHAC_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNNCCNN/ct12";

	private string path_I_sum_KHAC_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruNCCNN/ct12";

	private string path_II_sum_HH_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuHKD/ct16";

	private string path_II_sum_HH_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTHKD/ct16";

	private string path_II_sum_HH_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNHKD/ct16";

	private string path_II_sum_HH_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruHKD/ct16";

	private string path_II_sum_DV_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuHKD/ct17";

	private string path_II_sum_DV_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTHKD/ct17";

	private string path_II_sum_DV_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNHKD/ct17";

	private string path_II_sum_DV_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruHKD/ct17";

	private string path_II_sum_VT_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuHKD/ct18";

	private string path_II_sum_VT_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTHKD/ct18";

	private string path_II_sum_VT_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNHKD/ct18";

	private string path_II_sum_VT_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruHKD/ct18";

	private string path_II_sum_KHAC_ct4_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuHKD/ct19";

	private string path_II_sum_KHAC_ct5_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueGTGTHKD/ct19";

	private string path_II_sum_KHAC_ct6_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/SoThueTNCNHKD/ct19";

	private string path_II_sum_KHAC_ct7_tk = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/TongSoThueDaKhauTruHKD/ct19";

	private string pl01_1_table1_pathCt2 = "ct02";

	private string pl01_1_table1_pathCt3 = "ct03";

	private string pl01_1_table1_pathCt4 = "ct04";

	private string pl01_1_table1_pathCt5 = "ct05";

	private string pl01_1_table1_pathCt6 = "ct06";

	private string pl01_1_table1_pathCt7Ma = "ct07_ma";

	private string pl01_1_table1_pathCt7Ten = "ct07_ten";

	private string pl01_1_table1_pathCt8 = "ct08";

	private string pl01_1_table1_pathCt9 = "ct09";

	private string pl01_1_table1_pathCt10 = "ct10";

	private string pl01_1_table1_pathCt11 = "ct11";

	private string pl01_1_table1_sum_HH_8 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongHangHoa/ct12";

	private string pl01_1_table1_sum_HH_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongHangHoa/ct13";

	private string pl01_1_table1_sum_HH_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongHangHoa/ct14";

	private string pl01_1_table1_sum_HH_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongHangHoa/ct15";

	private string pl01_1_table1_sum_DV_8 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongDichVu/ct12";

	private string pl01_1_table1_sum_DV_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongDichVu/ct13";

	private string pl01_1_table1_sum_DV_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongDichVu/ct14";

	private string pl01_1_table1_sum_DV_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongDichVu/ct15";

	private string pl01_1_table1_sum_VT_8 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongVanTai/ct12";

	private string pl01_1_table1_sum_VT_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongVanTai/ct13";

	private string pl01_1_table1_sum_VT_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongVanTai/ct14";

	private string pl01_1_table1_sum_VT_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongVanTai/ct15";

	private string pl01_1_table1_sum_KHAC_8 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongKinhDoanhKhac/ct12";

	private string pl01_1_table1_sum_KHAC_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongKinhDoanhKhac/ct13";

	private string pl01_1_table1_sum_KHAC_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongKinhDoanhKhac/ct14";

	private string pl01_1_table1_sum_KHAC_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/ToChucNuocNgoai/TongCongKinhDoanhKhac/ct15";

	private string pl01_1_table2_pathCt2 = "ct02";

	private string pl01_1_table2_pathCt3 = "ct03";

	private string pl01_1_table2_pathCt4 = "ct04";

	private string pl01_1_table2_pathCt5 = "ct05";

	private string pl01_1_table2_pathCt6 = "ct06";

	private string pl01_1_table2_pathCt7 = "ct07";

	private string pl01_1_table2_pathCt8_ten = "ct08_ten";

	private string pl01_1_table2_pathCt8_ma = "ct08_ma";

	private string pl01_1_table2_pathCt9 = "ct09";

	private string pl01_1_table2_pathCt10 = "ct10";

	private string pl01_1_table2_pathCt11 = "ct11";

	private string pl01_1_table2_pathCt12 = "ct12";

	private string pl01_1_table2_sum_HH_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongHangHoa/ct13";

	private string pl01_1_table2_sum_HH_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongHangHoa/ct14";

	private string pl01_1_table2_sum_HH_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongHangHoa/ct15";

	private string pl01_1_table2_sum_HH_12 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongHangHoa/ct16";

	private string pl01_1_table2_sum_DV_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongDichVu/ct13";

	private string pl01_1_table2_sum_DV_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongDichVu/ct14";

	private string pl01_1_table2_sum_DV_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongDichVu/ct15";

	private string pl01_1_table2_sum_DV_12 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongDichVu/ct16";

	private string pl01_1_table2_sum_VT_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongVanTai/ct13";

	private string pl01_1_table2_sum_VT_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongVanTai/ct14";

	private string pl01_1_table2_sum_VT_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongVanTai/ct15";

	private string pl01_1_table2_sum_VT_12 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongVanTai/ct16";

	private string pl01_1_table2_sum_KHAC_9 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongKinhDoanhKhac/ct13";

	private string pl01_1_table2_sum_KHAC_10 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongKinhDoanhKhac/ct14";

	private string pl01_1_table2_sum_KHAC_11 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongKinhDoanhKhac/ct15";

	private string pl01_1_table2_sum_KHAC_12 = "PLuc_01_1_BK_CNKD_NCCNN_TMDT/HoKinhDoanh/TongCongKinhDoanhKhac/ct16";

	private int colMaHangHoa = HTKKFile.ColumnLettersToIndex("AD");

	private string pl01_1_nameColEnd = "AK";

	private bool dataChange = false;

	private string MA_NHOM_NGANH_MA_HANG_HOA = "01";

	private string MA_NHOM_NGANH_MA_DICH_VU = "02";

	private string MA_NHOM_NGANH_MA_VAN_TAI = "03";

	private string MA_NHOM_NGANH_MA_KHAC = "04";

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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int sheetIndex = 1;
			GeneralFunctions._page = 0;
			GeneralFunctions.ShowSheetTablePage(winFpSpread, sheetIndex, hTKKCMCSheetInfo, GeneralFunctions._page);
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
			int num = HTKKFile.ColumnLettersToIndex("U");
			int column = HTKKFile.ColumnLettersToIndex("Y");
			int column2 = HTKKFile.ColumnLettersToIndex("AR");
			int column3 = HTKKFile.ColumnLettersToIndex("AC");
			int column4 = HTKKFile.ColumnLettersToIndex("AS");
			if (iCol == num)
			{
				sheetView.Cells[(int)iRow - 1, column].Value = sheetView.Cells[(int)iRow - 1, column2].Value;
				sheetView.Cells[(int)iRow - 1, column3].Value = sheetView.Cells[(int)iRow - 1, column4].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Invalid comparison between Unknown and I4
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Invalid comparison between Unknown and I4
		try
		{
			dataChange = true;
			if ((int)GeneralFunctions.KEY_PRESS == 116 || (int)GeneralFunctions.KEY_PRESS == 117)
			{
				return;
			}
			int num = 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			int colIndex = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt7Ten).CellTemplate.ColIndex;
			int colIndex2 = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt8).CellTemplate.ColIndex;
			int colIndex3 = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt9).CellTemplate.ColIndex;
			int colIndex4 = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt10).CellTemplate.ColIndex;
			int colIndex5 = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt11).CellTemplate.ColIndex;
			int colIndex6 = hTKKCMCTableInfo2.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table2_pathCt8_ten).CellTemplate.ColIndex;
			int colIndex7 = hTKKCMCTableInfo2.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table2_pathCt9).CellTemplate.ColIndex;
			int colIndex8 = hTKKCMCTableInfo2.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table2_pathCt10).CellTemplate.ColIndex;
			int colIndex9 = hTKKCMCTableInfo2.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table2_pathCt11).CellTemplate.ColIndex;
			int colIndex10 = hTKKCMCTableInfo2.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table2_pathCt12).CellTemplate.ColIndex;
			HTKKCMCTableInfo hTKKCMCTableInfo3 = null;
			if (GeneralFunctions._pageSize == 0)
			{
				GeneralFunctions._pageSize = GeneralFunctions.tinhPageSize_Large(hTKKCMCSheetInfo, "X6");
			}
			int startRowSView = GeneralFunctions._page * GeneralFunctions._pageSize + (int)iRow;
			int num2 = GeneralFunctions.timBangBatDauSView(hTKKCMCSheetInfo, startRowSView);
			if (num2 >= 0)
			{
				hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[num2];
			}
			if (hTKKCMCTableInfo3 != null)
			{
				if (num2 == 0 && (iCol == colIndex3 || iCol == colIndex4))
				{
					sheetView.Cells[(int)iRow - 1, colIndex5 - 1].Value = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, colIndex3 - 1].Value) + Convert.ToDouble(sheetView.Cells[(int)iRow - 1, colIndex4 - 1].Value);
				}
				if (num2 > 0 && (iCol == colIndex8 || iCol == colIndex9))
				{
					sheetView.Cells[(int)iRow - 1, colIndex10 - 1].Value = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, colIndex4 - 1].Value) + Convert.ToDouble(sheetView.Cells[(int)iRow - 1, colIndex5 - 1].Value);
				}
				GeneralFunctions.UpdateFromSViewToSInfoLarge(sheetView, (int)iRow, -1, hTKKCMCTableInfo3.RowsInfo, hTKKCMCSheetInfo, num2);
				if ((num2 == 0 && iCol >= colIndex) || (num2 > 0 && iCol >= colIndex6))
				{
					tinhTong_Pl01_1(winFpSpread, ToKhaiInfo);
					tinhTk(winFpSpread, ToKhaiInfo);
				}
				if (num2 > 0)
				{
					GeneralFunctions.ShowSheetTablePage(winFpSpread, num, hTKKCMCSheetInfo, GeneralFunctions._page);
					GeneralFunctions.showRowsErrorLarge(sheetView, hTKKCMCSheetInfo);
				}
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
			if (activeRowIndex == 12 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_CNKD_NCCNN_TMDT.xlsx");
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
				GeneralFunctions._pageSize = GeneralFunctions.tinhPageSize_Large(hTKKCMCSheetInfo, "X6");
			}
			int num2 = -1;
			int num3 = 0;
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
				string tableName = num3 switch
				{
					2 => "Mục II.2", 
					1 => "Mục II.1", 
					0 => "Mục I", 
					_ => "", 
				};
				if (num3 == 0)
				{
					int index = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt2);
					int index2 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt3);
					int index3 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt4);
					int index4 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt5);
					int index5 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt6);
					int index6 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt7Ma);
					int index7 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt7Ten);
					int index8 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt8);
					int num4 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt9);
					int num5 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt10);
					int num6 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pl01_1_table1_pathCt11);
					StringBuilder stringBuilder = new StringBuilder();
					string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
					string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
					string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int num7 = 0; num7 < table.RowsValues.Count; num7++)
					{
						List<HTKKCMCRowValue> rowsValues = table.RowsValues;
						HTKKCMCRowValue hTKKCMCRowValue = rowsValues[num7];
						hTKKCMCRowValue.RowValue = hTKKCMCRowValue.RowValue.GetRange(0, cells.Count);
						hTKKCMCRowValue.RowValue.Take(table.RowsInfo.Cells.Count).ToList();
						if (GeneralFunctions.checkHaveNewDataOfTable(table.RowsInfo, hTKKCMCRowValue.RowValue))
						{
							string value2 = rowsValues[num7].RowValue[index];
							if (string.IsNullOrEmpty(value2))
							{
								hTKKCMCRowValue.RowValue.Add("error");
								hTKKCMCRowValue.RowValue.Add(cells[index].CellTemplate.ColIndex.ToString());
								string text = "ct02 " + message;
								hTKKCMCRowValue.RowValue.Add(text);
								value = 2;
								hTKKCMCRowValue.RowValue.Add(value.ToString());
								GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
							}
							string value3 = rowsValues[num7].RowValue[index2];
							if (string.IsNullOrEmpty(value3))
							{
								hTKKCMCRowValue.RowValue.Add("error");
								hTKKCMCRowValue.RowValue.Add(cells[index2].CellTemplate.ColIndex.ToString());
								string text = "ct03 " + message;
								hTKKCMCRowValue.RowValue.Add(text);
								value = 2;
								hTKKCMCRowValue.RowValue.Add(value.ToString());
								GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
							}
							string value4 = rowsValues[num7].RowValue[index3];
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
							string value5 = rowsValues[num7].RowValue[index4];
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
							string text2 = rowsValues[num7].RowValue[index8];
							if (text2 == "0")
							{
								hTKKCMCRowValue.RowValue.Add("error");
								hTKKCMCRowValue.RowValue.Add(cells[index8].CellTemplate.ColIndex.ToString());
								string text = "ct08 " + message;
								hTKKCMCRowValue.RowValue.Add(text);
								value = 2;
								hTKKCMCRowValue.RowValue.Add(value.ToString());
								GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
							}
							string value6 = rowsValues[num7].RowValue[index6];
							if (string.IsNullOrEmpty(value6))
							{
								hTKKCMCRowValue.RowValue.Add("error");
								hTKKCMCRowValue.RowValue.Add(cells[index7].CellTemplate.ColIndex.ToString());
								string text = "ct07 " + message;
								hTKKCMCRowValue.RowValue.Add(text);
								value = 2;
								hTKKCMCRowValue.RowValue.Add(value.ToString());
								GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
							}
							string text3 = rowsValues[num7].RowValue[index5];
							if (!string.IsNullOrEmpty(text3) && !ValidateBase.CheckMaSoThueCCCD(text3))
							{
								hTKKCMCRowValue.RowValue.Add("error");
								hTKKCMCRowValue.RowValue.Add(cells[index5].CellTemplate.ColIndex.ToString());
								string text = "ct06 " + message2;
								hTKKCMCRowValue.RowValue.Add(text);
								value = 2;
								hTKKCMCRowValue.RowValue.Add(value.ToString());
								GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text, hTKKCMCRowValue.RowValue[0]);
							}
						}
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
			GeneralFunctions.ShowSheetTablePage(winFpSpread, num, hTKKCMCSheetInfo, GeneralFunctions._page);
			if (num2 >= 0)
			{
				GeneralFunctions.showRowsErrorLarge(sheetView, hTKKCMCSheetInfo);
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			SheetView activeSheet = winFpSpread.ActiveSheet;
			GeneralFunctions._page = 0;
			tinhTong_Pl01_1(winFpSpread, ToKhaiInfo);
			tinhTk(winFpSpread, ToKhaiInfo);
			GeneralFunctions.ShowSheetTablePage(winFpSpread, activeSheetIndex, sInfo, GeneralFunctions._page);
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
			GeneralFunctions.ShowSheetTablePage(winFpSpread, num, sInfo, GeneralFunctions._page);
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
			GeneralFunctions.ShowSheetTablePage(winFpSpread, num, hTKKCMCSheetInfo, page);
			GeneralFunctions.showRowsErrorLarge(winFpSpread.Sheets[num], hTKKCMCSheetInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		return true;
	}

	public void tinhTong_Pl01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int num = 1;
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table1_pathCt8, pl01_1_table1_sum_HH_8);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table1_pathCt9, pl01_1_table1_sum_HH_9);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table1_pathCt10, pl01_1_table1_sum_HH_10);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table1_pathCt11, pl01_1_table1_sum_HH_11);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table1_pathCt8, pl01_1_table1_sum_DV_8);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table1_pathCt9, pl01_1_table1_sum_DV_9);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table1_pathCt10, pl01_1_table1_sum_DV_10);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table1_pathCt11, pl01_1_table1_sum_DV_11);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table1_pathCt8, pl01_1_table1_sum_VT_8);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table1_pathCt9, pl01_1_table1_sum_VT_9);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table1_pathCt10, pl01_1_table1_sum_VT_10);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table1_pathCt11, pl01_1_table1_sum_VT_11);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table1_pathCt8, pl01_1_table1_sum_KHAC_8);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table1_pathCt9, pl01_1_table1_sum_KHAC_9);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table1_pathCt10, pl01_1_table1_sum_KHAC_10);
			tong1LoaiTrongBang1(sInfo, 0, pl01_1_table1_pathCt7Ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table1_pathCt11, pl01_1_table1_sum_KHAC_11);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table2_pathCt9, pl01_1_table2_sum_HH_9);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table2_pathCt10, pl01_1_table2_sum_HH_10);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table2_pathCt11, pl01_1_table2_sum_HH_11);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_HANG_HOA, pl01_1_table2_pathCt12, pl01_1_table2_sum_HH_12);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table2_pathCt9, pl01_1_table2_sum_DV_9);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table2_pathCt10, pl01_1_table2_sum_DV_10);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table2_pathCt11, pl01_1_table2_sum_DV_11);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_DICH_VU, pl01_1_table2_pathCt12, pl01_1_table2_sum_DV_12);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table2_pathCt9, pl01_1_table2_sum_VT_9);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table2_pathCt10, pl01_1_table2_sum_VT_10);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table2_pathCt11, pl01_1_table2_sum_VT_11);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_VAN_TAI, pl01_1_table2_pathCt12, pl01_1_table2_sum_VT_12);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table2_pathCt9, pl01_1_table2_sum_KHAC_9);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table2_pathCt10, pl01_1_table2_sum_KHAC_10);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table2_pathCt11, pl01_1_table2_sum_KHAC_11);
			tong1LoaiTrongBang2_3(sInfo, pl01_1_table2_pathCt8_ma, MA_NHOM_NGANH_MA_KHAC, pl01_1_table2_pathCt12, pl01_1_table2_sum_KHAC_12);
			GeneralFunctions.ShowSheetTablePage(winFpSpread, num, sInfo, GeneralFunctions._page);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void tong1LoaiTrongBang1(HTKKCMCSheetInfo sInfo, int tableIndex, string pathMa, string ma, string pathNumber, string pathSum)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
		int maIndex = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathMa);
		int index = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathNumber);
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathSum);
		double num = 0.0;
		List<HTKKCMCRowValue> list = hTKKCMCTableInfo.RowsValues.FindAll((HTKKCMCRowValue x) => x.RowValue[maIndex] == ma);
		foreach (HTKKCMCRowValue item in list)
		{
			num += Convert.ToDouble(item.RowValue[index]);
		}
		hTKKCMCCellInfo.Value = num.ToString();
	}

	public void tong1LoaiTrongBang2_3(HTKKCMCSheetInfo sInfo, string pathMa, string ma, string pathNumber, string pathSum)
	{
		double num = 0.0;
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table.TableName == "table_1")
			{
				continue;
			}
			int maIndex = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathMa);
			int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == pathNumber);
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathSum);
			List<HTKKCMCRowValue> list = table.RowsValues.FindAll((HTKKCMCRowValue x) => x.RowValue[maIndex] == ma);
			foreach (HTKKCMCRowValue item in list)
			{
				num += Convert.ToDouble(item.RowValue[index]);
			}
			hTKKCMCCellInfo.Value = num.ToString();
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
		HTKKCMCSheetInfo sInfo_Pl01_ = ToKhaiInfo.SheetInfos[1];
		HTKKCMCSheetInfo sInfoTk = ToKhaiInfo.SheetInfos[0];
		SheetView sViewTk = winFpSpread.Sheets[0];
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_HH_8, path_I_sum_HH_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_HH_9, path_I_sum_HH_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_HH_10, path_I_sum_HH_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_HH_11, path_I_sum_HH_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_DV_8, path_I_sum_DV_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_DV_9, path_I_sum_DV_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_DV_10, path_I_sum_DV_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_DV_11, path_I_sum_DV_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_VT_8, path_I_sum_VT_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_VT_9, path_I_sum_VT_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_VT_10, path_I_sum_VT_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_VT_11, path_I_sum_VT_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_KHAC_8, path_I_sum_KHAC_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_KHAC_9, path_I_sum_KHAC_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_KHAC_10, path_I_sum_KHAC_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table1_sum_KHAC_11, path_I_sum_KHAC_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_HH_9, path_II_sum_HH_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_HH_10, path_II_sum_HH_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_HH_11, path_II_sum_HH_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_HH_12, path_II_sum_HH_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_DV_9, path_II_sum_DV_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_DV_10, path_II_sum_DV_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_DV_11, path_II_sum_DV_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_DV_12, path_II_sum_DV_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_VT_9, path_II_sum_VT_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_VT_10, path_II_sum_VT_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_VT_11, path_II_sum_VT_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_VT_12, path_II_sum_VT_ct7_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_KHAC_9, path_II_sum_KHAC_ct4_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_KHAC_10, path_II_sum_KHAC_ct5_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_KHAC_11, path_II_sum_KHAC_ct6_tk);
		tongPhuLucSangTk(sInfo_Pl01_, sInfoTk, sViewTk, pl01_1_table2_sum_KHAC_12, path_II_sum_KHAC_ct7_tk);
	}

	public void tongPhuLucSangTk(HTKKCMCSheetInfo sInfo_Pl01_1, HTKKCMCSheetInfo sInfoTk, SheetView sViewTk, string pathPhuLuc, string pathTk)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo_Pl01_1.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathPhuLuc);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfoTk.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTk);
		hTKKCMCCellInfo2.Value = hTKKCMCCellInfo.Value;
		sViewTk.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = hTKKCMCCellInfo2.Value;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.SuspendLayout();
			GeneralFunctions.ShowSheetTablePage(winFpSpread, 1, ToKhaiInfo.SheetInfos[1], GeneralFunctions._page);
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCt10_Pl_Bang1(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow)
	{
		string text = sView.Cells[iRow - 1, colMaHangHoa - 1].Text;
		int num = HTKKFile.ColumnLettersToIndex("Q");
		int num2 = HTKKFile.ColumnLettersToIndex("R");
		if (text == MA_NHOM_NGANH_MA_HANG_HOA)
		{
			sView.Cells[iRow - 1, num2 - 1].Value = Convert.ToDouble(sView.Cells[iRow - 1, num - 1].Value) * 0.01;
		}
		if (text == MA_NHOM_NGANH_MA_DICH_VU)
		{
			sView.Cells[iRow - 1, num2 - 1].Value = Convert.ToDouble(sView.Cells[iRow - 1, num - 1].Value) * 0.05;
		}
		if (text == MA_NHOM_NGANH_MA_VAN_TAI)
		{
			sView.Cells[iRow - 1, num2 - 1].Value = Convert.ToDouble(sView.Cells[iRow - 1, num - 1].Value) * 0.03;
		}
		if (text == MA_NHOM_NGANH_MA_KHAC)
		{
			sView.Cells[iRow - 1, num2 - 1].Value = Convert.ToDouble(sView.Cells[iRow - 1, num - 1].Value) * 0.05;
		}
	}
}
