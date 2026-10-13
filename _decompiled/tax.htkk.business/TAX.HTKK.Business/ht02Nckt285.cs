using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class ht02Nckt285
{
	private int[] cellPhuLuc = GeneralFunctions.ArrayCellId("BJ_15");

	private int[] cellTenTruongHopHoanThue = GeneralFunctions.ArrayCellId("H_15");

	private int[] cellMaTruongHopHoanThue = GeneralFunctions.ArrayCellId("BC_15");

	private int[] cellLoaiKy = GeneralFunctions.ArrayCellId("BB_19");

	private int[] cellTuKy_tk = GeneralFunctions.ArrayCellId("BC_19");

	private int[] cellDenKy_tk = GeneralFunctions.ArrayCellId("BD_19");

	private int[] cellNgayPSNV = GeneralFunctions.ArrayCellId("H_21");

	private int[] cellNgayKTNV = GeneralFunctions.ArrayCellId("S_21");

	private int[] cellVietNam = GeneralFunctions.ArrayCellId("L_27");

	private int[] cellUyQuyen = GeneralFunctions.ArrayCellId("H_11");

	private int[] cellToChuc = GeneralFunctions.ArrayCellId("J_135");

	private int[] cellCaNhan = GeneralFunctions.ArrayCellId("Y_135");

	private int[] cell_I_1_DangKyKinhDoanh = GeneralFunctions.ArrayCellId("N_144");

	private int[] cell_I_1_CCCD = GeneralFunctions.ArrayCellId("Z_144");

	private int[] cell_I_1_So = GeneralFunctions.ArrayCellId("F_147");

	private int[] cell_I_3_start_cb = GeneralFunctions.ArrayCellId("M_160");

	private int[] cell_I_3_end_cb = GeneralFunctions.ArrayCellId("M_164");

	private int[] cell_I_4_cb = GeneralFunctions.ArrayCellId("H_181");

	private int[] cell_I_5_cb = GeneralFunctions.ArrayCellId("H_197");

	private int[] cell_I_7_start_cb = GeneralFunctions.ArrayCellId("H_229");

	private int[] cell_I_7_end_cb = GeneralFunctions.ArrayCellId("L_231");

	private int[] cell_II_1_cb = GeneralFunctions.ArrayCellId("N_266");

	private int[] cell_II_4_1_phap_nhan = GeneralFunctions.ArrayCellId("M_294");

	private int[] cell_II_4_1_phong_phap_nhan = GeneralFunctions.ArrayCellId("M_296");

	private int[] cell_II_4_1_ca_nhan_phu_thuoc = GeneralFunctions.ArrayCellId("Y_294");

	private int[] cell_II_4_1_khac = GeneralFunctions.ArrayCellId("Y_296");

	private int[] cell_II_4_1_dai_ly_thue = GeneralFunctions.ArrayCellId("M_298");

	private int[] cell_I_NgayCap = GeneralFunctions.ArrayCellId("S_147");

	private int[] cell_I_4_1_TenTinh = GeneralFunctions.ArrayCellId("H_171");

	private int[] cell_I_4_1_MaTinh = GeneralFunctions.ArrayCellId("BC_171");

	private int[] cell_I_4_1_TenXa = GeneralFunctions.ArrayCellId("H_173");

	private int[] cell_I_4_1_MaXa = GeneralFunctions.ArrayCellId("BC_173");

	private int[] cell_I_4_1_SoNha = GeneralFunctions.ArrayCellId("K_175");

	private int[] cell_I_4_1_SoDT = GeneralFunctions.ArrayCellId("H_177");

	private int[] cell_I_4_1_Fax = GeneralFunctions.ArrayCellId("O_177");

	private int[] cell_I_4_1_Email = GeneralFunctions.ArrayCellId("U_177");

	private int[] cell_I_4_1_TruSoChinh = GeneralFunctions.ArrayCellId("H_181");

	private int[] cell_I_4_1_NhaOThuongTru = GeneralFunctions.ArrayCellId("O_181");

	private int[] cell_I_4_1_Khac = GeneralFunctions.ArrayCellId("R_181");

	private int[] cell_I_6_Khac = GeneralFunctions.ArrayCellId("K_210");

	private int[] cell_I_9_TuNgay = GeneralFunctions.ArrayCellId("G_244");

	private int[] cell_I_9_DenNgay = GeneralFunctions.ArrayCellId("L_244");

	private int[] cell_II_Start = GeneralFunctions.ArrayCellId("O_261");

	private int[] cell_II_End = GeneralFunctions.ArrayCellId("O_298");

	private int[] cell_II_DangKyKD = GeneralFunctions.ArrayCellId("N_266");

	private int[] cell_II_HoChieu = GeneralFunctions.ArrayCellId("S_266");

	private int[] cell_II_CCCD = GeneralFunctions.ArrayCellId("Z_266");

	private int[] cell_II_1_So = GeneralFunctions.ArrayCellId("F_268");

	private int[] cell_II_NgayCap = GeneralFunctions.ArrayCellId("S_268");

	private int[] cell_II_TenTinh = GeneralFunctions.ArrayCellId("H_275");

	private int[] cell_II_MaTinh = GeneralFunctions.ArrayCellId("BC_275");

	private int[] cell_II_TenXa = GeneralFunctions.ArrayCellId("H_277");

	private int[] cell_II_MaXa = GeneralFunctions.ArrayCellId("BC_277");

	private int[] cell_II_SoNha = GeneralFunctions.ArrayCellId("K_279");

	private int[] cell_II_SoDT = GeneralFunctions.ArrayCellId("H_281");

	private int[] cell_II_Fax = GeneralFunctions.ArrayCellId("O_281");

	private int[] cell_II_Email = GeneralFunctions.ArrayCellId("U_281");

	private int[] cell_II_Mst = GeneralFunctions.ArrayCellId("H_284");

	private int[] cell_III_1_TuNam = GeneralFunctions.ArrayCellId("Q_316");

	private int[] cell_III_1_DenNam = GeneralFunctions.ArrayCellId("X_316");

	private int columnMaKieuKy_III_3_1 = HTKKFile.ColumnLettersToIndex("BI");

	private int columnTenKieuKy_III_3_1 = HTKKFile.ColumnLettersToIndex("AS");

	private int columnKyThue_III_3_1 = HTKKFile.ColumnLettersToIndex("AT");

	private int column2_III_3_1 = HTKKFile.ColumnLettersToIndex("D");

	private int columnMa4a_III_3_1 = HTKKFile.ColumnLettersToIndex("BD");

	private int columnMa4b_III_3_1 = HTKKFile.ColumnLettersToIndex("BE");

	private int columnKhac4c_III_3_1 = HTKKFile.ColumnLettersToIndex("AC");

	private int columnTen4a_III_3_1 = HTKKFile.ColumnLettersToIndex("Q");

	private int columnTen4b_III_3_1 = HTKKFile.ColumnLettersToIndex("W");

	private int colCt7_kyKKhaiTuNgay = HTKKFile.ColumnLettersToIndex("BL");

	private int colCt7_kyKKhaiDenNgay = HTKKFile.ColumnLettersToIndex("BM");

	private int colCt7_kyKKhaiTuThang = HTKKFile.ColumnLettersToIndex("BN");

	private int colCt7_kyKKhaiDenThang = HTKKFile.ColumnLettersToIndex("BO");

	private int[] cellMaNuocKhac = GeneralFunctions.ArrayCellId("BC_197");

	private int[] cellQuocTichKhac = GeneralFunctions.ArrayCellId("S_197");

	private int[] cellTuKy = GeneralFunctions.ArrayCellId("G_4");

	private int[] cellTuKyIsNumber = GeneralFunctions.ArrayCellId("N_4");

	private int[] cellDenKy = GeneralFunctions.ArrayCellId("I_4");

	private int[] cellDenKyIsNumber = GeneralFunctions.ArrayCellId("O_4");

	private int colKieuKy = HTKKFile.ColumnLettersToIndex("I");

	private int colKyKK = HTKKFile.ColumnLettersToIndex("J");

	private int colKyKKTemp = HTKKFile.ColumnLettersToIndex("T");

	private int colMaKieuKy = HTKKFile.ColumnLettersToIndex("O");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[cellLoaiKy[1], cellLoaiKy[0]].Value = ((PeriodInfo.KieuKy == "Y") ? "Y" : "D");
			sheetView.Cells[cellTenTruongHopHoanThue[1], cellTenTruongHopHoanThue[0]].Value = "Hoàn thuế theo Hiệp định tránh đánh thuế hai lần";
			sheetView.Cells[cellMaTruongHopHoanThue[1], cellMaTruongHopHoanThue[0]].Value = "01";
			if (PeriodInfo.KieuKy == "Y")
			{
				sheetView.Cells[cellTuKy_tk[1], cellTuKy_tk[0]].Value = PeriodInfo.TuThang;
				sheetView.Cells[cellDenKy_tk[1], cellDenKy_tk[0]].Value = PeriodInfo.DenThang;
				sheetView.Cells[cellNgayPSNV[1], cellNgayPSNV[0]].Locked = false;
				sheetView.Cells[cellNgayKTNV[1], cellNgayKTNV[0]].Locked = false;
			}
			if (PeriodInfo.KieuKy == "D_D")
			{
				sheetView.Cells[cellTuKy_tk[1], cellTuKy_tk[0]].Value = PeriodInfo.TuNgay;
				sheetView.Cells[cellDenKy_tk[1], cellDenKy_tk[0]].Value = PeriodInfo.DenNgay;
				sheetView.Cells[cellNgayPSNV[1], cellNgayPSNV[0]].Value = "";
				sheetView.Cells[cellNgayPSNV[1], cellNgayPSNV[0]].Locked = true;
				sheetView.Cells[cellNgayKTNV[1], cellNgayKTNV[0]].Value = "";
				sheetView.Cells[cellNgayKTNV[1], cellNgayKTNV[0]].Locked = true;
			}
			load4_1TheoUyQuyen(sheetView, TTinChung, hTKKCMCSheetInfo);
			int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[1].RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, columnKyThue_III_3_1 - 1].Text.Trim();
				string typeKeKhai = sheetView.Cells[dataLocationTemplateRowStart + i - 1, columnMaKieuKy_III_3_1 - 1].Text.Trim();
				GeneralFunctions.SetCellTypeTheoKieuKy(sheetView, typeKeKhai, dataLocationTemplateRowStart + i, columnKyThue_III_3_1);
			}
			bool flag = Convert.ToBoolean(sheetView.Cells[cellToChuc[1], cellToChuc[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellCaNhan[1], cellCaNhan[0]].Value);
			if (!flag && !flag2)
			{
				sheetView.Cells[cell_II_4_1_dai_ly_thue[1], cell_II_4_1_dai_ly_thue[0]].Value = !string.IsNullOrEmpty(TTinChung.mstDLyThue);
			}
			load_II_TheoToChucCaNhan(sheetView, hTKKCMCSheetInfo, TTinChung);
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
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

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_06aa: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == columnTenKieuKy_III_3_1)
			{
				string text = sheetView.Cells[(int)iRow - 1, columnKyThue_III_3_1 - 1].Text.Trim();
				string typeKeKhai = sheetView.Cells[(int)iRow - 1, columnMaKieuKy_III_3_1 - 1].Text.Trim();
				GeneralFunctions.SetCellTypeTheoKieuKy(sheetView, typeKeKhai, (int)iRow, columnKyThue_III_3_1);
			}
			if (iCol == columnTenKieuKy_III_3_1 || iCol == columnKyThue_III_3_1)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, columnMaKieuKy_III_3_1 - 1].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, columnKyThue_III_3_1 - 1].Text;
				frmMessage frmMessage = new frmMessage("0456", "3", "", "Thông báo");
				if (string.IsNullOrEmpty(text3))
				{
					return;
				}
				switch (text2)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text3, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuNgay - 1].Text = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenNgay - 1].Text = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuThang - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenThang - 1].Text = text3;
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuNgay - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenNgay - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuThang - 1].Text = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy").ToString("MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenThang - 1].Text = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy").ToString("MM/yyyy");
					break;
				case "Y":
					if (ConvertHelpers.ConvertToDateTime("01/01/" + text3, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuNgay - 1].Text = DateTime.ParseExact("01/01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenNgay - 1].Text = DateTime.ParseExact("01/12/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuThang - 1].Text = "01/" + text3;
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenThang - 1].Text = "12/" + text3;
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text3.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuNgay - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text3.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text3.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenNgay - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text3.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text3.Split(new char[1] { '/' })[1]), 1, 1).AddMonths(3).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiTuThang - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, colCt7_kyKKhaiDenThang - 1].Text = text3;
					break;
				default:
					sheetView.Cells[(int)iRow - 1, columnKyThue_III_3_1 - 1].Text = "";
					((Form)frmMessage).ShowDialog();
					break;
				}
			}
			if (cellUyQuyen[1] + 1 == iRow || cellVietNam[1] + 1 == iRow || cellToChuc[1] + 1 == iRow || (cell_I_3_start_cb[1] + 1 <= iRow && iRow <= cell_I_3_end_cb[1] + 1) || cell_I_5_cb[1] + 1 == iRow || iRow == cell_I_4_cb[1] + 1)
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
			SheetView toSView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo toSInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
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
			int value = HTKKFile.ColumnLettersToIndex("O");
			int num = HTKKFile.ColumnLettersToIndex("L");
			int key = HTKKFile.ColumnLettersToIndex("N");
			int num2 = HTKKFile.ColumnLettersToIndex("H");
			int value2 = HTKKFile.ColumnLettersToIndex("D");
			int value3 = HTKKFile.ColumnLettersToIndex("BC");
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(num2, value2);
			dictionary.Add(key, value3);
			dictionary.Add(num, value);
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			if (iCol == num || iCol == num2)
			{
				GeneralFunctions.CopyTableToTable4(sheetView, toSView, toSInfo, hTKKCMCSheetInfo.Tables[0], 0, dictionary, deleteDataOld: true, ref checkExists);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
		{
			return false;
		}
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_56"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_58"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_61"));
			string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_67"));
			string message6 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_20"));
			string message7 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_55"));
			string message8 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_82"));
			string message9 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_84"));
			string message10 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_85"));
			DateTime utcNow = DateTime.UtcNow;
			string text = sView.Cells[cell_I_NgayCap[1], cell_I_NgayCap[0]].Text;
			if (!string.IsNullOrEmpty(text) && text.Length == 10)
			{
				DateTime t = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cell_I_NgayCap[1], cell_I_NgayCap[0], message);
					value = 2;
				}
			}
			string text2 = sView.Cells[cell_I_9_TuNgay[1], cell_I_9_TuNgay[0]].Text;
			string text3 = sView.Cells[cell_I_9_DenNgay[1], cell_I_9_DenNgay[0]].Text;
			if (!string.IsNullOrEmpty(text2) && text2.Length == 10 && !string.IsNullOrEmpty(text3) && text3.Length == 10)
			{
				DateTime t2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime t3 = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t2, t3) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cell_I_9_TuNgay[1], cell_I_9_TuNgay[0], message5);
					HTKKFile.SetNodeBackGroundPink(sView, cell_I_9_DenNgay[1], cell_I_9_DenNgay[0], message5);
					value = 2;
				}
			}
			string text4 = sView.Cells[cell_I_1_So[1], cell_I_1_So[0]].Text;
			if (Convert.ToBoolean(sView.Cells[cell_I_1_CCCD[1], cell_I_1_CCCD[0]].Value) && !string.IsNullOrEmpty(text4) && !GeneralFunctions.checkCCCD_CMND(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sView, cell_I_1_So[1], cell_I_1_So[0], message6);
				value = 2;
			}
			string text5 = sView.Cells[cell_II_NgayCap[1], cell_II_NgayCap[0]].Text;
			if (!string.IsNullOrEmpty(text5) && text5.Length == 10)
			{
				DateTime t = DateTime.ParseExact(text5, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, cell_II_NgayCap[1], cell_II_NgayCap[0], message);
					value = 2;
				}
			}
			string text6 = sView.Cells[cell_II_1_So[1], cell_II_1_So[0]].Text;
			if (Convert.ToBoolean(sView.Cells[cell_II_CCCD[1], cell_II_CCCD[0]].Value) && !string.IsNullOrEmpty(text6) && !GeneralFunctions.checkCCCD_CMND(text6))
			{
				HTKKFile.SetNodeBackGroundPink(sView, cell_II_1_So[1], cell_II_1_So[0], message6);
				value = 2;
			}
			string text7 = ((PeriodInfo.KieuKy == "Y") ? PeriodInfo.TuThang : PeriodInfo.TuNgay.Substring(PeriodInfo.TuNgay.Length - 4, 4));
			string value2 = ((PeriodInfo.KieuKy == "Y") ? PeriodInfo.DenThang : PeriodInfo.DenNgay.Substring(PeriodInfo.DenNgay.Length - 4, 4));
			string text8 = sView.Cells[cell_III_1_TuNam[1], cell_III_1_TuNam[0]].Text;
			string text9 = sView.Cells[cell_III_1_DenNam[1], cell_III_1_DenNam[0]].Text;
			if (!string.IsNullOrEmpty(text8) && !string.IsNullOrEmpty(text9) && text7.Length == 4 && text8.Length == 4 && text9.Length == 4)
			{
				if (Convert.ToInt32(text8) > Convert.ToInt32(text9))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cell_III_1_TuNam[1], cell_III_1_TuNam[0], message7);
					HTKKFile.SetNodeBackGroundPink(sView, cell_III_1_DenNam[1], cell_III_1_DenNam[0], message7);
					value = 2;
				}
				else if (Convert.ToInt32(text8) > Convert.ToInt32(text7) || Convert.ToInt32(value2) > Convert.ToInt32(text9))
				{
					HTKKFile.SetNodeBackGroundPink(sView, cell_III_1_TuNam[1], cell_III_1_TuNam[0], message2);
					HTKKFile.SetNodeBackGroundPink(sView, cell_III_1_DenNam[1], cell_III_1_DenNam[0], message2);
					value = 2;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text10 = sView.Cells[dataLocationTemplateRowStart + i - 1, columnMaKieuKy_III_3_1 - 1].Text;
				string text11 = sView.Cells[dataLocationTemplateRowStart + i - 1, columnKyThue_III_3_1 - 1].Text;
				if (text10 == "D")
				{
					DateTime t4 = DateTime.ParseExact(text11, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (DateTime.Compare(t4, utcNow) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, columnKyThue_III_3_1 - 1, message);
						value = 2;
					}
				}
				if (text10 == "M")
				{
					DateTime t4 = DateTime.ParseExact("01/" + text11, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (DateTime.Compare(t4, utcNow) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, columnKyThue_III_3_1 - 1, message);
						value = 2;
					}
				}
				if (text10 == "Q")
				{
					DateTime t4 = ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text11.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text11.Split(new char[1] { '/' })[1], "dd/MM/yyyy");
					if (DateTime.Compare(t4, utcNow) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, columnKyThue_III_3_1 - 1, message);
						value = 2;
					}
				}
			}
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text12 = sView.Cells[dataLocationTemplateRowStart + i - 1, column2_III_3_1 - 1].Text;
				string text13 = sView.Cells[dataLocationTemplateRowStart + i - 1, columnMa4a_III_3_1 - 1].Text;
				string text14 = sView.Cells[dataLocationTemplateRowStart + i - 1, columnMa4b_III_3_1 - 1].Text;
				string text15 = sView.Cells[dataLocationTemplateRowStart + i - 1, columnKhac4c_III_3_1 - 1].Text;
				if (!string.IsNullOrEmpty(text12) && ((string.IsNullOrEmpty(text13) && string.IsNullOrEmpty(text14) && string.IsNullOrEmpty(text15)) || (!string.IsNullOrEmpty(text13) && !string.IsNullOrEmpty(text14) && !string.IsNullOrEmpty(text15)) || (!string.IsNullOrEmpty(text13) && !string.IsNullOrEmpty(text14) && string.IsNullOrEmpty(text15)) || (string.IsNullOrEmpty(text13) && !string.IsNullOrEmpty(text14) && !string.IsNullOrEmpty(text15)) || (!string.IsNullOrEmpty(text13) && string.IsNullOrEmpty(text14) && !string.IsNullOrEmpty(text15))))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, columnTen4a_III_3_1 - 1, message4);
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, columnTen4b_III_3_1 - 1, message4);
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, columnKhac4c_III_3_1 - 1, message4);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, columnTen4a_III_3_1 - 1);
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, columnTen4b_III_3_1 - 1);
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, columnKhac4c_III_3_1 - 1);
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("O");
			int num2 = HTKKFile.ColumnLettersToIndex("U");
			int num3 = HTKKFile.ColumnLettersToIndex("AB");
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num3 - 1);
				double num4 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart2 + i - 1, num - 1].Value);
				double num5 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart2 + i - 1, num2 - 1].Value);
				double num6 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart2 + i - 1, num3 - 1].Value);
				if (num5 <= 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1, message8);
					value = 2;
				}
				if (num6 <= 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num3 - 1, message8);
					value = 2;
				}
				if (num5 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1, message9);
					value = 2;
				}
				if (num6 > num5)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num3 - 1, message10);
					value = 2;
				}
			}
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			int num2 = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			bool flag = Convert.ToBoolean(sheetView.Cells[cellUyQuyen[1], cellUyQuyen[0]].Value);
			if (cellUyQuyen[1] + 1 == num2 && cellUyQuyen[0] + 1 == num)
			{
				load4_1TheoUyQuyen(sheetView, TTinChung, sInfo);
			}
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellToChuc[1], cellToChuc[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[cellCaNhan[1], cellCaNhan[0]].Value);
			if (flag2 && num == cellToChuc[0] + 1)
			{
				sheetView.Cells[cellCaNhan[1], cellCaNhan[0]].Value = false;
				flag3 = false;
			}
			if (flag3 && num == cellCaNhan[0] + 1)
			{
				sheetView.Cells[cellToChuc[1], cellToChuc[0]].Value = false;
				flag2 = false;
			}
			if (num2 == cellToChuc[1] + 1)
			{
				if (flag2)
				{
					sheetView.Cells[cell_I_4_1_TruSoChinh[1], cell_I_4_1_TruSoChinh[0]].Value = true;
					resetOtherCheckboxTK(sheetView, sInfo, cell_I_4_1_TruSoChinh[1] + 1, cell_I_4_1_TruSoChinh[0] + 1);
				}
				else
				{
					sheetView.Cells[cell_I_4_1_TruSoChinh[1], cell_I_4_1_TruSoChinh[0]].Value = false;
				}
				if (flag3)
				{
					sheetView.Cells[cell_I_4_1_NhaOThuongTru[1], cell_I_4_1_NhaOThuongTru[0]].Value = true;
					resetOtherCheckboxTK(sheetView, sInfo, cell_I_4_1_NhaOThuongTru[1] + 1, cell_I_4_1_NhaOThuongTru[0] + 1);
				}
			}
			if (flag2 && num == cellToChuc[0] + 1)
			{
				sheetView.Cells[cell_I_1_DangKyKinhDoanh[1], cell_I_1_DangKyKinhDoanh[0]].Value = true;
				resetOtherCheckboxTK(sheetView, sInfo, cell_I_1_DangKyKinhDoanh[1] + 1, cell_I_1_DangKyKinhDoanh[0] + 1);
			}
			if (!flag2 && num == cellCaNhan[0] + 1)
			{
				sheetView.Cells[cell_I_1_CCCD[1], cell_I_1_CCCD[0]].Value = true;
				resetOtherCheckboxTK(sheetView, sInfo, cell_I_1_CCCD[1] + 1, cell_I_1_CCCD[0] + 1);
			}
			if (!flag)
			{
				sheetView.Cells[cell_I_6_Khac[1], cell_I_6_Khac[0]].Value = TTinChung.mst;
			}
			if (flag && num2 == cellUyQuyen[1] + 1)
			{
				sheetView.Cells[cell_I_6_Khac[1], cell_I_6_Khac[0]].Value = "";
			}
			load_II_TheoToChucCaNhan(sheetView, sInfo, TTinChung);
			if (flag)
			{
				if (flag2)
				{
					GeneralFunctions.LockUnLockCol(sheetView, cell_II_4_1_phap_nhan[1] + 1, cell_II_4_1_phap_nhan[0] + 1, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView, cell_II_4_1_phong_phap_nhan[1] + 1, cell_II_4_1_phong_phap_nhan[0] + 1, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView, cell_II_4_1_ca_nhan_phu_thuoc[1] + 1, cell_II_4_1_ca_nhan_phu_thuoc[0] + 1, locked: true, deleteData: true);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView, cell_II_4_1_phap_nhan[1] + 1, cell_II_4_1_phap_nhan[0] + 1, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, cell_II_4_1_phong_phap_nhan[1] + 1, cell_II_4_1_phong_phap_nhan[0] + 1, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, cell_II_4_1_ca_nhan_phu_thuoc[1] + 1, cell_II_4_1_ca_nhan_phu_thuoc[0] + 1, locked: false, deleteData: false);
				}
			}
			resetOtherCheckboxTK(sheetView, sInfo, num2, num);
			if (cellUyQuyen[1] + 1 == num2 || cellVietNam[1] + 1 == num2 || cellToChuc[1] + 1 == num2 || (cell_I_3_start_cb[1] + 1 <= num2 && num2 <= cell_I_3_end_cb[1] + 1) || cell_I_5_cb[1] + 1 == num2 || num2 == cell_I_4_cb[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
				if (!Convert.ToBoolean(sheetView.Cells[cellQuocTichKhac[1], cellQuocTichKhac[0]].Value) || flag2)
				{
					sheetView.Cells[cellMaNuocKhac[1], cellMaNuocKhac[0]].Value = "";
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
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_02_1_HT_TT80.xls");
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

	public void load4_1TheoUyQuyen(SheetView sView, HTKKCMCTTinChungInfo TTinChung, HTKKCMCSheetInfo sInfo)
	{
		if (!Convert.ToBoolean(sView.Cells[cellUyQuyen[1], cellUyQuyen[0]].Value))
		{
			sView.Cells[cell_I_4_1_TenTinh[1], cell_I_4_1_TenTinh[0]].Value = TTinChung.tenTinhNNT;
			sView.Cells[cell_I_4_1_MaTinh[1], cell_I_4_1_MaTinh[0]].Value = TTinChung.maTinhNNT;
			sView.Cells[cell_I_4_1_TenXa[1], cell_I_4_1_TenXa[0]].Value = TTinChung.tenXaPhuongNNT;
			sView.Cells[cell_I_4_1_MaXa[1], cell_I_4_1_MaXa[0]].Value = TTinChung.maXaPhuongNNT;
			sView.Cells[cell_I_4_1_SoNha[1], cell_I_4_1_SoNha[0]].Value = TTinChung.dchiNNT;
			sView.Cells[cell_I_4_1_SoDT[1], cell_I_4_1_SoDT[0]].Value = TTinChung.dthoaiNNT;
			sView.Cells[cell_I_4_1_Fax[1], cell_I_4_1_Fax[0]].Value = TTinChung.faxNNT;
			sView.Cells[cell_I_4_1_Email[1], cell_I_4_1_Email[0]].Value = TTinChung.emailNNT;
		}
	}

	public void load_II_TheoToChucCaNhan(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTTinChungInfo TTinChung)
	{
		if (!Convert.ToBoolean(sView.Cells[cellUyQuyen[1], cellUyQuyen[0]].Value))
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cell_II_Start[1] + 1, cell_II_End[1] + 1, isLock: true, isDeleteData: true);
			return;
		}
		sView.Cells[cell_II_Start[1], cell_II_Start[0]].Value = true;
		GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cell_II_Start[1] + 1, cell_II_End[1] + 1, isLock: false, isDeleteData: false);
		bool flag = Convert.ToBoolean(sView.Cells[cellToChuc[1], cellToChuc[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellCaNhan[1], cellCaNhan[0]].Value);
		bool flag3 = Convert.ToBoolean(sView.Cells[cell_II_HoChieu[1], cell_II_HoChieu[0]].Value);
		bool flag4 = Convert.ToBoolean(sView.Cells[cell_II_CCCD[1], cell_II_CCCD[0]].Value);
		if (flag)
		{
			sView.Cells[cell_II_DangKyKD[1], cell_II_DangKyKD[0]].Value = true;
			sView.Cells[cell_II_HoChieu[1], cell_II_HoChieu[0]].Value = false;
			sView.Cells[cell_II_CCCD[1], cell_II_CCCD[0]].Value = false;
		}
		if (flag2 && !flag3 && !flag4)
		{
			sView.Cells[cell_II_DangKyKD[1], cell_II_DangKyKD[0]].Value = false;
			sView.Cells[cell_II_HoChieu[1], cell_II_HoChieu[0]].Value = false;
			sView.Cells[cell_II_CCCD[1], cell_II_CCCD[0]].Value = true;
		}
		sView.Cells[cell_II_TenTinh[1], cell_II_TenTinh[0]].Value = TTinChung.tenTinhNNT;
		sView.Cells[cell_II_MaTinh[1], cell_II_MaTinh[0]].Value = TTinChung.maTinhNNT;
		sView.Cells[cell_II_TenXa[1], cell_II_TenXa[0]].Value = TTinChung.tenXaPhuongNNT;
		sView.Cells[cell_II_MaXa[1], cell_II_MaXa[0]].Value = TTinChung.maXaPhuongNNT;
		sView.Cells[cell_II_SoNha[1], cell_II_SoNha[0]].Value = TTinChung.dchiNNT;
		sView.Cells[cell_II_SoDT[1], cell_II_SoDT[0]].Value = TTinChung.dthoaiNNT;
		sView.Cells[cell_II_Fax[1], cell_II_Fax[0]].Value = TTinChung.faxNNT;
		sView.Cells[cell_II_Email[1], cell_II_Email[0]].Value = TTinChung.emailNNT;
		sView.Cells[cell_II_Mst[1], cell_II_Mst[0]].Value = TTinChung.mst;
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

	public void resetOtherCheckboxTK(SheetView sView, HTKKCMCSheetInfo sInfo, int iRowClick, int iColClick)
	{
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cellVietNam[1] + 1, cellVietNam[1] + 1, iRowClick, iColClick);
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cellToChuc[1] + 1, cellCaNhan[1] + 1, iRowClick, iColClick);
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_I_1_DangKyKinhDoanh[1] + 1, cell_I_1_DangKyKinhDoanh[1] + 1, iRowClick, iColClick);
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_I_3_start_cb[1] + 1, cell_I_3_end_cb[1] + 1, iRowClick, iColClick);
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_I_4_cb[1] + 1, cell_I_4_cb[1] + 1, iRowClick, iColClick);
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_I_5_cb[1] + 1, cell_I_5_cb[1] + 1, iRowClick, iColClick);
		GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_I_7_start_cb[1] + 1, cell_I_7_end_cb[1] + 1, iRowClick, iColClick);
		if (Convert.ToBoolean(sView.Cells[cellUyQuyen[1], cellUyQuyen[0]].Value))
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_II_1_cb[1] + 1, cell_II_1_cb[1] + 1, iRowClick, iColClick);
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_II_4_1_phap_nhan[1] + 1, cell_II_4_1_khac[1] + 1, iRowClick, iColClick);
		}
	}
}
