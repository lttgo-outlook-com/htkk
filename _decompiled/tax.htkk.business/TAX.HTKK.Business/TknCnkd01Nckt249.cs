using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TknCnkd01Nckt249
{
	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private List<item> _dmBVMTData = new List<item>();

	private bool _isDmMoi = false;

	private string namKK = "";

	private string kieuKyKK = "";

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			namKK = PeriodInfo.Nam;
			kieuKyKK = PeriodInfo.KieuKy;
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
			int num = HTKKFile.ColumnLettersToIndex("L");
			int[] array = GeneralFunctions.ArrayCellId("F_46");
			int[] array2 = GeneralFunctions.ArrayCellId("F_52");
			int[] array3 = GeneralFunctions.ArrayCellId("E_94");
			int[] array4 = GeneralFunctions.ArrayCellId("J_94");
			int[] array5 = GeneralFunctions.ArrayCellId("Q_94");
			int[] array6 = GeneralFunctions.ArrayCellId("I_169");
			int[] array7 = GeneralFunctions.ArrayCellId("I_171");
			int[] array8 = GeneralFunctions.ArrayCellId("E_173");
			int[] array9 = GeneralFunctions.ArrayCellId("E_175");
			int[] array10 = GeneralFunctions.ArrayCellId("J_175");
			int[] array11 = GeneralFunctions.ArrayCellId("Q_175");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			int num3 = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AD");
			int num6 = 29;
			if (string.IsNullOrEmpty(sView.Cells[num6 - 1, num2 - 1].Text) && !string.IsNullOrEmpty(PeriodInfo.ToChucKhaiThay))
			{
				sView.Cells[num6 - 1, num2 - 1].Text = PeriodInfo.ToChucKhaiThay;
			}
			if (string.IsNullOrEmpty(sView.Cells[num6 - 1, num4 - 1].Text.Trim()) && !string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy))
			{
				sView.Cells[num6 - 1, num4 - 1].Text = PeriodInfo.tenCQTQuanLy;
				sView.Cells[num6 - 1, num5 - 1].Text = PeriodInfo.maCQTQuanLy;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SetValue(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		CheckValue(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
	}

	public void CheckValue(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		int num = 208;
		int num2 = 209;
		int num3 = 210;
		int num4 = 211;
		int num5 = 212;
		int num6 = 213;
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int num7 = HTKKFile.ColumnLettersToIndex("G");
		int num8 = HTKKFile.ColumnLettersToIndex("J");
		int num9 = HTKKFile.ColumnLettersToIndex("M");
		int num10 = HTKKFile.ColumnLettersToIndex("P");
		int num11 = HTKKFile.ColumnLettersToIndex("S");
		int num12 = HTKKFile.ColumnLettersToIndex("V");
		int num13 = HTKKFile.ColumnLettersToIndex("Y");
		int num14 = HTKKFile.ColumnLettersToIndex("E");
		int num15 = HTKKFile.ColumnLettersToIndex("AB");
		int num16 = HTKKFile.ColumnLettersToIndex("AE");
		int[] array = GeneralFunctions.ArrayCellId("H_67");
		int[] array2 = GeneralFunctions.ArrayCellId("AJ_67");
		string pathCt6b = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/phuongXa";
		string pathCt15a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct15";
		string pathCt16a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct16";
		string pathCt17a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct17";
		string pathCt18a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct18";
		string pathCt19a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct19";
		string pathCt15b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct15";
		string pathCt16b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct16";
		string pathCt17b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct17";
		string pathCt18b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct18";
		string pathCt19b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct19";
		string pathCt15c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct15";
		string pathCt16c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct16";
		string pathCt17c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct17";
		string pathCt18c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct18";
		string pathCt19c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct19";
		string pathCt15d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct15";
		string pathCt16d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct16";
		string pathCt17d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct17";
		string pathCt18d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct18";
		string pathCt19d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct19";
		string pathCt15dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct15";
		string pathCt16dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct16";
		string pathCt17dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct17";
		string pathCt18dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct18";
		string pathCt19dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct19";
		string pathCt10a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct10a_ngaySinh";
		string pathCt10k = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct10k";
		string pathCt10c1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct10c_1_ngayCap";
		string pathCt10d1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10d_1_ngayCap";
		string pathCt10e1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10e_1_ngayCap";
		string pathCt10dd1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10dd_1_ngayCap";
		string pathCt10f1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10f_1_ngayCap";
		string pathCt10i1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10i/ct10i_ngayCap";
		string pathCt14_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT14/ngayVanBanUQ";
		HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt6b));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15a));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16a));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17a));
		HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18a));
		HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19a));
		HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15b));
		HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16b));
		HTKKCMCCellInfo hTKKCMCCellInfo9 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17b));
		HTKKCMCCellInfo hTKKCMCCellInfo10 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18b));
		HTKKCMCCellInfo hTKKCMCCellInfo11 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19b));
		HTKKCMCCellInfo hTKKCMCCellInfo12 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15c));
		HTKKCMCCellInfo hTKKCMCCellInfo13 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16c));
		HTKKCMCCellInfo hTKKCMCCellInfo14 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17c));
		HTKKCMCCellInfo hTKKCMCCellInfo15 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18c));
		HTKKCMCCellInfo hTKKCMCCellInfo16 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19c));
		HTKKCMCCellInfo hTKKCMCCellInfo17 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15d));
		HTKKCMCCellInfo hTKKCMCCellInfo18 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16d));
		HTKKCMCCellInfo hTKKCMCCellInfo19 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17d));
		HTKKCMCCellInfo hTKKCMCCellInfo20 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18d));
		HTKKCMCCellInfo hTKKCMCCellInfo21 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19d));
		HTKKCMCCellInfo hTKKCMCCellInfo22 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15dd));
		HTKKCMCCellInfo hTKKCMCCellInfo23 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16dd));
		HTKKCMCCellInfo hTKKCMCCellInfo24 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17dd));
		HTKKCMCCellInfo hTKKCMCCellInfo25 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18dd));
		HTKKCMCCellInfo hTKKCMCCellInfo26 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19dd));
		double num17 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num18 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num19 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num20 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num21 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num22 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num23 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num24 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num25 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num26 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num27 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num28 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num29 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num30 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num31 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num32 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num33 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num34 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num35 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo20.CellTemplate.RowIndex - 1, hTKKCMCCellInfo20.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num36 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo21.CellTemplate.RowIndex - 1, hTKKCMCCellInfo21.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num37 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo22.CellTemplate.RowIndex - 1, hTKKCMCCellInfo22.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num38 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo23.CellTemplate.RowIndex - 1, hTKKCMCCellInfo23.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num39 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo24.CellTemplate.RowIndex - 1, hTKKCMCCellInfo24.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num40 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo25.CellTemplate.RowIndex - 1, hTKKCMCCellInfo25.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		double num41 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo26.CellTemplate.RowIndex - 1, hTKKCMCCellInfo26.CellTemplate.ColIndex - 1].Text.Replace(".", "").Trim());
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCCellInfo hTKKCMCCellInfo27 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10k));
			HTKKCMCCellInfo hTKKCMCCellInfo28 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10a));
			HTKKCMCCellInfo hTKKCMCCellInfo29 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10c1));
			HTKKCMCCellInfo hTKKCMCCellInfo30 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10d1));
			HTKKCMCCellInfo hTKKCMCCellInfo31 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10dd1));
			HTKKCMCCellInfo hTKKCMCCellInfo32 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10e1));
			HTKKCMCCellInfo hTKKCMCCellInfo33 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10f1));
			HTKKCMCCellInfo hTKKCMCCellInfo34 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt14_1));
			HTKKCMCCellInfo hTKKCMCCellInfo35 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10i1));
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo27.CellTemplate.RowIndex - 1, hTKKCMCCellInfo27.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo28.CellTemplate.RowIndex - 1, hTKKCMCCellInfo28.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo29.CellTemplate.RowIndex - 1, hTKKCMCCellInfo29.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo30.CellTemplate.RowIndex - 1, hTKKCMCCellInfo30.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo31.CellTemplate.RowIndex - 1, hTKKCMCCellInfo31.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo32.CellTemplate.RowIndex - 1, hTKKCMCCellInfo32.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, hTKKCMCCellInfo33.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, hTKKCMCCellInfo35.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, hTKKCMCCellInfo34.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1);
			string text7 = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			if (!HTKKFormater.isNullString(text7) && !string.IsNullOrEmpty(text7) && !TTinChung.maCQTQuanLy.StartsWith(text7))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text6);
				value = 2;
			}
			string value2 = sheetView.Cells[hTKKCMCCellInfo27.CellTemplate.RowIndex - 1, hTKKCMCCellInfo27.CellTemplate.ColIndex - 1].Text.Trim().Replace(".", "");
			if (Convert.ToDouble(value2) < 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo27.CellTemplate.RowIndex - 1, hTKKCMCCellInfo27.CellTemplate.ColIndex - 1, text2);
				value = 2;
			}
			string text8 = sheetView.Cells[hTKKCMCCellInfo28.CellTemplate.RowIndex - 1, hTKKCMCCellInfo28.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text8))
			{
				DateTime t = DateTime.ParseExact(text8, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo28.CellTemplate.RowIndex - 1, hTKKCMCCellInfo28.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text9 = sheetView.Cells[hTKKCMCCellInfo29.CellTemplate.RowIndex - 1, hTKKCMCCellInfo29.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text9))
			{
				DateTime t = DateTime.ParseExact(text9, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo29.CellTemplate.RowIndex - 1, hTKKCMCCellInfo29.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text10 = sheetView.Cells[hTKKCMCCellInfo30.CellTemplate.RowIndex - 1, hTKKCMCCellInfo30.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text10))
			{
				DateTime t = DateTime.ParseExact(text10, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo30.CellTemplate.RowIndex - 1, hTKKCMCCellInfo30.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text11 = sheetView.Cells[hTKKCMCCellInfo31.CellTemplate.RowIndex - 1, hTKKCMCCellInfo31.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text11))
			{
				DateTime t = DateTime.ParseExact(text11, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo31.CellTemplate.RowIndex - 1, hTKKCMCCellInfo31.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text12 = sheetView.Cells[hTKKCMCCellInfo32.CellTemplate.RowIndex - 1, hTKKCMCCellInfo32.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text12))
			{
				DateTime t = DateTime.ParseExact(text12, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo32.CellTemplate.RowIndex - 1, hTKKCMCCellInfo32.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text13 = sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, hTKKCMCCellInfo33.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text13))
			{
				DateTime t = DateTime.ParseExact(text13, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, hTKKCMCCellInfo33.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text14 = sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, hTKKCMCCellInfo35.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text14))
			{
				DateTime t = DateTime.ParseExact(text14, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, hTKKCMCCellInfo35.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			string text15 = sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, hTKKCMCCellInfo34.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text15))
			{
				DateTime t = DateTime.ParseExact(text15, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, hTKKCMCCellInfo34.CellTemplate.ColIndex - 1, text4);
					value = 2;
				}
			}
			double num42 = num17 + num18 + num19 + num20 + num21;
			if (num42 <= 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text3);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, text3);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, text3);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, text3);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, text3);
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

	public void SetValue(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		int num = 208;
		int num2 = 209;
		int num3 = 210;
		int num4 = 211;
		int num5 = 212;
		int num6 = 213;
		if (iRow < num)
		{
			return;
		}
		SheetView sheetView = winFpSpread.Sheets[0];
		int num7 = HTKKFile.ColumnLettersToIndex("G");
		int num8 = HTKKFile.ColumnLettersToIndex("J");
		int num9 = HTKKFile.ColumnLettersToIndex("M");
		int num10 = HTKKFile.ColumnLettersToIndex("P");
		int num11 = HTKKFile.ColumnLettersToIndex("S");
		int num12 = HTKKFile.ColumnLettersToIndex("V");
		int num13 = HTKKFile.ColumnLettersToIndex("Y");
		int num14 = HTKKFile.ColumnLettersToIndex("E");
		int num15 = HTKKFile.ColumnLettersToIndex("AB");
		int num16 = HTKKFile.ColumnLettersToIndex("AE");
		string pathCt15a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct15";
		string pathCt16a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct16";
		string pathCt17a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct17";
		string pathCt18a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct18";
		string pathCt19a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TongDoanhThuTrongNam/ct19";
		string pathCt15b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct15";
		string pathCt16b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct16";
		string pathCt17b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct17";
		string pathCt18b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct18";
		string pathCt19b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTPhaiNop/ct19";
		string pathCt15c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct15";
		string pathCt16c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct16";
		string pathCt17c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct17";
		string pathCt18c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct18";
		string pathCt19c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNPhaiNop/ct19";
		string pathCt15d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct15";
		string pathCt16d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct16";
		string pathCt17d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct17";
		string pathCt18d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct18";
		string pathCt19d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueGTGTDaKhauTruTamNop/ct19";
		string pathCt15dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct15";
		string pathCt16dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct16";
		string pathCt17dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct17";
		string pathCt18dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct18";
		string pathCt19dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ThueTNCNDaKhauTruTamNop/ct19";
		string text = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct10a_ngaySinh";
		string text2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct10k";
		string text3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct10c_1_ngayCap";
		string text4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10d_1_ngayCap";
		string text5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10e_1_ngayCap";
		string text6 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10dd_1_ngayCap";
		string text7 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct10f_1_ngayCap";
		string text8 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10i/ct10i_ngayCap";
		string text9 = "path_S_190";
		double num17 = Convert.ToDouble(sheetView.Cells[num - 1, num7 - 1].Text.Replace(".", "").Trim());
		double num18 = Convert.ToDouble(sheetView.Cells[num2 - 1, num7 - 1].Text.Replace(".", "").Trim());
		double num19 = Convert.ToDouble(sheetView.Cells[num3 - 1, num7 - 1].Text.Replace(".", "").Trim());
		double num20 = Convert.ToDouble(sheetView.Cells[num4 - 1, num7 - 1].Text.Replace(".", "").Trim());
		double num21 = Convert.ToDouble(sheetView.Cells[num5 - 1, num7 - 1].Text.Replace(".", "").Trim());
		double num22 = Convert.ToDouble(sheetView.Cells[num - 1, num8 - 1].Text.Replace(".", "").Trim());
		double num23 = Convert.ToDouble(sheetView.Cells[num2 - 1, num8 - 1].Text.Replace(".", "").Trim());
		double num24 = Convert.ToDouble(sheetView.Cells[num3 - 1, num8 - 1].Text.Replace(".", "").Trim());
		double num25 = Convert.ToDouble(sheetView.Cells[num4 - 1, num8 - 1].Text.Replace(".", "").Trim());
		double num26 = Convert.ToDouble(sheetView.Cells[num5 - 1, num8 - 1].Text.Replace(".", "").Trim());
		double num27 = Convert.ToDouble(sheetView.Cells[num - 1, num9 - 1].Text.Replace(".", "").Trim());
		double num28 = Convert.ToDouble(sheetView.Cells[num2 - 1, num9 - 1].Text.Replace(".", "").Trim());
		double num29 = Convert.ToDouble(sheetView.Cells[num3 - 1, num9 - 1].Text.Replace(".", "").Trim());
		double num30 = Convert.ToDouble(sheetView.Cells[num4 - 1, num9 - 1].Text.Replace(".", "").Trim());
		double num31 = Convert.ToDouble(sheetView.Cells[num5 - 1, num9 - 1].Text.Replace(".", "").Trim());
		double num32 = Convert.ToDouble(sheetView.Cells[num - 1, num10 - 1].Text.Replace(".", "").Trim());
		double num33 = Convert.ToDouble(sheetView.Cells[num2 - 1, num10 - 1].Text.Replace(".", "").Trim());
		double num34 = Convert.ToDouble(sheetView.Cells[num3 - 1, num10 - 1].Text.Replace(".", "").Trim());
		double num35 = Convert.ToDouble(sheetView.Cells[num4 - 1, num10 - 1].Text.Replace(".", "").Trim());
		double num36 = Convert.ToDouble(sheetView.Cells[num5 - 1, num10 - 1].Text.Replace(".", "").Trim());
		double num37 = Convert.ToDouble(sheetView.Cells[num - 1, num11 - 1].Text.Replace(".", "").Trim());
		double num38 = Convert.ToDouble(sheetView.Cells[num2 - 1, num11 - 1].Text.Replace(".", "").Trim());
		double num39 = Convert.ToDouble(sheetView.Cells[num3 - 1, num11 - 1].Text.Replace(".", "").Trim());
		double num40 = Convert.ToDouble(sheetView.Cells[num4 - 1, num11 - 1].Text.Replace(".", "").Trim());
		double num41 = Convert.ToDouble(sheetView.Cells[num5 - 1, num11 - 1].Text.Replace(".", "").Trim());
		HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15a));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16a));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17a));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18a));
		HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19a));
		HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15b));
		HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16b));
		HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17b));
		HTKKCMCCellInfo hTKKCMCCellInfo9 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18b));
		HTKKCMCCellInfo hTKKCMCCellInfo10 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19b));
		HTKKCMCCellInfo hTKKCMCCellInfo11 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15c));
		HTKKCMCCellInfo hTKKCMCCellInfo12 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16c));
		HTKKCMCCellInfo hTKKCMCCellInfo13 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17c));
		HTKKCMCCellInfo hTKKCMCCellInfo14 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18c));
		HTKKCMCCellInfo hTKKCMCCellInfo15 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19c));
		HTKKCMCCellInfo hTKKCMCCellInfo16 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15d));
		HTKKCMCCellInfo hTKKCMCCellInfo17 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16d));
		HTKKCMCCellInfo hTKKCMCCellInfo18 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17d));
		HTKKCMCCellInfo hTKKCMCCellInfo19 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18d));
		HTKKCMCCellInfo hTKKCMCCellInfo20 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19d));
		HTKKCMCCellInfo hTKKCMCCellInfo21 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt15dd));
		HTKKCMCCellInfo hTKKCMCCellInfo22 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt16dd));
		HTKKCMCCellInfo hTKKCMCCellInfo23 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt17dd));
		HTKKCMCCellInfo hTKKCMCCellInfo24 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt18dd));
		HTKKCMCCellInfo hTKKCMCCellInfo25 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt19dd));
		try
		{
			double num42 = num17 + num18 + num19 + num20 + num21;
			sheetView.Cells[num6 - 1, num7 - 1].Value = num42;
			if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo.CellTemplate.ColIndex)
			{
				num27 = ((!(num42 > 100000000.0)) ? 0.0 : (num17 * 0.05));
				sheetView.Cells[num - 1, num9 - 1].Value = num27;
			}
			if (iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo2.CellTemplate.ColIndex)
			{
				num28 = ((!(num42 > 100000000.0)) ? 0.0 : (num18 * 0.05));
				sheetView.Cells[num2 - 1, num9 - 1].Value = num28;
			}
			if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex)
			{
				num29 = ((!(num42 > 100000000.0)) ? 0.0 : (num19 * 0.05));
				sheetView.Cells[num3 - 1, num9 - 1].Value = num29;
			}
			double num43 = num22 - num32;
			if (num43 < 0.0)
			{
				num43 = 0.0;
			}
			sheetView.Cells[num - 1, num12 - 1].Value = num43;
			double num44 = num23 - num33;
			if (num44 < 0.0)
			{
				num44 = 0.0;
			}
			sheetView.Cells[num2 - 1, num12 - 1].Value = num44;
			double num45 = num24 - num34;
			if (num45 < 0.0)
			{
				num45 = 0.0;
			}
			sheetView.Cells[num3 - 1, num12 - 1].Value = num45;
			double num46 = num25 - num35;
			if (num46 < 0.0)
			{
				num46 = 0.0;
			}
			sheetView.Cells[num4 - 1, num12 - 1].Value = num46;
			double num47 = num26 - num36;
			if (num47 < 0.0)
			{
				num47 = 0.0;
			}
			sheetView.Cells[num5 - 1, num12 - 1].Value = num47;
			double num48 = num27 - num37;
			if (num48 < 0.0)
			{
				num48 = 0.0;
			}
			sheetView.Cells[num - 1, num13 - 1].Value = num48;
			double num49 = num28 - num38;
			if (num49 < 0.0)
			{
				num49 = 0.0;
			}
			sheetView.Cells[num2 - 1, num13 - 1].Value = num49;
			double num50 = num29 - num39;
			if (num50 < 0.0)
			{
				num50 = 0.0;
			}
			sheetView.Cells[num3 - 1, num13 - 1].Value = num50;
			double num51 = num30 - num40;
			if (num51 < 0.0)
			{
				num51 = 0.0;
			}
			sheetView.Cells[num4 - 1, num13 - 1].Value = num51;
			double num52 = num31 - num41;
			if (num52 < 0.0)
			{
				num52 = 0.0;
			}
			sheetView.Cells[num5 - 1, num13 - 1].Value = num52;
			double num53 = num32 - num22;
			if (num53 < 0.0)
			{
				num53 = 0.0;
			}
			sheetView.Cells[num - 1, num15 - 1].Value = num53;
			double num54 = num33 - num23;
			if (num54 < 0.0)
			{
				num54 = 0.0;
			}
			sheetView.Cells[num2 - 1, num15 - 1].Value = num54;
			double num55 = num34 - num24;
			if (num55 < 0.0)
			{
				num55 = 0.0;
			}
			sheetView.Cells[num3 - 1, num15 - 1].Value = num55;
			double num56 = num35 - num25;
			if (num56 < 0.0)
			{
				num56 = 0.0;
			}
			sheetView.Cells[num4 - 1, num15 - 1].Value = num56;
			double num57 = num36 - num26;
			if (num57 < 0.0)
			{
				num57 = 0.0;
			}
			sheetView.Cells[num5 - 1, num15 - 1].Value = num57;
			double num58 = num37 - num27;
			if (num58 < 0.0)
			{
				num58 = 0.0;
			}
			sheetView.Cells[num - 1, num16 - 1].Value = num58;
			double num59 = num38 - num28;
			if (num59 < 0.0)
			{
				num59 = 0.0;
			}
			sheetView.Cells[num2 - 1, num16 - 1].Value = num59;
			double num60 = num39 - num29;
			if (num60 < 0.0)
			{
				num60 = 0.0;
			}
			sheetView.Cells[num3 - 1, num16 - 1].Value = num60;
			double num61 = num40 - num30;
			if (num61 < 0.0)
			{
				num61 = 0.0;
			}
			sheetView.Cells[num4 - 1, num16 - 1].Value = num61;
			double num62 = num41 - num31;
			if (num62 < 0.0)
			{
				num62 = 0.0;
			}
			sheetView.Cells[num5 - 1, num16 - 1].Value = num62;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
