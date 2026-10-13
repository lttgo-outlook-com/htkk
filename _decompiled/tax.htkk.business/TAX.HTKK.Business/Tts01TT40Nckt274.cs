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
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tts01TT40Nckt274
{
	private Dictionary<string, string[]> dicDMHopDong = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMHopDongTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMMucDich = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMMucDichTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMTTHopDong = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMTTHopDongTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicFormulaTK = new Dictionary<string, string>();

	private Dictionary<string, double[]> dicTongCT20a = new Dictionary<string, double[]>();

	private Dictionary<string, double[]> dicTongCT20 = new Dictionary<string, double[]>();

	private Dictionary<string, string> dicHDXoa = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMPhuongXaTheoCQT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMPhuongXaTheoCQTTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMQuocTich = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicQuocTichTmp = new Dictionary<string, string>();

	private HTKKPeriodInfo _PeriodInfo = null;

	private HTKKHeaderInfo _HeaderInfo = null;

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private HTKKCMCTTinChungInfo _TTinChung = null;

	private long _iRowClick = -1L;

	private string MA_HOP_DONG = "01";

	private string MA_KY_THANH_TOAN = "02";

	private string MA_KY_THANH_TOAN_CUOI_CUNG = "03";

	private string MA_TRA_TIEN_CHO_CA_HD = "04";

	private string MA_HOP_DONG_KHAI_LAN_DAU = "01";

	private bool _nhapTuExcel = false;

	private int[] khaiTheoPLuatDanSu = GeneralFunctions.ArrayCellId("Q_10");

	private int[] khaiTheoPLuatDanThue = GeneralFunctions.ArrayCellId("Q_12");

	private int[] cellCoPhuLuc_01_3 = GeneralFunctions.ArrayCellId("AG_42");

	private int[] cellCoPhuLuc_01_4 = GeneralFunctions.ArrayCellId("AH_42");

	private int[] cellCoPhuLuc_01_5 = GeneralFunctions.ArrayCellId("AI_42");

	private string _valueMstOld = "";

	private XMLMessage _xmlMessage = new XMLMessage("../Project/Message.xml");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_PeriodInfo = PeriodInfo;
			_HeaderInfo = HeaderInfo;
			_ToKhaiInfo = ToKhaiInfo;
			_TTinChung = TTinChung;
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			int num3 = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AD");
			int[] array = GeneralFunctions.ArrayCellId("F_42");
			int[] array2 = GeneralFunctions.ArrayCellId("F_44");
			int[] array3 = GeneralFunctions.ArrayCellId("H_54");
			int[] array4 = GeneralFunctions.ArrayCellId("E_64");
			int[] array5 = GeneralFunctions.ArrayCellId("J_64");
			int[] array6 = GeneralFunctions.ArrayCellId("Q_64");
			int[] array7 = GeneralFunctions.ArrayCellId("I_186");
			int[] array8 = GeneralFunctions.ArrayCellId("H_194");
			int[] array9 = GeneralFunctions.ArrayCellId("S_194");
			int[] array10 = GeneralFunctions.ArrayCellId("F_40");
			string text = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
			int num6 = 37;
			if (string.IsNullOrEmpty(sheetView.Cells[num6 - 1, num2 - 1].Text) && !string.IsNullOrEmpty(PeriodInfo.ToChucKhaiThay))
			{
				sheetView.Cells[num6 - 1, num2 - 1].Text = PeriodInfo.ToChucKhaiThay;
			}
			if (string.IsNullOrEmpty(sheetView.Cells[num6 - 1, num4 - 1].Text.Trim()) && !string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy))
			{
				sheetView.Cells[num6 - 1, num4 - 1].Text = PeriodInfo.tenCQTQuanLy;
				sheetView.Cells[num6 - 1, num5 - 1].Text = PeriodInfo.maCQTQuanLy;
			}
			bool flag = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value);
			string text2 = sheetView.Cells[array[1], array[0]].Text;
			string text3 = sheetView.Cells[array7[1], array7[0]].Text;
			if (PeriodInfo.KieuKy == "D_D")
			{
				sheetView.Cells[array10[1], array10[0]].Locked = false;
			}
			else
			{
				sheetView.Cells[array10[1], array10[0]].Locked = true;
			}
			if (PeriodInfo.KieuKy == "Q" || PeriodInfo.KieuKy == "M")
			{
				sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value = false;
				sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value = true;
				sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Locked = true;
				sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Locked = true;
			}
			CheckChonLoaiNopThue(winFpSpread, ToKhaiInfo, PeriodInfo);
			if (flag)
			{
				LockUnlockTheoLoaiNopThue(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, string.IsNullOrEmpty(text2) ? true : false);
			}
			else
			{
				LockUnlockTheoLoaiNopThue(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, string.IsNullOrEmpty(text3) ? true : false);
			}
			sheetView.Cells[cellCoPhuLuc_01_3[1], cellCoPhuLuc_01_3[0]].Value = ((_ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
			sheetView.Cells[cellCoPhuLuc_01_4[1], cellCoPhuLuc_01_4[0]].Value = ((_ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("B");
				int colTemDongTuSinh = HTKKFile.ColumnLettersToIndex("D");
				int colCt06aMa = HTKKFile.ColumnLettersToIndex("E");
				int num2 = HTKKFile.ColumnLettersToIndex("F");
				int item = HTKKFile.ColumnLettersToIndex("G");
				int item2 = HTKKFile.ColumnLettersToIndex("J");
				int item3 = HTKKFile.ColumnLettersToIndex("K");
				int item4 = HTKKFile.ColumnLettersToIndex("L");
				int item5 = HTKKFile.ColumnLettersToIndex("M");
				int item6 = HTKKFile.ColumnLettersToIndex("N");
				int item7 = HTKKFile.ColumnLettersToIndex("O");
				int num3 = HTKKFile.ColumnLettersToIndex("P");
				int item8 = HTKKFile.ColumnLettersToIndex("S");
				int item9 = HTKKFile.ColumnLettersToIndex("T");
				int num4 = HTKKFile.ColumnLettersToIndex("U");
				int item10 = HTKKFile.ColumnLettersToIndex("V");
				int item11 = HTKKFile.ColumnLettersToIndex("W");
				int item12 = HTKKFile.ColumnLettersToIndex("AA");
				int num5 = HTKKFile.ColumnLettersToIndex("AB");
				int num6 = HTKKFile.ColumnLettersToIndex("AC");
				int item13 = HTKKFile.ColumnLettersToIndex("AD");
				int item14 = HTKKFile.ColumnLettersToIndex("AE");
				int item15 = HTKKFile.ColumnLettersToIndex("AI");
				int item16 = HTKKFile.ColumnLettersToIndex("AJ");
				int num7 = HTKKFile.ColumnLettersToIndex("AK");
				int colTongDoanhThuPhatSinhTrongNamTungCN = HTKKFile.ColumnLettersToIndex("AL");
				int num8 = HTKKFile.ColumnLettersToIndex("AM");
				int item17 = HTKKFile.ColumnLettersToIndex("AQ");
				int num9 = HTKKFile.ColumnLettersToIndex("AR");
				int num10 = HTKKFile.ColumnLettersToIndex("AS");
				int item18 = HTKKFile.ColumnLettersToIndex("AU");
				int num11 = HTKKFile.ColumnLettersToIndex("AT");
				int item19 = HTKKFile.ColumnLettersToIndex("AV");
				int colTemVatNam = HTKKFile.ColumnLettersToIndex("AF");
				int num12 = HTKKFile.ColumnLettersToIndex("AH");
				int num13 = HTKKFile.ColumnLettersToIndex("AG");
				int num14 = HTKKFile.ColumnLettersToIndex("BB");
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				}
				List<int> list = new List<int>();
				list.Add(item);
				list.Add(item2);
				list.Add(item3);
				list.Add(item4);
				list.Add(item5);
				list.Add(item6);
				list.Add(item7);
				list.Add(item8);
				list.Add(item9);
				list.Add(item10);
				list.Add(item15);
				List<int> list2 = new List<int>();
				list2.Add(item11);
				list2.Add(item18);
				list2.Add(item19);
				List<int> list3 = new List<int>();
				list3.Add(item16);
				List<int> list4 = new List<int>();
				list4.Add(item15);
				list4.Add(num8);
				list4.Add(item17);
				List<int> list5 = new List<int>();
				list5.Add(item13);
				List<int> list6 = new List<int>();
				list6.Add(item12);
				List<int> list7 = new List<int>();
				list7.Add(num5);
				list7.Add(num6);
				list7.Add(item14);
				List<int> list8 = new List<int>();
				list8.Add(num2);
				List<int> list9 = new List<int>();
				list9.Add(item16);
				Dictionary<string, List<int>> cacHopDong = LayCacHangCungHopDong(sView, hTKKCMCSheetInfo, num);
				LockUnLockDongHopDongHoacKy(sView, hTKKCMCSheetInfo, num, colCt06aMa, colTemDongTuSinh, colTemVatNam, list, list2, list6, list7, list3, list8, list4, num5, num6, list9, colTongDoanhThuPhatSinhTrongNamTungCN, num8, list5, cacHopDong);
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sView, hTKKCMCSheetInfo, num, colCt06aMa, num2, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("B");
				int colTemDongTuSinh = HTKKFile.ColumnLettersToIndex("D");
				int colCt06aMa = HTKKFile.ColumnLettersToIndex("E");
				int num2 = HTKKFile.ColumnLettersToIndex("F");
				int item = HTKKFile.ColumnLettersToIndex("G");
				int item2 = HTKKFile.ColumnLettersToIndex("J");
				int item3 = HTKKFile.ColumnLettersToIndex("L");
				int item4 = HTKKFile.ColumnLettersToIndex("O");
				int item5 = HTKKFile.ColumnLettersToIndex("P");
				int item6 = HTKKFile.ColumnLettersToIndex("Q");
				int item7 = HTKKFile.ColumnLettersToIndex("R");
				int item8 = HTKKFile.ColumnLettersToIndex("S");
				int item9 = HTKKFile.ColumnLettersToIndex("U");
				int num3 = HTKKFile.ColumnLettersToIndex("T");
				int item10 = HTKKFile.ColumnLettersToIndex("V");
				int item11 = HTKKFile.ColumnLettersToIndex("W");
				int item12 = HTKKFile.ColumnLettersToIndex("Y");
				int num4 = HTKKFile.ColumnLettersToIndex("X");
				int item13 = HTKKFile.ColumnLettersToIndex("Z");
				int item14 = HTKKFile.ColumnLettersToIndex("AA");
				int num5 = HTKKFile.ColumnLettersToIndex("AB");
				int num6 = HTKKFile.ColumnLettersToIndex("AC");
				int item15 = HTKKFile.ColumnLettersToIndex("AD");
				int item16 = HTKKFile.ColumnLettersToIndex("AE");
				int item17 = HTKKFile.ColumnLettersToIndex("AL");
				int item18 = HTKKFile.ColumnLettersToIndex("AM");
				int num7 = HTKKFile.ColumnLettersToIndex("AN");
				int colTongDoanhThuPhatSinhTrongNamTungCN = HTKKFile.ColumnLettersToIndex("AO");
				int num8 = HTKKFile.ColumnLettersToIndex("AP");
				int item19 = HTKKFile.ColumnLettersToIndex("AQ");
				int num9 = HTKKFile.ColumnLettersToIndex("AR");
				int num10 = HTKKFile.ColumnLettersToIndex("AS");
				int item20 = HTKKFile.ColumnLettersToIndex("AU");
				int num11 = HTKKFile.ColumnLettersToIndex("AT");
				int item21 = HTKKFile.ColumnLettersToIndex("AV");
				int colTemVatNam = HTKKFile.ColumnLettersToIndex("AI");
				int num12 = HTKKFile.ColumnLettersToIndex("AK");
				int num13 = HTKKFile.ColumnLettersToIndex("AJ");
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				}
				List<int> list = new List<int>();
				list.Add(item);
				list.Add(item2);
				list.Add(item3);
				list.Add(item4);
				list.Add(item5);
				list.Add(item6);
				list.Add(item7);
				list.Add(item8);
				list.Add(item9);
				list.Add(item10);
				list.Add(item11);
				list.Add(item12);
				list.Add(item17);
				List<int> list2 = new List<int>();
				list2.Add(item13);
				list2.Add(item20);
				list2.Add(item21);
				List<int> list3 = new List<int>();
				list3.Add(item18);
				List<int> list4 = new List<int>();
				list4.Add(item17);
				list4.Add(num8);
				list4.Add(item19);
				List<int> list5 = new List<int>();
				list5.Add(item15);
				List<int> list6 = new List<int>();
				list6.Add(num5);
				list6.Add(num6);
				list6.Add(item16);
				List<int> list7 = new List<int>();
				list7.Add(num2);
				List<int> list8 = new List<int>();
				list8.Add(item14);
				List<int> list9 = new List<int>();
				list9.Add(item18);
				Dictionary<string, List<int>> cacHopDong = LayCacHangCungHopDong(sView, hTKKCMCSheetInfo, num);
				LockUnLockDongHopDongHoacKy(sView, hTKKCMCSheetInfo, num, colCt06aMa, colTemDongTuSinh, colTemVatNam, list, list2, list8, list6, list3, list7, list4, num5, num6, list9, colTongDoanhThuPhatSinhTrongNamTungCN, num8, list5, cacHopDong);
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sView, hTKKCMCSheetInfo, num, colCt06aMa, num2, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_3[1], cellCoPhuLuc_01_3[0]].Value = ((_ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_4[1], cellCoPhuLuc_01_4[0]].Value = ((_ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
			int[] array = GeneralFunctions.ArrayCellId("B_2");
			string text = "";
			if (PeriodInfo.KieuKy == "D_D")
			{
				DateTime dateTime = new DateTime(2023, 7, 1);
				DateTime dateTime2 = new DateTime(2023, 12, 31);
				DateTime dateTime3 = new DateTime(2024, 6, 30);
				DateTime dateTime4 = new DateTime(2024, 1, 1);
				DateTime dateTime5 = DateTime.Now;
				DateTime dateTime6 = DateTime.Now;
				if (!string.IsNullOrEmpty(PeriodInfo.TuNgay) && PeriodInfo.TuNgay.Length > 8)
				{
					dateTime5 = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
				}
				if (!string.IsNullOrEmpty(PeriodInfo.DenNgay) && PeriodInfo.DenNgay.Length > 8)
				{
					dateTime6 = DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
				}
				if (((DateTime.Compare(dateTime5, dateTime3) <= 0 && DateTime.Compare(dateTime, dateTime5) <= 0) || (DateTime.Compare(dateTime6, dateTime3) <= 0 && DateTime.Compare(dateTime, dateTime6) <= 0) || (DateTime.Compare(dateTime5, dateTime) <= 0 && DateTime.Compare(dateTime3, dateTime6) <= 0)) && PeriodInfo.TuNgay.Length > 8 && PeriodInfo.DenNgay.Length > 8)
				{
					text = ((DateTime.Compare(dateTime6, dateTime3) <= 0 && DateTime.Compare(dateTime4, dateTime6) <= 0) ? ((DateTime.Compare(dateTime5, dateTime3) > 0 || DateTime.Compare(dateTime4, dateTime5) > 0) ? "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15  VÀ NGHỊ QUYẾT SỐ 110/2023/QH15" : "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15") : ((DateTime.Compare(dateTime5, dateTime4) < 0 && DateTime.Compare(dateTime3, dateTime6) < 0) ? "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15  VÀ NGHỊ QUYẾT SỐ 110/2023/QH15" : ((DateTime.Compare(dateTime5, dateTime4) < 0) ? "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15" : "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15")));
				}
			}
			if (PeriodInfo.KieuKy == "M")
			{
				if (PeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) <= 12)
				{
					text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
				}
				else if (PeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) <= 6)
				{
					text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				if (PeriodInfo.Nam.Equals("2023") && (PeriodInfo.Quy == "3" || PeriodInfo.Quy == "4"))
				{
					text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
				}
				else if (PeriodInfo.Nam.Equals("2024") && (PeriodInfo.Quy == "1" || PeriodInfo.Quy == "2"))
				{
					text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
				}
			}
			winFpSpread.Sheets[4].Cells[array[1], array[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_5[1], cellCoPhuLuc_01_5[0]].Value = ((_ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
			int[] array = GeneralFunctions.ArrayCellId("B_2");
			string text = "";
			if (PeriodInfo.KieuKy == "D_D")
			{
				DateTime dateTime = new DateTime(2024, 7, 1);
				DateTime dateTime2 = new DateTime(2024, 12, 31);
				DateTime dateTime3 = new DateTime(2025, 6, 30);
				DateTime dateTime4 = new DateTime(2025, 1, 1);
				DateTime dateTime5 = DateTime.Now;
				DateTime dateTime6 = DateTime.Now;
				if (!string.IsNullOrEmpty(PeriodInfo.TuNgay) && PeriodInfo.TuNgay.Length > 8)
				{
					dateTime5 = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
				}
				if (!string.IsNullOrEmpty(PeriodInfo.DenNgay) && PeriodInfo.DenNgay.Length > 8)
				{
					dateTime6 = DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
				}
				if (((DateTime.Compare(dateTime5, dateTime3) <= 0 && DateTime.Compare(dateTime, dateTime5) <= 0) || (DateTime.Compare(dateTime6, dateTime3) <= 0 && DateTime.Compare(dateTime, dateTime6) <= 0) || (DateTime.Compare(dateTime5, dateTime) <= 0 && DateTime.Compare(dateTime3, dateTime6) <= 0)) && PeriodInfo.TuNgay.Length > 8 && PeriodInfo.DenNgay.Length > 8)
				{
					text = ((DateTime.Compare(dateTime6, dateTime3) <= 0 && DateTime.Compare(dateTime4, dateTime6) <= 0) ? ((DateTime.Compare(dateTime5, dateTime3) > 0 || DateTime.Compare(dateTime4, dateTime5) > 0) ? "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15  VÀ NGHỊ QUYẾT SỐ 174/2024/QH15" : "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15") : ((DateTime.Compare(dateTime5, dateTime4) < 0 && DateTime.Compare(dateTime3, dateTime6) < 0) ? "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15  VÀ NGHỊ QUYẾT SỐ 174/2024/QH15" : ((DateTime.Compare(dateTime5, dateTime4) < 0) ? "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15" : "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15")));
				}
			}
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "Q")
			{
				if (PeriodInfo.Nam.Equals("2024"))
				{
					text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
				}
				else if (PeriodInfo.Nam.Equals("2025"))
				{
					text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
				}
			}
			winFpSpread.Sheets[5].Cells[array[1], array[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_0(FpSpread winFpSpread, int actionType)
	{
		try
		{
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
			CheckChonLoaiNopThue(winFpSpread, _ToKhaiInfo, _PeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType)
	{
		try
		{
			_nhapTuExcel = false;
			CheckChonLoaiNopThue(winFpSpread, _ToKhaiInfo, _PeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_3(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_3[1], cellCoPhuLuc_01_3[0]].Value = ((_ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
			TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
			TinhChiTieu26_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_4(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_4[1], cellCoPhuLuc_01_4[0]].Value = ((_ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
			TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
			TinhChiTieu26_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_5(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_5[1], cellCoPhuLuc_01_5[0]].Value = ((_ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
			TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
			TinhChiTieu26_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellEditchange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("U_211");
			if (iRow == array[1] + 1)
			{
				TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
				TinhChiTieu26_TK(winFpSpread);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int item = HTKKFile.ColumnLettersToIndex("G");
			int item2 = HTKKFile.ColumnLettersToIndex("J");
			int item3 = HTKKFile.ColumnLettersToIndex("K");
			int item4 = HTKKFile.ColumnLettersToIndex("L");
			int item5 = HTKKFile.ColumnLettersToIndex("M");
			int item6 = HTKKFile.ColumnLettersToIndex("N");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = HTKKFile.ColumnLettersToIndex("P");
			int item7 = HTKKFile.ColumnLettersToIndex("S");
			int item8 = HTKKFile.ColumnLettersToIndex("T");
			int num7 = HTKKFile.ColumnLettersToIndex("U");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int item9 = HTKKFile.ColumnLettersToIndex("W");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AB");
			int num11 = HTKKFile.ColumnLettersToIndex("AC");
			int num12 = HTKKFile.ColumnLettersToIndex("AD");
			int num13 = HTKKFile.ColumnLettersToIndex("AE");
			int num14 = HTKKFile.ColumnLettersToIndex("AI");
			int num15 = HTKKFile.ColumnLettersToIndex("AJ");
			int colDoanhThuBinhQuanThang = HTKKFile.ColumnLettersToIndex("AK");
			int num16 = HTKKFile.ColumnLettersToIndex("AL");
			int num17 = HTKKFile.ColumnLettersToIndex("AM");
			int num18 = HTKKFile.ColumnLettersToIndex("AQ");
			int num19 = HTKKFile.ColumnLettersToIndex("AR");
			int num20 = HTKKFile.ColumnLettersToIndex("AS");
			int num21 = HTKKFile.ColumnLettersToIndex("AU");
			int num22 = HTKKFile.ColumnLettersToIndex("AT");
			int num23 = HTKKFile.ColumnLettersToIndex("AV");
			int num24 = HTKKFile.ColumnLettersToIndex("AF");
			int num25 = HTKKFile.ColumnLettersToIndex("AH");
			int colTemTronThang = HTKKFile.ColumnLettersToIndex("AG");
			int num26 = HTKKFile.ColumnLettersToIndex("BB");
			string pathCt23ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct23";
			string pathCt24ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct24";
			string pathMaHDToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayPhaiLonHonTuNgay = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorTuNgayKyLonHonTuNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayKyNhoHonDenNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			list.Add(item5);
			list.Add(item6);
			list.Add(num5);
			list.Add(item7);
			list.Add(item8);
			list.Add(num8);
			list.Add(num14);
			List<int> list2 = new List<int>();
			list2.Add(item9);
			list2.Add(num21);
			list2.Add(num23);
			List<int> list3 = new List<int>();
			list3.Add(num15);
			List<int> list4 = new List<int>();
			list4.Add(num14);
			list4.Add(num17);
			list4.Add(num18);
			List<int> list5 = new List<int>();
			list5.Add(num12);
			List<int> list6 = new List<int>();
			list6.Add(num9);
			List<int> list7 = new List<int>();
			list7.Add(num10);
			list7.Add(num11);
			list7.Add(num13);
			List<int> list8 = new List<int>();
			list8.Add(num4);
			List<int> list9 = new List<int>();
			list9.Add(num15);
			HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
			HTKKCMCCellInfo rowCellCt23ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt23ToKhai));
			if (iCol == num5)
			{
				string[] array = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array2 = dicDMMucDich[sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim()];
				if (array2.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = array2[0];
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = array2[1];
				}
			}
			if (iCol == num21)
			{
				string[] array2 = dicDMTTHopDong[sheetView.Cells[(int)iRow - 1, num21 - 1].Text.Trim()];
				if (array2.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num22 - 1].Text = array2[0];
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = array2[1];
				}
			}
			if (iCol == num8)
			{
				string tenDonVi = sheetView.Cells[(int)iRow - 1, num8 - 1].Text.Trim();
				KeyValuePair<string, string> keyValuePair = dicDMPhuongXaTheoCQTTmp.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == tenDonVi);
				sheetView.Cells[(int)iRow - 1, num7 - 1].Text = keyValuePair.Key;
				sheetView.Cells[(int)iRow - 1, num8 - 1].Text = keyValuePair.Value;
			}
			string text = "";
			if (iCol == num4)
			{
				string valueTenHD = sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(valueTenHD))
				{
					text = dicDMHopDongTmp.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == valueTenHD).Key;
					if (!string.IsNullOrEmpty(text))
					{
						sheetView.Cells[(int)iRow - 1, num3 - 1].Text = text;
						CheckNeuTichChonHDKyThanhToan1Lan(sheetView, hTKKCMCSheetInfo, PeriodInfo, text, (int)iRow, num10, num11, num12, num24);
						TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
						if (!CheckChonHopDongTrungKy(sheetView, hTKKCMCSheetInfo, num, num3, num4, errorMessager, tuDongXoaThongBao: true) || !CheckChonHopDongDauTien(sheetView, hTKKCMCSheetInfo, num, num3, num4, errorMessager2, tuDongXoaThongBao: true))
						{
							return;
						}
					}
				}
			}
			if (string.IsNullOrEmpty(text))
			{
				text = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
			}
			string text2 = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
			string[] array3 = text2.Split(new char[1] { '.' });
			string key = array3[0];
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
			if (!dictionary.ContainsKey(key))
			{
				return;
			}
			int num27 = dictionary[key][0];
			string text3 = sheetView.Cells[(int)iRow - 1, num24 - 1].Text.Trim();
			if (iCol == num4)
			{
				if (PeriodInfo.KieuKy == "D_D" && dictionary.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num27 - 1, num4 - 1, msg);
					return;
				}
				HTKKFile.RemoveValidate(sheetView, num27 - 1, num4 - 1);
			}
			if ((iCol == num10 || iCol == num11) && !CheckDenNgayNhoHonTuNgay(sheetView, num27, (int)iRow, num10, num11, errorDenNgayPhaiLonHonTuNgay, errorTuNgayKyLonHonTuNgayHD, errorDenNgayKyNhoHonDenNgayHD, tuDongXoaThongBao: true, num24))
			{
				return;
			}
			if (text != MA_HOP_DONG && (iCol == num10 || iCol == num11))
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num24, num10, num11, hTKKCMCSheetInfo);
			}
			if (num27 == (int)iRow && (iCol == num21 || iCol == num4 || iCol == num10 || iCol == num11))
			{
				TinhMaHopDong(sheetView, TTinChung, num10, num3, num, num22, num23, (int)iRow, ToKhaiInfo.SheetInfos[1].Tables[0]);
			}
			if (PeriodInfo.KieuKy == "D_D" && (iCol == num4 || iCol == num21 || iCol == num23))
			{
				string text4 = sheetView.Cells[num27 - 1, num23 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text4))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaHDToKhai));
					winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text4;
				}
			}
			int num28 = 0;
			if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13)
			{
				num28 = TinhKyThanhToanOnDinh(sheetView, hTKKCMCSheetInfo, PeriodInfo, num9, (int)iRow, num3, num4, num10, num11, num12, num13, num, num24, list, list2, num2, num27);
				if (num28 == 0)
				{
					num28 = TuDongTachNeuVatNam(sheetView, hTKKCMCSheetInfo, (int)iRow, num3, num4, num10, num11, num13, num, num24, list, list2, num2, num9, num27);
				}
				if (num28 > 0)
				{
					TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
					dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
				}
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol < num15)))
			{
				num27 = dictionary[key][0];
				if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13)
				{
					sheetView.Cells[num27 - 1, num14 - 1].Value = 0;
				}
				sheetView.Cells[num27 - 1, num15 - 1].Value = 0;
				for (int num29 = 0; num29 < dictionary[key].Count; num29++)
				{
					int num30 = dictionary[key][num29];
					if (num29 == 0)
					{
						sheetView.Cells[num30 - 1, num25 - 1].Value = 0;
						continue;
					}
					string text5 = sheetView.Cells[num30 - 1, num2 - 1].Text.Trim();
					AutoShowYear(sheetView, num30, num3, num10, num11, num12);
					bool tuDongSoThang = ((iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13) ? true : false);
					TinhTongGiaTriHD(sheetView, num27, num30, num3, num10, num11, num13, num24, num15, num25, colTemTronThang, num14, tuDongSoThang);
				}
			}
			if (text != MA_HOP_DONG && iCol == num15)
			{
				SumTinhTongGiaTriHD(sheetView, dictionary[key], num15, num10, num11, num24);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol <= num15)))
			{
				DoanhThuBinhQuanThangCuaHD(sheetView, num15, colDoanhThuBinhQuanThang, num3, num14, dictionary[key]);
			}
			if (iCol == num4 || iCol == num15 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num13 || iCol == num14)
			{
				TongDoanhThuPhatSinhTrongNam(sheetView, dictionary, num12, num15, num16, num, -1);
			}
			if (iCol == num4 || iCol == num15 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num13 || iCol == num14 || iCol == num16)
			{
				CaNhanPhaiNopThue(sheetView, dictionary, PeriodInfo, num3, colDoanhThuBinhQuanThang, num16, num, num17, num10, num11, num24);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol < num18)))
			{
				DoanhThuTinhTheTrongKy(sheetView, dictionary, num3, num, num17, num18, num15, num10, num11);
			}
			if (text != MA_HOP_DONG && iCol == num18)
			{
				SumDoanhThuTinhTheTrongKy(sheetView, dictionary[key], num18, num10, num11);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol <= num15)))
			{
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt23ToKhai, dictionary, num24, num15, num10, num11, num3);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol <= num18)))
			{
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt24ToKhai, dictionary, num24, num18, num10, num11, num3);
			}
			if (iCol == num4)
			{
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sheetView, hTKKCMCSheetInfo, num, num3, num4, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
			if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num13 || iCol == num14 || iCol == num15 || iCol == num16)
			{
				LockUnLockDongHopDongHoacKy(sheetView, hTKKCMCSheetInfo, num, num3, num2, num24, list, list2, list6, list7, list3, list8, list4, num10, num11, list9, num16, num17, list5, dictionary);
			}
			if (iCol == num4 && !CheckChonLoaiNopThue(winFpSpread, ToKhaiInfo, PeriodInfo))
			{
				return;
			}
			if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num13 || iCol == num21 || iCol == num23)
			{
				string text6 = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text6))
				{
					foreach (KeyValuePair<string, List<int>> item10 in dictionary)
					{
						string valueCopy = sheetView.Cells[item10.Value[0] - 1, num22 - 1].Text.Trim();
						string valueCopy2 = sheetView.Cells[item10.Value[0] - 1, num21 - 1].Text.Trim();
						string valueCopy3 = sheetView.Cells[item10.Value[0] - 1, num23 - 1].Text.Trim();
						if (text6.IndexOf(".") >= 0)
						{
							text6 = text6.Substring(0, text6.IndexOf("."));
						}
						Set1RowHDToAllRowKy(sheetView, item10.Value, num, num22, valueCopy, text6);
						Set1RowHDToAllRowKy(sheetView, item10.Value, num, num21, valueCopy2, text6);
						Set1RowHDToAllRowKy(sheetView, item10.Value, num, num23, valueCopy3, text6);
					}
				}
			}
			if (num28 > 0)
			{
				CheckChiTieu23BoTichDongHD(sheetView, dictionary[key], num17);
			}
			GeneralFunctions.SetRowHeight(sheetView, (int)iRow);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int item = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("J");
			int item2 = HTKKFile.ColumnLettersToIndex("L");
			int item3 = HTKKFile.ColumnLettersToIndex("O");
			int item4 = HTKKFile.ColumnLettersToIndex("P");
			int item5 = HTKKFile.ColumnLettersToIndex("Q");
			int item6 = HTKKFile.ColumnLettersToIndex("R");
			int item7 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("U");
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int item8 = HTKKFile.ColumnLettersToIndex("V");
			int item9 = HTKKFile.ColumnLettersToIndex("W");
			int num8 = HTKKFile.ColumnLettersToIndex("Y");
			int num9 = HTKKFile.ColumnLettersToIndex("X");
			int item10 = HTKKFile.ColumnLettersToIndex("Z");
			int num10 = HTKKFile.ColumnLettersToIndex("AA");
			int num11 = HTKKFile.ColumnLettersToIndex("AB");
			int num12 = HTKKFile.ColumnLettersToIndex("AC");
			int num13 = HTKKFile.ColumnLettersToIndex("AD");
			int num14 = HTKKFile.ColumnLettersToIndex("AE");
			int num15 = HTKKFile.ColumnLettersToIndex("AL");
			int num16 = HTKKFile.ColumnLettersToIndex("AM");
			int num17 = HTKKFile.ColumnLettersToIndex("AN");
			int num18 = HTKKFile.ColumnLettersToIndex("AO");
			int num19 = HTKKFile.ColumnLettersToIndex("AP");
			int num20 = HTKKFile.ColumnLettersToIndex("AQ");
			int num21 = HTKKFile.ColumnLettersToIndex("AR");
			int num22 = HTKKFile.ColumnLettersToIndex("AS");
			int num23 = HTKKFile.ColumnLettersToIndex("AU");
			int num24 = HTKKFile.ColumnLettersToIndex("AT");
			int num25 = HTKKFile.ColumnLettersToIndex("AV");
			int num26 = HTKKFile.ColumnLettersToIndex("AI");
			int num27 = HTKKFile.ColumnLettersToIndex("AK");
			int colTemTronThang = HTKKFile.ColumnLettersToIndex("AJ");
			int num28 = HTKKFile.ColumnLettersToIndex("BC");
			string pathMaHDToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
			string pathCt23ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct23";
			string pathCt24ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct24";
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayPhaiLonHonTuNgay = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorTuNgayKyLonHonTuNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayKyNhoHonDenNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(num5);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			list.Add(item5);
			list.Add(item6);
			list.Add(item7);
			list.Add(num6);
			list.Add(item8);
			list.Add(item9);
			list.Add(num8);
			list.Add(num15);
			List<int> list2 = new List<int>();
			list2.Add(item10);
			list2.Add(num23);
			list2.Add(num25);
			List<int> list3 = new List<int>();
			list3.Add(num16);
			List<int> list4 = new List<int>();
			list4.Add(num15);
			list4.Add(num19);
			list4.Add(num20);
			List<int> list5 = new List<int>();
			list5.Add(num13);
			List<int> list6 = new List<int>();
			list6.Add(num11);
			list6.Add(num12);
			list6.Add(num14);
			List<int> list7 = new List<int>();
			list7.Add(num4);
			List<int> list8 = new List<int>();
			list8.Add(num10);
			List<int> list9 = new List<int>();
			list9.Add(num16);
			HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
			HTKKCMCCellInfo rowCellCt23ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt23ToKhai));
			Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
			HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables[0];
			if (iCol == num6)
			{
				string[] array = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array2 = dicDMMucDich[sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim()];
				if (array2.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Text = array2[0];
					sheetView.Cells[(int)iRow - 1, num6 - 1].Text = array2[1];
				}
			}
			if (iCol == num23)
			{
				string text = sheetView.Cells[(int)iRow - 1, num23 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					string[] array2 = dicDMTTHopDong[text];
					if (array2.Length == 2)
					{
						sheetView.Cells[(int)iRow - 1, num24 - 1].Text = array2[0];
						sheetView.Cells[(int)iRow - 1, num23 - 1].Text = array2[1];
					}
				}
			}
			if (iCol == num8)
			{
				string tenDonVi = sheetView.Cells[(int)iRow - 1, num8 - 1].Text.Trim();
				KeyValuePair<string, string> keyValuePair = dicDMPhuongXaTheoCQTTmp.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == tenDonVi);
				sheetView.Cells[(int)iRow - 1, num9 - 1].Text = keyValuePair.Key;
				sheetView.Cells[(int)iRow - 1, num8 - 1].Text = keyValuePair.Value;
			}
			string text2 = "";
			if (iCol == num4)
			{
				string valueTenHD = sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(valueTenHD))
				{
					text2 = dicDMHopDongTmp.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == valueTenHD).Key;
					if (!string.IsNullOrEmpty(text2))
					{
						sheetView.Cells[(int)iRow - 1, num3 - 1].Text = text2;
						CheckNeuTichChonHDKyThanhToan1Lan(sheetView, hTKKCMCSheetInfo, PeriodInfo, text2, (int)iRow, num11, num12, num13, num26);
						TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
						CheckChonHopDongTrungKy(sheetView, hTKKCMCSheetInfo, num, num3, num4, errorMessager, tuDongXoaThongBao: true);
						CheckChonHopDongDauTien(sheetView, hTKKCMCSheetInfo, num, num3, num4, errorMessager2, tuDongXoaThongBao: true);
					}
				}
			}
			if (string.IsNullOrEmpty(text2))
			{
				text2 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
			}
			string text3 = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
			string[] array3 = text3.Split(new char[1] { '.' });
			string key = array3[0];
			Dictionary<string, List<int>> dictionary2 = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
			if (!dictionary2.ContainsKey(key))
			{
				return;
			}
			int num29 = dictionary2[key][0];
			dictionary = TimHopDongCungMSTCungNam(sheetView, table, num29, num, num5, dictionary2);
			if (iCol == num5 && _valueMstOld != sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim())
			{
				Dictionary<string, List<int>> dictionary3 = CheckCacHopDongAnhHuongSauKhiDieuChinhMST(sheetView, table, num29, num, num5, _valueMstOld, dictionary2);
				if (dictionary3.Count > 0)
				{
					foreach (KeyValuePair<string, List<int>> item11 in dictionary3)
					{
						if (!dictionary.ContainsKey(item11.Key))
						{
							dictionary.Add(item11.Key, item11.Value);
						}
					}
				}
				_valueMstOld = sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim();
			}
			if ((iCol == num11 || iCol == num12) && !CheckDenNgayNhoHonTuNgay(sheetView, num29, (int)iRow, num11, num12, errorDenNgayPhaiLonHonTuNgay, errorTuNgayKyLonHonTuNgayHD, errorDenNgayKyNhoHonDenNgayHD, tuDongXoaThongBao: true, num26))
			{
				return;
			}
			if (text2 != MA_HOP_DONG && (iCol == num11 || iCol == num12))
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num26, num11, num12, hTKKCMCSheetInfo);
			}
			if (iCol == num4 || iCol == num11 || iCol == num12 || iCol == num16 || iCol == num18)
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num19, num26, num11, num12, num18);
			}
			string text4 = sheetView.Cells[(int)iRow - 1, num26 - 1].Text.Trim();
			if (iCol == num4)
			{
				if (PeriodInfo.KieuKy == "D_D" && dictionary2.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num29 - 1, num4 - 1, msg);
					return;
				}
				HTKKFile.RemoveValidate(sheetView, num29 - 1, num4 - 1);
			}
			if (num29 == (int)iRow && (iCol == num23 || iCol == num4 || iCol == num11))
			{
				TinhMaHopDong(sheetView, TTinChung, num11, num3, num, num24, num25, (int)iRow, ToKhaiInfo.SheetInfos[2].Tables[0]);
			}
			if (PeriodInfo.KieuKy == "D_D" && (iCol == num4 || iCol == num23 || iCol == num25))
			{
				string text5 = sheetView.Cells[num29 - 1, num25 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text5))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaHDToKhai));
					winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text5;
				}
			}
			int num30 = 0;
			if (iCol == num4 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14)
			{
				num30 = TinhKyThanhToanOnDinh(sheetView, hTKKCMCSheetInfo, PeriodInfo, num10, (int)iRow, num3, num4, num11, num12, num13, num14, num, num26, list, list2, num2, num29);
				if (num30 == 0)
				{
					num30 = TuDongTachNeuVatNam(sheetView, hTKKCMCSheetInfo, (int)iRow, num3, num4, num11, num12, num14, num, num26, list, list2, num2, num10, num29);
				}
				if (num30 > 0)
				{
					TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
					dictionary2 = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
					dictionary = TimHopDongCungMSTCungNam(sheetView, table, num29, num, num5, dictionary2);
				}
			}
			if (dictionary2.Count > 0 && (iCol == num4 || (iCol >= num10 && iCol < num16)))
			{
				num29 = dictionary2[key][0];
				if (iCol == num4 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14)
				{
					sheetView.Cells[num29 - 1, num15 - 1].Value = 0;
				}
				sheetView.Cells[num29 - 1, num16 - 1].Value = 0;
				for (int num31 = 0; num31 < dictionary2[key].Count; num31++)
				{
					int num32 = dictionary2[key][num31];
					if (num31 == 0)
					{
						sheetView.Cells[num32 - 1, num27 - 1].Value = 0;
						continue;
					}
					string text6 = sheetView.Cells[num32 - 1, num2 - 1].Text.Trim();
					AutoShowYear(sheetView, num32, num3, num11, num12, num13);
					bool tuDongSoThang = ((iCol == num4 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14) ? true : false);
					TinhTongGiaTriHD(sheetView, num29, num32, num3, num11, num12, num14, num26, num16, num27, colTemTronThang, num15, tuDongSoThang);
				}
			}
			if (text2 != MA_HOP_DONG && iCol == num16)
			{
				SumTinhTongGiaTriHD(sheetView, dictionary2[key], num16, num11, num12, num26);
			}
			if (dictionary2.Count > 0 && (iCol == num4 || (iCol >= num10 && iCol < num17)))
			{
				DoanhThuBinhQuanThangCuaHD(sheetView, num16, num17, num3, num15, dictionary2[key]);
			}
			if (iCol == num4 || iCol == num5 || iCol == num16 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num14 || iCol == num15)
			{
				TongDoanhThuPhatSinhTrongNam(sheetView, dictionary, num13, num16, num18, num, num5);
			}
			if (iCol == num4 || iCol == num5 || iCol == num16 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num14 || iCol == num15 || iCol == num18)
			{
				CaNhanPhaiNopThue(sheetView, dictionary, PeriodInfo, num3, num17, num18, num, num19, num11, num12, num26);
			}
			if (dictionary.Count > 0 && (iCol == num4 || iCol == num5 || (iCol >= num10 && iCol < num20)))
			{
				DoanhThuTinhTheTrongKy(sheetView, dictionary, num3, num, num19, num20, num16, num11, num12);
			}
			if (text2 != MA_HOP_DONG && iCol == num20)
			{
				SumDoanhThuTinhTheTrongKy(sheetView, dictionary2[key], num20, num11, num12);
			}
			if (dictionary2.Count > 0 && (iCol == num4 || iCol == num5 || (iCol >= num10 && iCol <= num16)))
			{
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt23ToKhai, dictionary2, num26, num16, num11, num12, num3);
			}
			if (dictionary2.Count > 0 && (iCol == num4 || iCol == num5 || (iCol >= num10 && iCol <= num20)))
			{
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt24ToKhai, dictionary2, num26, num20, num11, num12, num3);
			}
			if (iCol == num4)
			{
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sheetView, hTKKCMCSheetInfo, num, num3, num4, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
			if (iCol == num4 || iCol == num5 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num14 || iCol == num15 || iCol == num16 || iCol == num18)
			{
				LockUnLockDongHopDongHoacKy(sheetView, hTKKCMCSheetInfo, num, num3, num2, num26, list, list2, list8, list6, list3, list7, list4, num11, num12, list9, num18, num19, list5, dictionary);
			}
			if (iCol == num4 && !CheckChonLoaiNopThue(winFpSpread, ToKhaiInfo, PeriodInfo))
			{
				return;
			}
			if (iCol == num4 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num14 || iCol == num23 || iCol == num25)
			{
				string text7 = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text7))
				{
					foreach (KeyValuePair<string, List<int>> item12 in dictionary)
					{
						string valueCopy = sheetView.Cells[item12.Value[0] - 1, num24 - 1].Text.Trim();
						string valueCopy2 = sheetView.Cells[item12.Value[0] - 1, num23 - 1].Text.Trim();
						string valueCopy3 = sheetView.Cells[item12.Value[0] - 1, num25 - 1].Text.Trim();
						if (text7.IndexOf(".") >= 0)
						{
							text7 = text7.Substring(0, text7.IndexOf("."));
						}
						Set1RowHDToAllRowKy(sheetView, item12.Value, num, num24, valueCopy, text7);
						Set1RowHDToAllRowKy(sheetView, item12.Value, num, num23, valueCopy2, text7);
						Set1RowHDToAllRowKy(sheetView, item12.Value, num, num25, valueCopy3, text7);
					}
				}
			}
			if (iCol == num4 || iCol == num5 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num14 || iCol == num15 || iCol == num16 || iCol == num18)
			{
				foreach (KeyValuePair<string, List<int>> item13 in dictionary)
				{
					CheckChiTieu23BoTichDongHD(sheetView, item13.Value, num19);
				}
			}
			GeneralFunctions.SetRowHeight(sheetView, (int)iRow);
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
			int num = HTKKFile.ColumnLettersToIndex("D");
			if (iCol == num)
			{
				TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
				TinhChiTieu26_TK(winFpSpread);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("D");
			if (iCol == num)
			{
				TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
				TinhChiTieu26_TK(winFpSpread);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("D");
			if (iCol == num)
			{
				TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
				TinhChiTieu26_TK(winFpSpread);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, List<int>> TimHopDongCungMSTCungNam(SheetView sView, HTKKCMCTableInfo table, int iRow, int colStt, int colMst, Dictionary<string, List<int>> cacHopDong)
	{
		Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
		try
		{
			string text = sView.Cells[iRow - 1, colMst - 1].Text.Trim();
			string item = sView.Cells[iRow - 1, colStt - 1].Text.Trim();
			List<int> list = new List<int>();
			List<string> list2 = new List<string>();
			if (!string.IsNullOrEmpty(text))
			{
				list.Add(colMst);
				list2.Add(text);
			}
			else
			{
				list.Add(colStt);
				list2.Add(item);
			}
			List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sView, table, list, list2);
			foreach (int item2 in rowIndexInSViewByCompares)
			{
				string text2 = sView.Cells[item2 - 1, colStt - 1].Text.Trim();
				string[] array = text2.Split(new char[1] { '.' });
				if (!dictionary.ContainsKey(array[0]))
				{
					dictionary.Add(array[0], cacHopDong[array[0]]);
				}
			}
			if (dictionary.Count > 0)
			{
				return dictionary;
			}
			return cacHopDong;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return cacHopDong;
		}
	}

	public Dictionary<string, List<int>> CheckCacHopDongAnhHuongSauKhiDieuChinhMST(SheetView sView, HTKKCMCTableInfo table, int iRow, int colStt, int colMst, string valueColMstOld, Dictionary<string, List<int>> cacHopDong)
	{
		Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
		try
		{
			string text = sView.Cells[iRow - 1, colStt - 1].Text.Trim();
			if (string.IsNullOrEmpty(valueColMstOld))
			{
				return cacHopDong;
			}
			List<int> list = new List<int>();
			List<string> list2 = new List<string>();
			list.Add(colMst);
			list2.Add(valueColMstOld);
			List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sView, table, list, list2);
			foreach (int item in rowIndexInSViewByCompares)
			{
				string text2 = sView.Cells[item - 1, colStt - 1].Text.Trim();
				string[] array = text2.Split(new char[1] { '.' });
				if (!dictionary.ContainsKey(array[0]))
				{
					dictionary.Add(array[0], cacHopDong[array[0]]);
				}
			}
			if (dictionary.Count > 0)
			{
				return dictionary;
			}
			return new Dictionary<string, List<int>>();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return new Dictionary<string, List<int>>();
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

	public void CellClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			_iRowClick = iRow;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			_iRowClick = iRow;
			int num = HTKKFile.ColumnLettersToIndex("J");
			if (iCol == num)
			{
				_valueMstOld = winFpSpread.Sheets[2].Cells[(int)iRow - 1, num - 1].Text.Trim();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		//IL_0137: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0176: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = (DateTime.Now.Year - 1) * 12 + DateTime.Now.Month;
			int num2 = (DateTime.Now.Year - 1) * 4 + DateTime.Now.Month / 3;
			int num3 = int.Parse(PeriodInfo.Nam);
			if (PeriodInfo.KieuKy == "M")
			{
				int num4 = (num3 - 1) * 12 + int.Parse(PeriodInfo.Thang);
				if (num4 <= 0)
				{
					return false;
				}
				if (num4 > num)
				{
					frmMessage frmMessage = new frmMessage("0062", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				int num5 = (num3 - 1) * 4 + int.Parse(PeriodInfo.Quy);
				if (num5 > num2)
				{
					frmMessage frmMessage = new frmMessage("0045", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				int num6 = 8059;
				if (num5 < num6)
				{
					frmMessage frmMessage = new frmMessage("01XSBHDC0007", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
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
		HTKKFile.deleteLogErrorBigData();
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		SheetView sheetView = winFpSpread.Sheets[0];
		int[] array = GeneralFunctions.ArrayCellId("I_186");
		int[] array2 = GeneralFunctions.ArrayCellId("E_188");
		int[] array3 = GeneralFunctions.ArrayCellId("E_190");
		int num = HTKKFile.ColumnLettersToIndex("G");
		int num2 = HTKKFile.ColumnLettersToIndex("J");
		int num3 = HTKKFile.ColumnLettersToIndex("M");
		int num4 = HTKKFile.ColumnLettersToIndex("P");
		int num5 = HTKKFile.ColumnLettersToIndex("S");
		int num6 = HTKKFile.ColumnLettersToIndex("V");
		int num7 = HTKKFile.ColumnLettersToIndex("Y");
		int num8 = HTKKFile.ColumnLettersToIndex("E");
		int num9 = HTKKFile.ColumnLettersToIndex("AB");
		int num10 = HTKKFile.ColumnLettersToIndex("AE");
		string text = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/tenNNT";
		string text2 = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/mst";
		string text3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct12k";
		string pathCt12a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct12a_ngaySinh";
		string pathCt12c1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct12c_1_ngayCap";
		string pathCt12d1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct12d_1_ngayCap";
		string pathCt12dd1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct12dd_1_ngayCap";
		string pathCt12e1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct12e_1_ngayCap";
		string pathCt12f1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct12f_1_ngayCap";
		string pathCt12i1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12i/ct12i_ngayCap";
		string pathCt22_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT22/ngay";
		int[] array4 = GeneralFunctions.ArrayCellId("F_40");
		try
		{
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			bool flag = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value);
			if (!CheckChonLoaiNopThue(winFpSpread, _ToKhaiInfo, _PeriodInfo))
			{
				value = 2;
			}
			HTKKFile.RemoveValidate(sheetView, array4[1], array4[0]);
			if (PeriodInfo.KieuKy == "D_D")
			{
				string value2 = sheetView.Cells[array4[1], array4[0]].Text.Trim();
				if (!string.IsNullOrEmpty(value2) && !CheckSaiCauTrungMaHD(sheetView, TTinChung, array4[0] + 1, array4[1] + 1))
				{
					value = 2;
				}
			}
			string value3 = sheetView.Cells[array[1], array[0]].Text.Trim();
			string text5 = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			string value4 = sheetView.Cells[array3[1], array3[0]].Text.Trim();
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
			if (!string.IsNullOrEmpty(value3) || !string.IsNullOrEmpty(text5) || !string.IsNullOrEmpty(value4))
			{
				if (string.IsNullOrEmpty(value3) && !sheetView.Cells[array[1], array[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], msg);
					value = 2;
				}
				if (string.IsNullOrEmpty(text5) && !sheetView.Cells[array2[1], array2[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], msg);
					value = 2;
				}
				if (string.IsNullOrEmpty(value4) && !sheetView.Cells[array3[1], array3[0]].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], msg);
					value = 2;
				}
			}
			if (!string.IsNullOrEmpty(text5) && !ValidateBase.CheckMaSoThue(text5) && !sheetView.Cells[array2[1], array2[0]].Locked)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], msg2);
				value = 2;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12a));
			string text6 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text6))
			{
				DateTime t = DateTime.ParseExact(text6, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12c1));
			string text7 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text7))
			{
				DateTime t = DateTime.ParseExact(text7, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12d1));
			string text8 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text8))
			{
				DateTime t = DateTime.ParseExact(text8, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12dd1));
			string text9 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text9))
			{
				DateTime t = DateTime.ParseExact(text9, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12e1));
			string text10 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text10))
			{
				DateTime t = DateTime.ParseExact(text10, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12f1));
			string text11 = sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text11))
			{
				DateTime t = DateTime.ParseExact(text11, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12i1));
			string text12 = sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text12))
			{
				DateTime t = DateTime.ParseExact(text12, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt22_1));
			string text13 = sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text13))
			{
				DateTime t = DateTime.ParseExact(text13, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime utcNow = DateTime.UtcNow;
				if (DateTime.Compare(t, utcNow) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, msg3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1);
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
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		SheetView sheetView = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("D");
		int num2 = HTKKFile.ColumnLettersToIndex("B");
		int num3 = HTKKFile.ColumnLettersToIndex("E");
		int num4 = HTKKFile.ColumnLettersToIndex("F");
		int num5 = HTKKFile.ColumnLettersToIndex("G");
		int num6 = HTKKFile.ColumnLettersToIndex("J");
		int num7 = HTKKFile.ColumnLettersToIndex("K");
		int num8 = HTKKFile.ColumnLettersToIndex("L");
		int num9 = HTKKFile.ColumnLettersToIndex("M");
		int num10 = HTKKFile.ColumnLettersToIndex("N");
		int num11 = HTKKFile.ColumnLettersToIndex("O");
		int num12 = HTKKFile.ColumnLettersToIndex("P");
		int num13 = HTKKFile.ColumnLettersToIndex("S");
		int num14 = HTKKFile.ColumnLettersToIndex("T");
		int num15 = HTKKFile.ColumnLettersToIndex("U");
		int num16 = HTKKFile.ColumnLettersToIndex("V");
		int num17 = HTKKFile.ColumnLettersToIndex("W");
		int num18 = HTKKFile.ColumnLettersToIndex("AA");
		int num19 = HTKKFile.ColumnLettersToIndex("AB");
		int num20 = HTKKFile.ColumnLettersToIndex("AC");
		int colNamTinhThue = HTKKFile.ColumnLettersToIndex("AD");
		int num21 = HTKKFile.ColumnLettersToIndex("AE");
		int num22 = HTKKFile.ColumnLettersToIndex("AI");
		int num23 = HTKKFile.ColumnLettersToIndex("AJ");
		int num24 = HTKKFile.ColumnLettersToIndex("AK");
		int num25 = HTKKFile.ColumnLettersToIndex("AL");
		int num26 = HTKKFile.ColumnLettersToIndex("AM");
		int num27 = HTKKFile.ColumnLettersToIndex("AQ");
		int num28 = HTKKFile.ColumnLettersToIndex("AR");
		int num29 = HTKKFile.ColumnLettersToIndex("AS");
		int num30 = HTKKFile.ColumnLettersToIndex("AU");
		int num31 = HTKKFile.ColumnLettersToIndex("AT");
		int num32 = HTKKFile.ColumnLettersToIndex("AV");
		int num33 = HTKKFile.ColumnLettersToIndex("AF");
		int num34 = HTKKFile.ColumnLettersToIndex("AH");
		int num35 = HTKKFile.ColumnLettersToIndex("AG");
		int num36 = HTKKFile.ColumnLettersToIndex("BB");
		string text = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
		try
		{
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[42, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[43, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorKyThanhToanPhaiLienTiepNhau = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorKyKeTiepLonHonKyBatDau = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[45, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayPhaiLonHonTuNgay = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorTuNgayKyLonHonTuNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayKyNhoHonDenNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num2);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					if (sheetView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Locked)
					{
						HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
					}
				}
			}
			string value2 = sheetView.Cells[dataLocationTemplateRowStart - 1, num4 - 1].Text.Trim();
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart - 1, num4 - 1, msg);
				value = 2;
			}
			foreach (KeyValuePair<string, List<int>> item in dictionary)
			{
				int num37 = item.Value[0];
				foreach (int item2 in item.Value)
				{
					string text5 = sheetView.Cells[item2 - 1, num33 - 1].Text.Trim();
					string text6 = sheetView.Cells[item2 - 1, num3 - 1].Text.Trim();
					string value3 = sheetView.Cells[item2 - 1, num4 - 1].Text.Trim();
					if (text5 != "1" && string.IsNullOrEmpty(value3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num4 - 1, msg);
						value = 2;
					}
					string value4 = sheetView.Cells[item2 - 1, num5 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num5 - 1, msg5);
						value = 2;
					}
					string text7 = sheetView.Cells[item2 - 1, num6 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && !string.IsNullOrEmpty(text7) && !ValidateBase.CheckMaSoThue(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num6 - 1, msg2);
						value = 2;
					}
					bool flag = Convert.ToBoolean(sheetView.Cells[item2 - 1, num7 - 1].Value);
					bool flag2 = Convert.ToBoolean(sheetView.Cells[item2 - 1, num8 - 1].Value);
					if (text6 == MA_HOP_DONG && !flag && !flag2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num7 - 1, msg6);
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num8 - 1, msg6);
						value = 2;
					}
					else if (text6 == MA_HOP_DONG && flag && flag2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num7 - 1, msg7);
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num8 - 1, msg7);
						value = 2;
					}
					string value5 = sheetView.Cells[item2 - 1, num9 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value5))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num9 - 1, msg);
						value = 2;
					}
					string value6 = sheetView.Cells[item2 - 1, num10 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num10 - 1, msg5);
						value = 2;
					}
					string value7 = sheetView.Cells[item2 - 1, num11 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && flag && string.IsNullOrEmpty(value7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num11 - 1, msg4);
						value = 2;
					}
					string value8 = sheetView.Cells[item2 - 1, num14 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && flag && string.IsNullOrEmpty(value8))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num14 - 1, msg3);
						value = 2;
					}
					string value9 = sheetView.Cells[item2 - 1, num16 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && flag && string.IsNullOrEmpty(value9))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num16 - 1, msg3);
						value = 2;
					}
					double num38 = Convert.ToDouble(sheetView.Cells[item2 - 1, num17 - 1].Value);
					if (text6 == MA_HOP_DONG && flag && num38 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num17 - 1, msg4);
						value = 2;
					}
					int num39 = Convert.ToInt32(sheetView.Cells[item2 - 1, num18 - 1].Text.Trim());
					string value10 = sheetView.Cells[item2 - 1, num19 - 1].Text.Trim();
					if (string.IsNullOrEmpty(value10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num19 - 1, msg);
						value = 2;
					}
					string value11 = sheetView.Cells[item2 - 1, num20 - 1].Text.Trim();
					if (string.IsNullOrEmpty(value11))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num20 - 1, msg);
						value = 2;
					}
					if (!CheckDenNgayNhoHonTuNgay(sheetView, num37, item2, num19, num20, errorDenNgayPhaiLonHonTuNgay, errorTuNgayKyLonHonTuNgayHD, errorDenNgayKyNhoHonDenNgayHD, tuDongXoaThongBao: false, num33))
					{
						value = 2;
					}
					if (!CheckCungHDCacKyPhaiLienTiep(sheetView, hTKKCMCSheetInfo, item2, num3, num2, num19, num20, num33, colNamTinhThue, errorKyKeTiepLonHonKyBatDau, errorKyThanhToanPhaiLienTiepNhau))
					{
						value = 2;
					}
					double num40 = Convert.ToDouble(sheetView.Cells[item2 - 1, num21 - 1].Value);
					if (text6 == MA_HOP_DONG && num39 > 0 && num40 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num21 - 1, msg9);
						value = 2;
					}
					else if (text6 != MA_HOP_DONG && num40 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num21 - 1, msg8);
						value = 2;
					}
					string value12 = sheetView.Cells[item2 - 1, num30 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value12))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num30 - 1, msg);
						value = 2;
					}
					string text8 = sheetView.Cells[item2 - 1, num31 - 1].Text.Trim();
					string value13 = sheetView.Cells[item2 - 1, num32 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item2 - 1, num32 - 1, msg);
						value = 2;
					}
					if (text6 == MA_HOP_DONG && !string.IsNullOrEmpty(value13) && !CheckSaiCauTrungMaHD(sheetView, TTinChung, num32, item2))
					{
						value = 2;
					}
				}
				if (PeriodInfo.KieuKy == "D_D" && dictionary.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num37 - 1, num4 - 1, msg10);
					value = 2;
				}
			}
			if (!CheckChonHopDongTrungKy(sheetView, hTKKCMCSheetInfo, num2, num3, num4, errorMessager, tuDongXoaThongBao: false))
			{
				value = 2;
			}
			if (!CheckChonHopDongDauTien(sheetView, hTKKCMCSheetInfo, num2, num3, num4, errorMessager2, tuDongXoaThongBao: false))
			{
				value = 2;
			}
			string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			if (!Check2HopDongKhongDuocLienTiep(sheetView, hTKKCMCSheetInfo, num2, num3, num4, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: false))
			{
				value = 2;
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
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		SheetView sheetView = winFpSpread.Sheets[2];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("D");
		int num2 = HTKKFile.ColumnLettersToIndex("B");
		int num3 = HTKKFile.ColumnLettersToIndex("E");
		int num4 = HTKKFile.ColumnLettersToIndex("F");
		int num5 = HTKKFile.ColumnLettersToIndex("G");
		int num6 = HTKKFile.ColumnLettersToIndex("J");
		int num7 = HTKKFile.ColumnLettersToIndex("L");
		int num8 = HTKKFile.ColumnLettersToIndex("O");
		int num9 = HTKKFile.ColumnLettersToIndex("P");
		int num10 = HTKKFile.ColumnLettersToIndex("Q");
		int num11 = HTKKFile.ColumnLettersToIndex("R");
		int num12 = HTKKFile.ColumnLettersToIndex("S");
		int num13 = HTKKFile.ColumnLettersToIndex("U");
		int num14 = HTKKFile.ColumnLettersToIndex("T");
		int num15 = HTKKFile.ColumnLettersToIndex("V");
		int num16 = HTKKFile.ColumnLettersToIndex("W");
		int num17 = HTKKFile.ColumnLettersToIndex("Y");
		int num18 = HTKKFile.ColumnLettersToIndex("X");
		int num19 = HTKKFile.ColumnLettersToIndex("Z");
		int num20 = HTKKFile.ColumnLettersToIndex("AA");
		int num21 = HTKKFile.ColumnLettersToIndex("AB");
		int num22 = HTKKFile.ColumnLettersToIndex("AC");
		int colNamTinhThue = HTKKFile.ColumnLettersToIndex("AD");
		int num23 = HTKKFile.ColumnLettersToIndex("AE");
		int num24 = HTKKFile.ColumnLettersToIndex("AL");
		int num25 = HTKKFile.ColumnLettersToIndex("AM");
		int num26 = HTKKFile.ColumnLettersToIndex("AN");
		int num27 = HTKKFile.ColumnLettersToIndex("AO");
		int num28 = HTKKFile.ColumnLettersToIndex("AP");
		int num29 = HTKKFile.ColumnLettersToIndex("AQ");
		int num30 = HTKKFile.ColumnLettersToIndex("AR");
		int num31 = HTKKFile.ColumnLettersToIndex("AS");
		int num32 = HTKKFile.ColumnLettersToIndex("AU");
		int num33 = HTKKFile.ColumnLettersToIndex("AT");
		int num34 = HTKKFile.ColumnLettersToIndex("AV");
		int num35 = HTKKFile.ColumnLettersToIndex("AI");
		int num36 = HTKKFile.ColumnLettersToIndex("AK");
		int num37 = HTKKFile.ColumnLettersToIndex("AJ");
		string text = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
		try
		{
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorMessager2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[42, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[43, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorKyThanhToanPhaiLienTiepNhau = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorKyKeTiepLonHonKyBatDau = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[45, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[50, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayPhaiLonHonTuNgay = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorTuNgayKyLonHonTuNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string errorDenNgayKyNhoHonDenNgayHD = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num2);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					if (sheetView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Locked)
					{
						HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
					}
				}
			}
			foreach (KeyValuePair<string, List<int>> item in dictionary)
			{
				int num38 = item.Value[0];
				if (PeriodInfo.KieuKy == "D_D" && dictionary.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num38 - 1, num4 - 1, msg10);
					value = 2;
				}
			}
			string value2 = sheetView.Cells[dataLocationTemplateRowStart - 1, num4 - 1].Text.Trim();
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart - 1, num4 - 1, msg);
				value = 2;
			}
			foreach (KeyValuePair<string, List<int>> item2 in dictionary)
			{
				int num38 = item2.Value[0];
				foreach (int item3 in item2.Value)
				{
					string text5 = sheetView.Cells[item3 - 1, num35 - 1].Text.Trim();
					string text6 = sheetView.Cells[item3 - 1, num3 - 1].Text.Trim();
					string value3 = sheetView.Cells[item3 - 1, num4 - 1].Text.Trim();
					if (text5 != "1" && string.IsNullOrEmpty(value3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num4 - 1, msg);
						value = 2;
					}
					string value4 = sheetView.Cells[item3 - 1, num5 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num5 - 1, msg5);
						value = 2;
					}
					string text7 = sheetView.Cells[item3 - 1, num6 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && HTKKFormater.isNullString(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num6 - 1, msg);
						value = 2;
					}
					else if (text6 == MA_HOP_DONG && !string.IsNullOrEmpty(text7) && !ValidateBase.CheckMaSoThue(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num6 - 1, msg2);
						value = 2;
					}
					string value5 = sheetView.Cells[item3 - 1, num7 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value5))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num7 - 1, msg5);
						value = 2;
					}
					string text8 = sheetView.Cells[item3 - 1, num8 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && !string.IsNullOrEmpty(text8) && !ValidateBase.CheckMaSoThue(text8))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num8 - 1, msg2);
						value = 2;
					}
					bool flag = Convert.ToBoolean(sheetView.Cells[item3 - 1, num9 - 1].Value);
					bool flag2 = Convert.ToBoolean(sheetView.Cells[item3 - 1, num10 - 1].Value);
					if (text6 == MA_HOP_DONG && !flag && !flag2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num9 - 1, msg6);
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num10 - 1, msg6);
						value = 2;
					}
					else if (text6 == MA_HOP_DONG && flag && flag2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num9 - 1, msg7);
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num10 - 1, msg7);
						value = 2;
					}
					string value6 = sheetView.Cells[item3 - 1, num11 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num11 - 1, msg);
						value = 2;
					}
					string value7 = sheetView.Cells[item3 - 1, num12 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num12 - 1, msg5);
						value = 2;
					}
					string value8 = sheetView.Cells[item3 - 1, num13 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && flag && string.IsNullOrEmpty(value8))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num13 - 1, msg4);
						value = 2;
					}
					string value9 = sheetView.Cells[item3 - 1, num16 - 1].Text.Trim();
					if (flag && string.IsNullOrEmpty(value9))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num16 - 1, msg3);
						value = 2;
					}
					string value10 = sheetView.Cells[item3 - 1, num17 - 1].Text.Trim();
					if (flag && string.IsNullOrEmpty(value10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num17 - 1, msg3);
						value = 2;
					}
					double num39 = Convert.ToDouble(sheetView.Cells[item3 - 1, num19 - 1].Value);
					if (text6 == MA_HOP_DONG && flag && num39 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num19 - 1, msg4);
						value = 2;
					}
					int num40 = Convert.ToInt32(sheetView.Cells[item3 - 1, num20 - 1].Value);
					string value11 = sheetView.Cells[item3 - 1, num21 - 1].Text.Trim();
					if (string.IsNullOrEmpty(value11))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num21 - 1, msg);
						value = 2;
					}
					string value12 = sheetView.Cells[item3 - 1, num22 - 1].Text.Trim();
					if (string.IsNullOrEmpty(value12))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num22 - 1, msg);
						value = 2;
					}
					if (!CheckDenNgayNhoHonTuNgay(sheetView, num38, item3, num21, num22, errorDenNgayPhaiLonHonTuNgay, errorTuNgayKyLonHonTuNgayHD, errorDenNgayKyNhoHonDenNgayHD, tuDongXoaThongBao: false, num35))
					{
						value = 2;
					}
					if (!CheckCungHDCacKyPhaiLienTiep(sheetView, hTKKCMCSheetInfo, item3, num3, num2, num21, num22, num35, colNamTinhThue, errorKyKeTiepLonHonKyBatDau, errorKyThanhToanPhaiLienTiepNhau))
					{
						value = 2;
					}
					double num41 = Convert.ToDouble(sheetView.Cells[item3 - 1, num23 - 1].Value);
					if (text6 == MA_HOP_DONG && num40 > 0 && num41 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num23 - 1, msg9);
						value = 2;
					}
					else if (text6 != MA_HOP_DONG && num41 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num23 - 1, msg8);
						value = 2;
					}
					string value13 = sheetView.Cells[item3 - 1, num32 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num32 - 1, msg);
						value = 2;
					}
					string text9 = sheetView.Cells[item3 - 1, num33 - 1].Text.Trim();
					string value14 = sheetView.Cells[item3 - 1, num34 - 1].Text.Trim();
					if (text6 == MA_HOP_DONG && string.IsNullOrEmpty(value14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, item3 - 1, num34 - 1, msg);
						value = 2;
					}
					if (text6 == MA_HOP_DONG && !string.IsNullOrEmpty(value14) && !CheckSaiCauTrungMaHD(sheetView, TTinChung, num34, item3))
					{
						value = 2;
					}
				}
				if (PeriodInfo.KieuKy == "D_D" && dictionary.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num38 - 1, num4 - 1, msg10);
					value = 2;
				}
			}
			if (!CheckChonHopDongTrungKy(sheetView, hTKKCMCSheetInfo, num2, num3, num4, errorMessager, tuDongXoaThongBao: false))
			{
				value = 2;
			}
			if (!CheckChonHopDongDauTien(sheetView, hTKKCMCSheetInfo, num2, num3, num4, errorMessager2, tuDongXoaThongBao: false))
			{
				value = 2;
			}
			string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			if (!Check2HopDongKhongDuocLienTiep(sheetView, hTKKCMCSheetInfo, num2, num3, num4, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: false))
			{
				value = 2;
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

	public Dictionary<string, string> GetDMMucDich(string pathFile)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			if (dicDMMucDichTmp.Count > 0)
			{
				return dicDMMucDichTmp;
			}
			string text = pathFile + "\\Catalogue_DM_Muc_dich_SD.xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string value2 = array[1];
				dictionary2.Add(key, value2);
				dictionary.Add(array[1], array);
			}
			dicDMMucDich = dictionary;
			dicDMMucDichTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMTrangThai(string pathFile)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			string text = pathFile + "\\Catalogue_DM_trang_thai_HD.xml";
			if (dicDMTTHopDongTmp.Count > 0)
			{
				return dicDMTTHopDongTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary2.Add(key, value2);
					dictionary.Add(array[1], array);
				}
			}
			dicDMTTHopDong = dictionary;
			dicDMTTHopDongTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMHopDong1(string pathFile)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			string text = pathFile + "\\Catalogue_HopDongVaKy.xml";
			if (dicDMHopDongTmp.Count > 0)
			{
				return dicDMHopDongTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary2.Add(key, value2);
					dictionary.Add(array[1], array);
				}
			}
			dicDMHopDong = dictionary;
			dicDMHopDongTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetPhuongXaTheoCQT(string pathFile)
	{
		try
		{
			dicDMPhuongXaTheoCQTTmp = HTKKWinDictionaryDiaDanh.dicPhuongXa.Where((KeyValuePair<string, string> x) => x.Key.Substring(0, 5) == _PeriodInfo.maCQTQuanLy).ToDictionary((KeyValuePair<string, string> p) => p.Key, (KeyValuePair<string, string> p) => p.Value);
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			foreach (KeyValuePair<string, string> item in dicDMPhuongXaTheoCQTTmp)
			{
				dictionary.Add(value: new string[2] { item.Key, item.Value }, key: item.Value);
			}
			return dicDMPhuongXaTheoCQTTmp;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("U_211");
			if (idxRow == khaiTheoPLuatDanSu[1] || idxRow == khaiTheoPLuatDanThue[1])
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				bool flag = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value);
				bool flag2 = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value);
				if (idxRow == khaiTheoPLuatDanSu[1] && flag)
				{
					sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value = false;
					flag2 = false;
				}
				if (idxRow == khaiTheoPLuatDanThue[1] && flag2)
				{
					sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value = false;
					flag = false;
				}
				if (flag)
				{
					LockUnlockTheoLoaiNopThue(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, isDefault: true);
				}
				else
				{
					LockUnlockTheoLoaiNopThue(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, isDefault: true);
				}
				if (CheckChonLoaiNopThue(winFpSpread, ToKhaiInfo, PeriodInfo))
				{
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo sheet = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int colTuNgay = HTKKFile.ColumnLettersToIndex("AB");
			int colDenNgay = HTKKFile.ColumnLettersToIndex("AC");
			int colTongGiaTriHopDong = HTKKFile.ColumnLettersToIndex("AJ");
			int num3 = HTKKFile.ColumnLettersToIndex("AL");
			int num4 = HTKKFile.ColumnLettersToIndex("AM");
			int num5 = HTKKFile.ColumnLettersToIndex("AQ");
			int num6 = HTKKFile.ColumnLettersToIndex("AF");
			string pathCt24ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct24";
			HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
			string value = "";
			if (string.IsNullOrEmpty(value))
			{
				value = sheetView.Cells[(int)idxRow, num - 1].Text.Trim();
			}
			string text = sheetView.Cells[(int)idxRow, num2 - 1].Text.Trim();
			string[] array = text.Split(new char[1] { '.' });
			string key = array[0];
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, sheet, num2);
			if (dictionary.ContainsKey(key))
			{
				int num7 = dictionary[key][0];
				if (num7 == (int)idxRow + 1 && idxCol + 1 == num4)
				{
					CaNhanPhaiNopThueNeuTichHD(sheetView, dictionary[key], num, num2, num4, num3, num6, colTuNgay, colDenNgay);
				}
				if (idxCol + 1 == num4)
				{
					CheckDongConCaNhanPhaiNopThue(sheetView, num2, (int)idxRow + 1, num4, num6, colTuNgay, colDenNgay, num3);
				}
				if (dictionary.Count > 0 && idxCol + 1 == num4)
				{
					DoanhThuTinhTheTrongKy(sheetView, dictionary, num, num2, num4, num5, colTongGiaTriHopDong, colTuNgay, colDenNgay);
				}
				if (dictionary.Count > 0 && idxCol + 1 == num4)
				{
					TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt24ToKhai, dictionary, num6, num5, colTuNgay, colDenNgay, num);
				}
				GeneralFunctions.SetRowHeight(sheetView, (int)idxRow + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_ND15_TTS_TT40.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_GiamThue_GTGT_23_24_TTS_TT40.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_GiamThue_GTGT_NQ142_TTS_TT40.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("B");
				int num2 = HTKKFile.ColumnLettersToIndex("D");
				int num3 = HTKKFile.ColumnLettersToIndex("E");
				int num4 = HTKKFile.ColumnLettersToIndex("F");
				int num5 = HTKKFile.ColumnLettersToIndex("AA");
				int num6 = HTKKFile.ColumnLettersToIndex("AB");
				int num7 = HTKKFile.ColumnLettersToIndex("AC");
				int num8 = HTKKFile.ColumnLettersToIndex("AE");
				int num9 = HTKKFile.ColumnLettersToIndex("AF");
				int num10 = HTKKFile.ColumnLettersToIndex("AH");
				int num11 = HTKKFile.ColumnLettersToIndex("AG");
				int num12 = HTKKFile.ColumnLettersToIndex("BB");
				Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
				foreach (KeyValuePair<string, List<int>> item in dictionary)
				{
					int num13 = item.Value[0];
					string text = sheetView.Cells[num13 - 1, num - 1].Text.Trim();
					string text2 = sheetView.Cells[num13 - 1, num3 - 1].Text.Trim();
					string text3 = sheetView.Cells[num13 - 1, num4 - 1].Text.Trim();
					double num14 = Convert.ToDouble(sheetView.Cells[num13 - 1, num5 - 1].Value);
					string value = sheetView.Cells[num13 - 1, num6 - 1].Text.Trim();
					string value2 = sheetView.Cells[num13 - 1, num7 - 1].Text.Trim();
					string value3 = sheetView.Cells[num13 - 1, num8 - 1].Text.Trim();
					if (num14 > 0.0 && !string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(value2) && !string.IsNullOrEmpty(value3))
					{
						if (text.IndexOf(".") >= 0)
						{
							text = text.Substring(0, text.IndexOf("."));
						}
						Set1RowHDToAllRowKy(sheetView, item.Value, num, num2, "1", text);
					}
					if (item.Value.Count <= 1)
					{
						continue;
					}
					for (int i = 1; i < item.Value.Count; i++)
					{
						int num15 = item.Value[i];
						sheetView.Cells[num15 - 1, num9 - 1].Value = 0;
						sheetView.Cells[num15 - 1, num2 - 1].Value = 0;
						string text4 = sheetView.Cells[num15 - 1, num3 - 1].Text.Trim();
						string value4 = sheetView.Cells[num15 - 1, num4 - 1].Text.Trim();
						if (text4 != MA_HOP_DONG && string.IsNullOrEmpty(value4))
						{
							sheetView.Cells[num15 - 1, num9 - 1].Value = 1;
							sheetView.Cells[num15 - 1, num2 - 1].Value = 1;
						}
					}
				}
			}
			SheetView sheetView2 = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo2 != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("B");
				int num2 = HTKKFile.ColumnLettersToIndex("D");
				int num3 = HTKKFile.ColumnLettersToIndex("E");
				int num4 = HTKKFile.ColumnLettersToIndex("F");
				int num16 = HTKKFile.ColumnLettersToIndex("AA");
				int num17 = HTKKFile.ColumnLettersToIndex("AB");
				int num18 = HTKKFile.ColumnLettersToIndex("AC");
				int num19 = HTKKFile.ColumnLettersToIndex("AE");
				int num9 = HTKKFile.ColumnLettersToIndex("AI");
				Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView2, hTKKCMCSheetInfo2, num);
				foreach (KeyValuePair<string, List<int>> item2 in dictionary)
				{
					int num13 = item2.Value[0];
					string text = sheetView2.Cells[num13 - 1, num - 1].Text.Trim();
					string text2 = sheetView2.Cells[num13 - 1, num3 - 1].Text.Trim();
					string text3 = sheetView2.Cells[num13 - 1, num4 - 1].Text.Trim();
					double num20 = Convert.ToDouble(sheetView2.Cells[num13 - 1, num16 - 1].Value);
					string value5 = sheetView2.Cells[num13 - 1, num17 - 1].Text.Trim();
					string value6 = sheetView2.Cells[num13 - 1, num18 - 1].Text.Trim();
					string value7 = sheetView2.Cells[num13 - 1, num19 - 1].Text.Trim();
					if (num20 > 0.0 && !string.IsNullOrEmpty(value5) && !string.IsNullOrEmpty(value6) && !string.IsNullOrEmpty(value7))
					{
						if (text.IndexOf(".") >= 0)
						{
							text = text.Substring(0, text.IndexOf("."));
						}
						Set1RowHDToAllRowKy(sheetView2, item2.Value, num, num2, "1", text);
					}
					if (item2.Value.Count <= 1)
					{
						continue;
					}
					for (int i = 1; i < item2.Value.Count; i++)
					{
						int num15 = item2.Value[i];
						sheetView2.Cells[num15 - 1, num9 - 1].Value = 0;
						sheetView2.Cells[num15 - 1, num2 - 1].Value = 0;
						string text4 = sheetView2.Cells[num15 - 1, num3 - 1].Text.Trim();
						string value4 = sheetView2.Cells[num15 - 1, num4 - 1].Text.Trim();
						if (text4 != MA_HOP_DONG && string.IsNullOrEmpty(value4))
						{
							sheetView2.Cells[num15 - 1, num9 - 1].Value = 1;
							sheetView2.Cells[num15 - 1, num2 - 1].Value = 1;
						}
					}
				}
			}
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				AutoBindData_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, _HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				AutoBindData_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, _HeaderInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKHeaderInfo header)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiSTT(SheetView sView, HTKKCMCSheetInfo sheet, int colSTT, int colMaHD)
	{
		try
		{
			for (int i = 0; i < sheet.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sheet.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				int num = 0;
				int num2 = 0;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sView.Cells[dataLocationTemplateRowStart + j - 1, colMaHD - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text))
					{
						if (text == MA_HOP_DONG || j == 0)
						{
							num++;
							num2 = 0;
							sView.Cells[dataLocationTemplateRowStart + j - 1, colSTT - 1].Text = num.ToString();
						}
						else
						{
							num2++;
							sView.Cells[dataLocationTemplateRowStart + j - 1, colSTT - 1].Text = num + "." + num2;
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

	public void CheckNeuTichChonHDKyThanhToan1Lan(SheetView sView, HTKKCMCSheetInfo sheet, HTKKPeriodInfo PeriodInfo, string maHD, int row, int colTuNgay, int colDenNgay, int colNam, int colTemVatNam)
	{
		try
		{
			if (!(maHD == MA_TRA_TIEN_CHO_CA_HD))
			{
				return;
			}
			int num = checkVatNam(PeriodInfo.TuNgay, PeriodInfo.DenNgay);
			if (num == 0)
			{
				sView.Cells[row - 1, colTuNgay - 1].Text = PeriodInfo.TuNgay;
				sView.Cells[row - 1, colDenNgay - 1].Text = PeriodInfo.DenNgay;
				sView.Cells[row - 1, colNam - 1].Text = (string.IsNullOrEmpty(PeriodInfo.TuNgay) ? "" : splitDay(PeriodInfo.TuNgay)[2]);
				return;
			}
			for (int i = 0; i <= num; i++)
			{
				int num2 = Convert.ToInt32(splitDay(PeriodInfo.TuNgay)[2]);
				string text = firstDayOfYear(num2 + i).ToString("dd/MM/yyyy");
				string text2 = lastDayOfYear(num2 + i).ToString("dd/MM/yyyy");
				sView.Cells[row + i - 1, colTuNgay - 1].Text = ((i == 0) ? PeriodInfo.TuNgay : text);
				sView.Cells[row + i - 1, colDenNgay - 1].Text = ((!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(PeriodInfo.DenNgay) && splitDay(text2)[2] == splitDay(PeriodInfo.DenNgay)[2]) ? PeriodInfo.DenNgay : text2);
				sView.Cells[row + i - 1, colNam - 1].Text = ((!string.IsNullOrEmpty(text2)) ? splitDay(text2)[2] : "");
				sView.Cells[row + i, colTemVatNam - 1].Text = "1";
				if (i < num)
				{
					GeneralFunctions.InsertTableRow(sView, row + i, -1, row + i, -1, -1, sheet, 0);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int checkVatNam(string tuNgay, string denNgay)
	{
		return (!string.IsNullOrEmpty(denNgay) && !string.IsNullOrEmpty(tuNgay)) ? (Convert.ToInt32(splitDay(denNgay)[2]) - Convert.ToInt32(splitDay(tuNgay)[2])) : 0;
	}

	private string[] splitDay(string day)
	{
		try
		{
			return day.Split(new char[1] { '/' });
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private DateTime lastDayOfYear(int year)
	{
		return new DateTime(year + 1, 1, 1).AddDays(-1.0);
	}

	private DateTime firstDayOfYear(int year)
	{
		return new DateTime(year, 1, 1);
	}

	public int DayOfMonth(DateTime time)
	{
		return DateTime.DaysInMonth(Convert.ToInt32(time.ToString("yyyy")), Convert.ToInt32(time.ToString("MM")));
	}

	public double GetMonths(DateTime fromDate, DateTime toDate)
	{
		DateTime dateTime = new DateTime(1, 1, 1);
		TimeSpan timeSpan = toDate - fromDate;
		double num = (dateTime + timeSpan).Year - 1;
		double num2 = (dateTime + timeSpan).Month - 1;
		double num3 = (dateTime + timeSpan).Day - 1;
		return num * 12.0 + num2 + num3 / (double)DayOfMonth(toDate);
	}

	public bool CheckChonHopDongTrungKy(SheetView sView, HTKKCMCSheetInfo sheet, int colSTT, int colMaHD, int colTenHD, string errorMessager, bool tuDongXoaThongBao)
	{
		bool result = true;
		try
		{
			List<string> list = new List<string>();
			for (int i = 0; i < sheet.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sheet.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sView.Cells[dataLocationTemplateRowStart + j - 1, colSTT - 1].Text.Trim();
					string text2 = sView.Cells[dataLocationTemplateRowStart + j - 1, colMaHD - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text2) && text2 == MA_KY_THANH_TOAN_CUOI_CUNG)
					{
						string[] array = text.Split(new char[1] { '.' });
						string item = text2 + "_" + array[0];
						if (list.Contains(item))
						{
							result = false;
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + j - 1, colTenHD - 1, errorMessager);
							break;
						}
						list.Add(item);
						if (tuDongXoaThongBao)
						{
							HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, colTenHD - 1);
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public bool CheckChonHopDongDauTien(SheetView sView, HTKKCMCSheetInfo sheet, int colSTT, int colMaHD, int colTenHD, string errorMessager, bool tuDongXoaThongBao)
	{
		bool result = true;
		try
		{
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sView, sheet, colSTT);
			using Dictionary<string, List<int>>.Enumerator enumerator = dictionary.GetEnumerator();
			if (enumerator.MoveNext())
			{
				int num = enumerator.Current.Value[0];
				string text = sView.Cells[num - 1, colMaHD - 1].Text.Trim();
				if (text != MA_HOP_DONG)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num - 1, colTenHD - 1, errorMessager);
					result = false;
				}
				else if (tuDongXoaThongBao)
				{
					HTKKFile.RemoveValidate(sView, num - 1, colTenHD - 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public int TinhKyThanhToanOnDinh(SheetView sView, HTKKCMCSheetInfo sheet, HTKKPeriodInfo PeriodInfo, int colKyThanhToanOnDinh, int iRow, int colMaHopDongKy, int colTenHopDongKy, int colTuNgay, int colDenNgay, int colNam, int colGiaChoThue1Thang, int colSTT, int colTemVatNam, List<int> colAnKyXoaData, List<int> colAnKyKhongXoa, int colTemDongTuSinh, int rowHD)
	{
		int num = 0;
		try
		{
			int num2 = Convert.ToInt32(sView.Cells[iRow - 1, colKyThanhToanOnDinh - 1].Text);
			if (num2 <= 0)
			{
				return num;
			}
			string text = sView.Cells[iRow - 1, colMaHopDongKy - 1].Text.Trim();
			if (string.IsNullOrEmpty(text) || text != MA_HOP_DONG)
			{
				return num;
			}
			string text2 = sView.Cells[iRow - 1, colTuNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text2))
			{
				return num;
			}
			string text3 = sView.Cells[iRow - 1, colDenNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text3))
			{
				return num;
			}
			int num3 = Convert.ToInt32(sView.Cells[iRow - 1, colGiaChoThue1Thang - 1].Text.Trim().Replace(".", ""));
			if (num3 <= 0)
			{
				return num;
			}
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && splitDay(text2)[2] == splitDay(text3)[2])
			{
				sView.Cells[iRow - 1, colNam - 1].Text = splitDay(text2)[2];
			}
			DateTime t = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t2 = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (DateTime.Compare(t, t2) > 0)
			{
				return num;
			}
			XoaHangTheoHopDong(sView, sheet, colSTT, iRow);
			List<int> hangVatNam = new List<int>();
			Dictionary<DateTime, DateTime> dictionary = TuDongTachKy(Convert.ToInt32(num2), text2, text3, iRow, ref hangVatNam);
			foreach (KeyValuePair<DateTime, DateTime> item in dictionary)
			{
				GeneralFunctions.InsertTableRow(sView, iRow + num, -1, iRow + num, -1, -1, sheet, 0);
				sView.Cells[iRow + num, colTuNgay - 1].Text = item.Key.ToString("dd/MM/yyyy");
				sView.Cells[iRow + num, colDenNgay - 1].Text = item.Value.ToString("dd/MM/yyyy");
				sView.Cells[iRow + num, colNam - 1].Text = ((Convert.ToInt32(item.Key.ToString("yyyy")) == Convert.ToInt32(item.Value.ToString("yyyy"))) ? item.Value.ToString("yyyy") : "");
				sView.Cells[iRow + num, colGiaChoThue1Thang - 1].Value = sView.Cells[iRow + num - 1, colGiaChoThue1Thang - 1].Value;
				if (num < dictionary.Count - 1)
				{
					sView.Cells[iRow + num, colTenHopDongKy - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN];
					sView.Cells[iRow + num, colMaHopDongKy - 1].Text = MA_KY_THANH_TOAN;
				}
				else
				{
					sView.Cells[iRow + num, colTenHopDongKy - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN_CUOI_CUNG];
					sView.Cells[iRow + num, colMaHopDongKy - 1].Text = MA_KY_THANH_TOAN_CUOI_CUNG;
				}
				sView.Cells[iRow + num, colTemDongTuSinh - 1].Text = "1";
				GeneralFunctions.SetRowHeight(sView, iRow + num + 1);
				num++;
			}
			int num4 = 0;
			foreach (int item2 in hangVatNam)
			{
				sView.Cells[item2 - 1, colTemVatNam - 1].Text = "1";
				num4++;
				if (num4 >= hangVatNam.Count)
				{
					ResetDongThanhToanCuoiCung(sView, item2, colTenHopDongKy, colMaHopDongKy, colTemVatNam, colTemDongTuSinh, colDenNgay, rowHD);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public void XoaHangTheoHopDong(SheetView sView, HTKKCMCSheetInfo sheet, int colSTT, int startRow)
	{
		try
		{
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sView, sheet, colSTT);
			foreach (KeyValuePair<string, List<int>> item in dictionary)
			{
				string text = sView.Cells[startRow - 1, colSTT - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					string text2 = sView.Cells[startRow, colSTT - 1].Text.Trim();
					string[] array = text.Split(new char[1] { '.' });
					string[] array2 = text2.Split(new char[1] { '.' });
					if (array[0] != array2[0])
					{
						break;
					}
					if (item.Key == array[0] && item.Value.Count > 1)
					{
						GeneralFunctions.RemoveDataRows(sView, 0, item.Value[0], item.Value.Count - 1, sheet);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<DateTime, DateTime> TuDongTachKy(int soThangOnDinh, string tuNgay, string denNgay, int row, ref List<int> hangVatNam)
	{
		Dictionary<DateTime, DateTime> dictionary = new Dictionary<DateTime, DateTime>();
		try
		{
			DateTime dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime2 = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			double num = Convert.ToDouble(Math.Round((double)dateTime2.Subtract(dateTime).Days / 30.4375 / (double)soThangOnDinh, 1));
			int num2 = 1;
			DateTime now = DateTime.Now;
			DateTime dateTime3 = DateTime.Now;
			for (int i = 0; (double)i < num; i++)
			{
				now = ((i == 0) ? dateTime : dateTime3.AddDays(1.0));
				if (i == 0)
				{
					dateTime3 = dateTime.AddMonths(soThangOnDinh).AddDays(-1.0);
				}
				else
				{
					dateTime3 = now.AddMonths(soThangOnDinh);
					if (DateTime.Compare(dateTime2, dateTime3) > 0)
					{
						if (now.Day == 1 && dateTime3.Day - now.Day == 0)
						{
							dateTime3 = dateTime3.AddDays(-1.0);
						}
						else if (now.Day == 1 && dateTime3.Day - now.Day > 0)
						{
							dateTime3 = dateTime3.AddDays(1.0);
						}
						if (now.Day != 1 && dateTime3.Day - now.Day == 0)
						{
							dateTime3 = dateTime3.AddDays(-1.0);
						}
						else if (now.Day != 1 && dateTime3.Day - now.Day > 0)
						{
							dateTime3 = dateTime3.AddDays(1.0);
						}
					}
					else
					{
						dateTime3 = dateTime2;
					}
				}
				dictionary.Add(now, dateTime3);
				Dictionary<DateTime, DateTime> dictionary2 = TachKyNhoHonNeuVatNam(now, dateTime3);
				if (dictionary2.Count > 0)
				{
					foreach (KeyValuePair<DateTime, DateTime> item in dictionary2)
					{
						dictionary.Add(item.Key.AddMilliseconds(1.0), item.Value);
						hangVatNam.Add(row + i + num2 + 1);
						num2++;
					}
				}
				if ((double)i != num - 1.0 || DateTime.Compare(dateTime3, dateTime2) >= 0)
				{
					continue;
				}
				dictionary.Add(dateTime3.AddDays(1.0), dateTime2);
				Dictionary<DateTime, DateTime> dictionary3 = TachKyNhoHonNeuVatNam(dateTime3.AddDays(1.0), dateTime2);
				if (dictionary3.Count <= 0)
				{
					continue;
				}
				foreach (KeyValuePair<DateTime, DateTime> item2 in dictionary3)
				{
					dictionary.Add(item2.Key.AddMilliseconds(1.0), item2.Value);
					hangVatNam.Add(row + i + num2 + 1);
					num2++;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public Dictionary<DateTime, DateTime> TachKyNhoHonNeuVatNam(DateTime kyBatDau, DateTime kyKeTiep)
	{
		Dictionary<DateTime, DateTime> dictionary = new Dictionary<DateTime, DateTime>();
		int num = Convert.ToInt32(kyBatDau.ToString("yyyy"));
		int num2 = Convert.ToInt32(kyKeTiep.ToString("yyyy"));
		if (num != num2)
		{
			for (int i = num; i <= num2; i++)
			{
				if (i == num)
				{
					DateTime value = lastDayOfYear(num);
					dictionary.Add(kyBatDau, value);
				}
				else if (i == num2)
				{
					DateTime key = firstDayOfYear(num2);
					dictionary.Add(key, kyKeTiep);
				}
				else
				{
					DateTime key2 = new DateTime(i, 1, 1);
					DateTime value2 = new DateTime(i, 12, 31);
					dictionary.Add(key2, value2);
				}
			}
		}
		return dictionary;
	}

	public int TuDongTachNeuVatNam(SheetView sView, HTKKCMCSheetInfo sheet, int row, int colMaHopDongKy, int colTenHopDongKy, int colTuNgay, int colDenNgay, int colGiaChoThue1Thang, int colSTT, int colTemVatNam, List<int> colAnKyXoaData, List<int> colAnKyKhongXoa, int colTemDongTuSinh, int colKyThanhToanOnDinh, int rowHD)
	{
		Dictionary<DateTime, DateTime> dictionary = new Dictionary<DateTime, DateTime>();
		int num = 0;
		try
		{
			string text = sView.Cells[row - 1, colMaHopDongKy - 1].Text.Trim();
			if (string.IsNullOrEmpty(text) || text == MA_HOP_DONG)
			{
				return num;
			}
			string text2 = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text2))
			{
				return num;
			}
			string text3 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text3))
			{
				return num;
			}
			int num2 = Convert.ToInt32(sView.Cells[row - 1, colGiaChoThue1Thang - 1].Text.Trim().Replace(".", ""));
			if (num2 <= 0)
			{
				return num;
			}
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && splitDay(text2)[2] == splitDay(text3)[2])
			{
				return num;
			}
			DateTime dateTime = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime2 = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (DateTime.Compare(dateTime, dateTime2) > 0)
			{
				return num;
			}
			XoaHangVatNamNeuTuSinh(sView, sheet, colMaHopDongKy, colTemVatNam, rowHD, row, colSTT);
			List<int> list = new List<int>();
			Dictionary<DateTime, DateTime> dictionary2 = TachKyNhoHonNeuVatNam(dateTime, dateTime2);
			if (dictionary2.Count > 0)
			{
				int num3 = 0;
				foreach (KeyValuePair<DateTime, DateTime> item in dictionary2)
				{
					dictionary.Add(item.Key, item.Value);
					list.Add(row + num3 + 1);
					num3++;
				}
			}
			foreach (KeyValuePair<DateTime, DateTime> item2 in dictionary)
			{
				GeneralFunctions.InsertTableRow(sView, row + num, -1, row + num, -1, -1, sheet, 0);
				sView.Cells[row + num, colTuNgay - 1].Text = item2.Key.ToString("dd/MM/yyyy");
				sView.Cells[row + num, colDenNgay - 1].Text = item2.Value.ToString("dd/MM/yyyy");
				sView.Cells[row + num, colGiaChoThue1Thang - 1].Value = num2;
				if (num < dictionary.Count)
				{
					sView.Cells[row + num, colTenHopDongKy - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN];
					sView.Cells[row + num, colMaHopDongKy - 1].Text = MA_KY_THANH_TOAN;
				}
				sView.Cells[row + num, colTemDongTuSinh - 1].Text = "1";
				num++;
			}
			int num4 = 0;
			foreach (int item3 in list)
			{
				sView.Cells[item3 - 1, colTemVatNam - 1].Text = "1";
				num4++;
				if (num4 >= list.Count)
				{
					ResetDongThanhToanCuoiCung(sView, item3, colTenHopDongKy, colMaHopDongKy, colTemVatNam, colTemDongTuSinh, colDenNgay, rowHD);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public void XoaHangVatNamNeuTuSinh(SheetView sView, HTKKCMCSheetInfo sheet, int colMaHD, int colTemVatNam, int rowHD, int startRow, int colSTT)
	{
		try
		{
			string text = sView.Cells[rowHD - 1, colSTT - 1].Text.Trim();
			int num = 0;
			for (int i = 1; i < 100; i++)
			{
				string text2 = sView.Cells[startRow + i - 1, colSTT - 1].Text.Trim();
				string text3 = sView.Cells[startRow + i - 1, colMaHD - 1].Text.Trim();
				int num2 = Convert.ToInt32(sView.Cells[startRow + i - 1, colTemVatNam - 1].Value);
				if (string.IsNullOrEmpty(text3) || text3 == MA_HOP_DONG || num2 != 1)
				{
					break;
				}
				if (text2.IndexOf(".") >= 0)
				{
					text2 = text2.Substring(0, text2.IndexOf("."));
				}
				if (string.IsNullOrEmpty(text2) || text2 != text)
				{
					break;
				}
				num = i;
			}
			if (num > 0)
			{
				GeneralFunctions.RemoveDataRows(sView, 0, startRow, num, sheet);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckCungHDCacKyPhaiLienTiep(SheetView sView, HTKKCMCSheetInfo sheet, int row, int colMaHD, int colSTT, int colTuNgay, int colDenNgay, int colTemVatNam, int colNamTinhThue, string errorKyKeTiepLonHonKyBatDau, string errorKyThanhToanPhaiLienTiepNhau)
	{
		try
		{
			string text = sView.Cells[row, colMaHD - 1].Text.Trim();
			if (text == MA_HOP_DONG)
			{
				return true;
			}
			string text2 = sView.Cells[row, colSTT - 1].Text;
			if (text2.IndexOf(".") >= 0)
			{
				text2 = text2.Substring(0, text2.IndexOf("."));
			}
			string text3 = sView.Cells[row + 1, colSTT - 1].Text;
			if (text3.IndexOf(".") >= 0)
			{
				text3 = text3.Substring(0, text3.IndexOf("."));
			}
			if (!string.IsNullOrEmpty(sView.Cells[row + 1, colSTT - 1].Text) && text2 != text3)
			{
				return true;
			}
			string text4 = sView.Cells[row, colTuNgay - 1].Text.Trim();
			string text5 = sView.Cells[row, colDenNgay - 1].Text.Trim();
			string text6 = sView.Cells[row + 1, colTuNgay - 1].Text.Trim();
			string text7 = sView.Cells[row + 1, colDenNgay - 1].Text.Trim();
			string text8 = sView.Cells[row + 1, colTemVatNam - 1].Text.Trim();
			string text9 = sView.Cells[row + 1, colNamTinhThue - 1].Text.Trim();
			if (string.IsNullOrEmpty(text4) || string.IsNullOrEmpty(text5) || string.IsNullOrEmpty(text6) || string.IsNullOrEmpty(text7))
			{
				return true;
			}
			DateTime t = DateTime.ParseExact(text4, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t2 = DateTime.ParseExact(text5, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime = DateTime.ParseExact(text6, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t3 = DateTime.ParseExact(text7, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (DateTime.Compare(t, t2) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, row, colTuNgay - 1, errorKyKeTiepLonHonKyBatDau);
				HTKKFile.SetNodeBackGroundPink(sView, row, colDenNgay - 1, errorKyKeTiepLonHonKyBatDau);
				return false;
			}
			if (text8 != "1" && DateTime.Compare(t2.AddDays(1.0), dateTime) != 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, row, colDenNgay - 1, errorKyThanhToanPhaiLienTiepNhau);
				HTKKFile.SetNodeBackGroundPink(sView, row + 1, colTuNgay - 1, errorKyThanhToanPhaiLienTiepNhau);
				return false;
			}
			if (DateTime.Compare(dateTime, t3) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, row, colTuNgay - 1, errorKyKeTiepLonHonKyBatDau);
				HTKKFile.SetNodeBackGroundPink(sView, row, colDenNgay - 1, errorKyKeTiepLonHonKyBatDau);
				return false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public bool CheckDenNgayNhoHonTuNgay(SheetView sView, int rowHD, int rowKy, int colTuNgay, int colDenNgay, string errorDenNgayPhaiLonHonTuNgay, string errorTuNgayKyLonHonTuNgayHD, string errorDenNgayKyNhoHonDenNgayHD, bool tuDongXoaThongBao, int colTemVatNam)
	{
		try
		{
			string text = sView.Cells[rowHD - 1, colTuNgay - 1].Text.Trim();
			string text2 = sView.Cells[rowHD - 1, colDenNgay - 1].Text.Trim();
			string text3 = sView.Cells[rowKy - 1, colTuNgay - 1].Text.Trim();
			string text4 = sView.Cells[rowKy - 1, colDenNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text3) || string.IsNullOrEmpty(text4) || string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2))
			{
				return true;
			}
			string text5 = sView.Cells[rowKy - 1, colTemVatNam - 1].Text.Trim();
			if (tuDongXoaThongBao)
			{
				if (text5 != "1")
				{
					HTKKFile.RemoveValidate(sView, rowKy - 1, colTuNgay - 1);
					HTKKFile.RemoveValidate(sView, rowKy - 1, colDenNgay - 1);
				}
				else
				{
					HTKKFile.RemoveValidateBackGround(sView, rowKy - 1, colTuNgay - 1);
					HTKKFile.RemoveValidateBackGround(sView, rowKy - 1, colDenNgay - 1);
				}
			}
			DateTime t = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime2 = DateTime.ParseExact(text4, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (DateTime.Compare(t, dateTime) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, rowKy - 1, colTuNgay - 1, errorTuNgayKyLonHonTuNgayHD);
				return false;
			}
			if (DateTime.Compare(dateTime2, t2) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, rowKy - 1, colDenNgay - 1, errorDenNgayKyNhoHonDenNgayHD);
				return false;
			}
			if (DateTime.Compare(dateTime, dateTime2) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, rowKy - 1, colTuNgay - 1, errorDenNgayPhaiLonHonTuNgay);
				HTKKFile.SetNodeBackGroundPink(sView, rowKy - 1, colDenNgay - 1, errorDenNgayPhaiLonHonTuNgay);
				return false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public Dictionary<string, List<int>> LayCacHangCungHopDong(SheetView sView, HTKKCMCSheetInfo sheet, int colSTT)
	{
		Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
		if (sheet == null || sView == null)
		{
			return dictionary;
		}
		try
		{
			for (int i = 0; i < sheet.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sheet.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					int num = dataLocationTemplateRowStart + j;
					string text = sView.Cells[num - 1, colSTT - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text))
					{
						string[] array = text.Split(new char[1] { '.' });
						if (dictionary.ContainsKey(array[0]))
						{
							dictionary[array[0]].Add(num);
							continue;
						}
						List<int> list = new List<int>();
						list.Add(num);
						dictionary.Add(array[0], list);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public double TinhTongGiaTriHD(SheetView sView, int rowHD, int row, int colMaHD, int colTuNgay, int colDenNgay, int colGiaChoThue1Thang, int colTemVatNam, int colTongGiaTriHD, int colTemSoThang, int colTemTronThang, int colSoThang, bool tuDongSoThang)
	{
		double num = 0.0;
		double num2 = 0.0;
		try
		{
			num = Convert.ToDouble(sView.Cells[rowHD - 1, colTongGiaTriHD - 1].Value);
			num2 = Convert.ToDouble(sView.Cells[rowHD - 1, colTemSoThang - 1].Value);
			double num3 = 0.0;
			string text = sView.Cells[row - 1, colMaHD - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				return num;
			}
			double num4 = Convert.ToDouble(sView.Cells[row - 1, colGiaChoThue1Thang - 1].Value);
			string text2 = sView.Cells[row - 1, colTemVatNam - 1].Text.Trim();
			int num5 = ((!string.IsNullOrEmpty(text2)) ? Convert.ToInt32(text2) : 0);
			string text3 = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			string text4 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text3) || string.IsNullOrEmpty(text4) || num4 <= 0.0 || string.IsNullOrEmpty(text))
			{
				return num;
			}
			if (text != MA_HOP_DONG)
			{
				DateTime dateTime = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime2 = DateTime.ParseExact(text4, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				string text5 = dateTime.ToString("yyyy");
				string text6 = dateTime2.ToString("yyyy");
				double num6 = 0.0;
				string text7 = "0";
				int num7 = 0;
				int num8 = 0;
				string text8 = "";
				string text9 = "";
				int num9 = row;
				if (text2 == "1")
				{
					while (text2 == "1")
					{
						num9--;
						text2 = sView.Cells[num9 - 1, colTemVatNam - 1].Text.Trim();
						num8++;
					}
					num7 = num9;
					if (num7 > 0)
					{
						text9 = sView.Cells[num7 - 1, colTuNgay - 1].Text.Trim();
						text8 = sView.Cells[num7 - 1, colDenNgay - 1].Text.Trim();
					}
					text2 = sView.Cells[row - 1, colTemVatNam - 1].Text.Trim();
					num5 = ((!string.IsNullOrEmpty(text2)) ? Convert.ToInt32(text2) : 0);
				}
				DateTime dateTime3 = DateTime.MinValue;
				DateTime dateTime4 = DateTime.MinValue;
				if (!string.IsNullOrEmpty(text8) && !string.IsNullOrEmpty(text9))
				{
					dateTime3 = DateTime.ParseExact(text9, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime4 = DateTime.ParseExact(text8, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				}
				if (num5 == 1 && !string.IsNullOrEmpty(text8) && !string.IsNullOrEmpty(text9) && ((dateTime3.Day == 1 && dateTime4.Day == DateTime.DaysInMonth(dateTime4.Year, dateTime4.Month)) || (dateTime4.Day != DateTime.DaysInMonth(dateTime4.Year, dateTime4.Month) && dateTime3.Day - dateTime4.Day == 1)))
				{
					if (dateTime.Day == 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month))
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month + 1));
						num3 *= num4;
						text7 = "1";
					}
					else if (dateTime.Day != 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month))
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)(DateTime.DaysInMonth(dateTime3.Year, dateTime3.Month) - dateTime.Day + 1) / ((double)DateTime.DaysInMonth(dateTime3.Year, dateTime3.Month) - (double)dateTime3.Day + (double)dateTime4.Day + 1.0));
						num3 *= num4;
						text7 = "3.1";
					}
					else if (dateTime.Day == 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month))
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)dateTime2.Day / ((double)(DateTime.DaysInMonth(dateTime3.Year, dateTime3.Month) - dateTime3.Day + 1) + (double)dateTime4.Day));
						num3 *= num4;
						text7 = "3.2";
					}
					sView.Cells[row - 1, colTemTronThang - 1].Text = text7.ToString();
					sView.Cells[row - 1, colTemSoThang - 1].Text = num6.ToString();
					sView.Cells[row - 1, colTongGiaTriHD - 1].Text = Math.Round(num3, 0).ToString();
					num9 = num7 + 1;
					double num10 = 0.0;
					double num11 = 0.0;
					for (; sView.Cells[num9 - 1, colTemVatNam - 1].Text.Trim() == "1"; num9++)
					{
						num10 += Convert.ToDouble(sView.Cells[num9 - 1, colTongGiaTriHD - 1].Value);
						num11 += Convert.ToDouble(sView.Cells[num9 - 1, colTemSoThang - 1].Value);
					}
					if (num7 > 0)
					{
						sView.Cells[num7 - 1, colTemSoThang - 1].Text = Convert.ToString(num11);
						sView.Cells[num7 - 1, colTongGiaTriHD - 1].Text = Convert.ToString(num10);
					}
					num += num3;
					num2 += num6;
					sView.Cells[rowHD - 1, colTongGiaTriHD - 1].Text = Math.Round(num, 0).ToString();
					if (tuDongSoThang)
					{
						sView.Cells[rowHD - 1, colTemSoThang - 1].Value = num2;
						sView.Cells[rowHD - 1, colSoThang - 1].Text = Math.Round(num2, 0).ToString();
					}
				}
				else
				{
					if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text4) && splitDay(text3)[2] != splitDay(text4)[2])
					{
						return num;
					}
					if (dateTime.Day == 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month))
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month + 1));
						num3 *= num4;
						text7 = "1";
					}
					else if (dateTime.Day != 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day == 1)
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month));
						num3 *= num4;
						text7 = "2";
					}
					else if (dateTime.Day != 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day > 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 -= (double)(dateTime.Day - dateTime2.Day - 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text7 = "4";
					}
					else if (dateTime.Day != 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)(dateTime2.Day - dateTime.Day + 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text7 = "5";
					}
					else if (dateTime.Day != 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)(dateTime2.Day - dateTime.Day + 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text7 = "6.1";
					}
					else if (dateTime.Day == 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month) + ConvertHelpers.ConvertToDouble1(dateTime2.Day) / ConvertHelpers.ConvertToDouble1(DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month)));
						num3 *= num4;
						text7 = "6.2";
					}
					else
					{
						num6 = (num3 = (dateTime2.Year - dateTime.Year) * 12 + dateTime2.Month - dateTime.Month + 1);
						num3 *= num4;
						text7 = "7";
					}
					num += num3;
					num2 += num6;
					sView.Cells[row - 1, colTemTronThang - 1].Text = text7.ToString();
					sView.Cells[row - 1, colTemSoThang - 1].Text = num6.ToString();
					sView.Cells[row - 1, colTongGiaTriHD - 1].Text = Math.Round(num3, 0).ToString();
					sView.Cells[rowHD - 1, colTongGiaTriHD - 1].Text = Math.Round(num, 0).ToString();
					num9 = num7 + 1;
					double num10 = 0.0;
					double num11 = 0.0;
					for (; sView.Cells[num9 - 1, colTemVatNam - 1].Text.Trim() == "1"; num9++)
					{
						num10 += Convert.ToDouble(sView.Cells[num9 - 1, colTongGiaTriHD - 1].Value);
						num11 += Convert.ToDouble(sView.Cells[num9 - 1, colTemSoThang - 1].Value);
					}
					if (num7 > 0)
					{
						sView.Cells[num7 - 1, colTemSoThang - 1].Text = Convert.ToString(num11);
						sView.Cells[num7 - 1, colTongGiaTriHD - 1].Text = Convert.ToString(num10);
					}
					if (tuDongSoThang)
					{
						sView.Cells[rowHD - 1, colTemSoThang - 1].Value = num2;
						sView.Cells[rowHD - 1, colSoThang - 1].Text = Math.Round(num2, 0).ToString();
					}
				}
				if (sView.Cells[row - 1, colTemVatNam - 1].Text.Trim() == "1")
				{
					num9 = row;
					string text10 = sView.Cells[num9 - 1, colTemVatNam - 1].Text.Trim();
					double num12 = 0.0;
					double num13 = 0.0;
					while (text10 == "1")
					{
						num12 += Convert.ToDouble(sView.Cells[num9 - 1, colTongGiaTriHD - 1].Value);
						num13 += Convert.ToDouble(sView.Cells[row - 1, colTemSoThang - 1].Value);
						num9--;
						text10 = sView.Cells[num9 - 1, colTemVatNam - 1].Text.Trim();
					}
					sView.Cells[num9 - 1, colTemSoThang - 1].Text = Convert.ToString(num13);
					sView.Cells[num9 - 1, colTongGiaTriHD - 1].Text = Convert.ToString(num12);
				}
				if (text != MA_HOP_DONG)
				{
					GeneralFunctions.LockUnLockCol(sView, row, colTongGiaTriHD, locked: false, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sView, rowHD, colTongGiaTriHD, locked: true, deleteData: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public void SumTinhTongGiaTriHD(SheetView sView, List<int> hopDong, int colTongGiaTriHD, int colTuNgay, int colDenNgay, int colTemVatNam)
	{
		try
		{
			double num = 0.0;
			for (int i = 1; i < hopDong.Count; i++)
			{
				int num2 = hopDong[i];
				string text = sView.Cells[num2 - 1, colTuNgay - 1].Text.Trim();
				string text2 = sView.Cells[num2 - 1, colDenNgay - 1].Text.Trim();
				if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2) || string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2) || !(splitDay(text)[2] == splitDay(text2)[2]))
				{
					continue;
				}
				num += Convert.ToDouble(sView.Cells[num2 - 1, colTongGiaTriHD - 1].Value);
				int num3 = 0;
				if (sView.Cells[num2 - 1, colTemVatNam - 1].Text.Trim() == "1")
				{
					int num4 = num2;
					double num5 = 0.0;
					while (sView.Cells[num4 - 1, colTemVatNam - 1].Text.Trim() == "1")
					{
						num4--;
					}
					num3 = num4;
					for (num4++; sView.Cells[num4 - 1, colTemVatNam - 1].Text.Trim() == "1"; num4++)
					{
						num5 += Convert.ToDouble(sView.Cells[num4 - 1, colTongGiaTriHD - 1].Value);
					}
					if (num3 > 0)
					{
						sView.Cells[num3 - 1, colTongGiaTriHD - 1].Value = num5;
					}
				}
			}
			sView.Cells[hopDong[0] - 1, colTongGiaTriHD - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void DoanhThuBinhQuanThangCuaHD(SheetView sView, int colTongGiaTriHD, int colDoanhThuBinhQuanThang, int colMaHD, int colSoThangChoThueCuaHD, List<int> hopDong)
	{
		try
		{
			int num = hopDong[0];
			for (int i = 1; i < hopDong.Count; i++)
			{
				int num2 = hopDong[i];
				string text = sView.Cells[num2 - 1, colMaHD - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text) && !(text == MA_HOP_DONG))
				{
					double num3 = Convert.ToDouble(sView.Cells[num - 1, colTongGiaTriHD - 1].Value);
					double num4 = Convert.ToDouble(sView.Cells[num - 1, colSoThangChoThueCuaHD - 1].Value);
					if (num3 == 0.0 || num4 == 0.0)
					{
						sView.Cells[num2 - 1, colDoanhThuBinhQuanThang - 1].Value = 0;
						sView.Cells[num - 1, colDoanhThuBinhQuanThang - 1].Value = 0;
					}
					else
					{
						double num5 = Math.Round(num3 / num4, 0);
						sView.Cells[num2 - 1, colDoanhThuBinhQuanThang - 1].Value = num5;
						sView.Cells[num - 1, colDoanhThuBinhQuanThang - 1].Value = num5;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongDoanhThuPhatSinhTrongNam(SheetView sView, Dictionary<string, List<int>> cacHopDong, int colNamTinhThue, int colTongGiaTriHD, int colTongDoanhThuPhatSinhTrongNam, int colSTT, int colMST)
	{
		try
		{
			Dictionary<string, double> dictionary = new Dictionary<string, double>();
			foreach (KeyValuePair<string, List<int>> item in cacHopDong)
			{
				int num = item.Value[0];
				string text = ((colMST > 0) ? sView.Cells[num - 1, colMST - 1].Text.Trim() : "");
				for (int i = 1; i < item.Value.Count; i++)
				{
					int num2 = item.Value[i];
					string text2 = sView.Cells[num2 - 1, colNamTinhThue - 1].Text.Trim();
					if (string.IsNullOrEmpty(text2))
					{
						continue;
					}
					double num3 = Convert.ToDouble(sView.Cells[num2 - 1, colTongGiaTriHD - 1].Value);
					if (num3 != 0.0)
					{
						string key = ((!string.IsNullOrEmpty(text)) ? GeneralFunctions.FormatKey(text, text2) : text2);
						if (dictionary.ContainsKey(key))
						{
							dictionary[key] += num3;
						}
						else
						{
							dictionary.Add(key, num3);
						}
					}
				}
			}
			if (dictionary.Count <= 0)
			{
				return;
			}
			foreach (KeyValuePair<string, List<int>> item2 in cacHopDong)
			{
				int num = item2.Value[0];
				string text = ((colMST > 0) ? sView.Cells[num - 1, colMST - 1].Text.Trim() : "");
				for (int i = 1; i < item2.Value.Count; i++)
				{
					int num2 = item2.Value[i];
					string text2 = sView.Cells[num2 - 1, colNamTinhThue - 1].Text.Trim();
					if (string.IsNullOrEmpty(text2))
					{
						sView.Cells[num2 - 1, colTongDoanhThuPhatSinhTrongNam - 1].Text = "0";
						continue;
					}
					string key = ((!string.IsNullOrEmpty(text)) ? GeneralFunctions.FormatKey(text, text2) : text2);
					if (dictionary.ContainsKey(key))
					{
						sView.Cells[num2 - 1, colTongDoanhThuPhatSinhTrongNam - 1].Text = dictionary[key].ToString();
					}
					else
					{
						sView.Cells[num2 - 1, colTongDoanhThuPhatSinhTrongNam - 1].Text = "0";
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CaNhanPhaiNopThue(SheetView sView, Dictionary<string, List<int>> cacHopDong, HTKKPeriodInfo PeriodInfo, int colMaHD, int colDoanhThuBinhQuanThang, int colTongDoanhThuPhatSinhTrongNam, int colSTT, int colCaNhanThuocDienPhaiNopThue, int colTuNgay, int colDenNgay, int colTemVatNam)
	{
		try
		{
			Dictionary<string, double> dictionary = new Dictionary<string, double>();
			foreach (KeyValuePair<string, List<int>> item in cacHopDong)
			{
				int num = item.Value[0];
				for (int i = 1; i < item.Value.Count; i++)
				{
					int num2 = item.Value[i];
					string text = sView.Cells[num2 - 1, colMaHD - 1].Text.Trim();
					if (string.IsNullOrEmpty(text))
					{
						continue;
					}
					string text2 = sView.Cells[num2 - 1, colTemVatNam - 1].Text.Trim();
					if (text != MA_HOP_DONG)
					{
						double num3 = Convert.ToDouble(sView.Cells[num2 - 1, colDoanhThuBinhQuanThang - 1].Value);
						double num4 = Convert.ToDouble(sView.Cells[num2 - 1, colTongDoanhThuPhatSinhTrongNam - 1].Value);
						if (num3 == 0.0 || num4 == 0.0)
						{
							continue;
						}
						string s = sView.Cells[num2 - 1, colTuNgay - 1].Text.Trim();
						string s2 = sView.Cells[num2 - 1, colDenNgay - 1].Text.Trim();
						DateTime dateTime = DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						DateTime dateTime2 = new DateTime(DateTime.ParseExact(s2, "dd/MM/yyyy", CultureInfo.CurrentCulture).Year, 1, 1).AddDays(-1.0);
						if (num4 > 100000000.0)
						{
							sView.Cells[num2 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
						}
						else
						{
							sView.Cells[num2 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = _nhapTuExcel && Convert.ToBoolean(sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value);
						}
					}
					if (!(text2 == "1"))
					{
						continue;
					}
					int num5 = num2;
					int num6 = 0;
					bool flag = true;
					while (sView.Cells[num5 - 1, colTemVatNam - 1].Text == "1")
					{
						num5--;
					}
					num6 = num5;
					for (num5++; sView.Cells[num5 - 1, colTemVatNam - 1].Text == "1"; num5++)
					{
						if (!Convert.ToBoolean(sView.Cells[num5 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value))
						{
							flag = false;
						}
					}
					if (text != MA_HOP_DONG && flag)
					{
						sView.Cells[num6 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
					}
				}
			}
			foreach (KeyValuePair<string, List<int>> item2 in cacHopDong)
			{
				for (int i = 1; i < item2.Value.Count; i++)
				{
					int num2 = item2.Value[i];
					Neu1Trong2DongVatNamTichThiDongTrenAnTheo(sView, colSTT, colTemVatNam, num2, colCaNhanThuocDienPhaiNopThue, colTuNgay, colDenNgay);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CaNhanPhaiNopThueNeuTichHD(SheetView sView, List<int> hopDong, int colMaHD, int colSTT, int colCaNhanThuocDienPhaiNopThue, int colTongDoanhThuPhatSinhTrongNamTungCN, int colTemVatNam, int colTuNgay, int colDenNgay)
	{
		try
		{
			int num = hopDong[0];
			bool flag = Convert.ToBoolean(sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value);
			int num2 = 0;
			for (int i = 1; i < hopDong.Count; i++)
			{
				int num3 = hopDong[i];
				string text = sView.Cells[num3 - 1, colMaHD - 1].Text.Trim();
				if (string.IsNullOrEmpty(text))
				{
					continue;
				}
				string[] array = sView.Cells[num3 - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				if (text != MA_HOP_DONG)
				{
					if (Convert.ToDouble(sView.Cells[num3 - 1, colTongDoanhThuPhatSinhTrongNamTungCN - 1].Value) >= 100000000.0)
					{
						sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
					}
					else
					{
						sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = flag;
					}
				}
				num2++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void DoanhThuTinhTheTrongKy(SheetView sView, Dictionary<string, List<int>> cacHopDong, int colMaHD, int colSTT, int colCaNhanThuocDienPhaiNopThue, int colDoanhThuTinhThueTrongKy, int colTongGiaTriHopDong, int colTuNgay, int colDenNgay)
	{
		try
		{
			foreach (KeyValuePair<string, List<int>> item in cacHopDong)
			{
				int num = item.Value[0];
				sView.Cells[num - 1, colDoanhThuTinhThueTrongKy - 1].Value = 0;
				double num2 = 0.0;
				for (int i = 1; i < item.Value.Count; i++)
				{
					int num3 = item.Value[i];
					string text = sView.Cells[num3 - 1, colMaHD - 1].Text.Trim();
					if (string.IsNullOrEmpty(text))
					{
						continue;
					}
					bool flag = Convert.ToBoolean(sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value);
					double num4 = Convert.ToDouble(sView.Cells[num3 - 1, colTongGiaTriHopDong - 1].Value);
					string text2 = sView.Cells[num3 - 1, colTuNgay - 1].Text.Trim();
					string text3 = sView.Cells[num3 - 1, colDenNgay - 1].Text.Trim();
					if (string.IsNullOrEmpty(text2) || string.IsNullOrEmpty(text3))
					{
						continue;
					}
					if (text != MA_HOP_DONG && flag)
					{
						sView.Cells[num3 - 1, colDoanhThuTinhThueTrongKy - 1].Value = num4;
						if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && splitDay(text2)[2] == splitDay(text3)[2])
						{
							num2 += num4;
						}
					}
					else
					{
						sView.Cells[num3 - 1, colDoanhThuTinhThueTrongKy - 1].Value = 0;
					}
				}
				sView.Cells[num - 1, colDoanhThuTinhThueTrongKy - 1].Value = num2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SumDoanhThuTinhTheTrongKy(SheetView sView, List<int> hopDong, int colDoanhThuTinhThueTrongKy, int colTuNgay, int colDenNgay)
	{
		try
		{
			double num = 0.0;
			for (int i = 1; i < hopDong.Count; i++)
			{
				int num2 = hopDong[i];
				string text = sView.Cells[num2 - 1, colTuNgay - 1].Text.Trim();
				string text2 = sView.Cells[num2 - 1, colDenNgay - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && splitDay(text)[2] == splitDay(text2)[2])
				{
					num += Convert.ToDouble(sView.Cells[num2 - 1, colDoanhThuTinhThueTrongKy - 1].Value);
				}
			}
			sView.Cells[hopDong[0] - 1, colDoanhThuTinhThueTrongKy - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhMaHopDong(SheetView sView, HTKKCMCTTinChungInfo TTinChung, int colTuNgay, int colMaHD1, int colSTT, int colMaTrangThaiHD, int colMaHopDong2, int row, HTKKCMCTableInfo table)
	{
		try
		{
			string[] array = sView.Cells[row - 1, colSTT - 1].Text.Split(new char[1] { '.' });
			string text = sView.Cells[row - 1, colMaTrangThaiHD - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				string text2 = sView.Cells[row - 1, colTuNgay - 1].Text.Replace("/", "");
				if (!string.IsNullOrEmpty(text2) && !(text != MA_HOP_DONG_KHAI_LAN_DAU))
				{
					string text3 = TTinChung.maCQTNoiNop + "_" + ((TTinChung.mst.Length == 10) ? (TTinChung.mst + "000_") : (TTinChung.mst + "_")) + text2 + "_" + ((array[0].Length == 1) ? ("0" + array[0]) : array[0]);
					sView.Cells[row - 1, colMaHopDong2 - 1].Text = text3;
					GeneralFunctions.SetRowHeight(sView, row);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongDoanhThuPhatSinhTrongKyCuaToKhai(SheetView sViewToKhai, SheetView sViewPL, HTKKPeriodInfo PeriodInfo, HTKKCMCCellInfo rowCellCt23ToKhai, Dictionary<string, List<int>> cacHopDong, int colTemVatNam, int colTongGiaTriHopDong, int colTuNgay, int colDenNgay, int colMaHD)
	{
		try
		{
			double num = 0.0;
			if (_ToKhaiInfo.SheetInfos[1] != null && (PeriodInfo.KieuKy == "Y" || PeriodInfo.KieuKy == "D_D"))
			{
				foreach (KeyValuePair<string, List<int>> item in cacHopDong)
				{
					foreach (int item2 in item.Value)
					{
						string text = sViewPL.Cells[item2 - 1, colTemVatNam - 1].Text.Trim();
						if (text == "1")
						{
							continue;
						}
						string text2 = sViewPL.Cells[item2 - 1, colTuNgay - 1].Text.Trim();
						string text3 = sViewPL.Cells[item2 - 1, colDenNgay - 1].Text.Trim();
						string text4 = sViewPL.Cells[item2 - 1, colMaHD - 1].Text.Trim();
						double num2 = Convert.ToDouble(sViewPL.Cells[item2 - 1, colTongGiaTriHopDong - 1].Value);
						if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && !(text4 == MA_HOP_DONG) && num2 != 0.0)
						{
							if (!string.IsNullOrEmpty(text2) && PeriodInfo.KieuKy == "Y" && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "D_D" && PeriodInfo.TuNgay == text2 && PeriodInfo.DenNgay == text3)
							{
								num = num2;
							}
						}
					}
				}
			}
			if (_ToKhaiInfo.SheetInfos[2] != null)
			{
				foreach (KeyValuePair<string, List<int>> item3 in cacHopDong)
				{
					foreach (int item4 in item3.Value)
					{
						string text = sViewPL.Cells[item4 - 1, colTemVatNam - 1].Text.Trim();
						if (text == "1")
						{
							continue;
						}
						string text2 = sViewPL.Cells[item4 - 1, colTuNgay - 1].Text.Trim();
						string text3 = sViewPL.Cells[item4 - 1, colDenNgay - 1].Text.Trim();
						string text4 = sViewPL.Cells[item4 - 1, colMaHD - 1].Text.Trim();
						double num2 = Convert.ToDouble(sViewPL.Cells[item4 - 1, colTongGiaTriHopDong - 1].Value);
						if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && !(text4 == MA_HOP_DONG) && num2 != 0.0)
						{
							List<string> quyByNamTC = GeneralFunctions.GetQuyByNamTC(int.Parse(splitDay(text2)[1]), int.Parse(splitDay(text2)[2]), "01/01");
							if (PeriodInfo.KieuKy == "Y" && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "M" && PeriodInfo.Thang == splitDay(text2)[1] && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "Q" && PeriodInfo.Nam == splitDay(text2)[2] && int.Parse(PeriodInfo.Quy) == int.Parse(quyByNamTC[0]))
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "D_D" && PeriodInfo.TuNgay == text2 && PeriodInfo.DenNgay == text3)
							{
								num = num2;
							}
						}
					}
				}
			}
			sViewToKhai.Cells[rowCellCt23ToKhai.CellTemplate.RowIndex - 1, rowCellCt23ToKhai.CellTemplate.ColIndex - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckDongConCaNhanPhaiNopThue(SheetView sView, int colSTT, int row, int colCaNhanThuocDienPhaiNopThue, int colVatNam, int colTuNgay, int colDenNgay, int colTongDoanhThuPhatSinhTrongNamTungCaNhan)
	{
		try
		{
			string text = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			string text2 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2))
			{
				return;
			}
			int num = 0;
			int num2 = 0;
			string text3 = sView.Cells[row - 1, colVatNam - 1].Text.Trim();
			if (text3 == "0" && splitDay(text)[2] != splitDay(text2)[2])
			{
				num = row;
				num2 = num + 1;
				text3 = sView.Cells[num2 - 1, colVatNam - 1].Text.Trim();
				while (text3 == "1")
				{
					sView.Cells[num2 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = Convert.ToDouble(sView.Cells[num2 - 1, colTongDoanhThuPhatSinhTrongNamTungCaNhan - 1].Value) > 100000000.0 || Convert.ToBoolean(sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value);
					num2++;
					text3 = sView.Cells[num2 - 1, colVatNam - 1].Text.Trim();
				}
				return;
			}
			num2 = row;
			bool flag = true;
			int num3 = 0;
			while (text3 == "1")
			{
				num2--;
				text3 = sView.Cells[num2 - 1, colVatNam - 1].Text.Trim();
			}
			num = num2++;
			text3 = sView.Cells[num2 - 1, colVatNam - 1].Text.Trim();
			while (text3 == "1")
			{
				if (!Convert.ToBoolean(sView.Cells[num2 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value))
				{
					flag = false;
				}
				num3++;
				num2++;
				text3 = sView.Cells[num2 - 1, colVatNam - 1].Text.Trim();
			}
			if (num3 > 0)
			{
				sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = flag;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongDoanhThuTinhThueCuaToKhai(SheetView sViewToKhai, SheetView sViewPL, HTKKPeriodInfo PeriodInfo, HTKKCMCCellInfo rowCellCt24ToKhai, Dictionary<string, List<int>> cacHopDong, int colTemVatNam, int colDoanhThuTinhTheTrongKy, int colTuNgay, int colDenNgay, int colMaHD)
	{
		try
		{
			double num = 0.0;
			if (_ToKhaiInfo.SheetInfos[1] != null && (PeriodInfo.KieuKy == "Y" || PeriodInfo.KieuKy == "D_D"))
			{
				foreach (KeyValuePair<string, List<int>> item in cacHopDong)
				{
					foreach (int item2 in item.Value)
					{
						string text = sViewPL.Cells[item2 - 1, colTemVatNam - 1].Text.Trim();
						if (text == "1")
						{
							continue;
						}
						string text2 = sViewPL.Cells[item2 - 1, colTuNgay - 1].Text.Trim();
						string text3 = sViewPL.Cells[item2 - 1, colDenNgay - 1].Text.Trim();
						string text4 = sViewPL.Cells[item2 - 1, colMaHD - 1].Text.Trim();
						double num2 = Convert.ToDouble(sViewPL.Cells[item2 - 1, colDoanhThuTinhTheTrongKy - 1].Value);
						if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && !(text4 == MA_HOP_DONG) && num2 != 0.0)
						{
							if (PeriodInfo.KieuKy == "Y" && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "D_D" && PeriodInfo.TuNgay == text2 && PeriodInfo.DenNgay == text3)
							{
								num = num2;
							}
						}
					}
				}
			}
			if (_ToKhaiInfo.SheetInfos[2] != null)
			{
				foreach (KeyValuePair<string, List<int>> item3 in cacHopDong)
				{
					foreach (int item4 in item3.Value)
					{
						string text = sViewPL.Cells[item4 - 1, colTemVatNam - 1].Text.Trim();
						if (text == "1")
						{
							continue;
						}
						string text2 = sViewPL.Cells[item4 - 1, colTuNgay - 1].Text.Trim();
						string text3 = sViewPL.Cells[item4 - 1, colDenNgay - 1].Text.Trim();
						string text4 = sViewPL.Cells[item4 - 1, colMaHD - 1].Text.Trim();
						double num2 = Convert.ToDouble(sViewPL.Cells[item4 - 1, colDoanhThuTinhTheTrongKy - 1].Value);
						if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && !(text4 == MA_HOP_DONG) && num2 != 0.0)
						{
							List<string> quyByNamTC = GeneralFunctions.GetQuyByNamTC(int.Parse(splitDay(text2)[1]), int.Parse(splitDay(text2)[2]), "01/01");
							if (PeriodInfo.KieuKy == "Y" && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "D_D" && PeriodInfo.TuNgay == text2 && PeriodInfo.DenNgay == text3)
							{
								num = num2;
							}
							else if (PeriodInfo.KieuKy == "M" && PeriodInfo.Thang == splitDay(text2)[1] && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "Q" && PeriodInfo.Nam == splitDay(text2)[2] && int.Parse(PeriodInfo.Quy) == int.Parse(quyByNamTC[0]))
							{
								num += num2;
							}
						}
					}
				}
			}
			sViewToKhai.Cells[rowCellCt24ToKhai.CellTemplate.RowIndex - 1, rowCellCt24ToKhai.CellTemplate.ColIndex - 1].Value = num;
			TinhChiTieu25_TK(sViewToKhai.FpSpread, _ToKhaiInfo);
			TinhChiTieu26_TK(sViewToKhai.FpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e, Keys keyData)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0005: Invalid comparison between Unknown and I4
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		if ((int)keyData == 117)
		{
			idxRow--;
			int idxCol2 = HTKKFile.ColumnLettersToIndex("F");
			if (idxSheet == 1)
			{
				KeyPressF6_Pl01_1(winFpSpread, ToKhaiInfo, idxSheet, idxRow, idxCol2, e, keyData);
			}
			if (idxSheet == 2)
			{
				KeyPressF6_Pl01_2(winFpSpread, ToKhaiInfo, idxSheet, idxRow, idxCol2, e, keyData);
			}
		}
	}

	public bool KeyPressF6_Pl01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e, Keys keyData)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo sheet = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int item = HTKKFile.ColumnLettersToIndex("G");
			int item2 = HTKKFile.ColumnLettersToIndex("J");
			int item3 = HTKKFile.ColumnLettersToIndex("K");
			int item4 = HTKKFile.ColumnLettersToIndex("L");
			int item5 = HTKKFile.ColumnLettersToIndex("M");
			int item6 = HTKKFile.ColumnLettersToIndex("N");
			int item7 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int item8 = HTKKFile.ColumnLettersToIndex("S");
			int item9 = HTKKFile.ColumnLettersToIndex("T");
			int num6 = HTKKFile.ColumnLettersToIndex("U");
			int item10 = HTKKFile.ColumnLettersToIndex("V");
			int item11 = HTKKFile.ColumnLettersToIndex("W");
			int item12 = HTKKFile.ColumnLettersToIndex("AA");
			int num7 = HTKKFile.ColumnLettersToIndex("AB");
			int num8 = HTKKFile.ColumnLettersToIndex("AC");
			int num9 = HTKKFile.ColumnLettersToIndex("AD");
			int num10 = HTKKFile.ColumnLettersToIndex("AE");
			int num11 = HTKKFile.ColumnLettersToIndex("AI");
			int item13 = HTKKFile.ColumnLettersToIndex("AJ");
			int num12 = HTKKFile.ColumnLettersToIndex("AK");
			int num13 = HTKKFile.ColumnLettersToIndex("AL");
			int num14 = HTKKFile.ColumnLettersToIndex("AM");
			int num15 = HTKKFile.ColumnLettersToIndex("AQ");
			int num16 = HTKKFile.ColumnLettersToIndex("AR");
			int num17 = HTKKFile.ColumnLettersToIndex("AS");
			int item14 = HTKKFile.ColumnLettersToIndex("AU");
			int num18 = HTKKFile.ColumnLettersToIndex("AT");
			int num19 = HTKKFile.ColumnLettersToIndex("AF");
			int item15 = HTKKFile.ColumnLettersToIndex("AV");
			int num20 = HTKKFile.ColumnLettersToIndex("AH");
			int num21 = HTKKFile.ColumnLettersToIndex("AG");
			int num22 = HTKKFile.ColumnLettersToIndex("BB");
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			list.Add(item5);
			list.Add(item6);
			list.Add(item7);
			list.Add(item8);
			list.Add(item9);
			list.Add(item10);
			List<int> list2 = new List<int>();
			list2.Add(item11);
			list2.Add(item14);
			list2.Add(item15);
			List<int> list3 = new List<int>();
			list3.Add(item13);
			List<int> list4 = new List<int>();
			list4.Add(num4);
			list4.Add(item12);
			list4.Add(item13);
			string text = sheetView.Cells[(int)_iRowClick - 1, num - 1].Text.Trim();
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, sheet, num);
			if (dictionary.Count > 0)
			{
				foreach (KeyValuePair<string, List<int>> item16 in dictionary)
				{
					int num23 = item16.Value[0];
					string text2 = sheetView.Cells[num23 - 1, num - 1].Text.Trim();
					string text3 = sheetView.Cells[(int)_iRowClick - 2, num - 1].Text.Trim();
					if (text2.IndexOf(".") >= 0)
					{
						text2 = text2.Substring(0, text2.IndexOf("."));
					}
					if (text3.IndexOf(".") >= 0)
					{
						text3 = text3.Substring(0, text3.IndexOf("."));
					}
					if (text2.Length > 0 && text3.Length > 0 && text2 == text3)
					{
						CellChange_1(winFpSpread, ToKhaiInfo, _TTinChung, _PeriodInfo, num23, num4);
					}
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool KeyPressF6_Pl01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e, Keys keyData)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int item = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("J");
			int item2 = HTKKFile.ColumnLettersToIndex("L");
			int item3 = HTKKFile.ColumnLettersToIndex("O");
			int item4 = HTKKFile.ColumnLettersToIndex("P");
			int item5 = HTKKFile.ColumnLettersToIndex("Q");
			int item6 = HTKKFile.ColumnLettersToIndex("R");
			int item7 = HTKKFile.ColumnLettersToIndex("S");
			int item8 = HTKKFile.ColumnLettersToIndex("U");
			int num6 = HTKKFile.ColumnLettersToIndex("T");
			int item9 = HTKKFile.ColumnLettersToIndex("V");
			int item10 = HTKKFile.ColumnLettersToIndex("W");
			int item11 = HTKKFile.ColumnLettersToIndex("Y");
			int num7 = HTKKFile.ColumnLettersToIndex("X");
			int num8 = HTKKFile.ColumnLettersToIndex("Z");
			int item12 = HTKKFile.ColumnLettersToIndex("AA");
			int num9 = HTKKFile.ColumnLettersToIndex("AB");
			int num10 = HTKKFile.ColumnLettersToIndex("AC");
			int num11 = HTKKFile.ColumnLettersToIndex("AD");
			int num12 = HTKKFile.ColumnLettersToIndex("AE");
			int item13 = HTKKFile.ColumnLettersToIndex("AL");
			int item14 = HTKKFile.ColumnLettersToIndex("AM");
			int num13 = HTKKFile.ColumnLettersToIndex("AN");
			int num14 = HTKKFile.ColumnLettersToIndex("AO");
			int num15 = HTKKFile.ColumnLettersToIndex("AP");
			int num16 = HTKKFile.ColumnLettersToIndex("AQ");
			int num17 = HTKKFile.ColumnLettersToIndex("AR");
			int num18 = HTKKFile.ColumnLettersToIndex("AS");
			int item15 = HTKKFile.ColumnLettersToIndex("AU");
			int num19 = HTKKFile.ColumnLettersToIndex("AT");
			int item16 = HTKKFile.ColumnLettersToIndex("AV");
			int num20 = HTKKFile.ColumnLettersToIndex("AI");
			int num21 = HTKKFile.ColumnLettersToIndex("AK");
			int num22 = HTKKFile.ColumnLettersToIndex("AJ");
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(num5);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			list.Add(item5);
			list.Add(item6);
			list.Add(item7);
			list.Add(item8);
			list.Add(item9);
			list.Add(item10);
			list.Add(item11);
			list.Add(item13);
			List<int> list2 = new List<int>();
			list2.Add(item9);
			list2.Add(item15);
			list2.Add(item16);
			List<int> list3 = new List<int>();
			list3.Add(item14);
			List<int> list4 = new List<int>();
			list4.Add(num4);
			list4.Add(item12);
			list4.Add(item14);
			string text = sheetView.Cells[idxRow - 1, num - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				idxRow++;
			}
			HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables[0];
			string[] array = text.Split(new char[1] { '.' });
			string key = array[0];
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
			if (!dictionary.ContainsKey(key))
			{
				return false;
			}
			int num23 = dictionary[key][0];
			Dictionary<string, List<int>> dictionary2 = new Dictionary<string, List<int>>();
			dictionary2 = TimHopDongCungMSTCungNam(sheetView, table, num23, num, num5, dictionary);
			if (dictionary2.Count > 0)
			{
				foreach (KeyValuePair<string, List<int>> item17 in dictionary2)
				{
					string text2 = sheetView.Cells[num23 - 1, num - 1].Text.Trim();
					string text3 = sheetView.Cells[idxRow - 2, num - 1].Text.Trim();
					if (text2.IndexOf(".") >= 0)
					{
						text2 = text2.Substring(0, text2.IndexOf("."));
					}
					if (text3.IndexOf(".") >= 0)
					{
						text3 = text3.Substring(0, text3.IndexOf("."));
					}
					if (text2.Length > 0 && text3.Length > 0 && text2 == text3)
					{
						CellChange_2(winFpSpread, ToKhaiInfo, _TTinChung, _PeriodInfo, num23, num4);
					}
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void AutoShowYear(SheetView sView, int row, int colMaHD, int colTuNgay, int colDenNgay, int colNamTinhThue)
	{
		try
		{
			string text = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			string text2 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
			{
				string text3 = sView.Cells[row - 1, colMaHD - 1].Text.Trim();
				if (splitDay(text)[2] == splitDay(text2)[2])
				{
					sView.Cells[row - 1, colNamTinhThue - 1].Text = splitDay(text)[2];
				}
				else
				{
					sView.Cells[row - 1, colNamTinhThue - 1].Text = "";
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
			int colSTT = HTKKFile.ColumnLettersToIndex("B");
			int num = HTKKFile.ColumnLettersToIndex("D");
			int colMaHD = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("L");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("O");
			int num10 = HTKKFile.ColumnLettersToIndex("P");
			int num11 = HTKKFile.ColumnLettersToIndex("S");
			int num12 = HTKKFile.ColumnLettersToIndex("T");
			int num13 = HTKKFile.ColumnLettersToIndex("U");
			int num14 = HTKKFile.ColumnLettersToIndex("V");
			int num15 = HTKKFile.ColumnLettersToIndex("W");
			int num16 = HTKKFile.ColumnLettersToIndex("AA");
			int colTuNgay = HTKKFile.ColumnLettersToIndex("AB");
			int colDenNgay = HTKKFile.ColumnLettersToIndex("AC");
			int num17 = HTKKFile.ColumnLettersToIndex("AD");
			int num18 = HTKKFile.ColumnLettersToIndex("AE");
			int num19 = HTKKFile.ColumnLettersToIndex("AI");
			int colTongGiaTriHopDong = HTKKFile.ColumnLettersToIndex("AJ");
			int num20 = HTKKFile.ColumnLettersToIndex("AK");
			int num21 = HTKKFile.ColumnLettersToIndex("AL");
			int num22 = HTKKFile.ColumnLettersToIndex("AM");
			int colDoanhThuTinhTheTrongKy = HTKKFile.ColumnLettersToIndex("AQ");
			int num23 = HTKKFile.ColumnLettersToIndex("AR");
			int num24 = HTKKFile.ColumnLettersToIndex("AS");
			int num25 = HTKKFile.ColumnLettersToIndex("AU");
			int num26 = HTKKFile.ColumnLettersToIndex("AT");
			int num27 = HTKKFile.ColumnLettersToIndex("AV");
			int colTemVatNam = HTKKFile.ColumnLettersToIndex("AF");
			int num28 = HTKKFile.ColumnLettersToIndex("AH");
			int num29 = HTKKFile.ColumnLettersToIndex("AG");
			int num30 = HTKKFile.ColumnLettersToIndex("BB");
			string pathCt23ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct23";
			string pathCt24ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct24";
			string pathMaHDToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
			HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
			HTKKCMCCellInfo rowCellCt23ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt23ToKhai));
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaHDToKhai));
			SheetView sheetView = winFpSpread.Sheets[1];
			if (sheetView != null)
			{
				HTKKCMCSheetInfo sheet = ToKhaiInfo.SheetInfos[1];
				Dictionary<string, List<int>> cacHopDong = LayCacHangCungHopDong(sheetView, sheet, colSTT);
				winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = "";
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, _PeriodInfo, rowCellCt23ToKhai, cacHopDong, colTemVatNam, colTongGiaTriHopDong, colTuNgay, colDenNgay, colMaHD);
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, _PeriodInfo, rowCellCt24ToKhai, cacHopDong, colTemVatNam, colDoanhThuTinhTheTrongKy, colTuNgay, colDenNgay, colMaHD);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_nhapTuExcel = false;
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo != null)
			{
				int colSTT = HTKKFile.ColumnLettersToIndex("B");
				int num = HTKKFile.ColumnLettersToIndex("D");
				int colMaHD = HTKKFile.ColumnLettersToIndex("E");
				int num2 = HTKKFile.ColumnLettersToIndex("F");
				int num3 = HTKKFile.ColumnLettersToIndex("G");
				int num4 = HTKKFile.ColumnLettersToIndex("J");
				int num5 = HTKKFile.ColumnLettersToIndex("L");
				int num6 = HTKKFile.ColumnLettersToIndex("O");
				int num7 = HTKKFile.ColumnLettersToIndex("P");
				int num8 = HTKKFile.ColumnLettersToIndex("Q");
				int num9 = HTKKFile.ColumnLettersToIndex("R");
				int num10 = HTKKFile.ColumnLettersToIndex("S");
				int num11 = HTKKFile.ColumnLettersToIndex("U");
				int num12 = HTKKFile.ColumnLettersToIndex("T");
				int num13 = HTKKFile.ColumnLettersToIndex("V");
				int num14 = HTKKFile.ColumnLettersToIndex("W");
				int num15 = HTKKFile.ColumnLettersToIndex("Y");
				int num16 = HTKKFile.ColumnLettersToIndex("X");
				int num17 = HTKKFile.ColumnLettersToIndex("Z");
				int num18 = HTKKFile.ColumnLettersToIndex("AA");
				int colTuNgay = HTKKFile.ColumnLettersToIndex("AB");
				int colDenNgay = HTKKFile.ColumnLettersToIndex("AC");
				int num19 = HTKKFile.ColumnLettersToIndex("AD");
				int num20 = HTKKFile.ColumnLettersToIndex("AE");
				int num21 = HTKKFile.ColumnLettersToIndex("AL");
				int colTongGiaTriHopDong = HTKKFile.ColumnLettersToIndex("AM");
				int num22 = HTKKFile.ColumnLettersToIndex("AN");
				int num23 = HTKKFile.ColumnLettersToIndex("AO");
				int num24 = HTKKFile.ColumnLettersToIndex("AP");
				int colDoanhThuTinhTheTrongKy = HTKKFile.ColumnLettersToIndex("AQ");
				int num25 = HTKKFile.ColumnLettersToIndex("AR");
				int num26 = HTKKFile.ColumnLettersToIndex("AS");
				int num27 = HTKKFile.ColumnLettersToIndex("AU");
				int num28 = HTKKFile.ColumnLettersToIndex("AT");
				int num29 = HTKKFile.ColumnLettersToIndex("AV");
				int colTemVatNam = HTKKFile.ColumnLettersToIndex("AI");
				int num30 = HTKKFile.ColumnLettersToIndex("AK");
				int num31 = HTKKFile.ColumnLettersToIndex("AJ");
				string pathMaHDToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/maHDong";
				string pathCt23ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct23";
				string pathCt24ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct24";
				HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
				HTKKCMCCellInfo rowCellCt23ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt23ToKhai));
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaHDToKhai));
				Dictionary<string, List<int>> cacHopDong = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, colSTT);
				winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = "";
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, _PeriodInfo, rowCellCt23ToKhai, cacHopDong, colTemVatNam, colTongGiaTriHopDong, colTuNgay, colDenNgay, colMaHD);
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, _PeriodInfo, rowCellCt24ToKhai, cacHopDong, colTemVatNam, colDoanhThuTinhTheTrongKy, colTuNgay, colDenNgay, colMaHD);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_3[1], cellCoPhuLuc_01_3[0]].Value = ((_ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
			TinhChiTieu25_TK(winFpSpread, ToKhaiInfo);
			TinhChiTieu26_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_4[1], cellCoPhuLuc_01_4[0]].Value = ((_ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
			TinhChiTieu25_TK(winFpSpread, ToKhaiInfo);
			TinhChiTieu26_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPhuLuc_01_5[1], cellCoPhuLuc_01_5[0]].Value = ((_ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
			TinhChiTieu25_TK(winFpSpread, ToKhaiInfo);
			TinhChiTieu26_TK(winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlockTheoLoaiNopThue(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, bool isDefault)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("AA");
			int num3 = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AD");
			int[] array = GeneralFunctions.ArrayCellId("F_42");
			int[] array2 = GeneralFunctions.ArrayCellId("F_44");
			int[] array3 = GeneralFunctions.ArrayCellId("H_54");
			int[] array4 = GeneralFunctions.ArrayCellId("H_50");
			int[] array5 = GeneralFunctions.ArrayCellId("H_48");
			int[] array6 = GeneralFunctions.ArrayCellId("E_64");
			int[] array7 = GeneralFunctions.ArrayCellId("J_64");
			int[] array8 = GeneralFunctions.ArrayCellId("Q_64");
			int[] array9 = GeneralFunctions.ArrayCellId("M_66");
			int[] array10 = GeneralFunctions.ArrayCellId("M_68");
			int[] array11 = GeneralFunctions.ArrayCellId("I_186");
			int[] array12 = GeneralFunctions.ArrayCellId("E_188");
			int[] array13 = GeneralFunctions.ArrayCellId("E_190");
			int[] array14 = GeneralFunctions.ArrayCellId("E_192");
			int[] array15 = GeneralFunctions.ArrayCellId("J_192");
			int[] array16 = GeneralFunctions.ArrayCellId("Q_192");
			int[] array17 = GeneralFunctions.ArrayCellId("H_194");
			int[] array18 = GeneralFunctions.ArrayCellId("S_194");
			bool flag = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value);
			if (flag)
			{
				sheetView.Cells[array9[1], array9[0]].Locked = false;
				sheetView.Cells[array10[1], array10[0]].Locked = false;
				sheetView.Cells[array11[1], array11[0]].Locked = false;
				sheetView.Cells[array12[1], array12[0]].Locked = false;
				sheetView.Cells[array13[1], array13[0]].Locked = false;
				sheetView.Cells[array14[1], array14[0]].Locked = false;
				sheetView.Cells[array15[1], array15[0]].Locked = false;
				sheetView.Cells[array16[1], array16[0]].Locked = false;
				sheetView.Cells[array[1], array[0]].Value = _HeaderInfo.tenNNT;
				sheetView.Cells[array2[1], array2[0]].Value = _HeaderInfo.mst;
				sheetView.Cells[array3[1], array3[0]].Value = _HeaderInfo.dchiNNT;
				sheetView.Cells[array4[1], array4[0]].Value = _HeaderInfo.tenHuyenNNT;
				sheetView.Cells[array5[1], array5[0]].Value = _HeaderInfo.tenTinhNNT;
				sheetView.Cells[array6[1], array6[0]].Value = _HeaderInfo.dthoaiNNT;
				sheetView.Cells[array7[1], array7[0]].Value = _HeaderInfo.faxNNT;
				sheetView.Cells[array8[1], array8[0]].Value = _HeaderInfo.emailNNT;
				if (isDefault)
				{
					sheetView.Cells[array11[1], array11[0]].Value = "";
					sheetView.Cells[array12[1], array12[0]].Value = "";
					sheetView.Cells[array13[1], array13[0]].Value = "";
					sheetView.Cells[array14[1], array14[0]].Value = "";
					sheetView.Cells[array15[1], array15[0]].Value = "";
					sheetView.Cells[array16[1], array16[0]].Value = "";
					sheetView.Cells[array17[1], array17[0]].Value = "";
					sheetView.Cells[array18[1], array18[0]].Value = "";
				}
				return;
			}
			sheetView.Cells[array3[1], array3[0]].Value = "";
			sheetView.Cells[array4[1], array4[0]].Value = "";
			sheetView.Cells[array5[1], array5[0]].Value = "";
			sheetView.Cells[array9[1], array9[0]].Locked = true;
			sheetView.Cells[array10[1], array10[0]].Locked = true;
			sheetView.Cells[array11[1], array11[0]].Locked = true;
			sheetView.Cells[array12[1], array12[0]].Locked = true;
			sheetView.Cells[array13[1], array13[0]].Locked = true;
			sheetView.Cells[array14[1], array14[0]].Locked = true;
			sheetView.Cells[array15[1], array15[0]].Locked = true;
			sheetView.Cells[array16[1], array16[0]].Locked = true;
			sheetView.Cells[array[1], array[0]].Value = "";
			sheetView.Cells[array2[1], array2[0]].Value = "";
			sheetView.Cells[array3[1], array3[0]].Value = "";
			sheetView.Cells[array4[1], array4[0]].Value = "";
			sheetView.Cells[array5[1], array5[0]].Value = "";
			sheetView.Cells[array6[1], array6[0]].Value = "";
			sheetView.Cells[array7[1], array7[0]].Value = "";
			sheetView.Cells[array8[1], array8[0]].Value = "";
			if (isDefault)
			{
				sheetView.Cells[array9[1], array9[0]].Value = "";
				sheetView.Cells[array10[1], array10[0]].Value = "";
				sheetView.Cells[array11[1], array11[0]].Value = _HeaderInfo.tenNNT;
				sheetView.Cells[array12[1], array12[0]].Value = _HeaderInfo.mst;
				sheetView.Cells[array13[1], array13[0]].Value = _HeaderInfo.dchiNNT + " - " + _HeaderInfo.tenHuyenNNT + " - " + _HeaderInfo.tenTinhNNT;
				sheetView.Cells[array14[1], array14[0]].Value = _HeaderInfo.dthoaiNNT;
				sheetView.Cells[array15[1], array15[0]].Value = _HeaderInfo.faxDLyThue;
				sheetView.Cells[array16[1], array16[0]].Value = _HeaderInfo.emailNNT;
				sheetView.Cells[array17[1], array17[0]].Value = "";
				sheetView.Cells[array18[1], array18[0]].Value = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckChonLoaiNopThue(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value);
			HTKKFile.RemoveValidate(sheetView, khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]);
			HTKKFile.RemoveValidate(sheetView, khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]);
			if ((PeriodInfo.KieuKy == "Q" || PeriodInfo.KieuKy == "M") && flag)
			{
				string message = _xmlMessage.GetMessage("0566");
				ThongBaoVeChonHinhThuc(sheetView, message, khaiTheoPLuatDanSuShow: false, khaiTheoPLuatDanThueShow: false, all: true);
				return false;
			}
			if (!flag2 && !flag)
			{
				string message = _xmlMessage.GetMessage("0535");
				ThongBaoVeChonHinhThuc(sheetView, message, khaiTheoPLuatDanSuShow: false, khaiTheoPLuatDanThueShow: false, all: true);
				return false;
			}
			if (ToKhaiInfo.SheetInfos[1] != null && ToKhaiInfo.SheetInfos[2] != null)
			{
				string message = _xmlMessage.GetMessage("0532");
				ThongBaoVeChonHinhThuc(sheetView, message, khaiTheoPLuatDanSuShow: true, khaiTheoPLuatDanThueShow: true, all: false);
				return false;
			}
			if (flag && ToKhaiInfo.SheetInfos[1] == null)
			{
				string message = _xmlMessage.GetMessage("0558");
				ThongBaoVeChonHinhThuc(sheetView, message, khaiTheoPLuatDanSuShow: true, khaiTheoPLuatDanThueShow: false, all: false);
				return false;
			}
			if (flag2 && ToKhaiInfo.SheetInfos[2] == null)
			{
				string message = _xmlMessage.GetMessage("0559");
				ThongBaoVeChonHinhThuc(sheetView, message, khaiTheoPLuatDanSuShow: false, khaiTheoPLuatDanThueShow: true, all: false);
				return false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return true;
	}

	public bool Check2HopDongKhongDuocLienTiep(SheetView sView, HTKKCMCSheetInfo sInfo, int colCt06, int colCt06aMa, int colCt06aTen, string duoiHopDongLaKyTinhThue, bool tuDongXoaThongBao)
	{
		try
		{
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sView, sInfo, colCt06);
			foreach (KeyValuePair<string, List<int>> item in dictionary)
			{
				int num = item.Value[0];
				if (tuDongXoaThongBao)
				{
					HTKKFile.RemoveValidate(sView, num - 1, colCt06aTen - 1);
				}
				string text = sView.Cells[num - 1, colCt06aMa - 1].Text.Trim();
				if (text == MA_HOP_DONG)
				{
					string text2 = sView.Cells[num - 2, colCt06aMa - 1].Text.Trim();
					if (text2 == text)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num - 1, colCt06aTen - 1, duoiHopDongLaKyTinhThue);
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return true;
	}

	public void LockUnLockDongHopDongHoacKy(SheetView sView, HTKKCMCSheetInfo sInfo, int colSTT, int colCt06aMa, int colTemDongTuSinh, int colTemVatNam, List<int> colAnKy, List<int> colAnNhungKhongXoaDuLieuKy, List<int> colAnNeuTuSinh, List<int> colAnNeuVatNam, List<int> colAnNeuLaHopDong, List<int> colKhongKhoaXoaDataNeuVatNamDongKy, List<int> colKhongAnNeuLaHopDong, int colTuNgay, int colDenNgay, List<int> colAnNeuTruocVatNam, int colTongDoanhThuPhatSinhTrongNamTungCN, int colCaNhanThuocDienPhaiNopThue, List<int> colAnKhoaNeuLaDongHD, Dictionary<string, List<int>> cacHopDong)
	{
		try
		{
			if (cacHopDong.Count <= 0)
			{
				return;
			}
			foreach (KeyValuePair<string, List<int>> item in cacHopDong)
			{
				int row = item.Value[0];
				for (int i = 1; i < item.Value.Count; i++)
				{
					int num = item.Value[i];
					string text = sView.Cells[num - 1, colCt06aMa - 1].Text.Trim();
					string text2 = sView.Cells[num - 1, colTemDongTuSinh - 1].Text.Trim();
					string text3 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
					string text4 = sView.Cells[num - 1, colTuNgay - 1].Text.Trim();
					string text5 = sView.Cells[num - 1, colDenNgay - 1].Text.Trim();
					double num2 = Convert.ToDouble(sView.Cells[num - 1, colTongDoanhThuPhatSinhTrongNamTungCN - 1].Value);
					sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Locked = num2 > 100000000.0;
					if (text != MA_HOP_DONG)
					{
						GeneralFunctions.LockUnLockCols(sView, num, colAnKy, locked: true, deleteData: true);
						GeneralFunctions.LockUnLockCols(sView, num, colAnNeuVatNam, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCols(sView, num, colAnNhungKhongXoaDuLieuKy, locked: true, deleteData: false);
						if (text2 == "1")
						{
							GeneralFunctions.LockUnLockCols(sView, num, colAnNeuTuSinh, locked: true, deleteData: false);
						}
						if (text3 == "1")
						{
							GeneralFunctions.LockUnLockCols(sView, num, colAnNeuVatNam, locked: true, deleteData: false);
							GeneralFunctions.LockUnLockCols(sView, num, colKhongKhoaXoaDataNeuVatNamDongKy, locked: false, deleteData: true);
						}
						GeneralFunctions.LockUnLockCols(sView, num, colAnNeuLaHopDong, locked: false, deleteData: false);
						if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text5) && splitDay(text4)[2] != splitDay(text5)[2])
						{
							GeneralFunctions.LockUnLockCols(sView, num, colAnNeuTruocVatNam, locked: true, deleteData: false);
						}
						Neu1Trong2DongVatNamTichThiDongTrenAnTheo(sView, colSTT, colTemVatNam, num, colCaNhanThuocDienPhaiNopThue, colTuNgay, colDenNgay);
					}
				}
				GeneralFunctions.LockUnLockCols(sView, row, colAnNeuLaHopDong, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colKhongAnNeuLaHopDong, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnKy, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnNhungKhongXoaDuLieuKy, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnNeuTuSinh, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnNeuVatNam, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colKhongKhoaXoaDataNeuVatNamDongKy, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnNeuTruocVatNam, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnKhoaNeuLaDongHD, locked: true, deleteData: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThongBaoVeChonHinhThuc(SheetView sView, string message, bool khaiTheoPLuatDanSuShow, bool khaiTheoPLuatDanThueShow, bool all)
	{
		if (khaiTheoPLuatDanSuShow || (all && !sView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Locked))
		{
			HTKKFile.SetNodeBackGroundPink(sView, khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0], message);
		}
		if (khaiTheoPLuatDanThueShow || (all && !sView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Locked))
		{
			HTKKFile.SetNodeBackGroundPink(sView, khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0], message);
		}
	}

	public bool CheckSaiCauTrungMaHD(SheetView sView, HTKKCMCTTinChungInfo TTinChung, int colMaHopdong, int rowMaHopDong)
	{
		try
		{
			string text = sView.Cells[rowMaHopDong - 1, colMaHopdong - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				return true;
			}
			if (text.IndexOf("_") > 0)
			{
				string[] array = text.Split(new char[1] { '_' });
				if (array.Length != 4)
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				if (HTKKFormater.isNullString(array[0]) || (!HTKKFormater.isNullString(array[0]) && !array[0].Equals(TTinChung.maCQTNoiNop)))
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				if (HTKKFormater.isNullString(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				if ((TTinChung.mst.Length >= array[1].Length && !array[1].Equals(TTinChung.mst)) || (TTinChung.mst.Length < array[1].Length && !array[1].Substring(0, TTinChung.mst.Length).Equals(TTinChung.mst)))
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				if (HTKKFormater.isNullString(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				DateTime dtDate = DateTime.Now;
				string strDate = HTKKFormater.toDateDDMMYYYY(array[2], isMess: false);
				if (!HTKKFormater.IsDateTime(strDate, "dd/MM/yyyy", out dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				if (HTKKFormater.isNullString(array[3]) || (!HTKKFormater.isNullString(array[3]) && ConvertHelpers.ConvertToInt(array[3]) == 0))
				{
					HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					return false;
				}
				return true;
			}
			HTKKFile.SetNodeBackGroundPink(sView, rowMaHopDong - 1, colMaHopdong - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void Set1RowHDToAllRowKy(SheetView sView, List<int> hopDong, int colStt, int colCopy, string valueCopy, string valueHD)
	{
		try
		{
			string text = sView.Cells[hopDong[0] - 1, colStt - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			if (text.IndexOf(".") >= 0)
			{
				text = text.Substring(0, text.IndexOf("."));
			}
			if (text == valueHD)
			{
				for (int i = 1; i < hopDong.Count; i++)
				{
					sView.Cells[hopDong[i] - 1, colCopy - 1].Text = valueCopy;
					GeneralFunctions.SetRowHeight(sView, hopDong[i]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ResetDongThanhToanCuoiCung(SheetView sView, int rowKy, int colTenHD, int colMaHD, int colTemVatNam, int colTemTuSinh, int colDenNgay, int rowHD)
	{
		try
		{
			string text = sView.Cells[rowKy - 1, colDenNgay - 1].Text.Trim();
			string text2 = sView.Cells[rowHD - 1, colDenNgay - 1].Text.Trim();
			string text3 = sView.Cells[rowKy - 1, colTemVatNam - 1].Text.Trim();
			string text4 = sView.Cells[rowKy, colTemVatNam - 1].Text.Trim();
			string text5 = sView.Cells[rowKy - 1, colTemTuSinh - 1].Text.Trim();
			int num = 0;
			if (text == text2 && text5 == "1" && text3 == "1" && text4 != "1")
			{
				string text6 = sView.Cells[rowKy, colMaHD - 1].Text.Trim();
				if (text6 != MA_KY_THANH_TOAN && text6 != MA_KY_THANH_TOAN_CUOI_CUNG)
				{
					num = rowKy;
					text3 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
					while (text3 == "1")
					{
						sView.Cells[num - 1, colMaHD - 1].Text = MA_KY_THANH_TOAN;
						sView.Cells[num - 1, colTenHD - 1].Text = "";
						num--;
						text3 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
					}
					sView.Cells[num - 3, colMaHD - 1].Text = MA_KY_THANH_TOAN_CUOI_CUNG;
					sView.Cells[num - 3, colTenHD - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN_CUOI_CUNG];
				}
			}
			num = rowKy;
			text3 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
			while (text3 == "1")
			{
				sView.Cells[num - 1, colTenHD - 1].Text = "";
				num--;
				text3 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckChiTieu23BoTichDongHD(SheetView sView, List<int> hopDong, int colCaNhanPhaiNopThue)
	{
		try
		{
			if (hopDong.Count == 0)
			{
				return;
			}
			int num = hopDong[0];
			int num2 = 0;
			for (int i = 1; i < hopDong.Count; i++)
			{
				int num3 = hopDong[i];
				if (Convert.ToBoolean(sView.Cells[num3 - 1, colCaNhanPhaiNopThue - 1].Value))
				{
					num2++;
				}
			}
			if (num2 == hopDong.Count - 1)
			{
				sView.Cells[num - 1, colCaNhanPhaiNopThue - 1].Value = true;
			}
			else
			{
				sView.Cells[num - 1, colCaNhanPhaiNopThue - 1].Value = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckDongConCaNhanPhaiNopThue(SheetView sView, int colSTT, int row, int colVatNam, int colTuNgay, int colDenNgay, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			string text = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			string text2 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
			{
				string[] array = sView.Cells[row - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array2 = sView.Cells[row, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array3 = sView.Cells[row + 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string text3 = sView.Cells[row, colVatNam - 1].Text.Trim();
				string text4 = sView.Cells[row + 1, colVatNam - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && splitDay(text)[2] == splitDay(text2)[2] && text3 == "1" && text4 == "1" && array[0] == array2[0] && array[0] == array3[0])
				{
					GeneralFunctions.RemoveDataRows(sView, 0, row, 2, sInfo);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Neu1Trong2DongVatNamTichThiDongTrenAnTheo(SheetView sView, int colSTT, int colTemVatNam, int rowKy, int colCaNhanThuocDienPhaiNopThue, int colTuNgay, int colDenNgay)
	{
		try
		{
			string[] array = sView.Cells[rowKy - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
			string[] array2 = sView.Cells[rowKy, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
			string[] array3 = sView.Cells[rowKy + 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
			string text = sView.Cells[rowKy, colTemVatNam - 1].Text.Trim();
			string text2 = sView.Cells[rowKy + 1, colTemVatNam - 1].Text.Trim();
			string text3 = sView.Cells[rowKy - 1, colTuNgay - 1].Text.Trim();
			string text4 = sView.Cells[rowKy - 1, colDenNgay - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(text4) && splitDay(text3)[2] != splitDay(text4)[2] && text == "1" && text2 == "1" && array[0] == array2[0] && array[0] == array3[0])
			{
				if (Convert.ToBoolean(sView.Cells[rowKy, colCaNhanThuocDienPhaiNopThue - 1].Value) && Convert.ToBoolean(sView.Cells[rowKy + 1, colCaNhanThuocDienPhaiNopThue - 1].Value))
				{
					sView.Cells[rowKy - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
				}
				else
				{
					sView.Cells[rowKy - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = false;
				}
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
			if (winFpSpread.ActiveSheetIndex == 2)
			{
				_nhapTuExcel = true;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
				SheetView sheetView = winFpSpread.Sheets[2];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				int count = hTKKCMCTableInfo.RowsValues.Count;
				int num = HTKKFile.ColumnLettersToIndex("B");
				int num2 = HTKKFile.ColumnLettersToIndex("F");
				int colMaHD = HTKKFile.ColumnLettersToIndex("E");
				int num3 = HTKKFile.ColumnLettersToIndex("D");
				int num4 = HTKKFile.ColumnLettersToIndex("AU");
				int num5 = HTKKFile.ColumnLettersToIndex("AI");
				int num6 = HTKKFile.ColumnLettersToIndex("AJ");
				int num7 = HTKKFile.ColumnLettersToIndex("AK");
				int num8 = HTKKFile.ColumnLettersToIndex("AE");
				int num9 = HTKKFile.ColumnLettersToIndex("AL");
				int num10 = HTKKFile.ColumnLettersToIndex("AP");
				int colTuNgay = HTKKFile.ColumnLettersToIndex("AB");
				int colDenNgay = HTKKFile.ColumnLettersToIndex("AC");
				int colTongGiaTriHopDong = HTKKFile.ColumnLettersToIndex("AM");
				int colDoanhThuTinhTheTrongKy = HTKKFile.ColumnLettersToIndex("AQ");
				string pathCt23ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct23";
				string pathCt24ToKhai = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKeKhai/ct24";
				HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
				HTKKCMCCellInfo rowCellCt23ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt23ToKhai));
				TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, colMaHD);
				Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
				Dictionary<string, bool> dictionary2 = new Dictionary<string, bool>();
				for (int num11 = count - 1; num11 >= 0; num11--)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + num11 - 1, num - 1].Text.Trim();
					if (text.IndexOf(".") < 0 && !dictionary2.ContainsKey(text))
					{
						dictionary2.Add(text, Convert.ToBoolean(sheetView.Cells[dataLocationTemplateRowStart + num11 - 1, num10 - 1].Value));
					}
					CellChange_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, dataLocationTemplateRowStart + num11, num2);
					CellChange_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, dataLocationTemplateRowStart + num11, num4);
				}
				dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
				if (dictionary.Count > 0)
				{
					foreach (KeyValuePair<string, List<int>> item in dictionary)
					{
						if (!dictionary.ContainsKey(item.Key))
						{
							continue;
						}
						int num12 = dictionary[item.Key][0];
						string text2 = sheetView.Cells[num12 - 1, num9 - 1].Text.Trim();
						if (text2 == "0")
						{
							foreach (int item2 in dictionary[item.Key])
							{
								if (sheetView.Cells[item2 - 1, num5 - 1].Text.Trim() == "")
								{
									sheetView.Cells[item2 - 1, num5 - 1].Value = 0;
								}
								if (sheetView.Cells[item2 - 1, num6 - 1].Text.Trim() == "")
								{
									sheetView.Cells[item2 - 1, num6 - 1].Value = 0;
								}
								if (sheetView.Cells[item2 - 1, num7 - 1].Text.Trim() == "")
								{
									sheetView.Cells[item2 - 1, num7 - 1].Value = 0;
								}
								CellChange_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, item2, num8);
							}
						}
						string text = sheetView.Cells[num12 - 1, num - 1].Text.Trim();
						if (text.IndexOf(".") < 0 && dictionary2.ContainsKey(text) && !Convert.ToBoolean(sheetView.Cells[num12 - 1, num10 - 1].Value))
						{
							sheetView.Cells[num12 - 1, num10 - 1].Value = dictionary2[text];
						}
						CellChange_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, num12, num10);
					}
				}
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt23ToKhai, dictionary, num5, colTongGiaTriHopDong, colTuNgay, colDenNgay, colMaHD);
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt24ToKhai, dictionary, num5, colDoanhThuTinhTheTrongKy, colTuNgay, colDenNgay, colMaHD);
			}
			if (winFpSpread.ActiveSheetIndex == 3 || winFpSpread.ActiveSheetIndex == 4 || winFpSpread.ActiveSheetIndex == 5)
			{
				TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
				TinhChiTieu26_TK(winFpSpread);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			if (idxRow == 7 && idxCol == 1)
			{
				GeneralFunctions.ShowBangKe("BangKe_01_2_TTS_TT40.xls");
			}
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int num = HTKKFile.ColumnLettersToIndex("AI");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("B");
			int colMst = HTKKFile.ColumnLettersToIndex("J");
			int colTuNgay = HTKKFile.ColumnLettersToIndex("AB");
			int colDenNgay = HTKKFile.ColumnLettersToIndex("AC");
			int colTongGiaTriHopDong = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AO");
			int num5 = HTKKFile.ColumnLettersToIndex("AP");
			int colDoanhThuTinhThueTrongKy = HTKKFile.ColumnLettersToIndex("AQ");
			string value = "";
			if (string.IsNullOrEmpty(value))
			{
				value = sheetView.Cells[(int)idxRow, num2 - 1].Text.Trim();
			}
			string text = sheetView.Cells[(int)idxRow, num3 - 1].Text.Trim();
			string[] array = text.Split(new char[1] { '.' });
			string key = array[0];
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num3);
			if (dictionary.ContainsKey(key))
			{
				int num6 = dictionary[key][0];
				if (num6 == (int)idxRow + 1 && idxCol + 1 == num5)
				{
					CaNhanPhaiNopThueNeuTichHD(sheetView, dictionary[key], num2, num3, num5, num4, num, colTuNgay, colDenNgay);
				}
				if (idxCol + 1 == num5)
				{
					CheckDongConCaNhanPhaiNopThue(sheetView, num3, (int)idxRow + 1, num5, num, colTuNgay, colDenNgay, num4);
				}
				Dictionary<string, List<int>> dictionary2 = new Dictionary<string, List<int>>();
				HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables[0];
				if (idxCol + 1 == num5)
				{
					dictionary2 = TimHopDongCungMSTCungNam(sheetView, table, (int)idxRow + 1, num3, colMst, dictionary);
				}
				if (dictionary2.Count > 0 && idxCol + 1 == num5)
				{
					DoanhThuTinhTheTrongKy(sheetView, dictionary2, num2, num3, num5, colDoanhThuTinhThueTrongKy, colTongGiaTriHopDong, colTuNgay, colDenNgay);
				}
				if (idxCol + 1 == num5)
				{
					CheckChiTieu23BoTichDongHD(sheetView, dictionary[key], num5);
				}
				GeneralFunctions.SetRowHeight(sheetView, (int)idxRow + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void TinhChiTieu25_TK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("U_212");
			int[] array2 = GeneralFunctions.ArrayCellId("AG_212");
			int[] array3 = GeneralFunctions.ArrayCellId("AH_212");
			int[] array4 = GeneralFunctions.ArrayCellId("AJ_212");
			int[] array5 = GeneralFunctions.ArrayCellId("AI_212");
			SheetView sheetView = winFpSpread.Sheets[0];
			if (ToKhaiInfo.SheetInfos[3] != null && ToKhaiInfo.SheetInfos[4] == null && ToKhaiInfo.SheetInfos[5] == null)
			{
				sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
			}
			else if (ToKhaiInfo.SheetInfos[3] == null && ToKhaiInfo.SheetInfos[4] != null && ToKhaiInfo.SheetInfos[5] == null)
			{
				sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array3[1], array3[0]].Value;
			}
			else if (ToKhaiInfo.SheetInfos[3] == null && ToKhaiInfo.SheetInfos[4] == null && ToKhaiInfo.SheetInfos[5] != null)
			{
				sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array4[1], array4[0]].Value;
			}
			else
			{
				sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array5[1], array5[0]].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhChiTieu26_TK(FpSpread winFpSpread)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("U_213");
			int[] array2 = GeneralFunctions.ArrayCellId("AG_213");
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TinhChiTieu25_TK(winFpSpread, _ToKhaiInfo);
		TinhChiTieu26_TK(winFpSpread);
		return true;
	}
}
