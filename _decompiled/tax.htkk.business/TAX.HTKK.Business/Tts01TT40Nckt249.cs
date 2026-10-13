using System;
using System.Collections.Generic;
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

public class Tts01TT40Nckt249
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

	private long _iRowClick = -1L;

	private string MA_HOP_DONG = "01";

	private string MA_KY_THANH_TOAN = "02";

	private string MA_KY_THANH_TOAN_CUOI_CUNG = "03";

	private string MA_TRA_TIEN_CHO_CA_HD = "04";

	private string MA_HOP_DONG_KHAI_LAN_DAU = "01";

	private string MA_COLOR_LOCKED = "#e8f5ff";

	private string MA_COLOR_HOP_DONG = "#F3F3F3";

	private string MA_COLOR_UNLOCK = "#FFFFFF";

	private string HOP_DONG_KHAI_LAN_DAU = "01";

	private int[] khaiTheoPLuatDanSu = GeneralFunctions.ArrayCellId("Q_10");

	private int[] khaiTheoPLuatDanThue = GeneralFunctions.ArrayCellId("Q_12");

	private XMLMessage _xmlMessage = new XMLMessage("../Project/Message.xml");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_PeriodInfo = PeriodInfo;
			_HeaderInfo = HeaderInfo;
			_ToKhaiInfo = ToKhaiInfo;
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
				int num7 = HTKKFile.ColumnLettersToIndex("AD");
				int item13 = HTKKFile.ColumnLettersToIndex("AE");
				int item14 = HTKKFile.ColumnLettersToIndex("AI");
				int item15 = HTKKFile.ColumnLettersToIndex("AJ");
				int num8 = HTKKFile.ColumnLettersToIndex("AK");
				int colTongDoanhThuPhatSinhTrongNamTungCN = HTKKFile.ColumnLettersToIndex("AL");
				int num9 = HTKKFile.ColumnLettersToIndex("AM");
				int item16 = HTKKFile.ColumnLettersToIndex("AQ");
				int num10 = HTKKFile.ColumnLettersToIndex("AR");
				int num11 = HTKKFile.ColumnLettersToIndex("AS");
				int item17 = HTKKFile.ColumnLettersToIndex("AU");
				int num12 = HTKKFile.ColumnLettersToIndex("AT");
				int item18 = HTKKFile.ColumnLettersToIndex("AV");
				int colTemVatNam = HTKKFile.ColumnLettersToIndex("AF");
				int num13 = HTKKFile.ColumnLettersToIndex("AH");
				int num14 = HTKKFile.ColumnLettersToIndex("AG");
				int num15 = HTKKFile.ColumnLettersToIndex("BB");
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
				list.Add(item14);
				List<int> list2 = new List<int>();
				list2.Add(item11);
				list2.Add(item17);
				list2.Add(item18);
				List<int> list3 = new List<int>();
				list3.Add(item15);
				List<int> list4 = new List<int>();
				list4.Add(item14);
				list4.Add(num9);
				list4.Add(item16);
				List<int> list5 = new List<int>();
				list5.Add(item12);
				List<int> list6 = new List<int>();
				list6.Add(num5);
				list6.Add(num6);
				list6.Add(item13);
				List<int> list7 = new List<int>();
				list7.Add(num2);
				List<int> list8 = new List<int>();
				list8.Add(item15);
				LockUnLockDongHopDongHoacKy(sView, hTKKCMCSheetInfo, num, colCt06aMa, colTemDongTuSinh, colTemVatNam, list, list2, list5, list6, list3, list7, list4, num5, num6, list8, colTongDoanhThuPhatSinhTrongNamTungCN, num9);
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
				int num7 = HTKKFile.ColumnLettersToIndex("AD");
				int item15 = HTKKFile.ColumnLettersToIndex("AE");
				int item16 = HTKKFile.ColumnLettersToIndex("AL");
				int item17 = HTKKFile.ColumnLettersToIndex("AM");
				int num8 = HTKKFile.ColumnLettersToIndex("AN");
				int colTongDoanhThuPhatSinhTrongNamTungCN = HTKKFile.ColumnLettersToIndex("AO");
				int num9 = HTKKFile.ColumnLettersToIndex("AP");
				int item18 = HTKKFile.ColumnLettersToIndex("AQ");
				int num10 = HTKKFile.ColumnLettersToIndex("AR");
				int num11 = HTKKFile.ColumnLettersToIndex("AS");
				int item19 = HTKKFile.ColumnLettersToIndex("AU");
				int num12 = HTKKFile.ColumnLettersToIndex("AT");
				int item20 = HTKKFile.ColumnLettersToIndex("AV");
				int colTemVatNam = HTKKFile.ColumnLettersToIndex("AI");
				int num13 = HTKKFile.ColumnLettersToIndex("AK");
				int num14 = HTKKFile.ColumnLettersToIndex("AJ");
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
				list.Add(item16);
				List<int> list2 = new List<int>();
				list2.Add(item13);
				list2.Add(item19);
				list2.Add(item20);
				List<int> list3 = new List<int>();
				list3.Add(item17);
				List<int> list4 = new List<int>();
				list4.Add(item16);
				list4.Add(num9);
				list4.Add(item18);
				List<int> list5 = new List<int>();
				list5.Add(num5);
				list5.Add(num6);
				list5.Add(item15);
				List<int> list6 = new List<int>();
				list6.Add(num2);
				List<int> list7 = new List<int>();
				list7.Add(item14);
				List<int> list8 = new List<int>();
				list8.Add(item17);
				LockUnLockDongHopDongHoacKy(sView, hTKKCMCSheetInfo, num, colCt06aMa, colTemDongTuSinh, colTemVatNam, list, list2, list7, list5, list3, list6, list4, num5, num6, list8, colTongDoanhThuPhatSinhTrongNamTungCN, num9);
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sView, hTKKCMCSheetInfo, num, colCt06aMa, num2, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
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
			CheckChonLoaiNopThue(winFpSpread, _ToKhaiInfo, _PeriodInfo);
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
		int[] array = GeneralFunctions.ArrayCellId("F_42");
		int[] array2 = GeneralFunctions.ArrayCellId("I_186");
		if (iRow == khaiTheoPLuatDanSu[1] + 1 || iRow == khaiTheoPLuatDanThue[1] + 1)
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value);
			if (iRow == khaiTheoPLuatDanSu[1] + 1 && flag)
			{
				sheetView.Cells[khaiTheoPLuatDanThue[1], khaiTheoPLuatDanThue[0]].Value = false;
				flag2 = false;
			}
			if (iRow == khaiTheoPLuatDanThue[1] + 1 && flag2)
			{
				sheetView.Cells[khaiTheoPLuatDanSu[1], khaiTheoPLuatDanSu[0]].Value = false;
				flag = false;
			}
			string text = sheetView.Cells[array[1], array[0]].Text;
			string text2 = sheetView.Cells[array2[1], array2[0]].Text;
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
			int colSoThueGTGTPhaiNop = HTKKFile.ColumnLettersToIndex("AR");
			int num19 = HTKKFile.ColumnLettersToIndex("AS");
			int num20 = HTKKFile.ColumnLettersToIndex("AU");
			int num21 = HTKKFile.ColumnLettersToIndex("AT");
			int num22 = HTKKFile.ColumnLettersToIndex("AV");
			int num23 = HTKKFile.ColumnLettersToIndex("AF");
			int colTemSoThang = HTKKFile.ColumnLettersToIndex("AH");
			int colTemTronThang = HTKKFile.ColumnLettersToIndex("AG");
			int num24 = HTKKFile.ColumnLettersToIndex("BB");
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
			list2.Add(num20);
			list2.Add(num22);
			List<int> list3 = new List<int>();
			list3.Add(num15);
			List<int> list4 = new List<int>();
			list4.Add(num14);
			list4.Add(num17);
			list4.Add(num18);
			List<int> list5 = new List<int>();
			list5.Add(num9);
			List<int> list6 = new List<int>();
			list6.Add(num10);
			list6.Add(num11);
			list6.Add(num13);
			List<int> list7 = new List<int>();
			list7.Add(num4);
			List<int> list8 = new List<int>();
			list8.Add(num15);
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
			if (iCol == num20)
			{
				string[] array2 = dicDMTTHopDong[sheetView.Cells[(int)iRow - 1, num20 - 1].Text.Trim()];
				if (array2.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = array2[0];
					sheetView.Cells[(int)iRow - 1, num20 - 1].Text = array2[1];
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
						CheckNeuTichChonHDKyThanhToan1Lan(sheetView, hTKKCMCSheetInfo, PeriodInfo, text, (int)iRow, num10, num11, num12, num23);
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
			int num25 = dictionary[key][0];
			string text3 = sheetView.Cells[(int)iRow - 1, num23 - 1].Text.Trim();
			if (iCol == num4)
			{
				if (PeriodInfo.KieuKy == "D_D" && dictionary.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num25 - 1, num4 - 1, msg);
					return;
				}
				HTKKFile.RemoveValidate(sheetView, num25 - 1, num4 - 1);
			}
			if ((iCol == num10 || iCol == num11) && !CheckDenNgayNhoHonTuNgay(sheetView, num25, (int)iRow, num10, num11, errorDenNgayPhaiLonHonTuNgay, errorTuNgayKyLonHonTuNgayHD, errorDenNgayKyNhoHonDenNgayHD, tuDongXoaThongBao: true, num23))
			{
				return;
			}
			if (text != MA_HOP_DONG && (iCol == num10 || iCol == num11))
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num23, num10, num11, hTKKCMCSheetInfo);
			}
			if (num25 == (int)iRow && iCol == num17)
			{
				CaNhanPhaiNopThueNeuTichHD(sheetView, dictionary[key], num3, num, num17, num16, num23, num10, num11);
			}
			if (iCol == num4 || iCol == num17 || iCol == num10 || iCol == num11 || iCol == num15 || iCol == num16)
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num17, num23, num10, num11, num16);
			}
			if (num25 == (int)iRow && (iCol == num20 || iCol == num4 || iCol == num10 || iCol == num11))
			{
				TinhMaHopDong(sheetView, TTinChung, num10, num3, num, num21, num22, (int)iRow, ToKhaiInfo.SheetInfos[1].Tables[0]);
			}
			if (PeriodInfo.KieuKy == "D_D" && (iCol == num4 || iCol == num20 || iCol == num22))
			{
				string text4 = sheetView.Cells[num25 - 1, num22 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text4))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaHDToKhai));
					winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text4;
				}
			}
			int num26 = 0;
			if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13)
			{
				num26 = TinhKyThanhToanOnDinh(sheetView, hTKKCMCSheetInfo, PeriodInfo, num9, (int)iRow, num3, num4, num10, num11, num12, num13, num, num23, list, list2, num2);
				if (num26 == 0)
				{
					num26 = TuDongTachNeuVatNam(sheetView, hTKKCMCSheetInfo, (int)iRow, num3, num4, num10, num11, num13, num, num23, list, list2, num2, num9, num25);
				}
				if (num26 > 0)
				{
					TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
					dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
				}
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol < num15)))
			{
				num25 = dictionary[key][0];
				if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13)
				{
					sheetView.Cells[num25 - 1, num14 - 1].Value = 0;
				}
				sheetView.Cells[num25 - 1, num15 - 1].Value = 0;
				for (int num27 = 1; num27 < dictionary[key].Count; num27++)
				{
					int num28 = dictionary[key][num27];
					string text5 = sheetView.Cells[num28 - 1, num2 - 1].Text.Trim();
					AutoShowYear(sheetView, num28, num3, num10, num11, num12);
					bool tuDongSoThang = ((iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num12 || iCol == num13) ? true : false);
					TinhTongGiaTriHD(sheetView, num25, num28, num3, num10, num11, num13, num23, num15, colTemSoThang, colTemTronThang, num14, tuDongSoThang);
				}
			}
			if (text != MA_HOP_DONG && iCol == num15)
			{
				SumTinhTongGiaTriHD(sheetView, dictionary[key], num15, num10, num11);
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
				CaNhanPhaiNopThue(sheetView, dictionary, PeriodInfo, num3, colDoanhThuBinhQuanThang, num16, num, num17, num10, num11, num23);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol < num18)))
			{
				DoanhThuTinhTheTrongKy(sheetView, dictionary, num3, num, num17, num18, num15, num10, num11);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol <= num19)))
			{
				foreach (KeyValuePair<string, List<int>> item10 in dictionary)
				{
					foreach (int item11 in item10.Value)
					{
						SoThueGTGTPhaiNop(sheetView, item11, num18, colSoThueGTGTPhaiNop, num);
						SoThueTNCNPhaiNop(sheetView, item11, num18, num19, num);
					}
				}
			}
			if (text != MA_HOP_DONG && iCol == num18)
			{
				SumDoanhThuTinhTheTrongKy(sheetView, dictionary[key], num18, num10, num11);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol <= num15)))
			{
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt23ToKhai, dictionary, num23, num15, num10, num11, num3);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num9 && iCol <= num18)))
			{
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt24ToKhai, dictionary, num23, num18, num10, num11, num3);
			}
			if (iCol == num4)
			{
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sheetView, hTKKCMCSheetInfo, num, num3, num4, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
			if (iCol == num4 && !CheckChonLoaiNopThue(winFpSpread, ToKhaiInfo, PeriodInfo))
			{
				return;
			}
			if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num13 || iCol == num14 || iCol == num15 || iCol == num16)
			{
				LockUnLockDongHopDongHoacKy(sheetView, hTKKCMCSheetInfo, num, num3, num2, num23, list, list2, list5, list6, list3, list7, list4, num10, num11, list8, num16, num17);
			}
			if (iCol == num4 || iCol == num9 || iCol == num10 || iCol == num11 || iCol == num13 || iCol == num20 || iCol == num22)
			{
				string text6 = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text6))
				{
					foreach (KeyValuePair<string, List<int>> item12 in dictionary)
					{
						string valueCopy = sheetView.Cells[item12.Value[0] - 1, num21 - 1].Text.Trim();
						string valueCopy2 = sheetView.Cells[item12.Value[0] - 1, num20 - 1].Text.Trim();
						string valueCopy3 = sheetView.Cells[item12.Value[0] - 1, num22 - 1].Text.Trim();
						Set1RowHDToAllRowKy(sheetView, item12, num, num21, valueCopy, text6.Substring(0, 1));
						Set1RowHDToAllRowKy(sheetView, item12, num, num20, valueCopy2, text6.Substring(0, 1));
						Set1RowHDToAllRowKy(sheetView, item12, num, num22, valueCopy3, text6.Substring(0, 1));
					}
				}
			}
			if (num26 > 0 || iCol == num17)
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
			int num8 = HTKKFile.ColumnLettersToIndex("W");
			int num9 = HTKKFile.ColumnLettersToIndex("Y");
			int num10 = HTKKFile.ColumnLettersToIndex("X");
			int item9 = HTKKFile.ColumnLettersToIndex("Z");
			int num11 = HTKKFile.ColumnLettersToIndex("AA");
			int num12 = HTKKFile.ColumnLettersToIndex("AB");
			int num13 = HTKKFile.ColumnLettersToIndex("AC");
			int num14 = HTKKFile.ColumnLettersToIndex("AD");
			int num15 = HTKKFile.ColumnLettersToIndex("AE");
			int num16 = HTKKFile.ColumnLettersToIndex("AL");
			int num17 = HTKKFile.ColumnLettersToIndex("AM");
			int num18 = HTKKFile.ColumnLettersToIndex("AN");
			int num19 = HTKKFile.ColumnLettersToIndex("AO");
			int num20 = HTKKFile.ColumnLettersToIndex("AP");
			int num21 = HTKKFile.ColumnLettersToIndex("AQ");
			int colSoThueGTGTPhaiNop = HTKKFile.ColumnLettersToIndex("AR");
			int num22 = HTKKFile.ColumnLettersToIndex("AS");
			int num23 = HTKKFile.ColumnLettersToIndex("AU");
			int num24 = HTKKFile.ColumnLettersToIndex("AT");
			int num25 = HTKKFile.ColumnLettersToIndex("AV");
			int num26 = HTKKFile.ColumnLettersToIndex("AI");
			int colTemSoThang = HTKKFile.ColumnLettersToIndex("AK");
			int colTemTronThang = HTKKFile.ColumnLettersToIndex("AJ");
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
			list.Add(num8);
			list.Add(num9);
			list.Add(num16);
			List<int> list2 = new List<int>();
			list2.Add(item9);
			list2.Add(num23);
			list2.Add(num25);
			List<int> list3 = new List<int>();
			list3.Add(num17);
			List<int> list4 = new List<int>();
			list4.Add(num16);
			list4.Add(num20);
			list4.Add(num21);
			List<int> list5 = new List<int>();
			list5.Add(num12);
			list5.Add(num13);
			list5.Add(num15);
			List<int> list6 = new List<int>();
			list6.Add(num4);
			List<int> list7 = new List<int>();
			list7.Add(num11);
			List<int> list8 = new List<int>();
			list8.Add(num17);
			HTKKCMCCellInfo rowCellCt24ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt24ToKhai));
			HTKKCMCCellInfo rowCellCt23ToKhai = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt23ToKhai));
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
				string[] array2 = dicDMTTHopDong[sheetView.Cells[(int)iRow - 1, num23 - 1].Text.Trim()];
				if (array2.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num24 - 1].Text = array2[0];
					sheetView.Cells[(int)iRow - 1, num23 - 1].Text = array2[1];
				}
			}
			if (iCol == num9)
			{
				string tenDonVi = sheetView.Cells[(int)iRow - 1, num9 - 1].Text.Trim();
				KeyValuePair<string, string> keyValuePair = dicDMPhuongXaTheoCQTTmp.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == tenDonVi);
				sheetView.Cells[(int)iRow - 1, num10 - 1].Text = keyValuePair.Key;
				sheetView.Cells[(int)iRow - 1, num9 - 1].Text = keyValuePair.Value;
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
						CheckNeuTichChonHDKyThanhToan1Lan(sheetView, hTKKCMCSheetInfo, PeriodInfo, text, (int)iRow, num12, num13, num14, num26);
						TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
						CheckChonHopDongTrungKy(sheetView, hTKKCMCSheetInfo, num, num3, num4, errorMessager, tuDongXoaThongBao: true);
						CheckChonHopDongDauTien(sheetView, hTKKCMCSheetInfo, num, num3, num4, errorMessager2, tuDongXoaThongBao: true);
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
			if ((iCol == num12 || iCol == num13) && !CheckDenNgayNhoHonTuNgay(sheetView, num27, (int)iRow, num12, num13, errorDenNgayPhaiLonHonTuNgay, errorTuNgayKyLonHonTuNgayHD, errorDenNgayKyNhoHonDenNgayHD, tuDongXoaThongBao: true, num26))
			{
				return;
			}
			if (text != MA_HOP_DONG && (iCol == num12 || iCol == num13))
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num26, num12, num13, hTKKCMCSheetInfo);
			}
			if (num27 == (int)iRow && iCol == num20)
			{
				CaNhanPhaiNopThueNeuTichHD(sheetView, dictionary[key], num3, num, num20, num19, num26, num12, num13);
			}
			if (iCol == num4 || iCol == num20 || iCol == num12 || iCol == num13 || iCol == num17 || iCol == num19)
			{
				CheckDongConCaNhanPhaiNopThue(sheetView, num, (int)iRow, num20, num26, num12, num13, num19);
			}
			string text3 = sheetView.Cells[(int)iRow - 1, num26 - 1].Text.Trim();
			if (iCol == num4)
			{
				if (PeriodInfo.KieuKy == "D_D" && dictionary.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num27 - 1, num4 - 1, msg);
					return;
				}
				HTKKFile.RemoveValidate(sheetView, num27 - 1, num4 - 1);
			}
			if (num27 == (int)iRow && (iCol == num23 || iCol == num4 || iCol == num12))
			{
				TinhMaHopDong(sheetView, TTinChung, num12, num3, num, num24, num25, (int)iRow, ToKhaiInfo.SheetInfos[2].Tables[0]);
			}
			if (PeriodInfo.KieuKy == "D_D" && (iCol == num4 || iCol == num23 || iCol == num25))
			{
				string text4 = sheetView.Cells[num27 - 1, num25 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text4))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMaHDToKhai));
					winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text4;
				}
			}
			int num28 = 0;
			if (iCol == num4 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14 || iCol == num15)
			{
				num28 = TinhKyThanhToanOnDinh(sheetView, hTKKCMCSheetInfo, PeriodInfo, num11, (int)iRow, num3, num4, num12, num13, num14, num15, num, num26, list, list2, num2);
				if (num28 == 0)
				{
					num28 = TuDongTachNeuVatNam(sheetView, hTKKCMCSheetInfo, (int)iRow, num3, num4, num12, num13, num15, num, num26, list, list2, num2, num11, num27);
				}
				if (num28 > 0)
				{
					TinhLaiSTT(sheetView, hTKKCMCSheetInfo, num, num3);
					dictionary = LayCacHangCungHopDong(sheetView, hTKKCMCSheetInfo, num);
				}
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num11 && iCol < num17)))
			{
				num27 = dictionary[key][0];
				if (iCol == num4 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14 || iCol == num15)
				{
					sheetView.Cells[num27 - 1, num16 - 1].Value = 0;
				}
				sheetView.Cells[num27 - 1, num17 - 1].Value = 0;
				for (int num29 = 1; num29 < dictionary[key].Count; num29++)
				{
					int num30 = dictionary[key][num29];
					string text5 = sheetView.Cells[num30 - 1, num2 - 1].Text.Trim();
					AutoShowYear(sheetView, num30, num3, num12, num13, num14);
					bool tuDongSoThang = ((iCol == num4 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num14 || iCol == num15) ? true : false);
					TinhTongGiaTriHD(sheetView, num27, num30, num3, num12, num13, num15, num26, num17, colTemSoThang, colTemTronThang, num16, tuDongSoThang);
				}
			}
			if (text != MA_HOP_DONG && iCol == num17)
			{
				SumTinhTongGiaTriHD(sheetView, dictionary[key], num17, num12, num13);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num11 && iCol < num18)))
			{
				DoanhThuBinhQuanThangCuaHD(sheetView, num17, num18, num3, num16, dictionary[key]);
			}
			if (iCol == num4 || iCol == num17 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num15 || iCol == num16)
			{
				TongDoanhThuPhatSinhTrongNam(sheetView, dictionary, num14, num17, num19, num, num5);
			}
			if (iCol == num4 || iCol == num17 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num15 || iCol == num16 || iCol == num19)
			{
				CaNhanPhaiNopThue(sheetView, dictionary, PeriodInfo, num3, num18, num19, num, num20, num12, num13, num26);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num11 && iCol < num21)))
			{
				DoanhThuTinhTheTrongKy(sheetView, dictionary, num3, num, num20, num21, num17, num12, num13);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num8 && iCol <= num22)))
			{
				foreach (KeyValuePair<string, List<int>> item10 in dictionary)
				{
					foreach (int item11 in dictionary[key])
					{
						SoThueGTGTPhaiNop(sheetView, item11, num21, colSoThueGTGTPhaiNop, num);
						SoThueTNCNPhaiNop(sheetView, item11, num21, num22, num);
					}
				}
			}
			if (text != MA_HOP_DONG && iCol == num21)
			{
				SumDoanhThuTinhTheTrongKy(sheetView, dictionary[key], num21, num12, num13);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num11 && iCol <= num17)))
			{
				TongDoanhThuPhatSinhTrongKyCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt23ToKhai, dictionary, num26, num17, num12, num13, num3);
			}
			if (dictionary.Count > 0 && (iCol == num4 || (iCol >= num11 && iCol <= num21)))
			{
				TongDoanhThuTinhThueCuaToKhai(winFpSpread.Sheets[0], sheetView, PeriodInfo, rowCellCt24ToKhai, dictionary, num26, num21, num12, num13, num3);
			}
			if (iCol == num4)
			{
				string duoiHopDongLaKyTinhThue = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
				Check2HopDongKhongDuocLienTiep(sheetView, hTKKCMCSheetInfo, num, num3, num4, duoiHopDongLaKyTinhThue, tuDongXoaThongBao: true);
			}
			if (iCol == num4 && !CheckChonLoaiNopThue(winFpSpread, ToKhaiInfo, PeriodInfo))
			{
				return;
			}
			if (iCol == num4 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num15 || iCol == num16 || iCol == num17 || iCol == num19)
			{
				LockUnLockDongHopDongHoacKy(sheetView, hTKKCMCSheetInfo, num, num3, num2, num26, list, list2, list7, list5, list3, list6, list4, num12, num13, list8, num19, num20);
			}
			if (iCol == num4 || iCol == num11 || iCol == num12 || iCol == num13 || iCol == num15 || iCol == num23 || iCol == num25)
			{
				string text6 = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text6))
				{
					foreach (KeyValuePair<string, List<int>> item12 in dictionary)
					{
						string valueCopy = sheetView.Cells[item12.Value[0] - 1, num24 - 1].Text.Trim();
						string valueCopy2 = sheetView.Cells[item12.Value[0] - 1, num23 - 1].Text.Trim();
						string valueCopy3 = sheetView.Cells[item12.Value[0] - 1, num25 - 1].Text.Trim();
						Set1RowHDToAllRowKy(sheetView, item12, num, num24, valueCopy, text6.Substring(0, 1));
						Set1RowHDToAllRowKy(sheetView, item12, num, num23, valueCopy2, text6.Substring(0, 1));
						Set1RowHDToAllRowKy(sheetView, item12, num, num25, valueCopy3, text6.Substring(0, 1));
					}
				}
			}
			if (num28 > 0 || iCol == num20)
			{
				CheckChiTieu23BoTichDongHD(sheetView, dictionary[key], num20);
			}
			GeneralFunctions.SetRowHeight(sheetView, (int)iRow);
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
						Set1RowHDToAllRowKy(sheetView, item, num, num2, "1", text.Substring(0, 1));
					}
					if (item.Value.Count <= 1)
					{
						continue;
					}
					for (int i = 1; i < item.Value.Count; i++)
					{
						int num15 = item.Value[i];
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
						Set1RowHDToAllRowKy(sheetView2, item2, num, num2, "1", text.Substring(0, 1));
					}
					if (item2.Value.Count <= 1)
					{
						continue;
					}
					for (int i = 1; i < item2.Value.Count; i++)
					{
						int num15 = item2.Value[i];
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
				sView.Cells[row - 1, colNam - 1].Text = splitDay(PeriodInfo.TuNgay)[2];
				return;
			}
			for (int i = 0; i <= num; i++)
			{
				int num2 = Convert.ToInt32(splitDay(PeriodInfo.TuNgay)[2]);
				string text = firstDayOfYear(num2 + i).ToString("dd/MM/yyyy");
				string text2 = lastDayOfYear(num2 + i).ToString("dd/MM/yyyy");
				sView.Cells[row + i - 1, colTuNgay - 1].Text = ((i == 0) ? PeriodInfo.TuNgay : text);
				sView.Cells[row + i - 1, colDenNgay - 1].Text = ((splitDay(text2)[2] == splitDay(PeriodInfo.DenNgay)[2]) ? PeriodInfo.DenNgay : text2);
				sView.Cells[row + i - 1, colNam - 1].Text = splitDay(text2)[2];
				sView.Cells[row + i - 1, colTemVatNam - 1].Text = "1";
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
		return Convert.ToInt32(splitDay(denNgay)[2]) - Convert.ToInt32(splitDay(tuNgay)[2]);
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

	public int TinhKyThanhToanOnDinh(SheetView sView, HTKKCMCSheetInfo sheet, HTKKPeriodInfo PeriodInfo, int colKyThanhToanOnDinh, int row, int colMaHopDongKy, int colTenHopDongKy, int colTuNgay, int colDenNgay, int colNam, int colGiaChoThue1Thang, int colSTT, int colTemVatNam, List<int> colAnKyXoaData, List<int> colAnKyKhongXoa, int colTemDongTuSinh)
	{
		int num = 0;
		try
		{
			int num2 = Convert.ToInt32(sView.Cells[row - 1, colKyThanhToanOnDinh - 1].Text);
			if (num2 <= 0)
			{
				return num;
			}
			string text = sView.Cells[row - 1, colMaHopDongKy - 1].Text.Trim();
			if (string.IsNullOrEmpty(text) || text != MA_HOP_DONG)
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
			int num3 = Convert.ToInt32(sView.Cells[row - 1, colGiaChoThue1Thang - 1].Text.Trim().Replace(".", ""));
			if (num3 <= 0)
			{
				return num;
			}
			if (splitDay(text2)[2] == splitDay(text3)[2])
			{
				sView.Cells[row - 1, colNam - 1].Text = splitDay(text2)[2];
			}
			DateTime t = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t2 = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (DateTime.Compare(t, t2) > 0)
			{
				return num;
			}
			XoaHangTheoHopDong(sView, sheet, colSTT, row);
			List<int> hangVatNam = new List<int>();
			Dictionary<DateTime, DateTime> dictionary = TuDongTachKy(Convert.ToInt32(num2), text2, text3, row, ref hangVatNam);
			foreach (KeyValuePair<DateTime, DateTime> item in dictionary)
			{
				GeneralFunctions.InsertTableRow(sView, row + num, -1, row + num, -1, -1, sheet, 0);
				sView.Cells[row + num, colTuNgay - 1].Text = item.Key.ToString("dd/MM/yyyy");
				sView.Cells[row + num, colDenNgay - 1].Text = item.Value.ToString("dd/MM/yyyy");
				sView.Cells[row + num, colNam - 1].Text = ((Convert.ToInt32(item.Key.ToString("yyyy")) == Convert.ToInt32(item.Value.ToString("yyyy"))) ? item.Value.ToString("yyyy") : "");
				sView.Cells[row + num, colGiaChoThue1Thang - 1].Value = sView.Cells[row + num - 1, colGiaChoThue1Thang - 1].Value;
				if (num < dictionary.Count - 1)
				{
					sView.Cells[row + num, colTenHopDongKy - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN];
					sView.Cells[row + num, colMaHopDongKy - 1].Text = MA_KY_THANH_TOAN;
				}
				else
				{
					sView.Cells[row + num, colTenHopDongKy - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN_CUOI_CUNG];
					sView.Cells[row + num, colMaHopDongKy - 1].Text = MA_KY_THANH_TOAN_CUOI_CUNG;
				}
				sView.Cells[row + num, colTemDongTuSinh - 1].Text = "1";
				GeneralFunctions.SetRowHeight(sView, row + num + 1);
				num++;
			}
			int num4 = 0;
			foreach (int item2 in hangVatNam)
			{
				sView.Cells[item2 - 1, colTemVatNam - 1].Text = "1";
				num4++;
				if (num4 >= hangVatNam.Count)
				{
					ResetDongThanhToanCuoiCung(sView, item2, colTenHopDongKy, colMaHopDongKy, colTemVatNam, colTemDongTuSinh);
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

	public Dictionary<DateTime, DateTime> TuDongTachKy(int soThangOnDinh, string tuNgay, string denNgay, int row, ref List<int> hangVatNam)
	{
		Dictionary<DateTime, DateTime> dictionary = new Dictionary<DateTime, DateTime>();
		try
		{
			DateTime dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime2 = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			int num = Convert.ToInt32(Math.Round((double)dateTime2.Subtract(dateTime).Days / 30.4375 / (double)soThangOnDinh, 0));
			int num2 = 1;
			DateTime now = DateTime.Now;
			DateTime dateTime3 = DateTime.Now;
			for (int i = 0; i < num; i++)
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
				if (i != num - 1 || DateTime.Compare(dateTime3, dateTime2) >= 0)
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
			DateTime key = firstDayOfYear(num2);
			DateTime value = lastDayOfYear(num);
			dictionary.Add(kyBatDau, value);
			dictionary.Add(key, kyKeTiep);
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
			if (splitDay(text2)[2] == splitDay(text3)[2])
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
					ResetDongThanhToanCuoiCung(sView, item3, colTenHopDongKy, colMaHopDongKy, colTemVatNam, colTemDongTuSinh);
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
		string text = sView.Cells[rowHD - 1, colSTT - 1].Text.Trim();
		int num = 0;
		for (int i = 1; i < 100; i++)
		{
			string text2 = sView.Cells[startRow + i - 1, colSTT - 1].Text.Trim();
			string text3 = sView.Cells[startRow + i - 1, colMaHD - 1].Text.Trim();
			int num2 = Convert.ToInt32(sView.Cells[startRow + i - 1, colTemVatNam - 1].Value);
			if (string.IsNullOrEmpty(text3) || text3 == MA_HOP_DONG || num2 != 1 || string.IsNullOrEmpty(text2) || text2.Substring(0, 1) != text)
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

	public bool CheckCungHDCacKyPhaiLienTiep(SheetView sView, HTKKCMCSheetInfo sheet, int row, int colMaHD, int colSTT, int colTuNgay, int colDenNgay, int colTemVatNam, int colNamTinhThue, string errorKyKeTiepLonHonKyBatDau, string errorKyThanhToanPhaiLienTiepNhau)
	{
		try
		{
			string text = sView.Cells[row, colMaHD - 1].Text.Trim();
			if (text == MA_HOP_DONG)
			{
				return true;
			}
			if (!string.IsNullOrEmpty(sView.Cells[row + 1, colSTT - 1].Text) && sView.Cells[row, colSTT - 1].Text.Substring(0, 1) != sView.Cells[row + 1, colSTT - 1].Text.Substring(0, 1))
			{
				return true;
			}
			string text2 = sView.Cells[row, colTuNgay - 1].Text.Trim();
			string text3 = sView.Cells[row, colDenNgay - 1].Text.Trim();
			string text4 = sView.Cells[row + 1, colTuNgay - 1].Text.Trim();
			string text5 = sView.Cells[row + 1, colDenNgay - 1].Text.Trim();
			string text6 = sView.Cells[row + 1, colTemVatNam - 1].Text.Trim();
			string text7 = sView.Cells[row + 1, colNamTinhThue - 1].Text.Trim();
			if (string.IsNullOrEmpty(text2) || string.IsNullOrEmpty(text3) || string.IsNullOrEmpty(text4) || string.IsNullOrEmpty(text5))
			{
				return true;
			}
			DateTime t = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t2 = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime = DateTime.ParseExact(text4, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime t3 = DateTime.ParseExact(text5, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (DateTime.Compare(t, t2) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, row, colTuNgay - 1, errorKyKeTiepLonHonKyBatDau);
				HTKKFile.SetNodeBackGroundPink(sView, row, colDenNgay - 1, errorKyKeTiepLonHonKyBatDau);
				return false;
			}
			if (text6 != "1" && DateTime.Compare(t2.AddDays(1.0), dateTime) != 0)
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

	public void TinhSoThangChoThueCuaHopDong(SheetView sView, int row, int colMaHD, int colTuNgay, int colDenNgay, int colSoThangChoThueCuaHD)
	{
		try
		{
			string text = sView.Cells[row - 1, colMaHD - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			string text2 = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			string text3 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3))
			{
				DateTime value = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				int num = Convert.ToInt32(Math.Round((double)DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture).Subtract(value).Days / 30.4375, 0));
				if (text == MA_HOP_DONG)
				{
					sView.Cells[row - 1, colSoThangChoThueCuaHD - 1].Text = num.ToString();
				}
				else
				{
					sView.Cells[row - 1, colSoThangChoThueCuaHD - 1].Text = "0";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public double TinhTongGiaTriHD(SheetView sView, int rowHD, int row, int colMaHD, int colTuNgay, int colDenNgay, int colGiaChoThue1Thang, int colTemVatNam, int colTongGiaTriHD, int colTemSoThang, int colTemTronThang, int colSoThang, bool tuDongSoThang)
	{
		double num = 0.0;
		double num2 = 0.0;
		try
		{
			num = Convert.ToDouble(sView.Cells[rowHD - 1, colTongGiaTriHD - 1].Value);
			num2 = Convert.ToDouble(sView.Cells[rowHD - 1, colSoThang - 1].Value);
			double num3 = 0.0;
			string text = sView.Cells[row - 1, colMaHD - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				return num;
			}
			double num4 = Convert.ToDouble(sView.Cells[row - 1, colGiaChoThue1Thang - 1].Value);
			int num5 = Convert.ToInt32(sView.Cells[row - 1, colTemVatNam - 1].Text.Trim());
			string text2 = sView.Cells[row - 1, colTuNgay - 1].Text.Trim();
			string text3 = sView.Cells[row - 1, colDenNgay - 1].Text.Trim();
			if (string.IsNullOrEmpty(text2) || string.IsNullOrEmpty(text3) || num4 <= 0.0 || string.IsNullOrEmpty(text))
			{
				return num;
			}
			if (text != MA_HOP_DONG)
			{
				DateTime dateTime = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime2 = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				string text4 = dateTime.ToString("yyyy");
				string text5 = dateTime2.ToString("yyyy");
				double num6 = 0.0;
				string text6 = "0";
				if (num5 == 1)
				{
					int num7 = 0;
					bool flag = false;
					string text7 = "";
					string text8 = "";
					if (sView.Cells[row - 2, colTemVatNam - 1].Text.Trim() == "0")
					{
						text8 = sView.Cells[row - 2, colTuNgay - 1].Text.Trim();
						text7 = sView.Cells[row - 2, colDenNgay - 1].Text.Trim();
						num7 = row - 1;
					}
					else if (sView.Cells[row - 3, colTemVatNam - 1].Text.Trim() == "0")
					{
						text8 = sView.Cells[row - 3, colTuNgay - 1].Text.Trim();
						text7 = sView.Cells[row - 3, colDenNgay - 1].Text.Trim();
						num7 = row - 2;
						flag = true;
					}
					if (string.IsNullOrEmpty(text7) || string.IsNullOrEmpty(text8))
					{
						return num;
					}
					DateTime dateTime3 = DateTime.ParseExact(text8, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime4 = DateTime.ParseExact(text7, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (dateTime.Day == 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month))
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month + 1));
						num3 *= num4;
						text6 = "9";
					}
					else if (dateTime.Day != 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)(dateTime2.Day - dateTime.Day + 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text6 = "10";
					}
					else if (dateTime.Day == 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month) + ConvertHelpers.ConvertToDouble1(dateTime2.Day) / ConvertHelpers.ConvertToDouble1(DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month)));
						num3 *= num4;
						text6 = "11";
					}
					else
					{
						num6 = (num3 = (dateTime2.Year - dateTime.Year) * 12 + dateTime2.Month - dateTime.Month + 1);
						num3 *= num4;
						text6 = "12";
					}
					sView.Cells[row - 1, colTemTronThang - 1].Text = text6.ToString();
					sView.Cells[row - 1, colTemSoThang - 1].Text = num6.ToString();
					sView.Cells[row - 1, colTongGiaTriHD - 1].Text = Math.Round(num3, 0).ToString();
					if (flag)
					{
						sView.Cells[num7 - 1, colTemSoThang - 1].Text = Convert.ToString(Convert.ToDouble(sView.Cells[num7, colTemSoThang - 1].Value) + Convert.ToDouble(sView.Cells[num7 + 1, colTemSoThang - 1].Value));
						sView.Cells[num7 - 1, colTongGiaTriHD - 1].Text = Convert.ToString(Convert.ToDouble(sView.Cells[num7, colTongGiaTriHD - 1].Value) + Convert.ToDouble(sView.Cells[num7 + 1, colTongGiaTriHD - 1].Value));
					}
					num += num3;
					num2 += num6;
					sView.Cells[rowHD - 1, colTongGiaTriHD - 1].Text = Math.Round(num, 0).ToString();
					sView.Cells[rowHD - 1, colTemSoThang - 1].Text = Math.Round(num2, 0).ToString();
					if (tuDongSoThang)
					{
						sView.Cells[rowHD - 1, colSoThang - 1].Text = Math.Round(num2, 0).ToString();
					}
				}
				else
				{
					if (splitDay(text2)[2] != splitDay(text3)[2])
					{
						return num;
					}
					if (dateTime.Day == 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month))
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month + 1));
						num3 *= num4;
						text6 = "1";
					}
					else if (dateTime.Day != 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day == 1)
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month));
						num3 *= num4;
						text6 = "2";
					}
					else if (dateTime.Day != 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day > 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 -= (double)(dateTime.Day - dateTime2.Day - 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text6 = "4";
					}
					else if (dateTime.Day != 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)(dateTime2.Day - dateTime.Day + 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text6 = "5";
					}
					else if (dateTime.Day != 1 && dateTime2.Day == DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month);
						num6 = (num3 += (double)(dateTime2.Day - dateTime.Day + 1) / (double)DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month));
						num3 *= num4;
						text6 = "6.1";
					}
					else if (dateTime.Day == 1 && dateTime2.Day != DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) && dateTime.Day - dateTime2.Day < 1)
					{
						num6 = (num3 = (double)((dateTime2.Year - dateTime.Year) * 12) + (double)(dateTime2.Month - dateTime.Month) + ConvertHelpers.ConvertToDouble1(dateTime2.Day) / ConvertHelpers.ConvertToDouble1(DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month)));
						num3 *= num4;
						text6 = "6.2";
					}
					else
					{
						num6 = (num3 = (dateTime2.Year - dateTime.Year) * 12 + dateTime2.Month - dateTime.Month + 1);
						num3 *= num4;
						text6 = "7";
					}
					num += num3;
					num2 += num6;
					sView.Cells[row - 1, colTemTronThang - 1].Text = text6.ToString();
					sView.Cells[row - 1, colTemSoThang - 1].Text = num6.ToString();
					sView.Cells[row - 1, colTongGiaTriHD - 1].Text = Math.Round(num3, 0).ToString();
					sView.Cells[rowHD - 1, colTongGiaTriHD - 1].Text = Math.Round(num, 0).ToString();
					sView.Cells[rowHD - 1, colTemSoThang - 1].Text = Math.Round(num2, 0).ToString();
					if (tuDongSoThang)
					{
						sView.Cells[rowHD - 1, colSoThang - 1].Text = Math.Round(num2, 0).ToString();
					}
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

	public void SumTinhTongGiaTriHD(SheetView sView, List<int> hopDong, int colTongGiaTriHD, int colTuNgay, int colDenNgay)
	{
		double num = 0.0;
		for (int i = 1; i < hopDong.Count; i++)
		{
			int num2 = hopDong[i];
			string text = sView.Cells[num2 - 1, colTuNgay - 1].Text.Trim();
			string text2 = sView.Cells[num2 - 1, colDenNgay - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && splitDay(text)[2] == splitDay(text2)[2])
			{
				num += Convert.ToDouble(sView.Cells[num2 - 1, colTongGiaTriHD - 1].Value);
			}
		}
		sView.Cells[hopDong[0] - 1, colTongGiaTriHD - 1].Value = num;
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
				for (int i = 1; i < item.Value.Count; i++)
				{
					int num = item.Value[i];
					string text = sView.Cells[num - 1, colMaHD - 1].Text.Trim();
					if (string.IsNullOrEmpty(text))
					{
						continue;
					}
					string text2 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
					if (text != MA_HOP_DONG)
					{
						double num2 = Convert.ToDouble(sView.Cells[num - 1, colDoanhThuBinhQuanThang - 1].Value);
						double num3 = Convert.ToDouble(sView.Cells[num - 1, colTongDoanhThuPhatSinhTrongNam - 1].Value);
						if (num2 == 0.0 || num3 == 0.0)
						{
							continue;
						}
						string s = sView.Cells[num - 1, colTuNgay - 1].Text.Trim();
						string s2 = sView.Cells[num - 1, colDenNgay - 1].Text.Trim();
						DateTime dateTime = DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						DateTime dateTime2 = new DateTime(DateTime.ParseExact(s2, "dd/MM/yyyy", CultureInfo.CurrentCulture).Year, 1, 1).AddDays(-1.0);
						if (num3 > 100000000.0)
						{
							sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
						}
						else
						{
							sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = false;
						}
					}
					if (text2 == "1" && sView.Cells[num - 1, colTemVatNam - 1].Text == "1" && sView.Cells[num - 2, colTemVatNam - 1].Text == "1" && Convert.ToBoolean(sView.Cells[num - 1, colCaNhanThuocDienPhaiNopThue - 1].Value) && Convert.ToBoolean(sView.Cells[num - 2, colCaNhanThuocDienPhaiNopThue - 1].Value) && text != MA_HOP_DONG)
					{
						sView.Cells[num - 3, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
					}
				}
			}
			foreach (KeyValuePair<string, List<int>> item2 in cacHopDong)
			{
				for (int i = 1; i < item2.Value.Count; i++)
				{
					int num = item2.Value[i];
					Neu1Trong2DongVatNamTichThiDongTrenAnTheo(sView, colSTT, colTemVatNam, num, colCaNhanThuocDienPhaiNopThue, colTuNgay, colDenNgay);
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
					double num4 = Convert.ToDouble(sView.Cells[num3 - 1, colTongDoanhThuPhatSinhTrongNamTungCN - 1].Value);
					if (flag)
					{
						sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = flag;
					}
					else if (num4 > 100000000.0)
					{
						sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = true;
					}
					else
					{
						string[] array2 = sView.Cells[num3 - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
						string[] array3 = sView.Cells[num3, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
						string[] array4 = sView.Cells[num3 + 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
						string text2 = sView.Cells[num3, colTemVatNam - 1].Text.Trim();
						string text3 = sView.Cells[num3 + 1, colTemVatNam - 1].Text.Trim();
						string day = sView.Cells[num3 - 1, colTuNgay - 1].Text.Trim();
						string day2 = sView.Cells[num3 - 1, colDenNgay - 1].Text.Trim();
						if (splitDay(day)[2] != splitDay(day2)[2] && text2 == "1" && text3 == "1" && array2[0] == array3[0] && array2[0] == array4[0])
						{
							double num5 = Convert.ToDouble(sView.Cells[num3, colTongDoanhThuPhatSinhTrongNamTungCN - 1].Value);
							double num6 = Convert.ToDouble(sView.Cells[num3 + 1, colTongDoanhThuPhatSinhTrongNamTungCN - 1].Value);
							if (num5 > 100000000.0 && num6 > 100000000.0)
							{
								GeneralFunctions.LockUnLockCol(sView, num3, colCaNhanThuocDienPhaiNopThue, locked: true, deleteData: false);
							}
							else
							{
								sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = false;
							}
						}
						else
						{
							sView.Cells[num3 - 1, colCaNhanThuocDienPhaiNopThue - 1].Value = false;
						}
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
						if (splitDay(text2)[2] == splitDay(text3)[2])
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
		double num = 0.0;
		for (int i = 1; i < hopDong.Count; i++)
		{
			int num2 = hopDong[i];
			string text = sView.Cells[num2 - 1, colTuNgay - 1].Text.Trim();
			string text2 = sView.Cells[num2 - 1, colDenNgay - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && splitDay(text)[2] == splitDay(text2)[2])
			{
				num += Convert.ToDouble(sView.Cells[num2 - 1, colDoanhThuTinhThueTrongKy - 1].Value);
			}
		}
		sView.Cells[hopDong[0] - 1, colDoanhThuTinhThueTrongKy - 1].Value = num;
	}

	public void SoThueGTGTPhaiNop(SheetView sView, int row, int colDoanhThuTinhThueTrongKy, int colSoThueGTGTPhaiNop, int colSTT)
	{
		try
		{
			double num = Convert.ToDouble(sView.Cells[row - 1, colDoanhThuTinhThueTrongKy - 1].Value);
			sView.Cells[row - 1, colSoThueGTGTPhaiNop - 1].Value = Math.Round(num * 5.0 / 100.0, 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SoThueTNCNPhaiNop(SheetView sView, int row, int colDoanhThuTinhThueTrongKy, int colSoThueTNCNPhaiNop, int colSTT)
	{
		try
		{
			double num = Convert.ToDouble(sView.Cells[row - 1, colDoanhThuTinhThueTrongKy - 1].Value);
			sView.Cells[row - 1, colSoThueTNCNPhaiNop - 1].Value = Math.Round(num * 5.0 / 100.0, 0);
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
					string text3 = TTinChung.maCQTNoiNop + "_" + ((TTinChung.mst.Length == 10) ? (TTinChung.mst + "000_") : TTinChung.mst) + text2 + "_" + ((array[0].Length == 1) ? ("0" + array[0]) : array[0]);
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
						double num2 = Convert.ToDouble(sViewPL.Cells[item4 - 1, colTongGiaTriHopDong - 1].Value);
						if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3) && !(text4 == MA_HOP_DONG) && num2 != 0.0)
						{
							Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(int.Parse(splitDay(text2)[1]), int.Parse(splitDay(text2)[2]), "01/01");
							if (PeriodInfo.KieuKy == "Y" && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "M" && PeriodInfo.Thang == splitDay(text2)[1] && PeriodInfo.Nam == splitDay(text2)[2])
							{
								num += num2;
							}
							else if (PeriodInfo.KieuKy == "Q" && PeriodInfo.Nam == splitDay(text2)[2] && int.Parse(PeriodInfo.Quy) == int.Parse(quyNamByThangNam.Keys.ToList()[0]))
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
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
			{
				string[] array = sView.Cells[row - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array2 = sView.Cells[row, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array3 = sView.Cells[row + 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string text3 = sView.Cells[row, colVatNam - 1].Text.Trim();
				string text4 = sView.Cells[row + 1, colVatNam - 1].Text.Trim();
				if (splitDay(text)[2] != splitDay(text2)[2] && text3 == "1" && text4 == "1" && array[0] == array2[0] && array[0] == array3[0])
				{
					sView.Cells[row, colCaNhanThuocDienPhaiNopThue - 1].Value = Convert.ToDouble(sView.Cells[row, colTongDoanhThuPhatSinhTrongNamTungCaNhan - 1].Value) > 100000000.0 || Convert.ToBoolean(sView.Cells[row - 1, colCaNhanThuocDienPhaiNopThue - 1].Value);
					sView.Cells[row + 1, colCaNhanThuocDienPhaiNopThue - 1].Value = ((Convert.ToDouble(sView.Cells[row + 1, colTongDoanhThuPhatSinhTrongNamTungCaNhan - 1].Value) > 100000000.0) ? ((object)true) : sView.Cells[row - 1, colCaNhanThuocDienPhaiNopThue - 1].Value);
				}
				string[] array4 = sView.Cells[row - 2, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array5 = sView.Cells[row - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array6 = sView.Cells[row, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string text5 = sView.Cells[row - 2, colVatNam - 1].Text.Trim();
				string text6 = sView.Cells[row, colVatNam - 1].Text.Trim();
				if (splitDay(text)[2] == splitDay(text2)[2] && text5 == "0" && text6 == "1" && array4[0] == array5[0] && array4[0] == array6[0])
				{
					sView.Cells[row - 2, colCaNhanThuocDienPhaiNopThue - 1].Value = ((Convert.ToBoolean(sView.Cells[row - 1, colCaNhanThuocDienPhaiNopThue - 1].Value) == Convert.ToBoolean(sView.Cells[row, colCaNhanThuocDienPhaiNopThue - 1].Value) && Convert.ToBoolean(sView.Cells[row - 1, colCaNhanThuocDienPhaiNopThue - 1].Value)) ? true : false);
				}
				string[] array7 = sView.Cells[row - 3, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array8 = sView.Cells[row - 2, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string[] array9 = sView.Cells[row - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
				string text7 = sView.Cells[row - 2, colVatNam - 1].Text.Trim();
				string text8 = sView.Cells[row - 1, colVatNam - 1].Text.Trim();
				if (splitDay(text)[2] == splitDay(text2)[2] && text7 == "1" && text8 == "1" && array7[0] == array8[0] && array7[0] == array9[0])
				{
					sView.Cells[row - 3, colCaNhanThuocDienPhaiNopThue - 1].Value = ((Convert.ToBoolean(sView.Cells[row - 2, colCaNhanThuocDienPhaiNopThue - 1].Value) == Convert.ToBoolean(sView.Cells[row - 1, colCaNhanThuocDienPhaiNopThue - 1].Value) && Convert.ToBoolean(sView.Cells[row - 2, colCaNhanThuocDienPhaiNopThue - 1].Value)) ? true : false);
				}
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
							Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(int.Parse(splitDay(text2)[1]), int.Parse(splitDay(text2)[2]), "01/01");
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
							else if (PeriodInfo.KieuKy == "Q" && PeriodInfo.Nam == splitDay(text2)[2] && int.Parse(PeriodInfo.Quy) == int.Parse(quyNamByThangNam.Keys.ToList()[0]))
							{
								num += num2;
							}
						}
					}
				}
			}
			sViewToKhai.Cells[rowCellCt24ToKhai.CellTemplate.RowIndex - 1, rowCellCt24ToKhai.CellTemplate.ColIndex - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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
					if (text2.Length > 0 && text3.Length > 0 && text2.Substring(0, 1) == text3.Substring(0, 1))
					{
						CellChange_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, num23, num4);
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

	public bool UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo sheet = ToKhaiInfo.SheetInfos[2];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int colMaHD = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			int item = HTKKFile.ColumnLettersToIndex("G");
			int item2 = HTKKFile.ColumnLettersToIndex("J");
			int item3 = HTKKFile.ColumnLettersToIndex("L");
			int item4 = HTKKFile.ColumnLettersToIndex("O");
			int item5 = HTKKFile.ColumnLettersToIndex("P");
			int item6 = HTKKFile.ColumnLettersToIndex("Q");
			int item7 = HTKKFile.ColumnLettersToIndex("R");
			int item8 = HTKKFile.ColumnLettersToIndex("S");
			int item9 = HTKKFile.ColumnLettersToIndex("U");
			int num4 = HTKKFile.ColumnLettersToIndex("T");
			int item10 = HTKKFile.ColumnLettersToIndex("V");
			int item11 = HTKKFile.ColumnLettersToIndex("W");
			int item12 = HTKKFile.ColumnLettersToIndex("Y");
			int num5 = HTKKFile.ColumnLettersToIndex("X");
			int num6 = HTKKFile.ColumnLettersToIndex("Z");
			int item13 = HTKKFile.ColumnLettersToIndex("AA");
			int num7 = HTKKFile.ColumnLettersToIndex("AB");
			int num8 = HTKKFile.ColumnLettersToIndex("AC");
			int num9 = HTKKFile.ColumnLettersToIndex("AD");
			int num10 = HTKKFile.ColumnLettersToIndex("AE");
			int item14 = HTKKFile.ColumnLettersToIndex("AL");
			int item15 = HTKKFile.ColumnLettersToIndex("AM");
			int num11 = HTKKFile.ColumnLettersToIndex("AN");
			int num12 = HTKKFile.ColumnLettersToIndex("AO");
			int num13 = HTKKFile.ColumnLettersToIndex("AP");
			int num14 = HTKKFile.ColumnLettersToIndex("AQ");
			int num15 = HTKKFile.ColumnLettersToIndex("AR");
			int num16 = HTKKFile.ColumnLettersToIndex("AS");
			int item16 = HTKKFile.ColumnLettersToIndex("AU");
			int num17 = HTKKFile.ColumnLettersToIndex("AT");
			int item17 = HTKKFile.ColumnLettersToIndex("AV");
			int num18 = HTKKFile.ColumnLettersToIndex("AI");
			int num19 = HTKKFile.ColumnLettersToIndex("AK");
			int num20 = HTKKFile.ColumnLettersToIndex("AJ");
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
			list.Add(item14);
			List<int> list2 = new List<int>();
			list2.Add(item10);
			list2.Add(item16);
			list2.Add(item17);
			List<int> list3 = new List<int>();
			list3.Add(item15);
			List<int> list4 = new List<int>();
			list4.Add(num3);
			list4.Add(item13);
			list4.Add(item15);
			string text = sheetView.Cells[(int)_iRowClick - 1, num - 1].Text.Trim();
			TinhLaiSTT(sheetView, sheet, num, colMaHD);
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sheetView, sheet, num);
			if (dictionary.Count > 0)
			{
				foreach (KeyValuePair<string, List<int>> item18 in dictionary)
				{
					int num21 = item18.Value[0];
					string text2 = sheetView.Cells[num21 - 1, num - 1].Text.Trim();
					string text3 = sheetView.Cells[(int)_iRowClick - 2, num - 1].Text.Trim();
					if (text2.Length > 0 && text3.Length > 0 && text2.Substring(0, 1) == text3.Substring(0, 1))
					{
						CellChange_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, num21, num3);
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

	public void LockUnLockDongHopDongHoacKy(SheetView sView, HTKKCMCSheetInfo sInfo, int colSTT, int colCt06aMa, int colTemDongTuSinh, int colTemVatNam, List<int> colAnKy, List<int> colAnNhungKhongXoaDuLieuKy, List<int> colAnNeuTuSinh, List<int> colAnNeuVatNam, List<int> colAnNeuLaHopDong, List<int> colKhongKhoaXoaDataNeuVatNamDongKy, List<int> colKhongAnNeuLaHopDong, int colTuNgay, int colDenNgay, List<int> colAnNeuTruocVatNam, int colTongDoanhThuPhatSinhTrongNamTungCN, int colCaNhanThuocDienPhaiNopThue)
	{
		try
		{
			Dictionary<string, List<int>> dictionary = LayCacHangCungHopDong(sView, sInfo, colSTT);
			if (dictionary.Count <= 0)
			{
				return;
			}
			foreach (KeyValuePair<string, List<int>> item in dictionary)
			{
				int row = item.Value[0];
				for (int i = 1; i < item.Value.Count; i++)
				{
					int num = item.Value[i];
					string text = sView.Cells[num - 1, colCt06aMa - 1].Text.Trim();
					string text2 = sView.Cells[num - 1, colTemDongTuSinh - 1].Text.Trim();
					string text3 = sView.Cells[num - 1, colTemVatNam - 1].Text.Trim();
					string day = sView.Cells[num - 1, colTuNgay - 1].Text.Trim();
					string day2 = sView.Cells[num - 1, colDenNgay - 1].Text.Trim();
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
						if (splitDay(day)[2] != splitDay(day2)[2])
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
				GeneralFunctions.LockUnLockCols(sView, row, colAnNeuLaHopDong, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCols(sView, row, colAnNeuTruocVatNam, locked: false, deleteData: false);
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

	public void Set1RowHDToAllRowKy(SheetView sView, KeyValuePair<string, List<int>> hopDong, int colStt, int colCopy, string valueCopy, string valueHD)
	{
		string text = sView.Cells[hopDong.Value[0] - 1, colStt - 1].Text.Trim();
		if (!string.IsNullOrEmpty(text) && text.Substring(0, 1) == valueHD)
		{
			for (int i = 1; i < hopDong.Value.Count; i++)
			{
				sView.Cells[hopDong.Value[i] - 1, colCopy - 1].Text = valueCopy;
				GeneralFunctions.SetRowHeight(sView, hopDong.Value[i]);
			}
		}
	}

	public void ResetDongThanhToanCuoiCung(SheetView sView, int row, int colTenHD, int colMaHD, int colTemVatNam, int colTemTuSinh)
	{
		string text = sView.Cells[row - 1, colTemVatNam - 1].Text.Trim();
		string text2 = sView.Cells[row, colTemVatNam - 1].Text.Trim();
		string text3 = sView.Cells[row, colTemTuSinh - 1].Text.Trim();
		if (text3 == "1" && text == "1" && text2 != "1")
		{
			string text4 = sView.Cells[row, colMaHD - 1].Text.Trim();
			if (text4 != MA_KY_THANH_TOAN && text4 != MA_KY_THANH_TOAN_CUOI_CUNG)
			{
				sView.Cells[row - 1, colMaHD - 1].Text = MA_KY_THANH_TOAN;
				sView.Cells[row - 1, colTenHD - 1].Text = "";
				sView.Cells[row - 3, colMaHD - 1].Text = MA_KY_THANH_TOAN_CUOI_CUNG;
				sView.Cells[row - 3, colTenHD - 1].Text = dicDMHopDongTmp[MA_KY_THANH_TOAN_CUOI_CUNG];
			}
		}
	}

	public void CheckChiTieu23BoTichDongHD(SheetView sView, List<int> hopDong, int colCaNhanPhaiNopThue)
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
				if (splitDay(text)[2] == splitDay(text2)[2] && text3 == "1" && text4 == "1" && array[0] == array2[0] && array[0] == array3[0])
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
		string[] array = sView.Cells[rowKy - 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
		string[] array2 = sView.Cells[rowKy, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
		string[] array3 = sView.Cells[rowKy + 1, colSTT - 1].Text.Trim().Split(new char[1] { '.' });
		string text = sView.Cells[rowKy, colTemVatNam - 1].Text.Trim();
		string text2 = sView.Cells[rowKy + 1, colTemVatNam - 1].Text.Trim();
		string day = sView.Cells[rowKy - 1, colTuNgay - 1].Text.Trim();
		string day2 = sView.Cells[rowKy - 1, colDenNgay - 1].Text.Trim();
		if (splitDay(day)[2] != splitDay(day2)[2] && text == "1" && text2 == "1" && array[0] == array2[0] && array[0] == array3[0])
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
}
