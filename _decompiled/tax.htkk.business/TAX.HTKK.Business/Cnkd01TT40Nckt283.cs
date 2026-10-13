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

public class Cnkd01TT40Nckt283
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMThueBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueBVMTTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMNganhNgheKD = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMNganhNgheKDTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMNhomNganhNgheKD = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMNhomNganhNgheKDTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMKSBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMKSBVMTTmp = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private List<item> _dmBVMTData = new List<item>();

	private bool _isDmMoi = false;

	private bool isToChuc = false;

	private string mst = "";

	private int[] cellPhuLuc01_3 = GeneralFunctions.ArrayCellId("AC_48");

	private int[] cellPhuLuc01_4 = GeneralFunctions.ArrayCellId("AD_48");

	private int[] cellPhuLuc01_5 = GeneralFunctions.ArrayCellId("AE_48");

	private int col3_PL01_3 = HTKKFile.ColumnLettersToIndex("E");

	private int col3_PL01_4 = HTKKFile.ColumnLettersToIndex("E");

	private int col4_PL01_4 = HTKKFile.ColumnLettersToIndex("F");

	private int col3_PL01_5 = HTKKFile.ColumnLettersToIndex("E");

	private int col4_PL01_5 = HTKKFile.ColumnLettersToIndex("F");

	private string namKK = "";

	private string kieuKyKK = "";

	private List<string> pathDonGiaTinh = new List<string>();

	private HTKKHeaderInfo _HeaderInfo = null;

	private HTKKPeriodInfo _PeriodInfo = null;

	private XMLMessage _xmlMessage = new XMLMessage("../Project/Message.xml");

	public Dictionary<string, List<double>> _dicMucThue = new Dictionary<string, List<double>>();

	private int[] HKD_CNKDNopThueKhoan = GeneralFunctions.ArrayCellId("B_6");

	private int[] CNKDNopPs = GeneralFunctions.ArrayCellId("B_8");

	private int[] TCCNKhaiNopThay = GeneralFunctions.ArrayCellId("B_10");

	private int[] HKD_CNKDNopKeKhai = GeneralFunctions.ArrayCellId("B_13");

	private int[] HKD_CNKD_NNXDDoanhThu = GeneralFunctions.ArrayCellId("B_15");

	private int[] HKDChuyenDoiPpTinhThue = GeneralFunctions.ArrayCellId("B_17");

	private int colB = HTKKFile.ColumnLettersToIndex("B");

	private int[] ct04 = GeneralFunctions.ArrayCellId("F_46");

	private int[] ct05 = GeneralFunctions.ArrayCellId("G_48");

	private int[] ct06 = GeneralFunctions.ArrayCellId("F_50");

	private int[] ct07 = GeneralFunctions.ArrayCellId("F_52");

	private int[] ct07_SoDinhDanh = GeneralFunctions.ArrayCellId("S_52");

	private int[] ct09 = GeneralFunctions.ArrayCellId("F_63");

	private int[] ct13a = GeneralFunctions.ArrayCellId("H_92");

	private int[] ct14 = GeneralFunctions.ArrayCellId("E_94");

	private int[] ct15 = GeneralFunctions.ArrayCellId("J_94");

	private int[] ct16 = GeneralFunctions.ArrayCellId("Q_94");

	private int[] ct17 = GeneralFunctions.ArrayCellId("J_96");

	private int[] ct17_1 = GeneralFunctions.ArrayCellId("S_96");

	private int[] ct22 = GeneralFunctions.ArrayCellId("I_169");

	private int[] ct23 = GeneralFunctions.ArrayCellId("I_171");

	private int[] ct24 = GeneralFunctions.ArrayCellId("E_173");

	private int[] ct25 = GeneralFunctions.ArrayCellId("E_175");

	private int[] ct26 = GeneralFunctions.ArrayCellId("J_175");

	private int[] ct27 = GeneralFunctions.ArrayCellId("Q_175");

	private int[] ct28a = GeneralFunctions.ArrayCellId("L_212");

	private int[] ct29a = GeneralFunctions.ArrayCellId("L_214");

	private int[] ct30a = GeneralFunctions.ArrayCellId("L_216");

	private int[] ct31a = GeneralFunctions.ArrayCellId("L_218");

	private int[] ct28b = GeneralFunctions.ArrayCellId("P_212");

	private int[] ct29b = GeneralFunctions.ArrayCellId("P_214");

	private int[] ct30b = GeneralFunctions.ArrayCellId("P_216");

	private int[] ct31b = GeneralFunctions.ArrayCellId("P_218");

	private int[] ct28c = GeneralFunctions.ArrayCellId("T_212");

	private int[] ct29c = GeneralFunctions.ArrayCellId("T_214");

	private int[] ct30c = GeneralFunctions.ArrayCellId("T_216");

	private int[] ct31c = GeneralFunctions.ArrayCellId("T_218");

	private int[] ct28d = GeneralFunctions.ArrayCellId("X_212");

	private int[] ct29d = GeneralFunctions.ArrayCellId("X_214");

	private int[] ct30d = GeneralFunctions.ArrayCellId("X_216");

	private int[] ct31d = GeneralFunctions.ArrayCellId("X_218");

	private int colTCKhaiThay = HTKKFile.ColumnLettersToIndex("AA");

	private int colAB = HTKKFile.ColumnLettersToIndex("AB");

	private int colAM = HTKKFile.ColumnLettersToIndex("AM");

	private int colAC = HTKKFile.ColumnLettersToIndex("AC");

	private int colAD = HTKKFile.ColumnLettersToIndex("AD");

	private int[] cellCoPhuLuc = GeneralFunctions.ArrayCellId("AE_36");

	private int rowTem36 = 36;

	private string pathCt18b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct18b_tenQuocTich";

	private string pathCt18c2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDChuaDangKyThue/ct18c_2_noiCap_ten";

	private string pathCt18d2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct18d_2_noiCap_ten";

	private string pathCt18e2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CNKDKhongCoCMND_CCCD/ct18e_2_noiCap_ten";

	private int colCongThucCot28_31_b = HTKKFile.ColumnLettersToIndex("AG");

	private int colCongThucCot28_31_d = HTKKFile.ColumnLettersToIndex("AH");

	private int colMaNhom = HTKKFile.ColumnLettersToIndex("AE");

	private int colCongThucCot28_31_b_PL01_4 = HTKKFile.ColumnLettersToIndex("AK");

	private int colCongThucCot28_31_b_PL01_5 = HTKKFile.ColumnLettersToIndex("AL");

	private static string fmlCt19 = "R40+AD40";

	private static string fmlCt20 = "R41+AD41";

	private static string fmlCt21 = "R42+AD42";

	private static string fmlCt22 = "R43+AD43";

	private static string fmlCt23 = "R44+AD44";

	private static string fmlCt24 = "R46+AD46";

	private static string fmlCt25 = "R47+AD47";

	private static string fmlCt26 = "R48+AD48";

	private static string fmlCt27 = "R49+AD49";

	private static string fmlCt28 = "R50+AD50";

	private static string fmlCt29 = "R52+AD52";

	private static string fmlCt31 = "R54+AD54";

	private static string fmlCt33 = "R57+AD57";

	private static string fmlCt34 = "R58+AD58";

	private static string fmlCt35 = "R59+AD59";

	private static string fmlCt36 = "R60+AD60";

	private static string fmlCt37 = "R61+AD61";

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
			_HeaderInfo = HeaderInfo;
			_PeriodInfo = PeriodInfo;
			if (PeriodInfo.Thang == "07" && PeriodInfo.Nam == "2022" && _dicMucThue.Count == 0)
			{
				_dicMucThue.Add("020201", new List<double> { 2000.0, 1000.0 });
				_dicMucThue.Add("020202", new List<double> { 1000.0, 500.0 });
				_dicMucThue.Add("020203", new List<double> { 300.0, 300.0 });
				_dicMucThue.Add("020205", new List<double> { 1500.0, 1000.0 });
				_dicMucThue.Add("020206", new List<double> { 2000.0, 1000.0 });
				_dicMucThue.Add("020207", new List<double> { 1000.0, 300.0 });
				_dicMucThue.Add("020208", new List<double> { 1000.0, 300.0 });
				_dicMucThue.Add("020209", new List<double> { 1000.0, 300.0 });
			}
			sView.Cells[rowTem36 - 1, colTCKhaiThay - 1].Text = PeriodInfo.ToChucKhaiThay;
			CheckPhuLuc(sView.FpSpread, ToKhaiInfo);
			if (string.IsNullOrEmpty(sView.Cells[rowTem36 - 1, colAC - 1].Text.Trim()) && !string.IsNullOrEmpty(PeriodInfo.maCQTQuanLy))
			{
				sView.Cells[rowTem36 - 1, colAC - 1].Text = PeriodInfo.tenCQTQuanLy;
				sView.Cells[rowTem36 - 1, colAD - 1].Text = PeriodInfo.maCQTQuanLy;
			}
			string value = sView.Cells[ct04[1], ct04[0]].Text.Trim();
			string value2 = sView.Cells[ct22[1], ct22[0]].Text.Trim();
			if (Convert.ToBoolean(sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value))
			{
				bool isDefault = (string.IsNullOrEmpty(value2) ? true : false);
				RangBuocKyTinhThueHinhThucKeKhai(sView, PeriodInfo, isDefault);
				ThayDoiCacHinhThucKeKhai(sView, ToKhaiInfo, PeriodInfo, isDefault);
				sView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Locked = true;
			}
			else
			{
				bool isDefault = (string.IsNullOrEmpty(value) ? true : false);
				RangBuocKyTinhThueHinhThucKeKhai(sView, PeriodInfo, isDefault);
				ThayDoiCacHinhThucKeKhai(sView, ToKhaiInfo, PeriodInfo, isDefault);
				sView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Locked = false;
			}
			CheckChonPhuLucVaHinhThucKhongThoaMan(sView, PeriodInfo, ToKhaiInfo);
			sView.Cells[cellPhuLuc01_3[1], cellPhuLuc01_3[0]].Value = ((ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
			sView.Cells[cellPhuLuc01_4[1], cellPhuLuc01_4[0]].Value = ((ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
			sView.Cells[cellPhuLuc01_5[1], cellPhuLuc01_5[0]].Value = ((ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
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
			CheckPhuLuc(sView.FpSpread, ToKhaiInfo);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_4(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("B_2");
			string text = "";
			if (PeriodInfo.KieuKy == "D" || PeriodInfo.KieuKy == "M")
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
			sView.Cells[array[1], array[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("B_2");
			string text = "";
			if (PeriodInfo.Nam.Equals("2024"))
			{
				text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
			}
			else if (PeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(PeriodInfo.Thang) <= 6 && !string.IsNullOrEmpty(PeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(PeriodInfo.Quy) <= 2 && !string.IsNullOrEmpty(PeriodInfo.Quy))))
			{
				text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
			}
			else if ((PeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(PeriodInfo.Thang) > 6 && !string.IsNullOrEmpty(PeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(PeriodInfo.Quy) > 2 && !string.IsNullOrEmpty(PeriodInfo.Quy)))) || PeriodInfo.Nam.Equals("2026"))
			{
				text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
			}
			sView.Cells[array[1], array[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetFormula(SheetView sView)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ResetFormula(SheetView sViewTK)
	{
		try
		{
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("AC");
			int num6 = HTKKFile.ColumnLettersToIndex("AB");
			int num7 = HTKKFile.ColumnLettersToIndex("K");
			int num8 = HTKKFile.ColumnLettersToIndex("M");
			int num9 = HTKKFile.ColumnLettersToIndex("U");
			int num10 = HTKKFile.ColumnLettersToIndex("H");
			int num11 = HTKKFile.ColumnLettersToIndex("J");
			int num12 = HTKKFile.ColumnLettersToIndex("W");
			int num13 = HTKKFile.ColumnLettersToIndex("AD");
			int num14 = HTKKFile.ColumnLettersToIndex("AE");
			if (iCol == num)
			{
				string[] array = dicDMNganhNgheKD[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 1)
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = array[0];
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "ttdb")
			{
				string[] array = dicDMThueDB[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 3)
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = array[1];
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = array[3];
					sheetView.Cells[(int)iRow - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(array[4]);
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = array[1];
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "TTDB");
					sheetView.Cells[(int)iRow - 1, num14 - 1].Value = maTieuMucByMaHH;
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "tain1")
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
				HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num3 - 1);
				string text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				foreach (KeyValuePair<string, string[]> item in dicDMKhoangSan)
				{
					string[] value = item.Value;
					if (value[2].Trim().Equals(text) && value.Length > 3 && value.Length > 2)
					{
						if (ConvertHelpers.ConvertToInt(value[4]) == 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
							winFpSpread.TabStrip[0].BackColor = Color.Pink;
						}
						sheetView.Cells[(int)iRow - 1, num10 - 1].Value = value[1];
						sheetView.Cells[(int)iRow - 1, num11 - 1].Value = value[3];
						sheetView.Cells[(int)iRow - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(value[4]);
						sheetView.Cells[(int)iRow - 1, num6 - 1].Value = value[1];
						sheetView.Cells[(int)iRow - 1, num13 - 1].Value = value[1];
						sheetView.Cells[(int)iRow - 1, num14 - 1].Value = ((value.Length == 8) ? value[7] : "");
						double donGia = getDonGia(item.Key);
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = donGia;
					}
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "bvmt")
			{
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 1)
				{
					sheetView.Cells[(int)iRow - 1, num10 - 1].Value = array[0];
					sheetView.Cells[(int)iRow - 1, num11 - 1].Value = array[2];
					sheetView.Cells[(int)iRow - 1, num4 - 1].Text = array[3];
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = array[0];
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = array[0];
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[0], "BVMT");
					sheetView.Cells[(int)iRow - 1, num14 - 1].Value = maTieuMucByMaHH;
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "tain2")
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[4];
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowsInfo.Cells.Find((HTKKCMCTableCellInfo obj) => obj.Path.Equals("ct2_ten"));
				bool flag = false;
				if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count() > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
				{
					string text2 = PeriodInfo.dicVar["IS_DM_MOI"].ToString();
					if (text2 == "1")
					{
						flag = true;
					}
				}
				string[] array = ((!flag) ? dicDMKhoangSan[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()] : dicDMKSBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()]);
				if (array.Length > 2)
				{
					sheetView.Cells[(int)iRow - 1, num10 - 1].Value = array[1];
					sheetView.Cells[(int)iRow - 1, num11 - 1].Value = array[3];
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = array[1];
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = array[1];
					string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "PBVMT");
					sheetView.Cells[(int)iRow - 1, num14 - 1].Value = maTieuMucByMaHH;
				}
			}
			if (iCol == num3 && iRow >= ToKhaiInfo.SheetInfos[0].Tables[2].Location.DataLocationTemplateRowStart && iRow <= ToKhaiInfo.SheetInfos[0].Tables[4].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[4].RowsValues.Count)
			{
				if (HTKKFormater.isNullString(sheetView.Cells[(int)iRow - 1, num3 - 1].Text))
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = false;
				}
			}
			if (iCol == num12)
			{
				object value2 = sheetView.Cells[(int)iRow - 1, num12 - 1].Value;
				sheetView.Cells[(int)iRow - 1, num12 - 1].Formula = "";
				sheetView.Cells[(int)iRow - 1, num12 - 1].Value = value2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("U");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("H");
			int num6 = HTKKFile.ColumnLettersToIndex("K");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("O");
			int num10 = HTKKFile.ColumnLettersToIndex("P");
			int num11 = HTKKFile.ColumnLettersToIndex("R");
			int num12 = HTKKFile.ColumnLettersToIndex("T");
			int num13 = HTKKFile.ColumnLettersToIndex("U");
			int num14 = HTKKFile.ColumnLettersToIndex("V");
			int num15 = HTKKFile.ColumnLettersToIndex("W");
			int num16 = HTKKFile.ColumnLettersToIndex("X");
			int num17 = HTKKFile.ColumnLettersToIndex("Y");
			int num18 = HTKKFile.ColumnLettersToIndex("P");
			if (iCol == num6)
			{
				string[] array = dicDMNhomNganhNgheKD[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 2)
				{
					sheetView.Cells[(int)iRow - 1, num15 - 1].Text = array[0];
					sheetView.Cells[(int)iRow - 1, num16 - 1].Value = Convert.ToDouble(array[2]);
					sheetView.Cells[(int)iRow - 1, num17 - 1].Value = Convert.ToDouble(array[3]);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AB");
			int num6 = HTKKFile.ColumnLettersToIndex("K");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("H");
			int num10 = HTKKFile.ColumnLettersToIndex("J");
			int num11 = HTKKFile.ColumnLettersToIndex("Q");
			int num12 = HTKKFile.ColumnLettersToIndex("W");
			int num13 = HTKKFile.ColumnLettersToIndex("L");
			string text = "ROUND(M###*Q###,0)";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[4];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
			{
				if (iCol == num12)
				{
					object value = sheetView.Cells[(int)iRow - 1, num12 - 1].Value;
					sheetView.Cells[(int)iRow - 1, num12 - 1].Formula = "";
					sheetView.Cells[(int)iRow - 1, num12 - 1].Value = value;
				}
				else if (iCol == num7 || iCol == num11)
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Formula = text.Replace("###", iRow.ToString());
				}
			}
			if (iCol == num11 && sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim() == "tain1")
			{
				validateTaiNguyen(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
			List<HTKKCMCTableInfo> tables = ToKhaiInfo.SheetInfos[0].Tables;
			for (int num14 = tables.Count - 1; num14 > tables.Count - 3; num14--)
			{
				SheetView sheetView2 = winFpSpread.Sheets[0];
				HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[0];
				fpSpReindexingTableTemplate(winFpSpread.Sheets[0], tables[num14]);
			}
			SheetView sheetView3 = winFpSpread.Sheets[0];
			string text2 = sheetView.Cells[(int)iRow - 1, colMaNhom - 1].Text.Trim();
			int num15 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			if (iCol == HTKKFile.ColumnLettersToIndex("L") || iCol == HTKKFile.ColumnLettersToIndex("T"))
			{
				TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
			}
			if (iCol == colB)
			{
				RangBuocGiuaHinhThucKeKhai(sheetView, PeriodInfo, (int)iRow);
				string text3 = sheetView.Cells[ct04[1], ct04[0]].Text.Trim();
				string text4 = sheetView.Cells[ct22[1], ct22[0]].Text.Trim();
				bool isDefault = ((iRow >= HKD_CNKDNopThueKhoan[1] + 1 && iRow < HKDChuyenDoiPpTinhThue[1] + 2) ? true : false);
				if (Convert.ToBoolean(sheetView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value))
				{
					ThayDoiCacHinhThucKeKhai(sheetView, ToKhaiInfo, PeriodInfo, isDefault);
					sheetView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Locked = true;
					sheetView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Text = "";
				}
				else
				{
					ThayDoiCacHinhThucKeKhai(sheetView, ToKhaiInfo, PeriodInfo, isDefault);
					sheetView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Locked = false;
				}
				CheckChonPhuLucVaHinhThucKhongThoaMan(sheetView, PeriodInfo, ToKhaiInfo);
			}
			if (iRow >= HKD_CNKDNopThueKhoan[1] + 1 || iRow <= HKDChuyenDoiPpTinhThue[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			}
			int num16 = 0;
			if (iCol == num3)
			{
				string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_78"));
				string text5 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
				if (text5.StartsWith("-----") || text5.StartsWith("---"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num3 - 1, message);
					num16 = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num3 - 1);
				}
				if (num16 == 2)
				{
					winFpSpread.TabStrip[0].BackColor = Color.Pink;
				}
				else
				{
					winFpSpread.TabStrip[0].BackColor = Color.White;
				}
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
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("AC");
			int num6 = HTKKFile.ColumnLettersToIndex("AB");
			int num7 = HTKKFile.ColumnLettersToIndex("K");
			int num8 = HTKKFile.ColumnLettersToIndex("M");
			int num9 = HTKKFile.ColumnLettersToIndex("U");
			int num10 = HTKKFile.ColumnLettersToIndex("H");
			int num11 = HTKKFile.ColumnLettersToIndex("J");
			int num12 = HTKKFile.ColumnLettersToIndex("N");
			int num13 = HTKKFile.ColumnLettersToIndex("O");
			int num14 = HTKKFile.ColumnLettersToIndex("P");
			int num15 = HTKKFile.ColumnLettersToIndex("R");
			int num16 = HTKKFile.ColumnLettersToIndex("W");
			int[] array = GeneralFunctions.ArrayCellId("X_" + iRow);
			int[] array2 = GeneralFunctions.ArrayCellId("Y_" + iRow);
			double num17 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array[1], array[0]].Value);
			double num18 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
			double num19 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num12, "NumberCellType"));
			if (iCol == num7 || iCol == num12)
			{
				double num20 = Math.Round(num19 * (num17 / 100.0), 0);
				sheetView.Cells[(int)iRow - 1, num13 - 1].Value = num20;
			}
			if (iCol == num7 || iCol == num12)
			{
				double num21 = Math.Round(num19 * (num18 / 100.0), 0);
				sheetView.Cells[(int)iRow - 1, num14 - 1].Value = num21;
			}
			ThayDoiToKhai28Den31_a_c(winFpSpread, ToKhaiInfo);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
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

	public void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		if (iCol == col3_PL01_3)
		{
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		}
	}

	public void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		if (iCol == col3_PL01_4 || iCol == col4_PL01_4)
		{
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		}
	}

	public void CellChange_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		if (iCol == col3_PL01_5 || iCol == col4_PL01_5)
		{
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
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
		int num = HTKKFile.ColumnLettersToIndex("C");
		int num2 = HTKKFile.ColumnLettersToIndex("E");
		int num3 = HTKKFile.ColumnLettersToIndex("F");
		int num4 = HTKKFile.ColumnLettersToIndex("Q");
		string text = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/tenNNT";
		string text2 = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/mst";
		string text3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct10";
		string text4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT11/tuGio";
		string text5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT11/tuPhut";
		string text6 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT11/denGio";
		string text7 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT11/denPhut";
		string pathCt12b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12/ct12b_soNha";
		string pathCt12c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12/ct12c_tenPhuong";
		string pathCt12d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12/ct12d_tenQuan";
		string pathCt12dd = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12/ct12d_tenTinh";
		string pathCt12ddMaTinh = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12/ct12d_maTinh";
		string pathCt12dMaQuan = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT12/ct12d_maQuan";
		int[] array = GeneralFunctions.ArrayCellId("L_212");
		int[] array2 = GeneralFunctions.ArrayCellId("L_214");
		int[] array3 = GeneralFunctions.ArrayCellId("L_216");
		int[] array4 = GeneralFunctions.ArrayCellId("L_218");
		int[] array5 = GeneralFunctions.ArrayCellId("T_212");
		int[] array6 = GeneralFunctions.ArrayCellId("T_214");
		int[] array7 = GeneralFunctions.ArrayCellId("T_216");
		int[] array8 = GeneralFunctions.ArrayCellId("T_218");
		int num5 = HTKKFile.ColumnLettersToIndex("AB");
		int num6 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[91, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[42, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[43, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[45, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[48, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text18 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[66, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text19 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[67, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text20 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[75, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text21 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[76, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_78"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_85"));
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12b));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12c));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12d));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12dd));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12ddMaTinh));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12dMaQuan));
			HTKKFile.RemoveValidate(sheetView, ct05[1], ct05[0]);
			HTKKFile.RemoveValidate(sheetView, ct09[1] + num6, ct09[0]);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, ct22[1] + num6, ct22[0]);
			HTKKFile.RemoveValidate(sheetView, ct23[1] + num6, ct23[0]);
			HTKKFile.RemoveValidate(sheetView, ct24[1] + num6, ct24[0]);
			HTKKFile.RemoveValidate(sheetView, array5[1] + num6, array5[0]);
			HTKKFile.RemoveValidate(sheetView, array6[1] + num6, array6[0]);
			HTKKFile.RemoveValidate(sheetView, array7[1] + num6, array7[0]);
			HTKKFile.RemoveValidate(sheetView, array8[1] + num6, array8[0]);
			string value2 = sheetView.Cells[ct05[1], ct05[0]].Text.Trim();
			if (string.IsNullOrEmpty(value2) && !sheetView.Cells[ct05[1], ct05[0]].Locked)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, ct05[1], ct05[0], text8);
				value = 2;
			}
			string text22 = sheetView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Text.Trim();
			HTKKFile.RemoveValidate(sheetView, ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]);
			if (string.IsNullOrEmpty(text22) && !sheetView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Locked)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, ct07_SoDinhDanh[1], ct07_SoDinhDanh[0], text8);
				value = 2;
			}
			else if (!ValidateBase.CheckMaSoThueCCCD(text22) && !sheetView.Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Locked)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, ct07_SoDinhDanh[1], ct07_SoDinhDanh[0], text9);
				value = 2;
			}
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart;
			for (int num7 = 0; num7 < ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; num7++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num7 - 1, num3 - 1);
			}
			for (int num7 = 0; num7 < ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; num7++)
			{
				if (!sheetView.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Locked && string.IsNullOrEmpty(sheetView.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Text.Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num7 - 1, num3 - 1, text8);
					value = 2;
				}
			}
			double num8 = Convert.ToDouble(sheetView.Cells[ct09[1] + num6, ct09[0]].Value);
			if (num8 == 0.0 && !sheetView.Cells[ct09[1] + num6, ct09[0]].Locked)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, ct09[1] + num6, ct09[0], text8);
				value = 2;
			}
			double num9 = Convert.ToDouble(sheetView.Cells[array[1] + num6, array[0]].Value);
			double num10 = Convert.ToDouble(sheetView.Cells[array2[1] + num6, array2[0]].Value);
			double num11 = Convert.ToDouble(sheetView.Cells[array3[1] + num6, array3[0]].Value);
			double num12 = Convert.ToDouble(sheetView.Cells[array4[1] + num6, array4[0]].Value);
			double num13 = Convert.ToDouble(sheetView.Cells[array5[1] + num6, array5[0]].Value);
			double num14 = Convert.ToDouble(sheetView.Cells[array6[1] + num6, array6[0]].Value);
			double num15 = Convert.ToDouble(sheetView.Cells[array7[1] + num6, array7[0]].Value);
			double num16 = Convert.ToDouble(sheetView.Cells[array8[1] + num6, array8[0]].Value);
			if (num13 < num9)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1] + num6, array[0], text14);
				HTKKFile.SetNodeBackGroundPink(sheetView, array5[1] + num6, array5[0], text14);
				value = 2;
			}
			if (num14 < num10)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1] + num6, array2[0], text15);
				HTKKFile.SetNodeBackGroundPink(sheetView, array6[1] + num6, array6[0], text15);
				value = 2;
			}
			if (num15 < num11)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array3[1] + num6, array3[0], text16);
				HTKKFile.SetNodeBackGroundPink(sheetView, array7[1] + num6, array7[0], text16);
				value = 2;
			}
			if (num16 < num12)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array4[1] + num6, array4[0], text17);
				HTKKFile.SetNodeBackGroundPink(sheetView, array8[1] + num6, array8[0], text17);
				value = 2;
			}
			if (Convert.ToBoolean(sheetView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value) || Convert.ToBoolean(sheetView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value) || Convert.ToBoolean(sheetView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value))
			{
				string str = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str) && !sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text8);
					value = 2;
				}
				string str2 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str2) && !sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text8);
					value = 2;
				}
				string str3 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str3) && !sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, text8);
					value = 2;
				}
				string str4 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str4) && !sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, text8);
					value = 2;
				}
				if ((HTKKFormater.isNullString(str3) || sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked || !(TTinChung.maCQTQuanLy != TTinChung.maCQTCapCuc)) && !HTKKFormater.isNullString(str4) && !sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked && TTinChung.maCQTQuanLy == TTinChung.maCQTCapCuc)
				{
					string value3 = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
					if (!string.IsNullOrEmpty(value3) && !PeriodInfo.MaTinh.StartsWith(value3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, text18);
						value = 2;
					}
				}
				string value4 = sheetView.Cells[ct22[1] + num6, ct22[0]].Text.Trim();
				string text23 = sheetView.Cells[ct23[1] + num6, ct23[0]].Text.Trim();
				string value5 = sheetView.Cells[ct24[1] + num6, ct24[0]].Text.Trim();
				if (!string.IsNullOrEmpty(value4) || !string.IsNullOrEmpty(text23) || !string.IsNullOrEmpty(value5))
				{
					if (string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, ct22[1] + num6, ct22[0], text8);
						value = 2;
					}
					if (string.IsNullOrEmpty(text23))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, ct23[1] + num6, ct23[0], text8);
						value = 2;
					}
					if (string.IsNullOrEmpty(value5))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, ct24[1] + num6, ct24[0], text8);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(text23) && !ValidateBase.CheckMaSoThueCCCD(text23))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, ct23[1] + num6, ct23[0], text11);
					value = 2;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[1];
			for (int num17 = 1; num17 < ToKhaiInfo.SheetInfos[0].Tables.Count; num17++)
			{
				int dataLocationTemplateRowStart2 = ToKhaiInfo.SheetInfos[0].Tables[num17].Location.DataLocationTemplateRowStart;
				for (int num7 = 0; num7 < ToKhaiInfo.SheetInfos[0].Tables[num17].RowsValues.Count; num7++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + num7 - 1, num - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart2 + num7 - 1, num4 - 1);
				}
			}
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num7 = 1; num7 < hTKKCMCTableInfo.RowsValues.Count; num7++)
			{
				string valueCompare = sheetView.Cells[dataLocationTemplateRowStart3 + num7 - 1, num - 1].Text.Trim();
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num, valueCompare);
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart3 + num7 - 1, num - 1, text13);
					value = 2;
				}
			}
			for (int num17 = 2; num17 < ToKhaiInfo.SheetInfos[0].Tables.Count; num17++)
			{
				int dataLocationTemplateRowStart4 = ToKhaiInfo.SheetInfos[0].Tables[num17].Location.DataLocationTemplateRowStart;
				for (int num7 = 1; num7 < ToKhaiInfo.SheetInfos[0].Tables[num17].RowsValues.Count; num7++)
				{
					string item = sheetView.Cells[dataLocationTemplateRowStart4 + num7 - 1, num - 1].Text.Trim();
					double num18 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart4 + num7 - 1, num4 - 1].Value);
					List<int> list = new List<int>();
					list.Add(num);
					list.Add(num4);
					List<string> list2 = new List<string>();
					list2.Add(item);
					list2.Add(num18.ToString());
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, ToKhaiInfo.SheetInfos[0].Tables[num17], list, list2);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart4 + num7 - 1, num - 1, text21);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart4 + num7 - 1, num4 - 1, text21);
						value = 2;
					}
				}
			}
			if (!CheckChonPhuLucVaHinhThucKhongThoaMan(sheetView, PeriodInfo, ToKhaiInfo))
			{
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				value = 2;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[2];
			int num19 = ConvertHelpers.ConvertToInt(hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart);
			for (int num20 = 0; num20 < hTKKCMCTableInfo2.RowsValues.Count; num20++)
			{
				string valColMa = GeneralFunctions.GetValueCustom(sheetView, num19 + num20, HTKKFile.ColumnLettersToIndex("AB"), "TextCellType");
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, num19 + num20, HTKKFile.ColumnLettersToIndex("Q"), "NumberCellType");
				if (!HTKKFormater.isNullString(valColMa) && _dmKSData != null && _dmKSData.Count > 0)
				{
					item item2 = _dmKSData.Find((item x) => x.id.Contains(valColMa));
					if (item2 != null && (HTKKFormater.isNullString(item2.thuesuat) || ConvertHelpers.ConvertToInt(item2.thuesuat) == 0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num19 + num20 - 1, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
				}
				if (!HTKKFormater.isNullString(valColMa) && _dmKSData != null && _dmKSData.Count > 0 && valueCustom.ToNumber() != 0m)
				{
					bool flag = true;
					item item2 = _dmKSData.Find((item x) => x.id.Contains(valColMa));
					if (item2 != null && flag && valueCustom.ToNumber() < item2.giatoithieu.ToNumber() - item2.giatoithieu.ToNumber() * 20m / 100m)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num19 + num20 - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
						value = 2;
					}
				}
			}
			foreach (HTKKCMCTableInfo table in ToKhaiInfo.SheetInfos[0].Tables)
			{
				int dataLocationTemplateRowStart5 = table.Location.DataLocationTemplateRowStart;
				for (int num21 = 0; num21 < table.RowsValues.Count; num21++)
				{
					string text24 = sheetView.Cells[dataLocationTemplateRowStart5 + num21 - 1, num - 1].Text.Trim();
					if (text24.StartsWith("-----") || text24.StartsWith("---"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart5 + num21 - 1, num - 1, message);
						value = 2;
					}
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = ToKhaiInfo.SheetInfos[0].Tables[ToKhaiInfo.SheetInfos[0].Tables.Count - 1];
			int dataLocationTemplateRowStart6 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			for (int num21 = 0; num21 < hTKKCMCTableInfo3.RowsValues.Count; num21++)
			{
				string maLoaiKhoangSan = sheetView.Cells[dataLocationTemplateRowStart6 + num21 - 1, num5 - 1].Text.Trim();
				object value6 = sheetView.Cells[dataLocationTemplateRowStart6 + num21 - 1, num4 - 1].Value;
				if (string.IsNullOrEmpty(maLoaiKhoangSan))
				{
					continue;
				}
				string[] array9 = dicDMKSBVMT.Values.FirstOrDefault((string[] x) => x[1] == maLoaiKhoangSan);
				if (array9 != null && array9.Count() > 9 && Convert.ToDouble(array9[8]) > 0.0 && Convert.ToDouble(array9[9]) > 0.0 && (Convert.ToDouble(value6) > Convert.ToDouble(array9[9]) || Convert.ToDouble(value6) < Convert.ToDouble(array9[8])))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart6 + num21 - 1, num4 - 1, message2);
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
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[74, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[73, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		List<int> list = new List<int>();
		SheetView sheetView = winFpSpread.Sheets[1];
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int item = HTKKFile.ColumnLettersToIndex("W");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
			}
			List<string> list2 = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text8 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text;
				string text9 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text;
				string text10 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text;
				double num5 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
				string text11 = "";
				string item2 = GeneralFunctions.FormatKey(text9, text10);
				if (!list2.Contains(item2))
				{
					list2.Add(item2);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text5);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text5);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(text8))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, text3);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(text9))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text3);
					list.Add(1);
				}
				else if (!ValidateBase.CheckMaSoThueCCCD(text9))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text);
					list.Add(1);
				}
				if (num5 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text3);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(text10))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text3);
					list.Add(1);
				}
			}
			List<int> list3 = new List<int>();
			list3.Add(num4);
			List<int> list4 = new List<int>();
			list4.Add(item);
			int num6 = GeneralFunctions.ValidateCombobox(sheetView, hTKKCMCTableInfo, list3, list4, text7, text6);
			if (num6 > 0)
			{
				list.Add(1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		for (int i = 0; i < list.Count; i++)
		{
			if (list[i] == 1)
			{
				value = 2;
				break;
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		List<int> list = new List<int>();
		SheetView sheetView = winFpSpread.Sheets[2];
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("I");
			int num5 = HTKKFile.ColumnLettersToIndex("L");
			int num6 = HTKKFile.ColumnLettersToIndex("N");
			int num7 = HTKKFile.ColumnLettersToIndex("O");
			int num8 = HTKKFile.ColumnLettersToIndex("Q");
			int num9 = HTKKFile.ColumnLettersToIndex("R");
			int num10 = HTKKFile.ColumnLettersToIndex("T");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string str = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					string str2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text.Trim();
					double num11 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Value);
					double num12 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Value);
					double num13 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Value);
					double num14 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Value);
					double num15 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num7 - 1].Value);
					double num16 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Value);
					double num17 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Value);
					double num18 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Value);
					string text6 = "";
					if (HTKKFormater.isNullString(str))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, text3);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(str2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text3);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						list.Add(0);
					}
					if (num17 < 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						list.Add(0);
					}
					if (num18 < 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
						list.Add(0);
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					value = 2;
					break;
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

	public Dictionary<string, string> GetDMKhoangSan(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
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
				string text = array[2];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKhoangSanNew(List<PathDm> pathDms, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0383: Unknown result type (might be due to invalid IL or missing references)
		//IL_038a: Expected O, but got Unknown
		//IL_03b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bf: Expected O, but got Unknown
		string text = "###0103###020101###020302###020303###0301###0302###0303###0304###0401###0402###0403###0404###0405###0702###07030101###07030102###07030201###07030202###090201###090202###090204###090205###";
		try
		{
			_isDmMoi = true;
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
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
						if (!namKK.Equals("2019") || !kieuKyKK.Equals("Y"))
						{
							list.Add(item3);
						}
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				List<item> dmTaiNguyenForTinh = GeneralFunctions.getDmTaiNguyenForTinh(pathDonGiaTinh, PeriodInfo);
				List<item> list3 = Serializer.DeepCopy(list);
				foreach (item iT in list)
				{
					Predicate<item> match2 = (item obj) => obj.ma.Equals(iT.ma);
					List<item> list4 = dmTaiNguyenForTinh.FindAll(match2);
					if (list4 != null && list4.Count > 0)
					{
						item item4 = list3.Find((item obj) => obj.id == iT.id);
						int num2 = list3.FindIndex((item obj) => obj.id == iT.id);
						if (item4 != null && PeriodInfo.KieuKy == "M")
						{
							list3.Remove(item4);
						}
						bool flag = false;
						for (int num3 = 0; num3 < list4.Count; num3++)
						{
							item item5 = Serializer.DeepCopy(iT);
							item5.cqtgia = list4[num3].cqtgia;
							item5.isCqt = true;
							list3.Insert(num2 + num3, item5);
						}
						list3.Remove(item4);
					}
				}
				list = Serializer.DeepCopy(list3);
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
					string text2 = item6.id + "_" + item6.ma + "_" + item6.thuesuat + "_" + item6.giatoithieu + "_" + item6.cqtgia;
					string[] array = new string[10] { text2, item6.ma, item6.ten, item6.dvt, item6.thuesuat, "0", "1", item6.matieumuc, item6.giatoithieu, item6.giatoida };
					item6.id = text2;
					if (!dictionary.ContainsKey(text2))
					{
						string value2 = array[2];
						dictionary2.Add(text2, value2);
						dictionary.Add(text2, array);
						_dmKSData.Add(item6);
					}
				}
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
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

	private bool kyKeKhaiTrongHL(item itemTinh, HTKKPeriodInfo PeriodInfo)
	{
		bool result = false;
		try
		{
			DateTime[] array = new DateTime[2];
			DateTime[] array2 = new DateTime[2];
			if (HTKKFormater.isNullString(itemTinh.cqthltu))
			{
				return false;
			}
			ref DateTime reference = ref array2[0];
			reference = DateTime.ParseExact(itemTinh.cqthltu, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (HTKKFormater.isNullString(itemTinh.cqthlden))
			{
				return false;
			}
			ref DateTime reference2 = ref array2[1];
			reference2 = DateTime.ParseExact(itemTinh.cqthlden, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			string nam = PeriodInfo.Nam;
			string tuNgay = PeriodInfo.TuNgay;
			DateTime dateTime;
			if (!HTKKFormater.isNullString(tuNgay))
			{
				dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			}
			else
			{
				tuNgay = PeriodInfo.TuThang;
				dateTime = (HTKKFormater.isNullString(tuNgay) ? DateTime.ParseExact("01/01" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture) : DateTime.ParseExact("01/" + tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture));
			}
			array[0] = dateTime;
			tuNgay = PeriodInfo.DenNgay;
			if (!HTKKFormater.isNullString(tuNgay))
			{
				dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			}
			else
			{
				tuNgay = PeriodInfo.DenThang;
				if (!HTKKFormater.isNullString(tuNgay))
				{
					dateTime = DateTime.ParseExact("01/" + tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				}
				else
				{
					dateTime = DateTime.ParseExact("01/12" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				}
			}
			array[1] = dateTime;
			if (array[0] >= array2[0] && array[1] <= array2[1])
			{
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return result;
	}

	public Dictionary<string, string> GetDMKSBVMT(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
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
				string text = array[2];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKSBVMTNew(List<PathDm> pathDms)
	{
		//IL_02ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f4: Expected O, but got Unknown
		//IL_0322: Unknown result type (might be due to invalid IL or missing references)
		//IL_0329: Expected O, but got Unknown
		try
		{
			_isDmMoi = true;
			if (dicDMKSBVMTTmp.Count > 0)
			{
				return dicDMKSBVMTTmp;
			}
			_dmBVMTData = new List<item>();
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
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
							{
								list.Add(item);
							}
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
				foreach (XmlNode item4 in elementsByTagName)
				{
					XmlNode val2 = item4;
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
						if (!namKK.Equals("2019") || !kieuKyKK.Equals("Y"))
						{
							list.Add(item3);
						}
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (item item5 in list)
				{
					string ten = item5.ten;
					List<item> list2 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list2 != null && list2.Count > 1)
					{
						ten = ten + "( " + item5.ma + " )";
					}
					string text = item5.id + "_" + item5.ma + "_" + item5.thuesuat + "_" + item5.giatoithieu + "_" + item5.cqtgia;
					string[] array = new string[10] { text, item5.ma, ten, item5.dvt, item5.thuesuat, "0", "1", item5.matieumuc, item5.giatoithieu, item5.giatoida };
					item5.id = text;
					if (!dictionary.ContainsKey(text))
					{
						string text2 = array[2];
						dictionary2.Add(text, text2);
						if (!dictionary.ContainsKey(text2))
						{
							dictionary.Add(text2, array);
						}
						_dmBVMTData.Add(item5);
					}
				}
			}
			dicDMKSBVMT = dictionary;
			dicDMKSBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<item> GetItemDm_GetDMKSBVMT()
	{
		return _dmBVMTData;
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

	public static void RemoveDataRows(SheetView sView, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		if (rowCount == 0)
		{
			return;
		}
		sView.Rows.Remove(frRow, rowCount);
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > frRow + 1)
			{
				table.Increment(rowCount * -1);
			}
		}
		sInfo.IncrementStaticCell(frRow, rowCount * -1);
	}

	public static void fpSpReindexingTableTemplate(SheetView sView, HTKKCMCTableInfo tableInfo)
	{
		try
		{
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				_ = tableInfo.iColumnSTT;
				bool flag = 1 == 0;
				if (tableInfo.CellTableSTT != null)
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = tableInfo.CellTableSTT.Value + "." + (i + 1);
				}
				else
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = (i + 1).ToString();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMThueBVMT(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0319: Unknown result type (might be due to invalid IL or missing references)
		//IL_0320: Expected O, but got Unknown
		//IL_0360: Unknown result type (might be due to invalid IL or missing references)
		//IL_0367: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012027#102026#052026#042026#032026#012026#072025#012025#012024#012023#072022#042022#112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			string[] array = text2.Split(new char[1] { '#' });
			DateTime t = new DateTime(1900, 1, 1);
			int num = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMThueBVMTTmp.Count > 0)
			{
				return dicDMThueBVMTTmp;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
				DateTime t2 = new DateTime(num, month, 1).AddMonths(1).AddDays(-1.0);
				for (int i = 0; i < array.Length; i++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = array[i];
					}
				}
			}
			else if (PeriodInfo.KieuKy == "Y")
			{
				DateTime t2 = new DateTime(num, 12, 1).AddMonths(1).AddDays(-1.0);
				if (PeriodInfo.Nam == "2020")
				{
					text = "012020";
				}
				else
				{
					for (int i = 0; i < array.Length; i++)
					{
						DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
						if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
						{
							t = dateTime;
							text = array[i];
						}
					}
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				int iQuy = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				DateTime t2 = GeneralFunctions.GetNgayDauQuy(iQuy, num, 1, 1).AddMonths(1).AddDays(-1.0);
				for (int i = 0; i < array.Length; i++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(array[i].Substring(2)), ConvertHelpers.ConvertToInt(array[i].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = array[i];
					}
				}
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string key = array2[0];
				string value2 = array2[1];
				dictionary2.Add(key, value2);
				dictionary.Add(array2[1], array2);
			}
			dicDMThueBVMT = dictionary;
			dicDMThueBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMNganhNgheKD(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (dicDMNganhNgheKDTmp.Count > 0)
			{
				return dicDMNganhNgheKDTmp;
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
				if (array.Length > 2 && array[2] == "1")
				{
					string key = array[0];
					string value2 = array[1];
					dictionary2.Add(key, value2);
					dictionary.Add(array[1], array);
				}
			}
			dicDMNganhNgheKD = dictionary;
			dicDMNganhNgheKDTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMNhomNganhNgheKD(string pathFile)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			string text = pathFile + "\\DM_NHOMNGANHNGHE.xml";
			if (dicDMNhomNganhNgheKDTmp.Count > 0)
			{
				return dicDMNhomNganhNgheKDTmp;
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
			dicDMNhomNganhNgheKD = dictionary;
			dicDMNhomNganhNgheKDTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
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
				if (array.Length > 4)
				{
					string key = array[1];
					string text = array[2].Trim() + " (" + array[4].Trim() + "%)";
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
		//IL_049d: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a4: Expected O, but got Unknown
		//IL_04d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d9: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> temp = new Dictionary<string, string[]>();
			Dictionary<string, string> dic = new Dictionary<string, string>();
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
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								if (!string.IsNullOrEmpty(item.thuesuat))
								{
									text = text + " (" + item.thuesuat + " %)";
								}
								GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
								if (!dic.ContainsKey(key))
								{
									dic.Add(key, text);
								}
								if (!temp.ContainsKey(text))
								{
									temp.Add(text, array);
								}
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							if (!string.IsNullOrEmpty(item.thuesuat))
							{
								text = text + " (" + item.thuesuat + " %)";
							}
							GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
							if (!dic.ContainsKey(key))
							{
								dic.Add(key, text);
							}
							if (!temp.ContainsKey(text))
							{
								temp.Add(text, array);
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
					if (array.Length > 2 && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						string text2 = array[4].Trim();
						if (!string.IsNullOrEmpty(text2))
						{
							text = text + " (" + text2 + " %)";
						}
						if (!dic.ContainsKey(key))
						{
							dic.Add(key, text);
						}
						if (!temp.ContainsKey(text))
						{
							temp.Add(text, array);
						}
					}
				}
			}
			dicDMThueDB = temp;
			dicDMThueDBTmp = dic;
			return dic;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void Refresh_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int actionType)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			CheckPhuLuc(winFpSpread, ToKhaiInfo);
			ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
			switch (actionType)
			{
			}
			CheckChonPhuLucVaHinhThucKhongThoaMan(winFpSpread.Sheets[0], _PeriodInfo, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int actionType)
	{
		try
		{
			CheckChonPhuLucVaHinhThucKhongThoaMan(winFpSpread.Sheets[0], _PeriodInfo, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellPhuLuc01_3[1], cellPhuLuc01_3[0]].Value = ((ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellPhuLuc01_4[1], cellPhuLuc01_4[0]].Value = ((ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellPhuLuc01_5[1], cellPhuLuc01_5[0]].Value = ((ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
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
			CheckPhuLuc(winFpSpread, ToKhaiInfo);
			ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc01_3[1], cellPhuLuc01_3[0]].Value = ((ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc01_4[1], cellPhuLuc01_4[0]].Value = ((ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
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
			winFpSpread.Sheets[0].Cells[cellPhuLuc01_5[1], cellPhuLuc01_5[0]].Value = ((ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("AB");
			int num2 = HTKKFile.ColumnLettersToIndex("AE");
			for (int i = 1; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					string value = "";
					if (!string.IsNullOrEmpty(text))
					{
						value = GeneralFunctions.GetMaTieuMucByMaHH(text, "TTDB");
					}
					sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Value = value;
				}
			}
			TinhLaiThueSuatPL01_4(winFpSpread.Sheets[4], ToKhaiInfo);
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

	public void validateTaiNguyen(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			if (iCol != HTKKFile.ColumnLettersToIndex("Q"))
			{
				return;
			}
			string valColMa = GeneralFunctions.GetValueCustom(sView, (int)iRow, HTKKFile.ColumnLettersToIndex("AB"), "TextCellType");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, (int)iRow, HTKKFile.ColumnLettersToIndex("Q"), "NumberCellType");
			HTKKFile.RemoveValidate(sView, (int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1);
			if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || !(valueCustom.ToNumber() != 0m))
			{
				return;
			}
			bool flag = true;
			item item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
			if (item == null)
			{
				return;
			}
			if (item.cqtgia.ToNumber() > 0m)
			{
				flag = false;
				if (valueCustom.ToNumber() < item.cqtgia.ToNumber())
				{
					HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")));
				}
			}
			if (flag && valueCustom.ToNumber() < item.giatoithieu.ToNumber())
			{
				HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_34")));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ThayDoiToKhai28Den31_a_c(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[1];
		if (sheetView != null && ToKhaiInfo.SheetInfos[1] != null)
		{
			SheetView sheetView2 = winFpSpread.Sheets[0];
			int num = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("W");
			int num8 = HTKKFile.ColumnLettersToIndex("AJ");
			int num9 = HTKKFile.ColumnLettersToIndex("AI");
			sheetView2.Cells[ct28a[1] + num, ct28a[0]].Value = sheetView2.Cells[ct28a[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct28c[1] + num, ct28c[0]].Value = sheetView2.Cells[ct28c[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct29a[1] + num, ct29a[0]].Value = sheetView2.Cells[ct29a[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct29c[1] + num, ct29c[0]].Value = sheetView2.Cells[ct29a[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct30a[1] + num, ct30a[0]].Value = sheetView2.Cells[ct30a[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct30c[1] + num, ct30c[0]].Value = sheetView2.Cells[ct30a[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct31a[1] + num, ct31a[0]].Value = sheetView2.Cells[ct31a[1] + num, num8 - 1].Value;
			sheetView2.Cells[ct31c[1] + num, ct31c[0]].Value = sheetView2.Cells[ct31a[1] + num, num8 - 1].Value;
		}
	}

	public void ClearDataCt28DenCt31(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		int num = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
		SheetView sheetView = winFpSpread.Sheets[0];
		sheetView.Cells[ct28a[1] + num, ct28a[0]].Value = 0;
		sheetView.Cells[ct29a[1] + num, ct29a[0]].Value = 0;
		sheetView.Cells[ct30a[1] + num, ct30a[0]].Value = 0;
		sheetView.Cells[ct31a[1] + num, ct31a[0]].Value = 0;
		sheetView.Cells[ct28b[1] + num, ct28b[0]].Value = 0;
		sheetView.Cells[ct29b[1] + num, ct29b[0]].Value = 0;
		sheetView.Cells[ct30b[1] + num, ct30b[0]].Value = 0;
		sheetView.Cells[ct31b[1] + num, ct31b[0]].Value = 0;
		sheetView.Cells[ct28c[1] + num, ct28c[0]].Value = 0;
		sheetView.Cells[ct29c[1] + num, ct29c[0]].Value = 0;
		sheetView.Cells[ct30c[1] + num, ct30c[0]].Value = 0;
		sheetView.Cells[ct31c[1] + num, ct31c[0]].Value = 0;
		sheetView.Cells[ct28d[1] + num, ct28d[0]].Value = 0;
		sheetView.Cells[ct29d[1] + num, ct29d[0]].Value = 0;
		sheetView.Cells[ct30d[1] + num, ct30d[0]].Value = 0;
		sheetView.Cells[ct31d[1] + num, ct31d[0]].Value = 0;
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
		ThayDoiToKhai28Den31_a_c(winFpSpread, ToKhaiInfo);
		TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		return true;
	}

	public bool UpdateFormula_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
		TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		SheetView sheetView = winFpSpread.Sheets[3];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
		if (hTKKCMCSheetInfo != null)
		{
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		}
		return true;
	}

	public bool UpdateFormula_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
		TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		return true;
	}

	public bool UpdateFormula_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
		TinhLaiThueSuatPL01_5(winFpSpread.Sheets[5], ToKhaiInfo);
		TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
		return true;
	}

	public void ThayDoiCacHinhThucKeKhai(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, bool isDefault)
	{
		int num = HTKKFile.ColumnLettersToIndex("F");
		int num2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
		if (PeriodInfo.KieuKy == "Y")
		{
			sView.Cells[ct22[1] + num2, ct22[0]].Locked = true;
			sView.Cells[ct23[1] + num2, ct23[0]].Locked = true;
			sView.Cells[ct24[1] + num2, ct24[0]].Locked = true;
			sView.Cells[ct25[1] + num2, ct25[0]].Locked = true;
			sView.Cells[ct26[1] + num2, ct26[0]].Locked = true;
			sView.Cells[ct27[1] + num2, ct27[0]].Locked = true;
		}
		if (Convert.ToBoolean(sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value))
		{
			LockAndDeleteNN(sView, ToKhaiInfo, isLock: true, isDelete: true);
			num2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			LockAndDeleteABC(sView, ToKhaiInfo, isLock: false, isDelete: false);
			for (int i = 0; i < ToKhaiInfo.SheetInfos[0].StaticCells.Count; i++)
			{
				int rowIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.RowIndex;
				int colIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.ColIndex;
				string defaultValue = ToKhaiInfo.SheetInfos[0].StaticCells[i].DefaultValue;
				if (rowIndex >= ct05[1] + 1 && rowIndex < ct17_1[1] + 2 + num2)
				{
					sView.Cells[rowIndex - 1, colIndex - 1].Value = defaultValue;
					sView.Cells[rowIndex - 1, colIndex - 1].Locked = true;
				}
			}
			if (isDefault)
			{
				sView.Cells[ct04[1], ct04[0]].Value = "";
				sView.Cells[ct07[1], ct07[0]].Value = "";
				sView.Cells[ct14[1] + num2, ct14[0]].Value = "";
				sView.Cells[ct15[1] + num2, ct15[0]].Value = "";
				sView.Cells[ct16[1] + num2, ct16[0]].Value = "";
				sView.Cells[ct22[1] + num2, ct22[0]].Value = _HeaderInfo.tenNNT;
				sView.Cells[ct23[1] + num2, ct23[0]].Value = _HeaderInfo.mst;
				sView.Cells[ct24[1] + num2, ct24[0]].Value = _HeaderInfo.dchiNNT + " - " + _HeaderInfo.tenHuyenNNT + " - " + _HeaderInfo.tenTinhNNT;
				sView.Cells[ct25[1] + num2, ct25[0]].Value = _HeaderInfo.dthoaiNNT;
				sView.Cells[ct26[1] + num2, ct26[0]].Value = _HeaderInfo.faxNNT;
				sView.Cells[ct27[1] + num2, ct27[0]].Value = _HeaderInfo.emailNNT;
			}
			sView.Cells[ct22[1] + num2, ct22[0]].Locked = true;
			sView.Cells[ct23[1] + num2, ct23[0]].Locked = true;
			sView.Cells[ct24[1] + num2, ct24[0]].Locked = true;
			sView.Cells[ct25[1] + num2, ct25[0]].Locked = true;
			sView.Cells[ct26[1] + num2, ct26[0]].Locked = true;
			sView.Cells[ct27[1] + num2, ct27[0]].Locked = true;
		}
		else if (Convert.ToBoolean(sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value) && Convert.ToBoolean(sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Value))
		{
			LockAndDeleteNN(sView, ToKhaiInfo, isLock: false, isDelete: false);
			num2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[0].StaticCells.Count; i++)
			{
				int rowIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.RowIndex;
				int colIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.ColIndex;
				string defaultValue = ToKhaiInfo.SheetInfos[0].StaticCells[i].DefaultValue;
				if (rowIndex >= ct05[1] + 1 && rowIndex < ct17_1[1] + 2 + num2)
				{
					sView.Cells[rowIndex - 1, colIndex - 1].Locked = false;
				}
			}
			LockAndDeleteABC(sView, ToKhaiInfo, isLock: true, isDelete: true);
			if (isDefault)
			{
				sView.Cells[ct04[1], ct04[0]].Value = _HeaderInfo.tenNNT;
				sView.Cells[ct07[1], ct07[0]].Value = _HeaderInfo.mst;
				sView.Cells[ct14[1] + num2, ct14[0]].Value = _HeaderInfo.dthoaiNNT;
				sView.Cells[ct15[1] + num2, ct15[0]].Value = _HeaderInfo.faxNNT;
				sView.Cells[ct16[1] + num2, ct16[0]].Value = _HeaderInfo.emailNNT;
				sView.Cells[ct22[1] + num2, ct22[0]].Value = "";
				sView.Cells[ct23[1] + num2, ct23[0]].Value = "";
				sView.Cells[ct24[1] + num2, ct24[0]].Value = "";
				sView.Cells[ct25[1] + num2, ct25[0]].Value = "";
				sView.Cells[ct26[1] + num2, ct26[0]].Value = "";
				sView.Cells[ct27[1] + num2, ct27[0]].Value = "";
			}
		}
		else if (Convert.ToBoolean(sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value) || Convert.ToBoolean(sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value) || Convert.ToBoolean(sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value))
		{
			num2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[0].StaticCells.Count; i++)
			{
				int rowIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.RowIndex;
				int colIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.ColIndex;
				string defaultValue = ToKhaiInfo.SheetInfos[0].StaticCells[i].DefaultValue;
				if (rowIndex >= ct05[1] + 1 && rowIndex < ct17_1[1] + 2 + num2)
				{
					sView.Cells[rowIndex - 1, colIndex - 1].Locked = false;
				}
			}
			sView.Cells[ct05[1], ct05[0]].Locked = false;
			sView.Cells[ct09[1], ct09[0]].Locked = false;
			LockAndDeleteNN(sView, ToKhaiInfo, isLock: false, isDelete: false);
			LockAndDeleteABC(sView, ToKhaiInfo, isLock: false, isDelete: false);
			if (isDefault)
			{
				sView.Cells[ct04[1], ct04[0]].Value = _HeaderInfo.tenNNT;
				sView.Cells[ct07[1], ct07[0]].Value = _HeaderInfo.mst;
				sView.Cells[ct14[1] + num2, ct14[0]].Value = _HeaderInfo.dthoaiNNT;
				sView.Cells[ct15[1] + num2, ct15[0]].Value = _HeaderInfo.faxNNT;
				sView.Cells[ct16[1] + num2, ct16[0]].Value = _HeaderInfo.emailNNT;
				sView.Cells[ct22[1] + num2, ct22[0]].Value = "";
				sView.Cells[ct23[1] + num2, ct23[0]].Value = "";
				sView.Cells[ct24[1] + num2, ct24[0]].Value = "";
				sView.Cells[ct25[1] + num2, ct25[0]].Value = "";
				sView.Cells[ct26[1] + num2, ct26[0]].Value = "";
				sView.Cells[ct27[1] + num2, ct27[0]].Value = "";
			}
		}
		sView.Cells[ct07[1], ct07[0]].Locked = true;
		sView.Cells[ct14[1] + num2, ct14[0]].Locked = true;
		sView.Cells[ct15[1] + num2, ct15[0]].Locked = true;
		sView.Cells[ct16[1] + num2, ct16[0]].Locked = true;
		int[] array = GeneralFunctions.ArrayCellId("H_76");
		int[] array2 = GeneralFunctions.ArrayCellId("H_88");
		sView.Cells[array[1] + num2, array[0]].Locked = true;
		sView.Cells[array2[1] + num2, array2[0]].Locked = true;
	}

	public void RangBuocKyTinhThueHinhThucKeKhai(SheetView sView, HTKKPeriodInfo PeriodInfo, bool isDefault)
	{
		if (PeriodInfo.KieuKy == "Y")
		{
			if (isDefault)
			{
				SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: true, valueCNKDNopPs: false, valueTCCNKhaiNopThay: false, valueHKD_CNKDNopKeKhai: false, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			}
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: true, TCCNKhaiNopThayShow: true, HKD_CNKDNopKeKhaiShow: true, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: false);
		}
		if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "Q")
		{
			if (isDefault)
			{
				SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: false, valueCNKDNopPs: false, valueTCCNKhaiNopThay: false, valueHKD_CNKDNopKeKhai: false, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			}
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: true, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: true);
		}
		if (PeriodInfo.KieuKy == "D")
		{
			if (isDefault)
			{
				SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: false, valueCNKDNopPs: false, valueTCCNKhaiNopThay: false, valueHKD_CNKDNopKeKhai: false, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			}
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: true, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: true);
		}
	}

	public void RangBuocGiuaHinhThucKeKhai(SheetView sView, HTKKPeriodInfo PeriodInfo, int iRow)
	{
		bool flag = Convert.ToBoolean(sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value);
		bool flag3 = Convert.ToBoolean(sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value);
		bool flag4 = Convert.ToBoolean(sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value);
		bool flag5 = Convert.ToBoolean(sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Value);
		bool flag6 = Convert.ToBoolean(sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Value);
		if (iRow == HKD_CNKDNopKeKhai[1] + 1 && flag4 && (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "Q"))
		{
			SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: false, valueCNKDNopPs: false, valueTCCNKhaiNopThay: false, valueHKD_CNKDNopKeKhai: true, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: true, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: true);
		}
		if (iRow == TCCNKhaiNopThay[1] + 1 && flag3 && (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "Q"))
		{
			SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: false, valueCNKDNopPs: false, valueTCCNKhaiNopThay: true, valueHKD_CNKDNopKeKhai: false, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: true, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: true);
		}
		if (iRow == CNKDNopPs[1] + 1 && flag2 && PeriodInfo.KieuKy == "D")
		{
			SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: false, valueCNKDNopPs: true, valueTCCNKhaiNopThay: false, valueHKD_CNKDNopKeKhai: false, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: true, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: true);
		}
		if (iRow == TCCNKhaiNopThay[1] + 1 && flag3 && PeriodInfo.KieuKy == "D")
		{
			SetHinhThucKeKhai(sView, valueHKD_CNKDNopThueKhoan: false, valueCNKDNopPs: false, valueTCCNKhaiNopThay: true, valueHKD_CNKDNopKeKhai: false, valueHKD_CNKD_NNXDDoanhThu: false, valueHKDChuyenDoiPpTinhThue: false);
			LockUnlockHinhThucKeKhai(sView, HKD_CNKDNopThueKhoanShow: true, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: true, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: true);
		}
	}

	public void LockUnlockHinhThucKeKhai(SheetView sView, bool HKD_CNKDNopThueKhoanShow, bool CNKDNopPsShow, bool TCCNKhaiNopThayShow, bool HKD_CNKDNopKeKhaiShow, bool HKD_CNKD_NNXDDoanhThuShow, bool HKDChuyenDoiPpTinhThueShow)
	{
		sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Locked = HKD_CNKDNopThueKhoanShow;
		sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Locked = CNKDNopPsShow;
		sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Locked = TCCNKhaiNopThayShow;
		sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Locked = HKD_CNKDNopKeKhaiShow;
		sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Locked = HKD_CNKD_NNXDDoanhThuShow;
		sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Locked = HKDChuyenDoiPpTinhThueShow;
	}

	public void SetHinhThucKeKhai(SheetView sView, bool valueHKD_CNKDNopThueKhoan, bool valueCNKDNopPs, bool valueTCCNKhaiNopThay, bool valueHKD_CNKDNopKeKhai, bool valueHKD_CNKD_NNXDDoanhThu, bool valueHKDChuyenDoiPpTinhThue)
	{
		sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value = valueHKD_CNKDNopThueKhoan;
		sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value = valueCNKDNopPs;
		sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value = valueTCCNKhaiNopThay;
		sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value = valueHKD_CNKDNopKeKhai;
		sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Value = valueHKD_CNKD_NNXDDoanhThu;
		sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Value = valueHKDChuyenDoiPpTinhThue;
	}

	public bool CheckChonPhuLucVaHinhThucKhongThoaMan(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		bool flag = Convert.ToBoolean(sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value);
		bool flag3 = Convert.ToBoolean(sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value);
		bool flag4 = Convert.ToBoolean(sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value);
		bool flag5 = Convert.ToBoolean(sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Value);
		bool flag6 = Convert.ToBoolean(sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Value);
		HTKKFile.RemoveValidate(sView, HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]);
		HTKKFile.RemoveValidate(sView, CNKDNopPs[1], CNKDNopPs[0]);
		HTKKFile.RemoveValidate(sView, HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]);
		HTKKFile.RemoveValidate(sView, HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]);
		HTKKFile.RemoveValidate(sView, HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]);
		if (ToKhaiInfo.SheetInfos[1] != null && ToKhaiInfo.SheetInfos[2] != null)
		{
			string message = _xmlMessage.GetMessage("0532");
			ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: true);
			return false;
		}
		if (!flag && !flag2 && !flag3 && !flag4 && !flag5 && !flag6)
		{
			string message = _xmlMessage.GetMessage("0565");
			ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: true);
			return false;
		}
		if (flag5 && !flag && !flag2 && !flag3 && !flag4 && !flag6)
		{
			string message = _xmlMessage.GetMessage("0560");
			ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: true, HKDChuyenDoiPpTinhThueShow: false, all: false);
			return false;
		}
		if (flag6 && !flag && !flag2 && !flag3 && !flag4 && !flag5)
		{
			string message = _xmlMessage.GetMessage("0561");
			ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: true, all: false);
			return false;
		}
		if (flag3)
		{
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				string message = _xmlMessage.GetMessage("0562");
				ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: true, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: false);
				return false;
			}
			if (ToKhaiInfo.SheetInfos[4] != null)
			{
				string message = _xmlMessage.GetMessage("0636");
				ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: true, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: false);
				return false;
			}
			if (ToKhaiInfo.SheetInfos[5] != null)
			{
				string message = _xmlMessage.GetMessage("0648");
				ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: true, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: false);
				return false;
			}
		}
		else if (!flag && !flag2 && !flag3 && flag4 && !flag5 && !flag6)
		{
			if (ToKhaiInfo.SheetInfos[2] == null)
			{
				string message = _xmlMessage.GetMessage("0563");
				ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: true, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: false);
				return false;
			}
		}
		else if ((flag || flag2 || flag3 || !flag4 || !flag5 || flag6) && (ToKhaiInfo.SheetInfos[1] != null || ToKhaiInfo.SheetInfos[2] != null))
		{
			string message = _xmlMessage.GetMessage("0564");
			ThongBaoVeChonHinhThuc(sView, message, HKD_CNKDNopThueKhoanShow: false, CNKDNopPsShow: false, TCCNKhaiNopThayShow: false, HKD_CNKDNopKeKhaiShow: false, HKD_CNKD_NNXDDoanhThuShow: false, HKDChuyenDoiPpTinhThueShow: false, all: true);
			return false;
		}
		return true;
	}

	public void ThongBaoVeChonHinhThuc(SheetView sView, string message, bool HKD_CNKDNopThueKhoanShow, bool CNKDNopPsShow, bool TCCNKhaiNopThayShow, bool HKD_CNKDNopKeKhaiShow, bool HKD_CNKD_NNXDDoanhThuShow, bool HKDChuyenDoiPpTinhThueShow, bool all)
	{
		bool flag = false;
		if (Convert.ToBoolean(sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value) || Convert.ToBoolean(sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value) || Convert.ToBoolean(sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value) || Convert.ToBoolean(sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value) || Convert.ToBoolean(sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Value) || Convert.ToBoolean(sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Value))
		{
			flag = true;
		}
		if (flag)
		{
			if (HKD_CNKDNopThueKhoanShow || (all && Convert.ToBoolean(sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Value)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0], message);
			}
			if (CNKDNopPsShow || (all && Convert.ToBoolean(sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Value)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CNKDNopPs[1], CNKDNopPs[0], message);
			}
			if (TCCNKhaiNopThayShow || (all && Convert.ToBoolean(sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Value)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, TCCNKhaiNopThay[1], TCCNKhaiNopThay[0], message);
			}
			if (HKD_CNKDNopKeKhaiShow || (all && Convert.ToBoolean(sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Value)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0], message);
			}
			if (HKD_CNKD_NNXDDoanhThuShow || (all && Convert.ToBoolean(sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Value)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0], message);
			}
			if (HKDChuyenDoiPpTinhThueShow || (all && Convert.ToBoolean(sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Value)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0], message);
			}
		}
		else
		{
			if (HKD_CNKDNopThueKhoanShow || (all && !sView.Cells[HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0]].Locked))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKD_CNKDNopThueKhoan[1], HKD_CNKDNopThueKhoan[0], message);
			}
			if (CNKDNopPsShow || (all && !sView.Cells[CNKDNopPs[1], CNKDNopPs[0]].Locked))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CNKDNopPs[1], CNKDNopPs[0], message);
			}
			if (TCCNKhaiNopThayShow || (all && !sView.Cells[TCCNKhaiNopThay[1], TCCNKhaiNopThay[0]].Locked))
			{
				HTKKFile.SetNodeBackGroundPink(sView, TCCNKhaiNopThay[1], TCCNKhaiNopThay[0], message);
			}
			if (HKD_CNKDNopKeKhaiShow || (all && !sView.Cells[HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0]].Locked))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKD_CNKDNopKeKhai[1], HKD_CNKDNopKeKhai[0], message);
			}
			if (HKD_CNKD_NNXDDoanhThuShow || (all && !sView.Cells[HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0]].Locked))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKD_CNKD_NNXDDoanhThu[1], HKD_CNKD_NNXDDoanhThu[0], message);
			}
			if (HKDChuyenDoiPpTinhThueShow || (all && !sView.Cells[HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0]].Locked))
			{
				HTKKFile.SetNodeBackGroundPink(sView, HKDChuyenDoiPpTinhThue[1], HKDChuyenDoiPpTinhThue[0], message);
			}
		}
	}

	public void LockAndDeleteNN(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, bool isLock, bool isDelete)
	{
		int num = HTKKFile.ColumnLettersToIndex("F");
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (isDelete)
		{
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				GeneralFunctions.RemoveDataRows(sView, 0, dataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, ToKhaiInfo.SheetInfos[0]);
			}
			sView.Cells[dataLocationTemplateRowStart - 1, num - 1].Value = "";
			sView.Cells[dataLocationTemplateRowStart - 1, colAC - 1].Value = "";
			sView.Cells[dataLocationTemplateRowStart - 1, num - 1].Locked = true;
		}
		if (isLock)
		{
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Locked = true;
			}
		}
		else
		{
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Locked = false;
			}
		}
	}

	public void LockAndDeleteABC(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, bool isLock, bool isDelete)
	{
		int num = HTKKFile.ColumnLettersToIndex("C");
		int num2 = HTKKFile.ColumnLettersToIndex("M");
		int num3 = HTKKFile.ColumnLettersToIndex("P");
		int num4 = HTKKFile.ColumnLettersToIndex("Q");
		int num5 = HTKKFile.ColumnLettersToIndex("X");
		int num6 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
		sView.Cells[ct28a[1] + num6, ct28a[0]].Locked = (isLock ? true : false);
		sView.Cells[ct29a[1] + num6, ct29a[0]].Locked = (isLock ? true : false);
		sView.Cells[ct30a[1] + num6, ct30a[0]].Locked = (isLock ? true : false);
		sView.Cells[ct31a[1] + num6, ct31a[0]].Locked = (isLock ? true : false);
		sView.Cells[ct29c[1] + num6, ct29c[0]].Locked = (isLock ? true : false);
		sView.Cells[ct30c[1] + num6, ct30c[0]].Locked = (isLock ? true : false);
		sView.Cells[ct31c[1] + num6, ct31c[0]].Locked = (isLock ? true : false);
		if (sView.Cells[ct28a[1] + num6, ct28a[0]].Locked)
		{
			sView.Cells[ct28a[1] + num6, ct28a[0]].Value = 0;
		}
		if (sView.Cells[ct29a[1] + num6, ct29a[0]].Locked)
		{
			sView.Cells[ct29a[1] + num6, ct29a[0]].Value = 0;
		}
		if (sView.Cells[ct30a[1] + num6, ct30a[0]].Locked)
		{
			sView.Cells[ct30a[1] + num6, ct30a[0]].Value = 0;
		}
		if (sView.Cells[ct31a[1] + num6, ct31a[0]].Locked)
		{
			sView.Cells[ct31a[1] + num6, ct31a[0]].Value = 0;
		}
		if (sView.Cells[ct29c[1] + num6, ct29c[0]].Locked)
		{
			sView.Cells[ct29c[1] + num6, ct29c[0]].Value = 0;
		}
		if (sView.Cells[ct30c[1] + num6, ct30c[0]].Locked)
		{
			sView.Cells[ct30c[1] + num6, ct30c[0]].Value = 0;
		}
		if (sView.Cells[ct31c[1] + num6, ct31c[0]].Locked)
		{
			sView.Cells[ct31c[1] + num6, ct31c[0]].Value = 0;
		}
		if (isDelete)
		{
			sView.Cells[ct28a[1] + num6, ct28a[0]].Value = 0;
			sView.Cells[ct29a[1] + num6, ct29a[0]].Value = 0;
			sView.Cells[ct30a[1] + num6, ct30a[0]].Value = 0;
			sView.Cells[ct31a[1] + num6, ct31a[0]].Value = 0;
			sView.Cells[ct28c[1] + num6, ct28c[0]].Value = 0;
			sView.Cells[ct29c[1] + num6, ct29c[0]].Value = 0;
			sView.Cells[ct30c[1] + num6, ct30c[0]].Value = 0;
			sView.Cells[ct31c[1] + num6, ct31c[0]].Value = 0;
		}
		if (isDelete)
		{
			for (int i = 1; i < ToKhaiInfo.SheetInfos[0].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				if (hTKKCMCTableInfo.RowsValues.Count > 1)
				{
					GeneralFunctions.RemoveDataRows(sView, 0, dataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, ToKhaiInfo.SheetInfos[0]);
				}
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
		}
		for (int i = 1; i < ToKhaiInfo.SheetInfos[0].Tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[i];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					if (isLock)
					{
						sView.Cells[dataLocationTemplateRowStart + j - 1, cell2.CellTemplate.ColIndex - 1].Locked = true;
					}
					else
					{
						sView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Locked = false;
						sView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Locked = false;
						sView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Locked = false;
						sView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Locked = false;
					}
					if (isDelete)
					{
						sView.Cells[dataLocationTemplateRowStart + j - 1, cell2.CellTemplate.ColIndex - 1].Value = cell2.DefaultValue;
					}
				}
			}
		}
	}

	public void CheckPhuLuc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		winFpSpread.Sheets[0].Cells[cellCoPhuLuc[1], cellCoPhuLuc[0]].Text = ((ToKhaiInfo.SheetInfos[1] == null) ? "0" : "1");
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
			if (activeRowIndex == 7 && activeColumnIndex == 1)
			{
				ShowBangKe("BangKe_01_CNKD_TT40.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeRowIndex == 1 && activeColumnIndex == 2)
			{
				ShowBangKe("BangKe_01_2_CNKD_TT40.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_ND15_CNKD_TT40.xls");
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
			if (idxRow == 0 && idxCol == 3)
			{
				ShowBangKe("Bang_Ke_01_GiamThue_GTGT_23_24_CNKD_TT40.xls");
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
			if (idxRow == 0 && idxCol == 3)
			{
				ShowBangKe("Bang_Ke_01_GiamThue_GTGT_NQ142_CNKD_TT40.xls");
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

	public void TinhBenToKhaiCot28Den31_b_d(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("W");
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			int num2 = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			if (ToKhaiInfo.SheetInfos[3] != null)
			{
				sheetView.Cells[ct28b[1] + num2, ct28b[0]].Value = sheetView.Cells[ct28b[1] + num2, colCongThucCot28_31_b - 1].Value;
				sheetView.Cells[ct29b[1] + num2, ct29b[0]].Value = sheetView.Cells[ct29b[1] + num2, colCongThucCot28_31_b - 1].Value;
				sheetView.Cells[ct30b[1] + num2, ct30b[0]].Value = sheetView.Cells[ct30b[1] + num2, colCongThucCot28_31_b - 1].Value;
				sheetView.Cells[ct31b[1] + num2, ct31b[0]].Value = sheetView.Cells[ct31b[1] + num2, colCongThucCot28_31_b - 1].Value;
			}
			if (ToKhaiInfo.SheetInfos[4] != null)
			{
				sheetView.Cells[ct28b[1] + num2, ct28b[0]].Value = sheetView.Cells[ct28b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
				sheetView.Cells[ct29b[1] + num2, ct29b[0]].Value = sheetView.Cells[ct29b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
				sheetView.Cells[ct30b[1] + num2, ct30b[0]].Value = sheetView.Cells[ct30b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
				sheetView.Cells[ct31b[1] + num2, ct31b[0]].Value = sheetView.Cells[ct31b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
			}
			if (ToKhaiInfo.SheetInfos[5] != null)
			{
				sheetView.Cells[ct28b[1] + num2, ct28b[0]].Value = sheetView.Cells[ct28b[1] + num2, colCongThucCot28_31_b_PL01_5 - 1].Value;
				sheetView.Cells[ct29b[1] + num2, ct29b[0]].Value = sheetView.Cells[ct29b[1] + num2, colCongThucCot28_31_b_PL01_5 - 1].Value;
				sheetView.Cells[ct30b[1] + num2, ct30b[0]].Value = sheetView.Cells[ct30b[1] + num2, colCongThucCot28_31_b_PL01_5 - 1].Value;
				sheetView.Cells[ct31b[1] + num2, ct31b[0]].Value = sheetView.Cells[ct31b[1] + num2, colCongThucCot28_31_b_PL01_5 - 1].Value;
			}
			if (ToKhaiInfo.SheetInfos[3] == null && ToKhaiInfo.SheetInfos[4] == null && ToKhaiInfo.SheetInfos[5] == null)
			{
				sheetView.Cells[ct28b[1] + num2, ct28b[0]].Value = sheetView.Cells[ct28b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
				sheetView.Cells[ct29b[1] + num2, ct29b[0]].Value = sheetView.Cells[ct29b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
				sheetView.Cells[ct30b[1] + num2, ct30b[0]].Value = sheetView.Cells[ct30b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
				sheetView.Cells[ct31b[1] + num2, ct31b[0]].Value = sheetView.Cells[ct31b[1] + num2, colCongThucCot28_31_b_PL01_4 - 1].Value;
			}
			sheetView.Cells[ct28d[1] + num2, ct28d[0]].Value = sheetView.Cells[ct28c[1] + num2, colCongThucCot28_31_d - 1].Value;
			sheetView.Cells[ct29d[1] + num2, ct29d[0]].Value = sheetView.Cells[ct29c[1] + num2, colCongThucCot28_31_d - 1].Value;
			sheetView.Cells[ct30d[1] + num2, ct30d[0]].Value = sheetView.Cells[ct30c[1] + num2, colCongThucCot28_31_d - 1].Value;
			sheetView.Cells[ct31d[1] + num2, ct31d[0]].Value = sheetView.Cells[ct31c[1] + num2, colCongThucCot28_31_d - 1].Value;
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
			ClearDataCt28DenCt31(winFpSpread, ToKhaiInfo);
			TinhLaiThueSuatPL01_4(winFpSpread.Sheets[4], ToKhaiInfo);
			TinhLaiThueSuatPL01_5(winFpSpread.Sheets[5], ToKhaiInfo);
			ThayDoiToKhai28Den31_a_c(winFpSpread, ToKhaiInfo);
			TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
			SheetView sheetView = winFpSpread.Sheets[3];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			if (hTKKCMCSheetInfo != null)
			{
				TinhBenToKhaiCot28Den31_b_d(winFpSpread, ToKhaiInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiThueSuatPL01_4(SheetView sViewPL, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[4];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sViewPL.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
				if (string.IsNullOrEmpty(text) || text == "0")
				{
					string text2 = sViewPL.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						sViewPL.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text = text2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiThueSuatPL01_5(SheetView sViewPL, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[5];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sViewPL.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
				if (string.IsNullOrEmpty(text) || text == "0")
				{
					string text2 = sViewPL.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						sViewPL.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text = text2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			string text = "HSoKhaiThue/CTieuTKhaiChinh/Header/tccnkhainopthay";
			XmlNode val = xmlNode.SelectSingleNode(text);
			if (val == null || string.IsNullOrEmpty(val.InnerText))
			{
				return;
			}
			string text2 = val.InnerText.Trim();
			if (ConvertHelpers.ConvertToBool(text2) || text2.Equals("1"))
			{
				isToChuc = true;
				GeneralFunctions.FillDataMstCu(xmlNode);
			}
			else
			{
				isToChuc = false;
				XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/mst");
				if (val2 == null || string.IsNullOrEmpty(val2.InnerText))
				{
					return;
				}
				string innerText = val2.InnerText.Trim();
				XmlNode val3 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/mst_cu");
				string innerText2 = (mst = val3.InnerText.Trim());
				if (val3 != null)
				{
					val3.InnerText = innerText;
				}
				if (val2 != null)
				{
					val2.InnerText = innerText2;
				}
			}
			if (PeriodInfo.Nam == "2026")
			{
				string path = "HSoKhaiThue/CTieuTKhaiChinh/KKhaiTBVMT_TN/ThueBVMT/CTietThueBVMT/ct2_ma";
				GeneralFunctions.ReplaceMaHH(xmlNode, path);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_6(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (!isToChuc)
			{
				XmlNode val = xmlNode.SelectSingleNode("Header/mst");
				if (val != null)
				{
					val.InnerText = mst;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_7(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (!isToChuc)
			{
				XmlNode val = xmlNode.SelectSingleNode("Header/mst");
				if (val != null)
				{
					val.InnerText = mst;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void EnterCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
		winFpSpread.Sheets[0].Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Note = "";
		if ((int)iRow == ct07_SoDinhDanh[1] + 1 && (int)iCol == ct07_SoDinhDanh[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[ct07_SoDinhDanh[1], ct07_SoDinhDanh[0]].Note = xMLMessage.GetMessage("0655");
		}
	}
}
