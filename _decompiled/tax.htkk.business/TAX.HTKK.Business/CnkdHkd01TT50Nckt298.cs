using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class CnkdHkd01TT50Nckt298
{
	private int[] cellHeader_Check_1 = GeneralFunctions.ArrayCellId("B_26");

	private int[] cellHeader_Check_2 = GeneralFunctions.ArrayCellId("B_28");

	private int[] cellHeader_Check_3 = GeneralFunctions.ArrayCellId("B_30");

	private int[] cellHeader_Check_4 = GeneralFunctions.ArrayCellId("B_33");

	private int[] cellHeader_Check_5 = GeneralFunctions.ArrayCellId("B_35");

	private int[] cellKieuKy = GeneralFunctions.ArrayCellId("BM_58");

	private int colC = HTKKFile.ColumnLettersToIndex("C");

	private int colBM = HTKKFile.ColumnLettersToIndex("BM");

	private int colBN = HTKKFile.ColumnLettersToIndex("BN");

	private int colBX = HTKKFile.ColumnLettersToIndex("BX");

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

	private List<string> pathDonGiaTinh = new List<string>();

	private string namKK = "";

	private string kieuKyKK = "";

	private HTKKHeaderInfo _HeaderInfo = null;

	private FpSpread _winFpSpread = null;

	private string maTieuMucGTGT = "1701";

	private string maTieuMucTNCN = "1003";

	private Dictionary<string, double> dicKeysTableD = new Dictionary<string, double>();

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			namKK = PeriodInfo.Nam;
			kieuKyKK = PeriodInfo.KieuKy;
			_HeaderInfo = HeaderInfo;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_winFpSpread = winFpSpread;
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[cellKieuKy[1], cellKieuKy[0]].Value = PeriodInfo.KieuKy;
			lockUnLockHeaderCheck(sheetView, sInfo, PeriodInfo);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				GeneralFunctions.lockUnlockMultiCellsStatic(sheetView, sInfo, cellHeader_Check_1[1] + 1, cellHeader_Check_4[1] + 1, isLock: true, isDeleteData: false);
			}
			lockUnLockA_TK(sheetView, sInfo, PeriodInfo);
			ThayDoiTieuMucA(sheetView, sInfo, TTinChung, PeriodInfo);
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == HTKKFile.ColumnLettersToIndex("L"))
			{
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Value;
			}
			if (iCol == colC)
			{
				int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
				if (num >= 0 && hTKKCMCSheetInfo.Tables[num].TableName == "table_4")
				{
					string[] array = dicDMThueDB[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
					if (array.Length > 3)
					{
						sheetView.Cells[(int)iRow - 1, colBM - 1].Value = array[1];
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Value = array[3];
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("X") - 1].Value = ConvertHelpers.ConvertToDouble(array[4]);
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "TTDB");
						sheetView.Cells[(int)iRow - 1, colBN - 1].Value = maTieuMucByMaHH;
						string tenTieuMucTuMa = GeneralFunctions.GetTenTieuMucTuMa(maTieuMucByMaHH, TTinChung, PeriodInfo);
						sheetView.Cells[(int)iRow - 1, colBX - 1].Value = tenTieuMucTuMa;
						TinhCotSttTTDB(sheetView, hTKKCMCSheetInfo.Tables[num], (int)iRow);
					}
				}
			}
			if (iCol == colC)
			{
				int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
				if (num >= 0 && hTKKCMCSheetInfo.Tables[num].TableName == "table_5")
				{
					string text = sheetView.Cells[(int)iRow - 1, colC - 1].Text;
					foreach (KeyValuePair<string, string[]> item in dicDMKhoangSan)
					{
						string[] value = item.Value;
						if (value[2].Trim().Equals(text) && value.Length > 3 && value.Length > 2)
						{
							if (ConvertHelpers.ConvertToInt(value[4]) == 0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, colC - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
								winFpSpread.TabStrip[0].BackColor = Color.Pink;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, colC - 1);
							}
							sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AL") - 1].Value = value[1];
							sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value = value[3];
							sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Value = ConvertHelpers.ConvertToDouble(value[4]);
							sheetView.Cells[(int)iRow - 1, colBM - 1].Value = value[1];
							string maTieuMucByMaHH = ((value.Length >= 8) ? value[7] : "");
							sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BP") - 1].Value = maTieuMucByMaHH;
							double donGia = getDonGia(item.Key);
							sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AT") - 1].Value = donGia;
							string tenTieuMucTuMa = GeneralFunctions.GetTenTieuMucTuMa(maTieuMucByMaHH, TTinChung, PeriodInfo);
							sheetView.Cells[(int)iRow - 1, colBX - 1].Value = tenTieuMucTuMa;
						}
					}
				}
			}
			if (iCol == colC)
			{
				int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
				if (num >= 0 && hTKKCMCSheetInfo.Tables[num].TableName == "table_6")
				{
					string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
					if (array.Length > 1)
					{
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AL") - 1].Value = array[0];
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value = array[2];
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AT") - 1].Text = array[3];
						sheetView.Cells[(int)iRow - 1, colBM - 1].Value = array[0];
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[0], "BVMT");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BP") - 1].Value = maTieuMucByMaHH;
						string tenTieuMucTuMa = GeneralFunctions.GetTenTieuMucTuMa(maTieuMucByMaHH, TTinChung, PeriodInfo);
						sheetView.Cells[(int)iRow - 1, colBX - 1].Value = tenTieuMucTuMa;
					}
				}
			}
			if (iCol == colC)
			{
				int num = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
				if (num >= 0 && hTKKCMCSheetInfo.Tables[num].TableName == "table_7")
				{
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
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AL") - 1].Value = array[1];
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value = array[3];
						sheetView.Cells[(int)iRow - 1, colBM - 1].Value = array[1];
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(array[1], "PBVMT");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BP") - 1].Value = maTieuMucByMaHH;
						string tenTieuMucTuMa = GeneralFunctions.GetTenTieuMucTuMa(maTieuMucByMaHH, TTinChung, PeriodInfo);
						sheetView.Cells[(int)iRow - 1, colBX - 1].Value = tenTieuMucTuMa;
					}
				}
			}
			int num2 = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			if (num2 >= 0)
			{
				SetCacDongCon(sheetView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, num2);
			}
			tinhCt14_tk(winFpSpread, ToKhaiInfo);
			if (iCol >= HTKKFile.ColumnLettersToIndex("L"))
			{
				tinhSumCt18_19_20_tk(sheetView, hTKKCMCSheetInfo);
			}
			if (num2 >= 0 && (iCol == HTKKFile.ColumnLettersToIndex("C") || iCol == HTKKFile.ColumnLettersToIndex("Q")))
			{
				tinhSumMucB_Ct5_7_tk(sheetView, hTKKCMCSheetInfo);
			}
			if (num2 >= 0)
			{
				GeneralFunctions.UpdateFromSViewToSInfo(sheetView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[num2].RowsInfo, hTKKCMCSheetInfo.Tables[num2]);
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("G") || num2 >= 0)
			{
				TinhTongSoThuePhaiNopTNCN_MucA_I_II(sheetView, hTKKCMCSheetInfo);
				TongHopBangD(sheetView, hTKKCMCSheetInfo, TTinChung, PeriodInfo);
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[hTKKCMCSheetInfo.Tables.Count - 1]);
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
			if (iCol == HTKKFile.ColumnLettersToIndex("Z"))
			{
				chonTiLeTinhThueTheoNhomNganhNghe(winFpSpread.Sheets[2], (int)iRow);
			}
			tinhCt14_tk(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			lockUnLockA_TK(sView, sInfo, PeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 11 && idxCol == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_CNKD_HKD.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 3)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_NQ142_CNKD_HKD_TT18.xls");
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29"));
			validateTableMucA(sheetView, hTKKCMCSheetInfo, ref value, PeriodInfo);
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				if (!(hTKKCMCSheetInfo.Tables[i].TableName == "table_4"))
				{
					continue;
				}
				int num = HTKKFile.ColumnLettersToIndex("C");
				int num2 = HTKKFile.ColumnLettersToIndex("L");
				int num3 = HTKKFile.ColumnLettersToIndex("U");
				bool flag = Convert.ToBoolean(sheetView.Cells[cellHeader_Check_4[1], cellHeader_Check_4[0]].Value);
				int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCSheetInfo.Tables[i].RowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, HTKKFile.ColumnLettersToIndex("BM") - 1].Text;
					if (flag && string.IsNullOrEmpty(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, message3);
						value = 2;
					}
				}
				int num4 = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sheetView, hTKKCMCSheetInfo.Tables[i], num, message);
				if (num4 > 0)
				{
					value = 2;
				}
			}
			GeneralFunctions.CheckMaDiaDiemKd(sheetView, hTKKCMCSheetInfo, "ma_diadiemkd", ref value);
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				if (hTKKCMCSheetInfo.Tables[i].TableName == "table_5" || hTKKCMCSheetInfo.Tables[i].TableName == "table_6" || hTKKCMCSheetInfo.Tables[i].TableName == "table_7")
				{
					int colCheck = HTKKFile.ColumnLettersToIndex("Q");
					int num4 = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sheetView, hTKKCMCSheetInfo.Tables[i], colCheck, message4);
					if (num4 > 0)
					{
						value = 2;
					}
				}
			}
			List<string> list = new List<string>();
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				if (!(hTKKCMCSheetInfo.Tables[i].TableName == "table_1"))
				{
					continue;
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.Tables[i].Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ma_diaDiemKD");
				string text2 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
				if (!string.IsNullOrEmpty(text2))
				{
					if (list.Contains(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message4);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
						list.Add(text2);
					}
				}
			}
			HTKKFile.RemoveValidate(sheetView, cellHeader_Check_2[1], cellHeader_Check_2[0]);
			if (PeriodInfo.KieuKy == "M" && !Convert.ToBoolean(sheetView.Cells[cellHeader_Check_2[1], cellHeader_Check_2[0]].Value))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cellHeader_Check_2[1], cellHeader_Check_2[0], message5);
				value = 2;
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

	public void validateTableMucA(SheetView sView, HTKKCMCSheetInfo sInfo, ref int iwarningColor, HTKKPeriodInfo PeriodInfo)
	{
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (!(sInfo.Tables[i].TableName == "table_1") && !(sInfo.Tables[i].TableName == "table_2") && !(sInfo.Tables[i].TableName == "table_3"))
			{
				continue;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_14"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_26"));
			int colCheck = HTKKFile.ColumnLettersToIndex("C");
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			int num4 = HTKKFile.ColumnLettersToIndex("U");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num5 = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sView, hTKKCMCTableInfo, colCheck, message);
			if (num5 > 0)
			{
				iwarningColor = 2;
			}
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, num2 - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, num3 - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, num4 - 1);
				double num6 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Value);
				double num7 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Value);
				double num8 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Value);
				double num9 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Value);
				if (num9 < num6)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + j - 1, num4 - 1, message2);
					iwarningColor = 2;
				}
				if (num7 + num8 > num6)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + j - 1, num2 - 1, message4);
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + j - 1, num3 - 1, message4);
					iwarningColor = 2;
				}
			}
		}
		bool flag = Convert.ToBoolean(sView.Cells[cellHeader_Check_1[1], cellHeader_Check_1[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellHeader_Check_2[1], cellHeader_Check_2[0]].Value);
		bool flag3 = Convert.ToBoolean(sView.Cells[cellHeader_Check_3[1], cellHeader_Check_3[0]].Value);
		bool flag4 = Convert.ToBoolean(sView.Cells[cellHeader_Check_4[1], cellHeader_Check_4[0]].Value);
		bool flag5 = Convert.ToBoolean(sView.Cells[cellHeader_Check_5[1], cellHeader_Check_5[0]].Value);
		string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30"));
		string message6 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_31"));
		int num10 = HTKKFile.ColumnLettersToIndex("C");
		int num11 = HTKKFile.ColumnLettersToIndex("S");
		int num12 = HTKKFile.ColumnLettersToIndex("AE");
		int num13 = 0;
		if (!(PeriodInfo.KieuKy == "M") && !(PeriodInfo.KieuKy == "Q"))
		{
			return;
		}
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (!(sInfo.Tables[i].TableName == "table_1") && !(sInfo.Tables[i].TableName == "table_2"))
			{
				continue;
			}
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, num10 - 1);
				double num14 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Value);
				double num15 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Value);
				if (num14 != 0.0 || num15 != 0.0)
				{
					num13++;
				}
			}
		}
		if (flag2 && !flag3 && num13 == 0)
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
				if (hTKKCMCTableInfo.TableName == "table_1" || hTKKCMCTableInfo.TableName == "table_2")
				{
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + j - 1, num10 - 1, message6);
						iwarningColor = 2;
					}
				}
			}
		}
		if (!flag2 || !flag3)
		{
			return;
		}
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (!(sInfo.Tables[i].TableName == "table_2"))
			{
				continue;
			}
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				double num14 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Value);
				double num15 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Value);
				if (num14 == 0.0 && num15 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + j - 1, num10 - 1, message5);
					iwarningColor = 2;
				}
			}
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			int num = HTKKFile.ColumnLettersToIndex("Y");
			int num2 = HTKKFile.ColumnLettersToIndex("BB");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			List<string> list = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text;
				list.Add(text);
			}
			if (list.Count == 0)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
					value = 2;
				}
			}
			int num3 = list.FindIndex((string x) => x == "00");
			if (num3 < 0)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message2);
					value = 2;
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24"));
			int num = HTKKFile.ColumnLettersToIndex("C");
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, num - 1);
				}
			}
			int num2 = 0;
			foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo.Tables)
			{
				if (!GeneralFunctions.HaveToValidateRow(table2.RowsInfo, table2.RowsValues[0].RowValue.ToArray()))
				{
					num2++;
				}
			}
			if (num2 >= 2)
			{
				foreach (HTKKCMCTableInfo table3 in hTKKCMCSheetInfo.Tables)
				{
					int dataLocationTemplateRowStart = table3.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < table3.RowsValues.Count; i++)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num - 1, message);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				removeTruSoChinhTuBangSau(sView, 0, hTKKCMCSheetInfo, "truSoChinh");
				int num = hTKKCMCSheetInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableName == "table_4" && x.TableID == "1");
				if (num >= 0)
				{
					removeTruSoChinhTuBangSau(sView, num, hTKKCMCSheetInfo, "truSoChinh");
				}
			}
		}
		catch (Exception)
		{
		}
	}

	public void lockUnLockHeaderCheck(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo)
	{
		if (PeriodInfo.KieuKy == "M")
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_2[1] + 1, cellHeader_Check_4[1] + 1, isLock: false, isDeleteData: false);
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_1[1] + 1, cellHeader_Check_1[1] + 1, isLock: true, isDeleteData: true);
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_5[1] + 1, cellHeader_Check_5[1] + 1, isLock: true, isDeleteData: true);
		}
		if (PeriodInfo.KieuKy == "D")
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_1[1] + 1, cellHeader_Check_4[1] + 1, isLock: true, isDeleteData: true);
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_5[1] + 1, cellHeader_Check_5[1] + 1, isLock: false, isDeleteData: false);
		}
		if (PeriodInfo.KieuKy == "Q")
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_1[1] + 1, cellHeader_Check_4[1] + 1, isLock: false, isDeleteData: false);
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, cellHeader_Check_5[1] + 1, cellHeader_Check_5[1] + 1, isLock: true, isDeleteData: true);
		}
	}

	public void lockUnLockA_TK(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo)
	{
		bool flag = Convert.ToBoolean(sView.Cells[cellHeader_Check_1[1], cellHeader_Check_1[0]].Value);
		bool flag2 = Convert.ToBoolean(sView.Cells[cellHeader_Check_2[1], cellHeader_Check_2[0]].Value);
		bool flag3 = Convert.ToBoolean(sView.Cells[cellHeader_Check_3[1], cellHeader_Check_3[0]].Value);
		bool flag4 = Convert.ToBoolean(sView.Cells[cellHeader_Check_4[1], cellHeader_Check_4[0]].Value);
		bool flag5 = Convert.ToBoolean(sView.Cells[cellHeader_Check_5[1], cellHeader_Check_5[0]].Value);
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhTMDT/DoanhThuThueGTGTTNCN/ten_diaDiemKD");
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhTMDT/DoanhThuThueGTGTTNCN/diaChi_diaDiemKD/diaChi");
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhLPS/DoanhThuThueGTGTTNCN/ten_diaDiemKD");
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhLPS/DoanhThuThueGTGTTNCN/diaChi_diaDiemKD/diaChi");
		List<int> list = new List<int>();
		list.Add(HTKKFile.ColumnLettersToIndex("C"));
		list.Add(HTKKFile.ColumnLettersToIndex("L"));
		list.Add(HTKKFile.ColumnLettersToIndex("Q"));
		List<int> list2 = new List<int>();
		list2.Add(HTKKFile.ColumnLettersToIndex("C"));
		list2.Add(HTKKFile.ColumnLettersToIndex("I"));
		list2.Add(HTKKFile.ColumnLettersToIndex("L"));
		list2.Add(HTKKFile.ColumnLettersToIndex("Q"));
		list2.Add(HTKKFile.ColumnLettersToIndex("U"));
		list2.Add(HTKKFile.ColumnLettersToIndex("AD"));
		list2.Add(HTKKFile.ColumnLettersToIndex("AH"));
		list2.Add(HTKKFile.ColumnLettersToIndex("AP"));
		list2.Add(HTKKFile.ColumnLettersToIndex("AT"));
		List<int> cacCotAn = new List<int>(list2);
		List<int> list3 = new List<int>(list2);
		list3.Add(HTKKFile.ColumnLettersToIndex("AX"));
		if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "Q")
		{
			if (flag2 && flag3)
			{
				foreach (HTKKCMCTableInfo table in sInfo.Tables)
				{
					List<int> colsLockMucA = getColsLockMucA(table);
					if (table.TableName == "table_1")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table, colsLockMucA, locked: true, deleteData: true, lockAllHeader: true);
					}
				}
				foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
				{
					List<int> colsLockMucA = getColsLockMucA(table2);
					if (table2.TableName == "table_2")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table2, colsLockMucA, locked: false, deleteData: false, lockAllHeader: true);
					}
				}
				GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
				foreach (HTKKCMCTableInfo table3 in sInfo.Tables)
				{
					List<int> colsLockMucA = getColsLockMucA(table3);
					if (table3.TableName == "table_3")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table3, colsLockMucA, locked: true, deleteData: true, lockAllHeader: true);
					}
				}
			}
			if (flag2 && !flag3)
			{
				foreach (HTKKCMCTableInfo table4 in sInfo.Tables)
				{
					List<int> colsLockMucA = getColsLockMucA(table4);
					if (table4.TableName == "table_1")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table4, colsLockMucA, locked: false, deleteData: false, lockAllHeader: true);
					}
				}
				foreach (HTKKCMCTableInfo table5 in sInfo.Tables)
				{
					List<int> colsLockMucA = getColsLockMucA(table5);
					if (table5.TableName == "table_2")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table5, colsLockMucA, locked: false, deleteData: false, lockAllHeader: true);
					}
				}
				GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
				foreach (HTKKCMCTableInfo table6 in sInfo.Tables)
				{
					List<int> colsLockMucA = getColsLockMucA(table6);
					if (table6.TableName == "table_3")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table6, colsLockMucA, locked: true, deleteData: true, lockAllHeader: true);
					}
				}
			}
			if (flag4)
			{
				foreach (HTKKCMCTableInfo table7 in sInfo.Tables)
				{
					if (table7.TableName == "table_4")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table7, list, locked: false, deleteData: false, lockAllHeader: true);
					}
					if (table7.TableName == "table_5")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table7, list2, locked: false, deleteData: false, lockAllHeader: true);
					}
					if (table7.TableName == "table_6")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table7, cacCotAn, locked: false, deleteData: false, lockAllHeader: true);
					}
					if (table7.TableName == "table_7")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table7, list3, locked: false, deleteData: false, lockAllHeader: true);
					}
				}
			}
			else
			{
				foreach (HTKKCMCTableInfo table8 in sInfo.Tables)
				{
					if (table8.TableName == "table_4")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table8, list, locked: true, deleteData: true, lockAllHeader: true);
					}
					if (table8.TableName == "table_5")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table8, list2, locked: true, deleteData: true, lockAllHeader: true);
					}
					if (table8.TableName == "table_6")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table8, cacCotAn, locked: true, deleteData: true, lockAllHeader: true);
					}
					if (table8.TableName == "table_7")
					{
						GeneralFunctions.LockUnLockColsInTable(sView, table8, list3, locked: true, deleteData: true, lockAllHeader: true);
					}
				}
			}
		}
		if (!(PeriodInfo.KieuKy == "D"))
		{
			return;
		}
		GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
		foreach (HTKKCMCTableInfo table9 in sInfo.Tables)
		{
			List<int> colsLockMucA = getColsLockMucA(table9);
			if (table9.TableName == "table_1" || table9.TableName == "table_2")
			{
				GeneralFunctions.LockUnLockColsInTable(sView, table9, colsLockMucA, locked: true, deleteData: true, lockAllHeader: true);
			}
		}
		GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.RowIndex, isLock: false, isDeleteData: false);
		foreach (HTKKCMCTableInfo table10 in sInfo.Tables)
		{
			List<int> colsLockMucA = getColsLockMucA(table10);
			if (table10.TableName == "table_3")
			{
				GeneralFunctions.LockUnLockColsInTable(sView, table10, colsLockMucA, locked: false, deleteData: false, lockAllHeader: true);
			}
		}
		foreach (HTKKCMCTableInfo table11 in sInfo.Tables)
		{
			if (table11.TableName == "table_4")
			{
				GeneralFunctions.LockUnLockColsInTable(sView, table11, list, locked: true, deleteData: true, lockAllHeader: true);
			}
			if (table11.TableName == "table_5")
			{
				GeneralFunctions.LockUnLockColsInTable(sView, table11, list2, locked: true, deleteData: true, lockAllHeader: true);
			}
			if (table11.TableName == "table_6")
			{
				GeneralFunctions.LockUnLockColsInTable(sView, table11, cacCotAn, locked: true, deleteData: true, lockAllHeader: true);
			}
			if (table11.TableName == "table_7")
			{
				GeneralFunctions.LockUnLockColsInTable(sView, table11, list3, locked: true, deleteData: true, lockAllHeader: true);
			}
		}
	}

	public List<int> getColsLockMucA(HTKKCMCTableInfo table)
	{
		List<int> list = new List<int>();
		list.Add(HTKKFile.ColumnLettersToIndex("C"));
		list.Add(HTKKFile.ColumnLettersToIndex("L"));
		list.Add(HTKKFile.ColumnLettersToIndex("O"));
		list.Add(HTKKFile.ColumnLettersToIndex("Q"));
		list.Add(HTKKFile.ColumnLettersToIndex("U"));
		if (table.TableName == "table_1")
		{
			list.Add(HTKKFile.ColumnLettersToIndex("Z"));
		}
		return list;
	}

	public void removeTruSoChinhTuBangSau(SheetView sView, int tableIndex, HTKKCMCSheetInfo sInfo, string pathTruSoChinh)
	{
		string tableName = sInfo.Tables[tableIndex].TableName;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (i != tableIndex && tableName == sInfo.Tables[i].TableName)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.Tables[i].Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTruSoChinh);
				if (hTKKCMCCellInfo != null)
				{
					GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
			}
		}
	}

	public void tinhSumCt18_19_20_tk(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		double num = 0.0;
		double num2 = 0.0;
		double num3 = 0.0;
		double num4 = 0.0;
		double num5 = 0.0;
		double num6 = 0.0;
		double num7 = 0.0;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (sInfo.Tables[i].TableName == "table_1" || sInfo.Tables[i].TableName == "table_2" || sInfo.Tables[i].TableName == "table_3")
			{
				int num8 = HTKKFile.ColumnLettersToIndex("C");
				int num9 = HTKKFile.ColumnLettersToIndex("L");
				int num10 = HTKKFile.ColumnLettersToIndex("O");
				int num11 = HTKKFile.ColumnLettersToIndex("Q");
				int num12 = HTKKFile.ColumnLettersToIndex("S");
				int num13 = HTKKFile.ColumnLettersToIndex("U");
				int num14 = HTKKFile.ColumnLettersToIndex("Z");
				int num15 = HTKKFile.ColumnLettersToIndex("AE");
				int dataLocationTemplateRowStart = sInfo.Tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < sInfo.Tables[i].RowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, HTKKFile.ColumnLettersToIndex("U"));
					double num16 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Value);
					double num17 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Value);
					double num18 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Value);
					double num19 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Value);
					double num20 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num13 - 1].Value);
					double num21 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num14 - 1].Value);
					double num22 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num15 - 1].Value);
					num += num16;
					num2 += num17;
					num3 += num18;
					num4 += num19;
					num5 += num20;
					num6 += num21;
					num7 += num22;
				}
			}
		}
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT11");
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT12");
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT13");
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT14");
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT15");
		HTKKCMCCellInfo hTKKCMCCellInfo6 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT16");
		HTKKCMCCellInfo hTKKCMCCellInfo7 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/TongCongCT18/tongCT17");
		HTKKCMCCellInfo hTKKCMCCellInfo8 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/SoThueConPhaiNopCT20/soThueGTGT_ct14");
		HTKKCMCCellInfo hTKKCMCCellInfo9 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/SoThueConPhaiNopCT20/soThueTNCN_ct17");
		sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
		sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
		sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = num3;
		sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = num4;
		sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = num5;
		sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = num6;
		sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = num7;
		sView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = num4;
		sView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = num7;
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

	public Dictionary<string, string> GetDMThueBVMT(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_03ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f6: Expected O, but got Unknown
		//IL_0436: Unknown result type (might be due to invalid IL or missing references)
		//IL_043d: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "012027#102026#052026#042026#032026#012026#072025#012025#012024#012023#072022#042022#112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			List<string> list = new List<string>();
			string iNTERFACE_INI = HTKKConfig.INTERFACE_INI;
			DirectoryInfo directoryInfo = new DirectoryInfo(iNTERFACE_INI);
			if (directoryInfo.Exists)
			{
				FileInfo[] files = directoryInfo.GetFiles();
				if (files.Length > 0)
				{
					FileInfo[] array = files;
					foreach (FileInfo fileInfo in array)
					{
						if (fileInfo.Name.Contains("Catalogue_DM_thue_BVMT_"))
						{
							list.Add(fileInfo.Name.Replace("Catalogue_DM_thue_BVMT_", "").Replace(".xml", ""));
						}
					}
				}
			}
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
				for (int j = 0; j < list.Count; j++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(list[j].Substring(2)), ConvertHelpers.ConvertToInt(list[j].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = list[j];
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
					for (int j = 0; j < list.Count; j++)
					{
						DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(list[j].Substring(2)), ConvertHelpers.ConvertToInt(list[j].Substring(0, 2)), 1);
						if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
						{
							t = dateTime;
							text = list[j];
						}
					}
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				int iQuy = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				DateTime t2 = GeneralFunctions.GetNgayDauQuy(iQuy, num, 1, 1).AddMonths(1).AddDays(-1.0);
				for (int j = 0; j < list.Count; j++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(list[j].Substring(2)), ConvertHelpers.ConvertToInt(list[j].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = list[j];
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

	public List<item> GetItemDm_GetDMKSBVMT()
	{
		return _dmBVMTData;
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

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		int num = winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1;
		int iCol = winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1;
		HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
		int tableIndex = GeneralFunctions.DetectIndexOfTable(sInfo, num);
		SetCacDongCon(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], num, iCol, tableIndex);
		return true;
	}

	public bool UpdateFormula_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TinhLaiThueSuatPL01_3(winFpSpread.Sheets[3], ToKhaiInfo);
		return true;
	}

	public void TinhLaiThueSuatPL01_3(SheetView sViewPL, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			TinhLaiThueSuatPL01_3(winFpSpread.Sheets[3], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			GeneralFunctions.FillDataMstCu(xmlNode);
			if (PeriodInfo.Nam == "2026")
			{
				string path = "HSoKhaiThue/CTieuTKhaiChinh/KKhaiTaiNguyen/ThueBVMT/CTietThueBVMT/ct2_ma";
				GeneralFunctions.ReplaceMaHH(xmlNode, path);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetCacDongCon(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, int tableIndex)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = null;
		HTKKCMCCellInfo hTKKCMCCellInfo2 = null;
		HTKKCMCCellInfo hTKKCMCCellInfo3 = null;
		HTKKCMCCellInfo hTKKCMCCellInfo4 = null;
		if (tableIndex >= 0 && sInfo.Tables[tableIndex].TableName == "table_1")
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "truSoChinh");
			hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ma_diaDiemKD");
			hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "diaChi_diaDiemKD/ten_xaPhuong");
			hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "diaChi_diaDiemKD/ma_xaPhuong");
			object value = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
			object value2 = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			object value3 = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
			object value4 = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value, HTKKFile.ColumnLettersToIndex("BS"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value2, HTKKFile.ColumnLettersToIndex("BN"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value3, HTKKFile.ColumnLettersToIndex("BO"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value4, HTKKFile.ColumnLettersToIndex("BP"));
		}
		if (tableIndex >= 0 && sInfo.Tables[tableIndex].TableName == "table_2")
		{
			hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhTMDT/DoanhThuThueGTGTTNCN/ma_diaDiemKD");
			hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhTMDT/DoanhThuThueGTGTTNCN/diaChi_diaDiemKD/ten_xaPhuong");
			hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhTMDT/DoanhThuThueGTGTTNCN/diaChi_diaDiemKD/ma_xaPhuong");
			if ((tableIndex >= 0 && sInfo.Tables[tableIndex].TableName == "table_2") || iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TableName == "table_2");
				object value2 = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
				object value3 = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
				object value4 = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
				GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value2, HTKKFile.ColumnLettersToIndex("BN"));
				GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value3, HTKKFile.ColumnLettersToIndex("BO"));
				GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value4, HTKKFile.ColumnLettersToIndex("BP"));
			}
		}
		hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhLPS/DoanhThuThueGTGTTNCN/ma_diaDiemKD");
		hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhLPS/DoanhThuThueGTGTTNCN/diaChi_diaDiemKD/ten_xaPhuong");
		hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KhaiThueGTGTTNCN/KinhDoanhLPS/DoanhThuThueGTGTTNCN/diaChi_diaDiemKD/ma_xaPhuong");
		if ((tableIndex >= 0 && sInfo.Tables[tableIndex].TableName == "table_3") || iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TableName == "table_3");
			object value2 = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			object value3 = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
			object value4 = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value2, HTKKFile.ColumnLettersToIndex("BN"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value3, HTKKFile.ColumnLettersToIndex("BO"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value4, HTKKFile.ColumnLettersToIndex("BP"));
		}
		if (tableIndex >= 0 && sInfo.Tables[tableIndex].TableName == "table_4")
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[tableIndex];
			hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ma_diaDiemKD");
			hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "diaChi_diaDiemKD/ten_xaPhuong");
			hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "diaChi_diaDiemKD/ma_xaPhuong");
			object value2 = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			object value3 = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
			object value4 = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value2, HTKKFile.ColumnLettersToIndex("BO"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value3, HTKKFile.ColumnLettersToIndex("BP"));
			GeneralFunctions.SetLaiDongConBangF5(sView, hTKKCMCTableInfo, value4, HTKKFile.ColumnLettersToIndex("BQ"));
		}
	}

	public List<string> GetAllKeyForTableD(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		List<string> list = new List<string>();
		for (int i = 0; i < sInfo.Tables.Count - 1; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			if (hTKKCMCTableInfo.TableName == "table_8")
			{
				continue;
			}
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				string text = sView.Cells[dataLocationTemplateRowStart + j - 1, HTKKFile.ColumnLettersToIndex("BV") - 1].Text;
				if (!string.IsNullOrEmpty(text) && !list.Contains(text))
				{
					List<string> list2 = GeneralFunctions.SplitString2(text, "###");
					if (list2.Count == 4 && !string.IsNullOrEmpty(list2[0]) && !string.IsNullOrEmpty(list2[1]) && !string.IsNullOrEmpty(list2[2]) && !string.IsNullOrEmpty(list2[3]))
					{
						list.Add(text);
					}
				}
			}
		}
		for (int i = 0; i < sInfo.Tables.Count - 1; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			if (!(hTKKCMCTableInfo.TableName == "table_1") && !(hTKKCMCTableInfo.TableName == "table_2") && !(hTKKCMCTableInfo.TableName == "table_3"))
			{
				continue;
			}
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				string text = sView.Cells[dataLocationTemplateRowStart + j - 1, HTKKFile.ColumnLettersToIndex("BW") - 1].Text;
				if (!string.IsNullOrEmpty(text) && !list.Contains(text))
				{
					List<string> list2 = GeneralFunctions.SplitString2(text, "###");
					if (list2.Count == 4 && !string.IsNullOrEmpty(list2[0]) && !string.IsNullOrEmpty(list2[1]) && !string.IsNullOrEmpty(list2[2]) && !string.IsNullOrEmpty(list2[3]))
					{
						list.Add(text);
					}
				}
			}
		}
		return list;
	}

	public void createKeysTableD(SheetView sView, HTKKCMCTableInfo table, string keyCompare, int colKey, int colSum)
	{
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < table.RowsValues.Count; i++)
		{
			string text = sView.Cells[dataLocationTemplateRowStart + i - 1, colKey - 1].Text;
			if (keyCompare.Equals(text))
			{
				double num = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, colSum - 1].Value);
				if (!dicKeysTableD.ContainsKey(text))
				{
					dicKeysTableD.Add(keyCompare, num);
				}
				else
				{
					dicKeysTableD[keyCompare] += num;
				}
			}
		}
	}

	public void TongHopBangD(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		dicKeysTableD.Clear();
		List<string> allKeyForTableD = GetAllKeyForTableD(sView, sInfo);
		if (allKeyForTableD.Count == 0)
		{
			return;
		}
		for (int i = 0; i < sInfo.Tables.Count - 1; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			if (hTKKCMCTableInfo.TableName == "table_8")
			{
				continue;
			}
			int num = -1;
			int num2 = -1;
			if (hTKKCMCTableInfo.TableName == "table_1" || hTKKCMCTableInfo.TableName == "table_2" || hTKKCMCTableInfo.TableName == "table_3")
			{
				num = HTKKFile.ColumnLettersToIndex("BV");
				num2 = HTKKFile.ColumnLettersToIndex("S");
			}
			else if (hTKKCMCTableInfo.TableName == "table_4")
			{
				num = HTKKFile.ColumnLettersToIndex("BV");
				num2 = HTKKFile.ColumnLettersToIndex("AD");
			}
			else
			{
				num = HTKKFile.ColumnLettersToIndex("BV");
				num2 = HTKKFile.ColumnLettersToIndex("BB");
			}
			foreach (string item in allKeyForTableD)
			{
				createKeysTableD(sView, hTKKCMCTableInfo, item, num, num2);
			}
		}
		for (int i = 0; i < sInfo.Tables.Count - 1; i++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
			int num = -1;
			int num2 = -1;
			if (!(hTKKCMCTableInfo.TableName == "table_1") && !(hTKKCMCTableInfo.TableName == "table_2") && !(hTKKCMCTableInfo.TableName == "table_3"))
			{
				continue;
			}
			num = HTKKFile.ColumnLettersToIndex("BW");
			num2 = HTKKFile.ColumnLettersToIndex("BT");
			foreach (string item2 in allKeyForTableD)
			{
				createKeysTableD(sView, hTKKCMCTableInfo, item2, num, num2);
			}
		}
		int num3 = sInfo.Tables.Count - 1;
		HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables[num3];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		if (hTKKCMCTableInfo2.RowsValues.Count > 1)
		{
			GeneralFunctions.RemoveDataRows(sView, num3, dataLocationTemplateRowStart, hTKKCMCTableInfo2.RowsValues.Count - 1, sInfo);
		}
		GeneralFunctions.setDefaultAllTable(sView, hTKKCMCTableInfo2);
		int num4 = 0;
		foreach (KeyValuePair<string, double> item3 in dicKeysTableD)
		{
			if (num4 > 0)
			{
				GeneralFunctions.InsertTableRow(sView, dataLocationTemplateRowStart + num4 - 1, -1, dataLocationTemplateRowStart + num4 - 1, -1, -1, sInfo, num3);
			}
			num4++;
		}
		DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
		int num5 = 0;
		foreach (KeyValuePair<string, double> item4 in dicKeysTableD)
		{
			List<string> list = GeneralFunctions.SplitString2(item4.Key, "###");
			string value = list[0];
			string text = list[1];
			string text2 = list[2];
			string text3 = list[3];
			double value2 = item4.Value;
			string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(text3, ngayDauKyKeKhai);
			string valueInDm = GeneralFunctions.GetValueInDm("..\\InterfaceIni\\Catalogue_CQT_Dia_Ban.xml", "Value", text2 + "###", "Value");
			string text4 = "";
			string text5 = "";
			if (!string.IsNullOrEmpty(valueInDm))
			{
				List<string> list2 = GeneralFunctions.SplitString2(valueInDm, "###");
				if (list2.Count == 3)
				{
					text4 = list2[1];
					text5 = list2[2];
				}
			}
			string valueInDm2 = GeneralFunctions.GetValueInDm("..\\InterfaceIni\\Catalogue_CQThu_Dia_Ban.xml", "MaXa", text2, "MaCQThu");
			string valueInDm3 = GeneralFunctions.GetValueInDm("..\\InterfaceIni\\Catalogue_CQThu_Dia_Ban.xml", "MaXa", text2, "TenCQThu");
			DateTime hanNopToKhai = GeneralFunctions.GetHanNopToKhai(PeriodInfo, _HeaderInfo);
			string valueInDm4 = GeneralFunctions.GetValueInDm("..\\InterfaceIni\\Catalogue_DBHC_Dia_Ban.xml", "MaXa", text2, "MaDBHC");
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Value = value;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("E") - 1].Value = tenTmByMaTm;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("I") - 1].Value = value2;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = text3;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("BW") - 1].Value = text2;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("BV") - 1].Value = item4.Key;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("AH") - 1].Value = text4 + " - " + text5;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("BN") - 1].Value = text4;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("Z") - 1].Value = valueInDm2 + " - " + valueInDm3;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("BM") - 1].Value = valueInDm2;
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Value = hanNopToKhai.ToString("dd/MM/yyyy");
			sView.Cells[dataLocationTemplateRowStart + num5 - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = valueInDm4;
			num5++;
		}
	}

	public void TinhCotSttTTDB(SheetView sView, HTKKCMCTableInfo table, int iRow)
	{
		int num = HTKKFile.ColumnLettersToIndex("BL");
		int num2 = HTKKFile.ColumnLettersToIndex("I");
		string text = sView.Cells[iRow - 1, num - 1].Text;
		string[] array = text.Split(new char[1] { '.' });
		if (array.Length == 2)
		{
			sView.Cells[iRow - 1, num2 - 1].Value = "[19" + GeneralFunctions.GetSttVietnameseLetter(Convert.ToInt32(array[1])) + "]";
		}
	}

	public void TinhTongSoThuePhaiNopTNCN_MucA_I_II(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		int[] array = GeneralFunctions.ArrayCellId("B_28");
		bool flag = Convert.ToBoolean(sView.Cells[array[1], array[0]].Value);
		double num = 0.0;
		double num2 = 0.0;
		int num3 = HTKKFile.ColumnLettersToIndex("BS");
		int num4 = HTKKFile.ColumnLettersToIndex("AE");
		int num5 = HTKKFile.ColumnLettersToIndex("BT");
		int num6 = HTKKFile.ColumnLettersToIndex("BL");
		int num7 = HTKKFile.ColumnLettersToIndex("B");
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (!(table.TableName == "table_1") && !(table.TableName == "table_2"))
			{
				continue;
			}
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				bool flag2 = Convert.ToBoolean(sView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
				sView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Value = sView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value;
				if (table.TableName == "table_1")
				{
					sView.Cells[dataLocationTemplateRowStart + i - 1, num7 - 1].Value = sView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Value;
				}
				if (!flag2)
				{
					double num8 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Value);
					if (table.TableName == "table_1")
					{
						num += num8;
					}
					else
					{
						num2 += num8;
					}
				}
			}
		}
		if (!flag)
		{
			return;
		}
		foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
		{
			if (!(table2.TableName == "table_1") && !(table2.TableName == "table_2"))
			{
				continue;
			}
			int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
			if (Convert.ToBoolean(sView.Cells[dataLocationTemplateRowStart - 1, num3 - 1].Value))
			{
				if (table2.TableName == "table_1")
				{
					sView.Cells[dataLocationTemplateRowStart - 1, num5 - 1].Value = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart - 1, num5 - 1].Value) + num;
				}
				else
				{
					sView.Cells[dataLocationTemplateRowStart - 1, num5 - 1].Value = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart - 1, num5 - 1].Value) + num2;
				}
			}
		}
	}

	public void ThayDoiTieuMucA(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table.TableName == "table_1" || table.TableName == "table_2")
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct14_maTieuMuc");
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct14_tenTieuMuc");
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo3 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct17_maTieuMuc");
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo4 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct17_tenTieuMuc");
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num = 0; num < table.RowsValues.Count; num++)
				{
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy == "D") ? "" : maTieuMucGTGT);
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy == "D") ? "" : GeneralFunctions.GetTenTieuMucTuMa(maTieuMucGTGT, TTinChung, PeriodInfo));
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy == "D") ? "" : maTieuMucTNCN);
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo4.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy == "D") ? "" : GeneralFunctions.GetTenTieuMucTuMa(maTieuMucTNCN, TTinChung, PeriodInfo));
				}
			}
			if (table.TableName == "table_3")
			{
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct14_maTieuMuc");
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct14_tenTieuMuc");
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo3 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct17_maTieuMuc");
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo4 = table.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Path == "ct17_tenTieuMuc");
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num = 0; num < table.RowsValues.Count; num++)
				{
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy != "D") ? "" : maTieuMucGTGT);
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy != "D") ? "" : GeneralFunctions.GetTenTieuMucTuMa(maTieuMucGTGT, TTinChung, PeriodInfo));
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy != "D") ? "" : maTieuMucTNCN);
					sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo4.CellTemplate.ColIndex - 1].Value = ((PeriodInfo.KieuKy != "D") ? "" : GeneralFunctions.GetTenTieuMucTuMa(maTieuMucTNCN, TTinChung, PeriodInfo));
				}
			}
		}
	}

	public double TimGiamThue_Pl_01_2(SheetView sView_01_2, HTKKCMCSheetInfo sInfo_01_2, string maMuc_TK, string maNhomNganhNghe_TK)
	{
		if (sInfo_01_2 == null)
		{
			return 0.0;
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo_01_2.Tables[1];
		int num = HTKKFile.ColumnLettersToIndex("AA");
		int num2 = HTKKFile.ColumnLettersToIndex("AC");
		int num3 = HTKKFile.ColumnLettersToIndex("H");
		double num4 = 0.0;
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string text = sView_01_2.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text;
			if (!(text != maMuc_TK))
			{
				string text2 = sView_01_2.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text;
				if (!(text2 != maNhomNganhNghe_TK))
				{
					double num5 = Convert.ToDouble(sView_01_2.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
					num4 += num5;
				}
			}
		}
		return num4;
	}

	public void tinhCt14_tk(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		SheetView sheetView = winFpSpread.Sheets[0];
		int num = HTKKFile.ColumnLettersToIndex("S");
		int num2 = HTKKFile.ColumnLettersToIndex("BZ");
		int num3 = HTKKFile.ColumnLettersToIndex("BM");
		HTKKCMCSheetInfo sInfo_01_ = ToKhaiInfo.SheetInfos[2];
		SheetView sView_01_ = winFpSpread.Sheets[2];
		foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
		{
			string tableName = table.TableName;
			if (tableName != "table_1" && tableName != "table_2" && tableName != "table_3")
			{
				break;
			}
			string text = tableName switch
			{
				"table_3" => "III", 
				"table_2" => "II", 
				"table_1" => "I", 
				_ => "", 
			};
			if (!string.IsNullOrEmpty(text) && !(table.TableName != tableName))
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text;
					double num4 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value);
					double num5 = TimGiamThue_Pl_01_2(sView_01_, sInfo_01_, text, text2);
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value = num4 - num5;
				}
			}
		}
	}

	public void tinhSumMucB_Ct5_7_tk(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		double num = 0.0;
		double num2 = 0.0;
		for (int i = 0; i < sInfo.Tables.Count; i++)
		{
			if (sInfo.Tables[i].TableName == "table_4")
			{
				int num3 = HTKKFile.ColumnLettersToIndex("Q");
				int num4 = HTKKFile.ColumnLettersToIndex("AD");
				int dataLocationTemplateRowStart = sInfo.Tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < sInfo.Tables[i].RowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + j - 1, HTKKFile.ColumnLettersToIndex("U"));
					double num5 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Value);
					double num6 = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Value);
					num += num5;
					num2 += num6;
				}
			}
		}
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKhaiThueTTDB/TongCongCT22/tongCT5");
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKhaiThueTTDB/TongCongCT22/tongCT7");
		sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
		sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
	}

	public void chonTiLeTinhThueTheoNhomNganhNghe(SheetView sView, int iRow)
	{
		int num = HTKKFile.ColumnLettersToIndex("AA");
		int num2 = HTKKFile.ColumnLettersToIndex("F");
		int num3 = HTKKFile.ColumnLettersToIndex("J");
		string text = "";
		text = sView.Cells[iRow - 1, num - 1].Text switch
		{
			"06" => "2", 
			"05" => "5", 
			"04" => "3", 
			"03" => "5", 
			"02" => "5", 
			"01" => "1", 
			_ => "", 
		};
		sView.Cells[iRow - 1, num2 - 1].Value = text;
		sView.Cells[iRow - 1, num3 - 1].Value = text;
	}
}
