using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt01TT80Nckt268
{
	private int idx01_2_GTGT = 1;

	private int idx01_3_GTGT = 2;

	private int idx01_6_CSSX_GTGT = 3;

	private int idx01_6_VT_GTGT = 4;

	private int[] cell11c = GeneralFunctions.ArrayCellId("D_22");

	private int[] cell11cMa = GeneralFunctions.ArrayCellId("P_22");

	private int[] cell11b = GeneralFunctions.ArrayCellId("D_24");

	private int[] cell11bMa = GeneralFunctions.ArrayCellId("P_24");

	private int[] cell11a = GeneralFunctions.ArrayCellId("D_26");

	private int[] cell11aMa = GeneralFunctions.ArrayCellId("P_26");

	private int[] cell09 = GeneralFunctions.ArrayCellId("B_16");

	private int[] cell10 = GeneralFunctions.ArrayCellId("H_18");

	private int sheetIdxPL_MGT = 5;

	private List<string> _listKeyAndRowTo_Pl01_2 = new List<string>();

	private List<string> _listRowAndKeyOld_Pl01_2 = new List<string>();

	private List<string> _listKeyAndRowTo_Pl_01_6CSSX = new List<string>();

	private List<string> _listRowAndKeyOld_Pl_01_6CSSX = new List<string>();

	private List<string> _listKeyAndRowTo_Pl_01_6DVVT = new List<string>();

	private List<string> _listRowAndKeyOld_Pl_01_6DVVT = new List<string>();

	private string cellId_maNNKD = "P_4";

	private string cellId_tenNNKD = "Q_4";

	private string cellId_maTieuMuc = "L_27";

	private string cellId_ckbGiaHan = "B_18";

	private string cellId_ct21 = "I_36";

	private string cellId_ct22 = "L_37";

	private string cellId_ct26 = "J_44";

	private string cellId_ct29 = "J_46";

	private string cellId_ct30 = "J_47";

	private string cellId_ct31 = "L_47";

	private string cellId_ct32 = "J_48";

	private string cellId_ct33 = "L_48";

	private string cellId_ct32a = "J_49";

	private string cellId_ct34 = "J_50";

	private string strCellId_01_5_tong6 = "M54";

	private int odrColC = HTKKFile.ColumnLettersToIndex("C");

	private int odrColD = HTKKFile.ColumnLettersToIndex("D");

	private int odrColE = HTKKFile.ColumnLettersToIndex("E");

	private int odrColF = HTKKFile.ColumnLettersToIndex("F");

	private int odrColL = HTKKFile.ColumnLettersToIndex("L");

	private int odrColG = HTKKFile.ColumnLettersToIndex("G");

	private int odrColH = HTKKFile.ColumnLettersToIndex("H");

	private int odrColI = HTKKFile.ColumnLettersToIndex("I");

	private int odrColJ = HTKKFile.ColumnLettersToIndex("J");

	private int odrColK = HTKKFile.ColumnLettersToIndex("K");

	private int odrColM = HTKKFile.ColumnLettersToIndex("M");

	private int odrColP = HTKKFile.ColumnLettersToIndex("P");

	private int[] cellTenTieuMuc = GeneralFunctions.ArrayCellId("L_28");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private int[] CellId_maNNKD => GeneralFunctions.ArrayCellId(cellId_maNNKD);

	private int[] CellId_tenNNKD => GeneralFunctions.ArrayCellId(cellId_tenNNKD);

	private int[] CellId_maTieuMuc => GeneralFunctions.ArrayCellId(cellId_maTieuMuc);

	private int[] CellId_ckbGiaHan => GeneralFunctions.ArrayCellId(cellId_ckbGiaHan);

	private int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private int[] CellId_ct32a => GeneralFunctions.ArrayCellId(cellId_ct32a);

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	public string CheckPeriod(HTKKPeriodInfo PeriodInfo, XmlNode TTinChungNode)
	{
		string empty = string.Empty;
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			empty = string.Empty;
		}
		return empty;
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[GeneralFunctions.ArrayCellId("K_27")[1], GeneralFunctions.ArrayCellId("K_27")[0]].Value = ((PeriodInfo.LoaiToKhai == "C") ? 1 : 0);
			sView.Cells[CellId_maNNKD[1], CellId_maNNKD[0]].Text = PeriodInfo.TieuMucHachToanVal;
			sView.Cells[CellId_tenNNKD[1], CellId_tenNNKD[0]].Text = getTenNNKD(PeriodInfo.TieuMucHachToanVal);
			if (PeriodInfo.TieuMucHachToanVal.Equals("01"))
			{
				sView.Cells[CellId_maTieuMuc[1], CellId_maTieuMuc[0]].Text = "1705";
			}
			else if (PeriodInfo.TieuMucHachToanVal.Equals("05"))
			{
				sView.Cells[CellId_maTieuMuc[1], CellId_maTieuMuc[0]].Text = "1704";
			}
			else
			{
				sView.Cells[CellId_maTieuMuc[1], CellId_maTieuMuc[0]].Text = "1701";
			}
			DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(TTinChung, PeriodInfo);
			string maTieuMuc = sView.Cells[CellId_maTieuMuc[1], CellId_maTieuMuc[0]].Text.Trim();
			string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(maTieuMuc, ngayDauKyKeKhai);
			sView.Cells[cellTenTieuMuc[1], cellTenTieuMuc[0]].Value = tenTmByMaTm;
			bool flag = false;
			XuLyCheckboxCt21(sView.FpSpread, isBindData: true, ToKhaiInfo, TTinChung.mst, DicSheet);
			int row = 36;
			int column = odrColL - 1;
			int column2 = odrColP - 1;
			double? preperiod_CT43Value = GetPreperiod_CT43Value(PeriodInfo, ToKhaiInfo, HeaderInfo);
			sView.Cells[row, column2].Value = preperiod_CT43Value ?? 0.0;
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile || PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[row, column].ResetFormula();
			}
			else
			{
				sView.Cells[row, column].Formula = "P37";
			}
			XuLyChiTieu40(sView, ToKhaiInfo.SheetInfos[0], isBindData: true);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, cell11c, cell11cMa, cell11b, cell11bMa, cell11a, cell11aMa, null, cell09, cell10);
			SetCoPhuLucHayKhong(sView, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			UpdateAllChoTongTheoTinhDuLieuCu(sView.FpSpread, ToKhaiInfo, idx01_3_GTGT);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		UpdateAllChoTongTheoTinhDuLieuCu(sView.FpSpread, ToKhaiInfo, idx01_6_CSSX_GTGT);
	}

	public void AutoBindData_4(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		UpdateAllChoTongTheoTinhDuLieuCu(sView.FpSpread, ToKhaiInfo, idx01_6_VT_GTGT);
	}

	public void AutoBindData_5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("C_1");
			if (sView.Cells[array[1], array[0]].CellType.ToString() != "ButtonCellType")
			{
				ButtonCellType buttonCellType = new ButtonCellType();
				buttonCellType.TwoState = false;
				buttonCellType.Text = "Mẫu tải bảng kê";
				buttonCellType.ShadowSize = 0;
				buttonCellType.TextColor = Color.Blue;
				sView.Cells[array[1], array[0]].CellType = buttonCellType;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_6(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("B_2");
			string text = "";
			if (PeriodInfo.KieuKy == "M")
			{
				if (PeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) >= 6 && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) <= 12)
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
				if (PeriodInfo.Nam.Equals("2023") && PeriodInfo.Quy != "1")
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIdxPL_MGT];
			SheetView sheetView = winFpSpread.Sheets[sheetIdxPL_MGT];
			if (hTKKCMCSheetInfo != null)
			{
				List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
				int column = HTKKFile.ColumnLettersToIndex("C") - 1;
				int num = HTKKFile.ColumnLettersToIndex("D") - 1;
				int column2 = HTKKFile.ColumnLettersToIndex("E") - 1;
				int column3 = HTKKFile.ColumnLettersToIndex("I") - 1;
				int num2 = 0;
				for (int i = 0; i < tables.Count; i++)
				{
					num2 = tables[i].Location.DataLocationTemplateRowStart - 1;
					HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
					List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
					for (int j = 0; j < rowsValues.Count; j++)
					{
						string text = sheetView.Cells[num2 + j, column].Text;
						string text2 = sheetView.Cells[num2 + j, column3].Text;
						sheetView.Cells[num2 + j, column2].Text = getTenThueSuat(text2);
					}
				}
			}
			if (ToKhaiInfo.SheetInfos[idx01_2_GTGT] != null)
			{
				int colHeader = HTKKFile.ColumnLettersToIndex("J");
				int colHeader2 = HTKKFile.ColumnLettersToIndex("T");
				int colChilden = HTKKFile.ColumnLettersToIndex("AO");
				int colChilden2 = HTKKFile.ColumnLettersToIndex("AP");
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[idx01_2_GTGT], ToKhaiInfo.SheetInfos[idx01_2_GTGT], -1, -1, colHeader2, colChilden);
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[idx01_2_GTGT], ToKhaiInfo.SheetInfos[idx01_2_GTGT], -1, -1, colHeader, colChilden2);
				foreach (HTKKCMCTableInfo table in ToKhaiInfo.SheetInfos[idx01_2_GTGT].Tables)
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int j = 0; j < table.RowsValues.Count; j++)
					{
						GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[idx01_2_GTGT], dataLocationTemplateRowStart + j, -1, table.RowsInfo, table);
					}
				}
				TinhLaiTongCot10_12_PL01_2(winFpSpread, ToKhaiInfo);
			}
			if (ToKhaiInfo.SheetInfos[idx01_3_GTGT] != null)
			{
				UpdateAllChoTongTheoTinhDuLieuCu(winFpSpread, ToKhaiInfo, idx01_3_GTGT);
				TongTheoTungTinhPL_01_2(winFpSpread, ToKhaiInfo, -1);
				int colHeader3 = HTKKFile.ColumnLettersToIndex("J");
				int colChilden3 = HTKKFile.ColumnLettersToIndex("AQ");
				GeneralFunctions.SetLaiDongCon(winFpSpread.Sheets[idx01_3_GTGT], ToKhaiInfo.SheetInfos[idx01_3_GTGT], -1, -1, colHeader3, colChilden3);
				foreach (HTKKCMCTableInfo table2 in ToKhaiInfo.SheetInfos[idx01_3_GTGT].Tables)
				{
					int dataLocationTemplateRowStart = table2.Location.DataLocationTemplateRowStart;
					for (int j = 0; j < table2.RowsValues.Count; j++)
					{
						GeneralFunctions.UpdateFromSViewToSInfo(winFpSpread.Sheets[idx01_3_GTGT], dataLocationTemplateRowStart + j, -1, table2.RowsInfo, table2);
					}
				}
			}
			if (ToKhaiInfo.SheetInfos[idx01_6_CSSX_GTGT] != null)
			{
				UpdateAllChoTongTheoTinhDuLieuCu(winFpSpread, ToKhaiInfo, idx01_6_CSSX_GTGT);
				TinhTongTheoTinhPl01_3(winFpSpread, ToKhaiInfo, -1);
				TinhCot_17_21_Pl_01_6CSSX(winFpSpread.Sheets[idx01_6_CSSX_GTGT], ToKhaiInfo.SheetInfos[idx01_6_CSSX_GTGT]);
				GeneralFunctions.fpSpReindexingTableTemplate(winFpSpread.Sheets[idx01_6_CSSX_GTGT], ToKhaiInfo.SheetInfos[idx01_6_CSSX_GTGT].Tables[1]);
				HTKKCMCTableInfo current = ToKhaiInfo.SheetInfos[idx01_6_CSSX_GTGT].Tables[0];
				int colMaCQT = HTKKFile.ColumnLettersToIndex("U") - 1;
				int colMaCucThue = HTKKFile.ColumnLettersToIndex("T") - 1;
				int colTenCucThue = HTKKFile.ColumnLettersToIndex("H") - 1;
				GeneralFunctions.BindDataCQTInTable(winFpSpread.Sheets[idx01_6_CSSX_GTGT], current, colMaCQT, colMaCucThue, colTenCucThue);
			}
			if (ToKhaiInfo.SheetInfos[idx01_6_VT_GTGT] != null)
			{
				UpdateAllChoTongTheoTinhDuLieuCu(winFpSpread, ToKhaiInfo, idx01_6_VT_GTGT);
				TinhTongTheoTinhPl01_4(winFpSpread, ToKhaiInfo, -1);
				HTKKCMCTableInfo current = ToKhaiInfo.SheetInfos[idx01_6_VT_GTGT].Tables[0];
				int colMaCQT = HTKKFile.ColumnLettersToIndex("T") - 1;
				int colMaCucThue = HTKKFile.ColumnLettersToIndex("S") - 1;
				int colTenCucThue = HTKKFile.ColumnLettersToIndex("H") - 1;
				GeneralFunctions.BindDataCQTInTable(winFpSpread.Sheets[idx01_6_VT_GTGT], current, colMaCQT, colMaCucThue, colTenCucThue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private string getTenThueSuat(string maThueSuat)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		string result = "";
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_thue_suat_01_GTGT.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 2 && maThueSuat == array[0])
			{
				result = array[1];
				break;
			}
		}
		return result;
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL1_2(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL1_3(winFpSpread.Sheets[2], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL1_6(winFpSpread.Sheets[3], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning, 3);
		checkPL1_6(winFpSpread.Sheets[4], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning, 4);
		checkPL_MGT_ND43(winFpSpread.Sheets[sheetIdxPL_MGT], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_28"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_40"));
			int num = odrColI - 1;
			int column = odrColJ - 1;
			int column2 = odrColL - 1;
			int column3 = odrColK - 1;
			int column4 = odrColH - 1;
			int row = 17;
			int num2 = 36;
			int row2 = 46;
			int row3 = 46;
			int row4 = 47;
			int row5 = 47;
			int row6 = 65;
			string text = sView.Cells[row, column4].Text;
			double num3 = Convert.ToDouble(sView.Cells[row2, column].Value);
			double num4 = Convert.ToDouble(sView.Cells[row3, column2].Value);
			double num5 = Convert.ToDouble(sView.Cells[row4, column].Value);
			double num6 = Convert.ToDouble(sView.Cells[row5, column2].Value);
			string text2 = sView.Cells[row6, column3].Text;
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, cell11c, cell11cMa, cell11b, cell11bMa, cell11a, cell11aMa, null, cell09, cell10))
			{
				value = 2;
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

	private void checkPL1_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int index = 1;
			int index2 = 2;
			string text = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_45"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_46"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_47"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13"));
			string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44"));
			string message6 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_43"));
			string message7 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_49"));
			string message8 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_50"));
			string message9 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_51"));
			int[] array = GeneralFunctions.ArrayCellId("AG_7");
			int[] array2 = GeneralFunctions.ArrayCellId("AG_9");
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("M");
			int num5 = HTKKFile.ColumnLettersToIndex("AH");
			int num6 = HTKKFile.ColumnLettersToIndex("P");
			int num7 = HTKKFile.ColumnLettersToIndex("R");
			int num8 = HTKKFile.ColumnLettersToIndex("T");
			int num9 = HTKKFile.ColumnLettersToIndex("U");
			int num10 = HTKKFile.ColumnLettersToIndex("X");
			int num11 = HTKKFile.ColumnLettersToIndex("AG");
			if (hTKKCMCSheetInfo != null)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num12 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array3 = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					int rowIndex = hTKKCMCTableInfo.Header.StaticCells[0].CellTemplate.RowIndex;
					HTKKFile.RemoveValidate(sView, rowIndex - 1, num9 - 1);
					HTKKFile.RemoveValidate(sView, rowIndex - 1, num2 - 1);
					int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[index].CellTemplate.RowIndex;
					int colIndex = hTKKCMCTableInfo.Header.StaticCells[index].CellTemplate.ColIndex;
					int rowIndex3 = hTKKCMCTableInfo.Header.StaticCells[index2].CellTemplate.RowIndex;
					int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[index2].CellTemplate.ColIndex;
					string text4 = sView.Cells[rowIndex2 - 1, colIndex - 1].Text;
					string text5 = sView.Cells[rowIndex3 - 1, colIndex2 - 1].Text;
					HTKKFile.RemoveValidate(sView, rowIndex3 - 1, colIndex2 - 1);
					if (string.IsNullOrEmpty(text5))
					{
						HTKKFile.SetNodeBackGroundPink(sView, rowIndex3 - 1, colIndex2 - 1, text);
						value = 2;
					}
					else if (text5.Length != 5 && text5.Length != 13 && text5.Length != 14)
					{
						HTKKFile.SetNodeBackGroundPink(sView, rowIndex3 - 1, colIndex2 - 1, text3);
						value = 2;
					}
					else if ((text5.Length == 13 || text5.Length == 14) && !ValidateBase.CheckMaSoThue(text5))
					{
						HTKKFile.SetNodeBackGroundPink(sView, rowIndex3 - 1, colIndex2 - 1, text2);
						value = 2;
					}
					double num13 = Convert.ToDouble(sView.Cells[rowIndex - 1, num9 - 1].Value);
					if (num13 != 100.0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, rowIndex - 1, num9 - 1, message7);
						value = 2;
					}
					if (hTKKCMCTableInfo.RowsValues.Count < 2)
					{
						HTKKFile.SetNodeBackGroundPink(sView, rowIndex - 1, num2 - 1, message8);
						value = 2;
					}
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						HTKKFile.RemoveValidate(sView, num12 + j, num - 1);
						HTKKFile.RemoveValidate(sView, num12 + j, num - 1);
						string text6 = sView.Cells[num12 + j, num11 - 1].Text;
						string text7 = sView.Cells[num12 + j, num - 1].Text;
						string text8 = sView.Cells[num12 + j, num2 - 1].Text;
						string text9 = sView.Cells[num12 + j, num3 - 1].Text;
						string text10 = sView.Cells[num12 + j, num4 - 1].Text;
						string text11 = sView.Cells[num12 + j, num5 - 1].Text;
						string text12 = sView.Cells[num12 + j, num6 - 1].Text;
						string text13 = sView.Cells[num12 + j, num7 - 1].Text;
						double num14 = ConvertHelpers.ConvertToDouble(sView.Cells[num12 + j, num8 - 1].Value);
						double num15 = ConvertHelpers.ConvertToDouble(sView.Cells[num12 + j, num9 - 1].Value);
						bool flag = true;
						if (!string.IsNullOrEmpty(text7) || !string.IsNullOrEmpty(text8) || !string.IsNullOrEmpty(text9) || !string.IsNullOrEmpty(text10) || !string.IsNullOrEmpty(text12) || !string.IsNullOrEmpty(text13) || num14 != 0.0 || num15 != 0.0)
						{
							flag = false;
						}
						if (!flag && string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sView, num12 + j, num - 1, text);
							value = 2;
						}
						if (text6 == "01")
						{
							if (string.IsNullOrEmpty(text10))
							{
								HTKKFile.SetNodeBackGroundPink(sView, num12 + j, num4 - 1, text);
								value = 2;
							}
							if (string.IsNullOrEmpty(text12))
							{
								HTKKFile.SetNodeBackGroundPink(sView, num12 + j, num6 - 1, text);
								value = 2;
							}
						}
						List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sView, hTKKCMCTableInfo, num11, "01");
						if (rowIndexInSViewByCompare.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num12 + j, num - 1, message9);
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL1_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AP");
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("AI");
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			if (!sView.FpSpread.Sheets[2].Visible)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_45"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_46"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_47"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13"));
			string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44"));
			string message6 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_43"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			string[] array = new string[11];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, num2 - 1);
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				string text2 = sView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
				if (text == "01")
				{
					if (!string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num2 - 1, message4);
						value = 2;
					}
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sView, hTKKCMCSheetInfo.Tables[0], num2, text2);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num2 - 1, message5);
						value = 2;
					}
				}
				if (text == "02" && !string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue5So(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num2 - 1, message6);
					value = 2;
				}
				string text3 = sView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
				string text4 = text3 + "00";
				List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sView, hTKKCMCSheetInfo.Tables[0], num4, text3);
				if (rowIndexInSViewByCompare2.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num3 - 1, message3);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL1_6(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning, int sheetIdx)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("F_13");
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("R");
			int value = 0;
			levelWarning.TryGetValue(sheetIdx, out value);
			if (!sView.FpSpread.Sheets[sheetIdx].Visible)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_45"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_46"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_78"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13"));
			string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_44"));
			string message6 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_43"));
			string message7 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_53"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIdx];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, num - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, num2 - 1);
				string text = sView.Cells[array[1], array[0]].Text.Trim();
				string text2 = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text2) && text2.Length != 5 && !ValidateBase.CheckMaSoThue(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num - 1, message4);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && text2.Length == 5 && !ValidateBase.CheckMaSoThue5So(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num - 1, message6);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && text2.Substring(0, 10) != text.Substring(0, 10))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num - 1, message7);
					value = 2;
				}
				List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sView, hTKKCMCSheetInfo.Tables[0], num, text2);
				if (rowIndexInSViewByCompare.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, num - 1, message5);
					value = 2;
				}
				string text3 = sView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text.Trim();
				string text4 = text3 + "00";
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[sheetIdx].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[sheetIdx].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[sheetIdx].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL_MGT_ND43(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(sheetIdxPL_MGT, out value);
			if (!sView.FpSpread.Sheets[sheetIdxPL_MGT].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIdxPL_MGT];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[7];
			double num2 = 0.0;
			double num3 = 0.0;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, odrColC - 1].Text;
				array[1] = sView.Cells[num + i, odrColD - 1].Value.ToString();
				array[2] = sView.Cells[num + i, odrColE - 1].Text;
				array[3] = sView.Cells[num + i, odrColF - 1].Value.ToString();
				array[4] = sView.Cells[num + i, odrColG - 1].Value.ToString();
				array[6] = sView.Cells[num + i, odrColI - 1].Text;
				HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
				HTKKFile.RemoveValidate(sView, num + i, odrColD - 1);
				HTKKFile.RemoveValidate(sView, num + i, odrColE - 1);
				if (array[6] == "5")
				{
					num2 += ConvertHelpers.ConvertToDouble(array[4]);
				}
				if (array[6] == "10")
				{
					num3 += ConvertHelpers.ConvertToDouble(array[4]);
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColC - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_32")));
					value = 2;
				}
				if (ConvertHelpers.ConvertToDouble(array[1]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColD - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_32")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColE - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_32")));
					value = 2;
				}
			}
			double num4 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_47")[1], GeneralFunctions.ArrayCellId("L_47")[0]].Value);
			double num5 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_48")[1], GeneralFunctions.ArrayCellId("L_48")[0]].Value);
			double num6 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("J_47")[1], GeneralFunctions.ArrayCellId("J_47")[0]].Value);
			double num7 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("J_48")[1], GeneralFunctions.ArrayCellId("J_48")[0]].Value);
			if (num4 == num6 * 0.05 - num2)
			{
				HTKKFile.RemoveValidate(sView.FpSpread.Sheets[0], GeneralFunctions.ArrayCellId("L_47")[1], GeneralFunctions.ArrayCellId("L_47")[0]);
			}
			else if (sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_47")[1], GeneralFunctions.ArrayCellId("L_47")[0]].BackColor != Color.Pink)
			{
				HTKKFile.SetNodeBackGroundYellow(sView.FpSpread.Sheets[0], GeneralFunctions.ArrayCellId("L_47")[1], GeneralFunctions.ArrayCellId("L_47")[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_33")));
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[sheetIdxPL_MGT].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[sheetIdxPL_MGT].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[sheetIdxPL_MGT].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckAddPhuLuc_1(FpSpread winFpSpread, SheetView sView)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value)
			{
				frmMessage frmMessage = new frmMessage("01GTGT0009", "", "3", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool CheckAddPhuLuc_2(FpSpread winFpSpread, SheetView sView)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value)
			{
				frmMessage frmMessage = new frmMessage("01GTGT0009", "", "3", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool CheckAddPhuLuc_3(FpSpread winFpSpread, SheetView sView)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value)
			{
				frmMessage frmMessage = new frmMessage("01GTGT0010", "", "3", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool CheckAddPhuLuc_4(FpSpread winFpSpread, SheetView sView)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value)
			{
				frmMessage frmMessage = new frmMessage("01GTGT0009", "", "3", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listRowAndKeyOld_Pl01_2.Clear();
			_listKeyAndRowTo_Pl01_2.Clear();
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listRowAndKeyOld_Pl01_2.Clear();
			_listKeyAndRowTo_Pl01_2.Clear();
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_3(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listRowAndKeyOld_Pl_01_6CSSX.Clear();
			_listKeyAndRowTo_Pl_01_6CSSX.Clear();
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_3(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listRowAndKeyOld_Pl_01_6CSSX.Clear();
			_listKeyAndRowTo_Pl_01_6CSSX.Clear();
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_4(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listRowAndKeyOld_Pl_01_6DVVT.Clear();
			_listKeyAndRowTo_Pl_01_6DVVT.Clear();
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_4(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			_listRowAndKeyOld_Pl_01_6DVVT.Clear();
			_listKeyAndRowTo_Pl_01_6DVVT.Clear();
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], ToKhaiInfo);
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
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], _ToKhaiInfo);
			TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
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
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], _ToKhaiInfo);
			TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], _ToKhaiInfo);
			TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_6(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SetCoPhuLucHayKhong(winFpSpread.Sheets[0], _ToKhaiInfo);
			TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet, out bool Obj)
	{
		Obj = false;
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if ((activeColumnIndex == CellId_ct21[0]) & (activeRowIndex == CellId_ct21[1]))
			{
				Obj = XuLyCheckboxCt21(winFpSpread, isBindData: false, ToKhaiInfo, TTinChung.mst, DicSheet);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_ND15_GTGT_TT80.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_GiamThue_GTGT_23_24_GTGT_TT80.xls");
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("P_47");
			XuLyChiTieu40(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], isBindData: false);
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				if (((iRow - 1 == CellId_ct22[1]) & (iCol - 1 == CellId_ct22[0])) && winFpSpread.Sheets[0].Cells[CellId_ct22[1], CellId_ct22[0]].Text == "")
				{
					winFpSpread.Sheets[0].Cells[CellId_ct22[1], CellId_ct22[0]].Value = 0;
				}
				if (iRow == CellId_ct30[1] + 1 && iCol == CellId_ct30[0] + 1)
				{
					winFpSpread.Sheets[0].Cells[CellId_ct31[1], CellId_ct31[0]].Value = winFpSpread.Sheets[0].Cells[array[1], array[0]].Value;
				}
				TongTheoTungTinhPL_01_2(winFpSpread, ToKhaiInfo, -1);
				if (iRow == CellId_ct32[1] + 1 && iCol == CellId_ct32[0] + 1)
				{
					TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("M");
			int num5 = HTKKFile.ColumnLettersToIndex("AH");
			int num6 = HTKKFile.ColumnLettersToIndex("P");
			int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("R");
			int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("AJ");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("S");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AK");
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("X");
			int colCheck = HTKKFile.ColumnLettersToIndex("AG");
			int colChilden = HTKKFile.ColumnLettersToIndex("AO");
			int colChilden2 = HTKKFile.ColumnLettersToIndex("AP");
			int num10 = HTKKFile.ColumnLettersToIndex("AM");
			int num11 = HTKKFile.ColumnLettersToIndex("AN");
			int[] array = GeneralFunctions.ArrayCellId("AG_7");
			int[] array2 = GeneralFunctions.ArrayCellId("AG_9");
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[50, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			if (iCol == num || iCol == num7)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, num7, colChilden);
			}
			if (iCol == num || iCol == num3)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, hTKKCMCSheetInfo, -1, -1, num3, colChilden2);
			}
			if (iCol == num7 || iCol == num8)
			{
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					if (table.TableName == "01_2_GTGT")
					{
						int rowIndex = table.Header.StaticCells[0].CellTemplate.RowIndex;
						int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
						double num12 = GeneralFunctions.CalculateColSumTable(sheetView, table, num8);
						sheetView.Cells[dataLocationTemplateRowStart - 2, num8 - 1].Value = num12;
					}
				}
			}
			if (iCol == num4)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc, isLock: false);
			}
			int num13 = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			List<int> list = new List<int>();
			if (num13 >= 0)
			{
				list = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCSheetInfo.Tables[num13], colCheck, "01");
			}
			if (iCol == num && list.Count > 1)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num - 1, msg);
			}
			if (iCol == num7 || iCol == num8)
			{
				TinhLaiTongCot10_12_PL01_2(winFpSpread, ToKhaiInfo);
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
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[2];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			int colChilden = HTKKFile.ColumnLettersToIndex("AQ");
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("AI");
			int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("R");
			int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("AK");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("T");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AL");
			int[] array = GeneralFunctions.ArrayCellId("J_16");
			if (iCol == num || iCol == num2)
			{
				GeneralFunctions.SetLaiDongCon(sheetView, sInfo, -1, -1, num2, colChilden);
			}
			if (iCol == num3)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc, isLock: false);
			}
			if (iRow == array[1] + 1)
			{
				TongTheoTungTinhPL_01_2(winFpSpread, ToKhaiInfo, -1);
			}
			else
			{
				TongTheoTungTinhPL_01_2(winFpSpread, ToKhaiInfo, (int)iCol);
			}
			if (iCol == num)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
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
			int[] array = GeneralFunctions.ArrayCellId("O_15");
			int[] array2 = GeneralFunctions.ArrayCellId("O_16");
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("R");
			int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("H");
			int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("T");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("I");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("U");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			int num5 = HTKKFile.ColumnLettersToIndex("Z");
			int num6 = HTKKFile.ColumnLettersToIndex("AA");
			int num7 = HTKKFile.ColumnLettersToIndex("D");
			int num8 = HTKKFile.ColumnLettersToIndex("R");
			if (iCol == num)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num2 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc, isLock: false);
			}
			if (iRow == array[1] + 1 || iRow == array2[1] + 1)
			{
				TinhTongTheoTinhPl01_3(winFpSpread, ToKhaiInfo, -1);
			}
			else
			{
				TinhTongTheoTinhPl01_3(winFpSpread, ToKhaiInfo, (int)iCol);
			}
			TinhCot_17_21_Pl_01_6CSSX(sheetView, hTKKCMCSheetInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[1]);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[4];
			SheetView sheetView = winFpSpread.Sheets[4];
			int[] array = GeneralFunctions.ArrayCellId("N_15");
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("Q");
			int colTenCQTCuc = HTKKFile.ColumnLettersToIndex("H");
			int colMaCQTCuc = HTKKFile.ColumnLettersToIndex("S");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("I");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("T");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			int num5 = HTKKFile.ColumnLettersToIndex("D");
			int num6 = HTKKFile.ColumnLettersToIndex("Q");
			if (iCol == num)
			{
				string maTinh = sheetView.Cells[(int)iRow - 1, num2 - 1].Text.Trim();
				GeneralFunctions.ThayDoiCQTTheoTinhDynamic(sheetView, maTinh, (int)iRow, colTenCQTCuc, colMaCQTCuc, colTenCQTChiCuc, colMaCQTChiCuc, isLock: false);
			}
			if (iRow == array[1] + 1)
			{
				TinhTongTheoTinhPl01_4(winFpSpread, ToKhaiInfo, -1);
			}
			else
			{
				TinhTongTheoTinhPl01_4(winFpSpread, ToKhaiInfo, (int)iCol);
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
			TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				TinhLaiChiTieu32_TK(winFpSpread, ToKhaiInfo);
				TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TinhTongTheoTinhPl01_4(winFpSpread, ToKhaiInfo, -1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool XuLyCheckboxCt21(FpSpread winFpSpread, bool isBindData, HTKKCMCToKhaiInfo ToKhaiInfo, string mstNNT, Dictionary<string, int> DicSheet)
	{
		//IL_009f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Invalid comparison between Unknown and I4
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("I") - 1;
			bool flag = (bool)winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value;
			string text = "J_40###L_40###J_41###L_41###L_42###J_44###J_46###J_47###L_47###J_48###L_48###J_49";
			string text2 = "###";
			string[] array = text.Split(new string[1] { text2 }, StringSplitOptions.None);
			SheetView sheetView = winFpSpread.Sheets[0];
			winFpSpread.Sheets[0].Protect = true;
			if (bool.Parse(flag.ToString()))
			{
				DialogResult val = (DialogResult)6;
				if (!isBindData)
				{
					val = MessageBox.Show("NNT chỉ được kê khai các chỉ tiêu: [22], [37], [38], [39a], [40b], [42]. Các chỉ tiêu còn lại khóa không cho kê khai. Bạn có tiếp tục thực hiện không?", "Thong bao", (MessageBoxButtons)4);
				}
				if ((int)val == 6)
				{
					for (int i = 0; i < array.Length; i++)
					{
						int[] array2 = GeneralFunctions.ArrayCellId(array[i]);
						sheetView.Cells[array2[1], array2[0]].Locked = true;
						sheetView.Cells[array2[1], array2[0]].Value = 0;
					}
					if (!isBindData)
					{
						return true;
					}
				}
				else
				{
					winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value = false;
				}
			}
			else
			{
				for (int i = 0; i < array.Length; i++)
				{
					int[] array2 = GeneralFunctions.ArrayCellId(array[i]);
					sheetView.Cells[array2[1], array2[0]].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
		return false;
	}

	private void XuLyChiTieu40(SheetView sView, HTKKCMCSheetInfo sInfoTK, bool isBindData)
	{
		try
		{
			double num = 0.0;
			int row = 58;
			int row2 = 60;
			int column = HTKKFile.ColumnLettersToIndex("L") - 1;
			num = ConvertHelpers.ConvertToDouble(sView.Cells[row, column].Value);
			if (num != 0.0)
			{
				sView.Cells[row2, column].Value = 0;
				sView.Cells[row2, column].Locked = true;
			}
			else
			{
				sView.Cells[row2, column].Locked = false;
			}
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	private double? GetPreperiod_CT43Value(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo)
	{
		try
		{
			double? result = null;
			List<string> preperiodFileName_ThangQuy = GeneralFunctions.GetPreperiodFileName_ThangQuy(periodInfo, toKhaiInfo);
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			bool flag = periodInfo.LoaiToKhai.Equals("C");
			if (directoryInfo.Exists && flag && preperiodFileName_ThangQuy.Count > 0)
			{
				FileInfo[] files = directoryInfo.GetFiles();
				if (files.Length > 0)
				{
					bool flag2 = false;
					foreach (string item in preperiodFileName_ThangQuy)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(item))
							{
								string valueInXml = GeneralFunctions.GetValueInXml(text + "\\" + fileInfo.Name, "L_62");
								result = Convert.ToDouble(valueInXml);
								flag2 = true;
								break;
							}
						}
						if (flag2)
						{
							break;
						}
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}

	private string getTenNNKD(string maNNKD)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		string result = "";
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_NNKD.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 3 && maNNKD == array[1] && array[0] == "01_GTGT_TT80")
			{
				result = array[2];
				break;
			}
		}
		return result;
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			TinhLaiTongCot10_12_PL01_2(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public bool UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			TongTheoTungTinhPL_01_2(winFpSpread, ToKhaiInfo, -1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public bool UpdateFormula_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			SheetView sView = winFpSpread.Sheets[3];
			TinhTongTheoTinhPl01_3(winFpSpread, ToKhaiInfo, -1);
			TinhCot_17_21_Pl_01_6CSSX(sView, hTKKCMCSheetInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return true;
	}

	public void TinhTongTheoTinhPl01_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			SheetView sheetView = winFpSpread.Sheets[3];
			if (hTKKCMCSheetInfo != null && sheetView != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("F");
				int num2 = HTKKFile.ColumnLettersToIndex("R");
				int num3 = HTKKFile.ColumnLettersToIndex("K");
				int num4 = HTKKFile.ColumnLettersToIndex("L");
				int num5 = HTKKFile.ColumnLettersToIndex("N");
				int num6 = HTKKFile.ColumnLettersToIndex("P");
				int num7 = HTKKFile.ColumnLettersToIndex("Z");
				int num8 = HTKKFile.ColumnLettersToIndex("AA");
				int value = HTKKFile.ColumnLettersToIndex("D");
				int num9 = HTKKFile.ColumnLettersToIndex("R");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(num, value);
				dictionary.Add(num2, num9);
				dictionary.Add(num7, num7);
				dictionary.Add(num8, num8);
				if (iCol < 0 || iCol == num)
				{
					Dictionary<string, int> checkExists = new Dictionary<string, int>();
					GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
				}
				if (iCol == num3 || iCol == num4)
				{
					List<int> list = new List<int>();
					list.Add(num7);
					list.Add(num8);
					GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(sheetView, hTKKCMCSheetInfo, "01_6_GTGT_1", "01_6_GTGT_2", num2, num9, list);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongTheoTungTinhPL_01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[2];
			if (hTKKCMCSheetInfo != null && sheetView != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("M");
				int num2 = HTKKFile.ColumnLettersToIndex("AI");
				int num3 = HTKKFile.ColumnLettersToIndex("U");
				int num4 = HTKKFile.ColumnLettersToIndex("Y");
				int value = HTKKFile.ColumnLettersToIndex("C");
				int num5 = HTKKFile.ColumnLettersToIndex("AL");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(num, value);
				dictionary.Add(num2, num5);
				dictionary.Add(num4, num4);
				if (iCol < 0 || iCol == num)
				{
					Dictionary<string, int> checkExists = new Dictionary<string, int>();
					GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
				}
				if (iCol == num3)
				{
					List<int> list = new List<int>();
					list.Add(num4);
					GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(sheetView, hTKKCMCSheetInfo, "01_3_GTGT_1", "01_3_GTGT_2", num2, num5, list);
				}
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[1]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhTongTheoTinhPl01_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iCol)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[4];
			SheetView sheetView = winFpSpread.Sheets[4];
			if (hTKKCMCSheetInfo != null && sheetView != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("F");
				int num2 = HTKKFile.ColumnLettersToIndex("Q");
				int num3 = HTKKFile.ColumnLettersToIndex("K");
				int num4 = HTKKFile.ColumnLettersToIndex("L");
				int num5 = HTKKFile.ColumnLettersToIndex("N");
				int num6 = HTKKFile.ColumnLettersToIndex("P");
				int value = HTKKFile.ColumnLettersToIndex("D");
				int num7 = HTKKFile.ColumnLettersToIndex("Q");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(num, value);
				dictionary.Add(num2, num7);
				dictionary.Add(num5, num5);
				if (iCol < 0 || iCol == num)
				{
					Dictionary<string, int> checkExists = new Dictionary<string, int>();
					GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
				}
				if (iCol == num4)
				{
					List<int> list = new List<int>();
					list.Add(num5);
					GeneralFunctions.TinhLaiCacGiaTriKhiThayDoiTuBangSangBang(sheetView, hTKKCMCSheetInfo, "01_6_GTGT_1", "01_6_GTGT_2", num2, num7, list);
				}
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[1]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetCoPhuLucHayKhong(SheetView sViewToKhai, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		int[] array = GeneralFunctions.ArrayCellId("P_9");
		int[] array2 = GeneralFunctions.ArrayCellId("Q_9");
		int[] array3 = GeneralFunctions.ArrayCellId("R_9");
		int[] array4 = GeneralFunctions.ArrayCellId("S_9");
		int[] array5 = GeneralFunctions.ArrayCellId("T_9");
		int[] array6 = GeneralFunctions.ArrayCellId("U_9");
		sViewToKhai.Cells[array[1], array[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		sViewToKhai.Cells[array2[1], array2[0]].Value = ((ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
		sViewToKhai.Cells[array3[1], array3[0]].Value = ((ToKhaiInfo.SheetInfos[3] != null) ? 1 : 0);
		sViewToKhai.Cells[array4[1], array4[0]].Value = ((ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
		sViewToKhai.Cells[array5[1], array5[0]].Value = ((ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
		sViewToKhai.Cells[array6[1], array6[0]].Value = ((ToKhaiInfo.SheetInfos[6] != null) ? 1 : 0);
	}

	public void TinhCot_17_21_Pl_01_6CSSX(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sView == null || sInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("N");
			int num2 = HTKKFile.ColumnLettersToIndex("Z");
			int num3 = HTKKFile.ColumnLettersToIndex("P");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("AA");
			string pathCell19 = "Pluc_01_6_GTGT_CSSX/CongCacTinh/ct19";
			string pathCell20 = "Pluc_01_6_GTGT_CSSX/SoThuePhaiNop/ct20";
			string pathCell23 = "Pluc_01_6_GTGT_CSSX/CongCacTinh/ct23";
			string pathCell24 = "Pluc_01_6_GTGT_CSSX/SoThuePhaiNop/ct24";
			string pathCell20Formular = "path_22";
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCell20Formular));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCell19));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCell20));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCell23));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCell24));
			double num6 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num7 = 0; num7 < table.RowsValues.Count; num7++)
				{
					sView.Cells[dataLocationTemplateRowStart + num7 - 1, num - 1].Value = ((num6 < 0.0) ? ((object)0) : sView.Cells[dataLocationTemplateRowStart + num7 - 1, num2 - 1].Value);
					sView.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Value = ((num6 < 0.0) ? sView.Cells[dataLocationTemplateRowStart + num7 - 1, num5 - 1].Value : ((object)0));
				}
			}
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, num - 1].Value = ((num6 < 0.0) ? ((object)0) : sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, num2 - 1].Value);
			sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, num - 1].Value = ((num6 < 0.0) ? ((object)0) : sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, num2 - 1].Value);
			sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, num4 - 1].Value = ((num6 < 0.0) ? sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, num5 - 1].Value : ((object)0));
			sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, num4 - 1].Value = ((num6 < 0.0) ? sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, num5 - 1].Value : ((object)0));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhCt33TrenToKhai(SheetView sViewTK)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("L_48");
			int[] array2 = GeneralFunctions.ArrayCellId("P_48");
			int[] array3 = GeneralFunctions.ArrayCellId("Q_48");
			int[] array4 = GeneralFunctions.ArrayCellId("T_9");
			int[] array5 = GeneralFunctions.ArrayCellId("U_9");
			string text = sViewTK.Cells[array4[1], array4[0]].Text.Trim();
			string text2 = sViewTK.Cells[array5[1], array5[0]].Text.Trim();
			if (text == "1")
			{
				sViewTK.Cells[array[1], array[0]].Value = sViewTK.Cells[array2[1], array2[0]].Value;
			}
			else if (text2 == "1")
			{
				sViewTK.Cells[array[1], array[0]].Value = sViewTK.Cells[array3[1], array3[0]].Value;
			}
			else
			{
				sViewTK.Cells[array[1], array[0]].Value = sViewTK.Cells[array3[1], array3[0]].Value;
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
			SheetView sheetView = winFpSpread.Sheets[5];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[5];
			if (hTKKCMCSheetInfo != null)
			{
				TinhCt33TrenToKhai(winFpSpread.Sheets[0]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateAllChoTongTheoTinhDuLieuCu(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndex)
	{
		try
		{
			if (sheetIndex == idx01_3_GTGT)
			{
				List<int> list = new List<int>();
				int item = HTKKFile.ColumnLettersToIndex("M");
				int item2 = HTKKFile.ColumnLettersToIndex("AI");
				list.Add(item);
				list.Add(item2);
				_listRowAndKeyOld_Pl01_2 = GeneralFunctions.GetRowOnSViewCreateDictionary(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2].Tables[0], list);
			}
			if (sheetIndex == idx01_6_CSSX_GTGT)
			{
				List<int> list = new List<int>();
				int item3 = HTKKFile.ColumnLettersToIndex("F");
				int item4 = HTKKFile.ColumnLettersToIndex("R");
				list.Add(item3);
				list.Add(item4);
				_listRowAndKeyOld_Pl_01_6CSSX = GeneralFunctions.GetRowOnSViewCreateDictionary(winFpSpread.Sheets[3], ToKhaiInfo.SheetInfos[3].Tables[0], list);
			}
			if (sheetIndex == idx01_6_VT_GTGT)
			{
				List<int> list = new List<int>();
				int item3 = HTKKFile.ColumnLettersToIndex("F");
				int item4 = HTKKFile.ColumnLettersToIndex("Q");
				list.Add(item3);
				list.Add(item4);
				_listRowAndKeyOld_Pl_01_6DVVT = GeneralFunctions.GetRowOnSViewCreateDictionary(winFpSpread.Sheets[4], ToKhaiInfo.SheetInfos[4].Tables[0], list);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiTongCot10_12_PL01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo != null)
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int colSum = HTKKFile.ColumnLettersToIndex("T");
			int colSum2 = HTKKFile.ColumnLettersToIndex("X");
			string pathSum10 = "PLuc_01_2_GTGT/ct13";
			string pathSum12 = "PLuc_01_2_GTGT/ct14";
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathSum10);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathSum12);
			double num = GeneralFunctions.CalculateSumHeaderByGroup(sheetView, hTKKCMCSheetInfo.Tables, colSum);
			double num2 = GeneralFunctions.CalculateSumByGroup(sheetView, hTKKCMCSheetInfo.Tables, colSum2, isView: false);
			sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
			sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
		}
	}

	public void TinhLaiChiTieu32_TK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		int[] array = GeneralFunctions.ArrayCellId("J_48");
		int[] array2 = GeneralFunctions.ArrayCellId("J_9");
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[6];
		if (hTKKCMCSheetInfo != null)
		{
			SheetView sheetView2 = winFpSpread.Sheets[6];
			sheetView.Cells[array[1], array[0]].Value = sheetView2.Cells[array2[1], array2[0]].Value;
		}
	}
}
