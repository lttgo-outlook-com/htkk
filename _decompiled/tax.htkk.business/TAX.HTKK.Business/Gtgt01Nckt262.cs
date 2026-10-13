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
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt01Nckt262
{
	private int[] cellPL_01_1TD = GeneralFunctions.ArrayCellId("Q_5");

	private int[] cellPL_01_2TD = GeneralFunctions.ArrayCellId("R_5");

	private int sheetIdxKHBS = 7;

	private int sheetIdxPL_MGT = 6;

	private string cellId_maTieuMuc = "L_14";

	private string cellId_ckbGiaHan = "B_18";

	private string cellId_cbbGiaHan = "D_20";

	private string cellId_maGiaHan = "P_20";

	private string cellId_ct21 = "I_23";

	private string cellId_ct22 = "L_24";

	private string cellId_ct26 = "J_30";

	private string cellId_ct29 = "J_32";

	private string cellId_ct30 = "J_33";

	private string cellId_ct31 = "L_33";

	private string cellId_ct32 = "J_34";

	private string cellId_ct33 = "L_34";

	private string cellId_ct32a = "J_35";

	private string cellId_ct34 = "J_36";

	private string cellId_ct39 = "L_41";

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

	private int[] CellId_maTieuMuc => GeneralFunctions.ArrayCellId(cellId_maTieuMuc);

	private int[] CellId_ckbGiaHan => GeneralFunctions.ArrayCellId(cellId_ckbGiaHan);

	private int[] CellId_cbbGiaHan => GeneralFunctions.ArrayCellId(cellId_cbbGiaHan);

	private int[] CellId_maGiaHan => GeneralFunctions.ArrayCellId(cellId_maGiaHan);

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

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId(cellId_ct39);

	public void TongHopDieuChinhSoThuePhaiNopTang_6(FpSpread winFpSpread, HTKKCMCSheetInfo KHBS_LanTruoc, HTKKCMCSheetInfo sKHBSInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[sheetIdxKHBS];
			SheetView sheetView2 = winFpSpread.Sheets[0];
			double? num = null;
			double? num2 = null;
			double? num3 = null;
			double? num4 = null;
			for (int i = 0; i < KHBS_LanTruoc.StaticCells.Count; i++)
			{
				if (KHBS_LanTruoc.StaticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct40")
				{
					num = (double)sheetView2.Cells[KHBS_LanTruoc.StaticCells[i].CellTemplate.RowIndex - 1, KHBS_LanTruoc.StaticCells[i].CellTemplate.ColIndex - 1].Value;
					num2 = double.Parse(KHBS_LanTruoc.StaticCells[i].Value);
				}
				else if (KHBS_LanTruoc.StaticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct43")
				{
					num3 = (double)sheetView2.Cells[KHBS_LanTruoc.StaticCells[i].CellTemplate.RowIndex - 1, KHBS_LanTruoc.StaticCells[i].CellTemplate.ColIndex - 1].Value;
					num4 = double.Parse(KHBS_LanTruoc.StaticCells[i].Value);
				}
			}
			if (num.HasValue)
			{
				for (int j = 0; j < sKHBSInfo.StaticCells.Count; j++)
				{
					if (sKHBSInfo.StaticCells[j].Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct40/ct4")
					{
						sheetView.Cells[sKHBSInfo.StaticCells[j].CellTemplate.RowIndex - 1, sKHBSInfo.StaticCells[j].CellTemplate.ColIndex - 1].Value = num2;
					}
					else if (sKHBSInfo.StaticCells[j].Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct40/ct5")
					{
						sheetView.Cells[sKHBSInfo.StaticCells[j].CellTemplate.RowIndex - 1, sKHBSInfo.StaticCells[j].CellTemplate.ColIndex - 1].Value = num;
					}
					else if (sKHBSInfo.StaticCells[j].Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct40/ct6")
					{
						sheetView.Cells[sKHBSInfo.StaticCells[j].CellTemplate.RowIndex - 1, sKHBSInfo.StaticCells[j].CellTemplate.ColIndex - 1].Value = num - num2;
					}
				}
			}
			if (!num3.HasValue)
			{
				return;
			}
			for (int j = 0; j < sKHBSInfo.StaticCells.Count; j++)
			{
				if (sKHBSInfo.StaticCells[j].Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct43/ct4")
				{
					sheetView.Cells[sKHBSInfo.StaticCells[j].CellTemplate.RowIndex - 1, sKHBSInfo.StaticCells[j].CellTemplate.ColIndex - 1].Value = num4;
				}
				else if (sKHBSInfo.StaticCells[j].Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct43/ct5")
				{
					sheetView.Cells[sKHBSInfo.StaticCells[j].CellTemplate.RowIndex - 1, sKHBSInfo.StaticCells[j].CellTemplate.ColIndex - 1].Value = num3;
				}
				else if (sKHBSInfo.StaticCells[j].Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct43/ct6")
				{
					sheetView.Cells[sKHBSInfo.StaticCells[j].CellTemplate.RowIndex - 1, sKHBSInfo.StaticCells[j].CellTemplate.ColIndex - 1].Value = num3 - num4;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

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
			bool flag = false;
			try
			{
				flag = (bool)sView.Cells[CellId_ckbGiaHan[1], CellId_ckbGiaHan[0]].Value;
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			}
			if (flag)
			{
				sView.Cells[CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]].Locked = false;
			}
			else
			{
				sView.Cells[CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]].Locked = true;
			}
			sView.Cells[GeneralFunctions.ArrayCellId("K_53")[1], GeneralFunctions.ArrayCellId("K_53")[0]].Value = ((PeriodInfo.LoaiToKhai == "C") ? 1 : 0);
			sView.Cells[GeneralFunctions.ArrayCellId("N_14")[1], GeneralFunctions.ArrayCellId("N_14")[0]].Text = PeriodInfo.TieuMucHachToanVal;
			sView.Cells[GeneralFunctions.ArrayCellId("M_14")[1], GeneralFunctions.ArrayCellId("M_14")[0]].Text = getTenNNKD(PeriodInfo.TieuMucHachToanVal);
			if (PeriodInfo.TieuMucHachToanVal.Equals("01") || PeriodInfo.TieuMucHachToanVal.Equals("09"))
			{
				sView.Cells[GeneralFunctions.ArrayCellId("L_14")[1], GeneralFunctions.ArrayCellId("L_14")[0]].Text = "1705";
			}
			else if (PeriodInfo.TieuMucHachToanVal.Equals("05") || PeriodInfo.TieuMucHachToanVal.Equals("08"))
			{
				sView.Cells[GeneralFunctions.ArrayCellId("L_14")[1], GeneralFunctions.ArrayCellId("L_14")[0]].Text = "1704";
			}
			else
			{
				sView.Cells[GeneralFunctions.ArrayCellId("L_14")[1], GeneralFunctions.ArrayCellId("L_14")[0]].Text = "1701";
			}
			bool flag2 = false;
			XuLyCheckboxCt21(sView.FpSpread, isBindData: true, ToKhaiInfo, TTinChung.mst, DicSheet);
			int row = 23;
			int column = odrColL - 1;
			int column2 = odrColP - 1;
			double? preperiod_CT43Value = GetPreperiod_CT43Value(PeriodInfo, ToKhaiInfo, HeaderInfo);
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) == 2021 && (ConvertHelpers.ConvertToInt(PeriodInfo.Thang) == 1 || ConvertHelpers.ConvertToInt(PeriodInfo.Quy) == 1))
			{
				sView.Cells[row, column2].Value = 0;
			}
			else
			{
				sView.Cells[row, column2].Value = preperiod_CT43Value ?? 0.0;
			}
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile || PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[row, column].ResetFormula();
			}
			else
			{
				sView.Cells[row, column].Formula = "P24";
			}
			if (ToKhaiInfo.SheetInfos[1] != null && PeriodInfo.LoaiToKhai == "C")
			{
				sView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = strCellId_01_5_tong6;
			}
			else if (PeriodInfo.LoaiToKhai == "B")
			{
				if (ToKhaiInfo.SheetInfos[1] != null)
				{
					sView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = strCellId_01_5_tong6;
				}
				sView.Cells[CellId_ct39[1], CellId_ct39[0]].Locked = true;
			}
			sView.FpSpread.Sheets[2].Cells[GeneralFunctions.ArrayCellId("G_12")[1], GeneralFunctions.ArrayCellId("G_12")[0]].Formula = "'" + sView.SheetName + "'!L45";
			sView.FpSpread.Sheets[2].Cells[GeneralFunctions.ArrayCellId("G_14")[1], GeneralFunctions.ArrayCellId("G_14")[0]].Formula = "'" + sView.SheetName + "'!J36";
			XuLyChiTieu40(sView, ToKhaiInfo.SheetInfos[0], isBindData: true);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				if (sView.FpSpread.Sheets[1].Visible)
				{
					double num = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[20].Value);
					sView.Cells[CellId_ct39[1], CellId_ct39[0]].ResetFormula();
					sView.Cells[CellId_ct39[1], CellId_ct39[0]].Value = num;
				}
				else
				{
					sView.Cells[CellId_ct39[1], CellId_ct39[0]].Locked = true;
				}
			}
			CheckPhuLucAnHien(sView.FpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[2].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				LayThemThongTinTinhHuyen_PL01_6(sView, dataLocationTemplateRowStart + i);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[3].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				LayThemThongTinTinhHuyen_PL_01_7(sView, dataLocationTemplateRowStart + i);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[5];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					LayThemThongTinCQTChiCuc_PL_01_2_TD(sView, dataLocationTemplateRowStart + j);
					LayThemThongTinTinhHuyen_PL_01_2_TD(sView, dataLocationTemplateRowStart + j);
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
				}
			}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.Cells[GeneralFunctions.ArrayCellId("B_18")[1], GeneralFunctions.ArrayCellId("B_18")[0]].Value = !string.IsNullOrEmpty(hTKKCMCSheetInfo.StaticCells[37].Value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[sheetIdxPL_MGT];
			SheetView sheetView2 = winFpSpread.Sheets[sheetIdxPL_MGT];
			if (hTKKCMCSheetInfo2 != null)
			{
				List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo2.Tables;
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
						string text = sheetView2.Cells[num2 + j, column].Text;
						string text2 = sheetView2.Cells[num2 + j, column3].Text;
						sheetView2.Cells[num2 + j, column2].Text = getTenThueSuat(text2);
					}
				}
			}
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo3 = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo3 != null)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo3.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					LayThemThongTinTinhHuyen_PL01_6(sView, dataLocationTemplateRowStart + j);
					GeneralFunctions.UpdateFromSViewToSInfo(sView, dataLocationTemplateRowStart + j, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
				}
			}
			SheetView sView2 = winFpSpread.Sheets[3];
			HTKKCMCSheetInfo hTKKCMCSheetInfo4 = ToKhaiInfo.SheetInfos[3];
			if (hTKKCMCSheetInfo4 != null)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo4.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo2.RowsValues.Count; j++)
				{
					LayThemThongTinTinhHuyen_PL_01_7(sView2, dataLocationTemplateRowStart + j);
					GeneralFunctions.UpdateFromSViewToSInfo(sView2, dataLocationTemplateRowStart + j, -1, hTKKCMCTableInfo2.RowsInfo, hTKKCMCTableInfo2);
				}
			}
			SheetView sView3 = winFpSpread.Sheets[5];
			HTKKCMCSheetInfo hTKKCMCSheetInfo5 = ToKhaiInfo.SheetInfos[5];
			if (hTKKCMCSheetInfo5 == null)
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo5.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo5.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo3.RowsValues.Count; j++)
				{
					LayThemThongTinCQTChiCuc_PL_01_2_TD(sView3, dataLocationTemplateRowStart + j);
					LayThemThongTinTinhHuyen_PL_01_2_TD(sView3, dataLocationTemplateRowStart + j);
					GeneralFunctions.SetLaiDongCon(sView3, hTKKCMCSheetInfo5, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
					GeneralFunctions.SetLaiDongCon(sView3, hTKKCMCSheetInfo5, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
					GeneralFunctions.UpdateFromSViewToSInfo(sView3, dataLocationTemplateRowStart + j, -1, hTKKCMCTableInfo3.RowsInfo, hTKKCMCTableInfo3);
				}
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
		checkPL1_5(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL1_6(winFpSpread.Sheets[2], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL1_7(winFpSpread.Sheets[3], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL1_1TD(winFpSpread.Sheets[4], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL1_2TD(winFpSpread.Sheets[5], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL_MGT_ND92(winFpSpread.Sheets[sheetIdxPL_MGT], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public bool SetActivePrintSheet_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			bool result = false;
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string text = "";
			int column = rowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, column].Text;
				if (HTKKFormater.isNullString(text))
				{
					result = true;
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			int num = odrColI - 1;
			int column = odrColJ - 1;
			int num2 = odrColL - 1;
			int column2 = odrColK - 1;
			int num3 = 23;
			int row = 32;
			int num4 = 32;
			int row2 = 33;
			int num5 = 33;
			int row3 = 50;
			double value2 = Convert.ToDouble(sView.Cells[row, column].Value);
			double value3 = Convert.ToDouble(sView.Cells[num4, num2].Value);
			double value4 = Convert.ToDouble(sView.Cells[row2, column].Value);
			double value5 = Convert.ToDouble(sView.Cells[num5, num2].Value);
			string text = sView.Cells[row3, column2].Text;
			bool flag = false;
			try
			{
				flag = (bool)sView.Cells[CellId_ckbGiaHan[1], CellId_ckbGiaHan[0]].Value;
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			}
			string text2 = sView.Cells[CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]].Text;
			if (flag & string.IsNullOrEmpty(text2))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_cbbGiaHan[1], CellId_cbbGiaHan[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_64")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sView, CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]);
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIdxPL_MGT];
			if (hTKKCMCSheetInfo == null)
			{
				double num6 = ((Math.Abs(value2) * 0.0001 < 100000.0) ? (Math.Abs(value2) * 0.0001) : 100000.0);
				if (Math.Abs(value3) > Math.Abs(value2) * 0.05 + num6 || Math.Abs(value3) < Math.Abs(value2) * 0.05 - num6)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num4, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num4, num2);
				}
				double num7 = ((Math.Abs(value4) * 0.0001 < 100000.0) ? (Math.Abs(value4) * 0.0001) : 100000.0);
				if (Math.Abs(value5) > Math.Abs(value4) * 0.1 + num7 || Math.Abs(value5) < Math.Abs(value4) * 0.1 - num7)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num5, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_16")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num5, num2);
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

	private void checkPL1_5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			if (!sView.FpSpread.Sheets[1].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int num2 = rowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			int num3 = rowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int num4 = rowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int num5 = rowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num6 = rowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int column = rowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int num7 = rowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int column2 = rowsInfo.Cells[7].CellTemplate.ColIndex - 1;
			string[] array = new string[8];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, num2].Text;
				array[1] = sView.Cells[num + i, num3].Text;
				array[2] = sView.Cells[num + i, num4].Text;
				array[3] = sView.Cells[num + i, num5].Text;
				array[4] = sView.Cells[num + i, num6].Text;
				array[5] = sView.Cells[num + i, column].Text;
				array[6] = sView.Cells[num + i, num7].Text;
				array[7] = sView.Cells[num + i, column2].Text;
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
					HTKKFile.RemoveValidate(sView, num + i, num3);
					HTKKFile.RemoveValidate(sView, num + i, num4);
					HTKKFile.RemoveValidate(sView, num + i, num5);
					HTKKFile.RemoveValidate(sView, num + i, num6);
					HTKKFile.RemoveValidate(sView, num + i, num7);
					continue;
				}
				if (HTKKFormater.isNullString(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
				}
				if (HTKKFormater.isNullString(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num4);
				}
				if (string.IsNullOrEmpty(sView.Cells[num + i, num5].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num5);
				}
				if (string.IsNullOrEmpty(sView.Cells[num + i, num6].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num6);
				}
				double num8 = Convert.ToDouble(sView.Cells[num + i, num7].Value);
				if (num8 > 0.0)
				{
					HTKKFile.RemoveValidate(sView, num + i, num7);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				string text = sView.Cells[num + i, num3].Text;
				if (!string.IsNullOrEmpty(text))
				{
					bool success = false;
					DateTime dateTime = HTKKFormater.ConvertToDateTime(text, out success);
					if (!success)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage("0071"));
						value = 2;
					}
					else if (dateTime > DateTime.Now.Date)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_55")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, num + i, num3);
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
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

	private void checkPL1_6(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			if (!sView.FpSpread.Sheets[2].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[11];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, odrColC - 1].Text;
				array[1] = sView.Cells[num + i, odrColD - 1].Text;
				array[2] = sView.Cells[num + i, odrColE - 1].Text;
				array[3] = sView.Cells[num + i, odrColF - 1].Text;
				array[4] = sView.Cells[num + i, odrColL - 1].Text;
				array[5] = sView.Cells[num + i, odrColG - 1].Text;
				array[6] = sView.Cells[num + i, odrColH - 1].Text;
				array[7] = sView.Cells[num + i, odrColI - 1].Text;
				array[8] = sView.Cells[num + i, odrColJ - 1].Text;
				array[9] = sView.Cells[num + i, odrColK - 1].Text;
				array[10] = sView.Cells[num + i, odrColM - 1].Text;
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
					HTKKFile.RemoveValidate(sView, num + i, odrColD - 1);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColC - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
				}
				if (!string.IsNullOrEmpty(array[1]))
				{
					if (!ValidateBase.CheckMaSoThue(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColD - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_82")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, num + i, odrColD - 1);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColD - 1);
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColE - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColE - 1);
				}
				if (string.IsNullOrEmpty(array[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColF - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColF - 1);
				}
			}
			string text = sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_45")[1], GeneralFunctions.ArrayCellId("L_45")[0]].Text;
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_6_GTGT/PBoGTGTKhongHTKT/ct16"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_6_GTGT/PBoGTGTKhongHTKT/ct18"));
			if (text == "0")
			{
				if (hTKKCMCCellInfo.Value != "0")
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_96")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, testColor: true);
				}
				if (ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo2.Value) != 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_97")));
					value = 2;
				}
				else if (sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[GeneralFunctions.ArrayCellId("B_47")[1], GeneralFunctions.ArrayCellId("B_47")[0]].Text == "1")
				{
					HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, testColor: true);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, testColor: true);
				if (sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[GeneralFunctions.ArrayCellId("B_47")[1], GeneralFunctions.ArrayCellId("B_47")[0]].Text == "1")
				{
					HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, testColor: true);
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

	private void checkPL1_7(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(3, out value);
			if (!sView.FpSpread.Sheets[3].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[8];
			string defaultValue = hTKKCMCTableInfo.RowsInfo.Cells[6].DefaultValue;
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, odrColC - 1].Text;
				array[1] = sView.Cells[num + i, odrColD - 1].Text;
				array[2] = sView.Cells[num + i, odrColE - 1].Text;
				array[3] = sView.Cells[num + i, odrColF - 1].Text;
				array[4] = sView.Cells[num + i, odrColI - 1].Text;
				array[5] = sView.Cells[num + i, odrColG - 1].Value.ToString();
				array[6] = defaultValue;
				array[7] = sView.Cells[num + i, odrColJ - 1].Text;
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColC - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
				}
				string text = sView.Cells[num + i, odrColJ - 1].Text.Trim();
				if (text == headerInfo.maCQTCapCuc)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColE - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_116")));
					value = 2;
				}
			}
			double num2 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_45")[1], GeneralFunctions.ArrayCellId("L_45")[0]].Value);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_7_GTGT/tongCong"));
			if (num2 == 0.0)
			{
				if (ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) != 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_98")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, testColor: true);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, testColor: true);
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[3].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[3].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[3].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL_MGT_ND92(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
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
				array[5] = sView.Cells[num + i, odrColH - 1].Text;
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
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColC - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
					value = 2;
				}
				if (ConvertHelpers.ConvertToDouble(array[1]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColD - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColE - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
					value = 2;
				}
			}
			double num4 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_33")[1], GeneralFunctions.ArrayCellId("L_33")[0]].Value);
			double num5 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_34")[1], GeneralFunctions.ArrayCellId("L_34")[0]].Value);
			double num6 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("J_33")[1], GeneralFunctions.ArrayCellId("J_33")[0]].Value);
			double num7 = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("J_34")[1], GeneralFunctions.ArrayCellId("J_34")[0]].Value);
			string text = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("B") - 1].Text;
			string text2 = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("B") - 1].Text;
			if (num4 == num6 * 0.05 - num2)
			{
				HTKKFile.RemoveValidate(sView.FpSpread.Sheets[0], GeneralFunctions.ArrayCellId("L_33")[1], GeneralFunctions.ArrayCellId("L_33")[0]);
			}
			else if (sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_33")[1], GeneralFunctions.ArrayCellId("L_33")[0]].BackColor != Color.Pink)
			{
				HTKKFile.SetNodeBackGroundYellow(sView.FpSpread.Sheets[0], GeneralFunctions.ArrayCellId("L_33")[1], GeneralFunctions.ArrayCellId("L_33")[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_103")));
			}
			if (num5 == num7 * 0.1 - num3)
			{
				HTKKFile.RemoveValidate(sView.FpSpread.Sheets[0], GeneralFunctions.ArrayCellId("L_34")[1], GeneralFunctions.ArrayCellId("L_34")[0]);
			}
			else if (sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_34")[1], GeneralFunctions.ArrayCellId("L_34")[0]].BackColor != Color.Pink)
			{
				HTKKFile.SetNodeBackGroundYellow(sView.FpSpread.Sheets[0], GeneralFunctions.ArrayCellId("L_34")[1], GeneralFunctions.ArrayCellId("L_34")[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_104")));
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

	private void checkPL1_1TD(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(4, out value);
			if (!sView.FpSpread.Sheets[4].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[4];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = hTKKCMCTableInfo.RowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int num3 = hTKKCMCTableInfo.RowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int num4 = 0;
			string[] array = new string[5];
			string defaultValue = hTKKCMCTableInfo.RowsInfo.Cells[4].DefaultValue;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, odrColC - 1].Text;
				array[1] = sView.Cells[num + i, num2].Text;
				array[2] = sView.Cells[num + i, odrColF - 1].Text;
				array[3] = sView.Cells[num + i, odrColH - 1].Text;
				array[4] = defaultValue;
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
					HTKKFile.RemoveValidate(sView, num + i, num2);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, odrColC - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, odrColC - 1);
				}
				if (!string.IsNullOrEmpty(array[1]))
				{
					if (!ValidateBase.CheckMaSoThue(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_82")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, num + i, num2);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
				}
				num4++;
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[4].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[4].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[4].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void checkPL1_2TD(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(5, out value);
			if (!sView.FpSpread.Sheets[5].Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[5];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			double num = ConvertHelpers.ConvertToDouble(sView.FpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("L_45")[1], GeneralFunctions.ArrayCellId("L_45")[0]].Value);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_2_TD_GTGT/tongCong"));
			if (tables.Count > 0)
			{
				HTKKCMCStaticOfTable header = tables[0].Header;
				string[] array = new string[8];
				string defaultValue = tables[0].Header.StaticCells[0].DefaultValue;
				string defaultValue2 = tables[0].Header.StaticCells[3].DefaultValue;
				string defaultValue3 = tables[0].Header.StaticCells[6].DefaultValue;
				string defaultValue4 = tables[0].Header.StaticCells[7].DefaultValue;
				array[0] = defaultValue;
				array[3] = defaultValue2;
				array[6] = defaultValue3;
				array[7] = defaultValue4;
				string[] array2 = new string[8];
				string defaultValue5 = tables[0].RowsInfo.Cells[0].DefaultValue;
				string defaultValue6 = tables[0].RowsInfo.Cells[6].DefaultValue;
				string defaultValue7 = tables[0].RowsInfo.Cells[7].DefaultValue;
				array2[0] = defaultValue5;
				array2[6] = defaultValue6;
				array2[7] = defaultValue7;
				HTKKCMCStaticOfTable hTKKCMCStaticOfTable = new HTKKCMCStaticOfTable();
				HTKKCMCRowInfo hTKKCMCRowInfo = new HTKKCMCRowInfo();
				int num2 = 0;
				for (int num3 = 0; num3 < tables.Count; num3++)
				{
					hTKKCMCStaticOfTable = tables[num3].Header;
					hTKKCMCRowInfo = tables[num3].RowsInfo;
					num2 = tables[num3].Location.DataLocationTemplateRowStart - 1;
					array[1] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1].Value);
					array[2] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1].Value);
					array[4] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.ColIndex - 1].Value);
					array[5] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1].Value);
					bool flag = GeneralFunctions.HaveToValidateHeaderRow(header, array);
					bool flag2 = false;
					double num4 = 0.0;
					for (int num5 = 0; num5 < tables[num3].RowsValues.Count; num5++)
					{
						array2[1] = ConvertHelpers.ConvertToString(sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Value);
						array2[2] = ConvertHelpers.ConvertToString(sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1].Value);
						array2[3] = ConvertHelpers.ConvertToString(sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1].Value);
						array2[4] = ConvertHelpers.ConvertToString(sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1].Value);
						array2[5] = ConvertHelpers.ConvertToString(sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1].Value);
						array2[6] = ConvertHelpers.ConvertToString(sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[6].CellTemplate.ColIndex - 1].Value);
						bool visible = sView.Cells[num2 + num5, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Row.Visible;
						if (GeneralFunctions.HaveToValidateRow(hTKKCMCRowInfo, array2) && visible)
						{
							flag2 = true;
						}
						else
						{
							if (!flag)
							{
								HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidateBackGround(sView, num2 + num5, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
								continue;
							}
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidateBackGround(sView, num2 + num5, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
						if (!string.IsNullOrEmpty(array2[2]))
						{
							if (!ValidateBase.CheckMaSoThue(array2[2]))
							{
								HTKKFile.SetNodeBackGroundPink(sView, num2 + num5, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_82")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
						}
						if (string.IsNullOrEmpty(array2[3]))
						{
							HTKKFile.SetNodeBackGroundPink(sView, num2 + num5, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
						}
						num4 += ConvertHelpers.ConvertToDouble(array2[4]);
						if (ConvertHelpers.ConvertToDouble(array2[5]) == 0.0 && num != 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num2 + num5, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidateBackGround(sView, num2 + num5, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
					}
					if (!flag2 && !flag)
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
						continue;
					}
					if (num4 != 100.0)
					{
						for (int num5 = 0; num5 < tables[num3].RowsValues.Count; num5++)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num2 + num5, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_94")));
							value = 2;
						}
					}
					else
					{
						for (int num5 = 0; num5 < tables[num3].RowsValues.Count; num5++)
						{
							HTKKFile.RemoveValidate(sView, num2 + num5, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
						}
					}
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
					}
					if (!string.IsNullOrEmpty(array[2]))
					{
						if (!ValidateBase.CheckMaSoThue(array[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_82")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
					}
					if (ConvertHelpers.ConvertToDouble(array[5]) == 0.0 && num != 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_81")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
					}
				}
			}
			if (num == 0.0 && ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) != 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_99")));
				value = 2;
			}
			if (PeriodInfo.TieuMucHachToanVal == "06" && ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) != num)
			{
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_113")));
				value = 2;
			}
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				string maCQTCapCuc = headerInfo.maCQTCapCuc;
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				int colCheck = HTKKFile.ColumnLettersToIndex("AH");
				int num6 = HTKKFile.ColumnLettersToIndex("P");
				string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_115"));
				int rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare2(sView, dataLocationTemplateRowStart, colCheck, maCQTCapCuc);
				if (rowIndexInSViewByCompare < 0)
				{
					for (int num7 = 0; num7 < table.RowsValues.Count; num7++)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + num7 - 1, num6 - 1, message);
					}
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[5].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[5].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[5].BackColor = Color.White;
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

	public void Refresh_1(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("M_54");
			if (actionType == 1)
			{
				string formula = "ROUND('" + winFpSpread.Sheets[1].SheetName + "'!I9,0)";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
				double num = (double)winFpSpread.Sheets[0].Cells[array[1], array[0]].Value;
				winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].Formula = "M54";
			}
			else
			{
				winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].Value = 0;
				if (PeriodInfo.LoaiToKhai == "B")
				{
					winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].Locked = true;
				}
			}
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
			winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].ResetFormula();
			winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].Value = 0;
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
			CheckPhuLucAnHien(winFpSpread, ToKhaiInfo);
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
			CheckPhuLucAnHien(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_5(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			CheckPhuLucAnHien(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_5(FpSpread winFpSpread, int actionType, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			CheckPhuLucAnHien(winFpSpread, ToKhaiInfo);
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
			if ((activeColumnIndex == CellId_ckbGiaHan[0]) & (activeRowIndex == CellId_ckbGiaHan[1]))
			{
				if ((bool)winFpSpread.Sheets[0].ActiveCell.Value)
				{
					winFpSpread.ActiveSheet.Cells[CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]].Locked = false;
					return;
				}
				winFpSpread.ActiveSheet.Cells[CellId_maGiaHan[1], CellId_maGiaHan[0]].Value = null;
				winFpSpread.ActiveSheet.Cells[CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]].ResetValue();
				winFpSpread.ActiveSheet.Cells[CellId_cbbGiaHan[1], CellId_cbbGiaHan[0]].Locked = true;
			}
			else if ((activeColumnIndex == CellId_ct21[0]) & (activeRowIndex == CellId_ct21[1]))
			{
				Obj = XuLyCheckboxCt21(winFpSpread, isBindData: false, ToKhaiInfo, TTinChung.mst, DicSheet);
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
			XuLyChiTieu40(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], isBindData: false);
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				if (((iRow - 1 == CellId_ct22[1]) & (iCol - 1 == CellId_ct22[0])) && winFpSpread.Sheets[0].Cells[CellId_ct22[1], CellId_ct22[0]].Text == "")
				{
					winFpSpread.Sheets[0].Cells[CellId_ct22[1], CellId_ct22[0]].Value = 0;
				}
				if ((iRow - 1 == CellId_ct32[1]) & (iCol - 1 == CellId_ct32[0]))
				{
					winFpSpread.Sheets[0].Cells[CellId_ct33[1], CellId_ct33[0]].Formula = "ROUND(J34*10/100,0)";
				}
				if ((iRow - 1 == CellId_ct30[1]) & (iCol - 1 == CellId_ct30[0]))
				{
					winFpSpread.Sheets[0].Cells[CellId_ct31[1], CellId_ct31[0]].Formula = "ROUND(J33*5/100,0)";
				}
				TinhLaiCt06_PL_01_6(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
				SetLaiChiTieu07_PL01_6(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
				if (ToKhaiInfo.SheetInfos[2] != null)
				{
					SetLaiCt15_17_PL01_6(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2]);
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
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart;
			if ((winFpSpread.ActiveSheetIndex == 1 && iRow >= dataLocationTemplateRowStart) & (iCol == odrColI))
			{
				XuLyChiTieu40(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], isBindData: false);
				winFpSpread.Sheets[0].Cells[CellId_ct39[1], CellId_ct39[0]].Formula = strCellId_01_5_tong6;
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				SetLaiCt15_17_PL01_6(winFpSpread.Sheets[2], ToKhaiInfo.SheetInfos[2]);
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
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = GeneralFunctions.DetectIndexOfTable(hTKKCMCSheetInfo, (int)iRow);
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("H");
			if (iCol == num || iCol == num2)
			{
				LayThemThongTinTinhHuyen_PL01_6(sView, (int)iRow);
			}
			if (num3 >= 0 && (iCol == num4 || iCol == num5))
			{
				SetLaiCt15_17_PL01_6(sView, hTKKCMCSheetInfo);
			}
			string pathCt07 = "PL01_6_GTGT/ct07";
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt07);
			if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
			{
				SetLaiCt15_17_PL01_6(sView, hTKKCMCSheetInfo);
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
			SheetView sView = winFpSpread.Sheets[3];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			if (iCol == num || iCol == num2)
			{
				LayThemThongTinTinhHuyen_PL_01_7(sView, (int)iRow);
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
			if (iCol == HTKKFile.ColumnLettersToIndex("U"))
			{
				double num = 0.0;
				bool flag = false;
				int index = -1;
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					if (iRow == hTKKCMCTableInfo.Header.StaticCells[5].CellTemplate.RowIndex)
					{
						index = i;
						flag = true;
					}
					num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
				}
				if (flag)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_2_TD_GTGT/tongCong"));
					winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[index];
					int num2 = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[index].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[index].RowsValues.Count;
					for (int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num3 < num2; num3++)
					{
						double num4 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num5 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num5 * (num4 / 100.0);
					}
				}
			}
			else if (iCol == HTKKFile.ColumnLettersToIndex("R"))
			{
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					double num6 = 0.0;
					int index = -1;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					int num2 = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].RowsValues.Count;
					for (int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num3 < num2; num3++)
					{
						double num4 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num5 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						num6 += num4;
						if (iRow == num3)
						{
							index = i;
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num5 * (num4 / 100.0);
						}
					}
					if (index > -1)
					{
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[index].Header.StaticCells[4].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Value = num6;
					}
				}
			}
			int num7 = HTKKFile.ColumnLettersToIndex("P");
			int num8 = HTKKFile.ColumnLettersToIndex("AJ");
			if (iCol == num7 || iCol == num8)
			{
				LayThemThongTinTinhHuyen_PL_01_2_TD(winFpSpread.Sheets[5], (int)iRow);
				LayThemThongTinCQTChiCuc_PL_01_2_TD(winFpSpread.Sheets[5], (int)iRow);
			}
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[5];
			SheetView sView = winFpSpread.Sheets[5];
			if (iCol == HTKKFile.ColumnLettersToIndex("C") || iCol == HTKKFile.ColumnLettersToIndex("H") || iCol == HTKKFile.ColumnLettersToIndex("P"))
			{
				int tableIndex = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
				GeneralFunctions.SetLaiDongCon(sView, sInfo, tableIndex, -1, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
				GeneralFunctions.SetLaiDongCon(sView, sInfo, tableIndex, -1, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			double num = 0.0;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
				num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_2_TD_GTGT/tongCong"));
			winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_7(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string GetFormulaTienPNC(int countRowDyn, double soTien, string strColRow, int soNgayNC, HTKKPeriodInfo periodInfor)
	{
		try
		{
			string text = "";
			string text2 = "~08~44~21~40~42~98~99~93~92~89~";
			DateTime hanNopToKhai = GeneralFunctions.GetHanNopToKhai(periodInfor);
			if (!periodInfor.NgayLapKHBS.Equals(""))
			{
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(periodInfor.NgayLapKHBS, "d/M/yyyy");
			}
			DateTime t = new DateTime(2013, 7, 1);
			DateTime dateTime2 = new DateTime(2013, 9, 28);
			DateTime dateTime3 = new DateTime(2014, 12, 31);
			DateTime dateTime4 = new DateTime(2016, 6, 30);
			if (DateTime.Compare(hanNopToKhai, t) < 0)
			{
				int days = (dateTime2 - hanNopToKhai).Days;
				int days2 = (dateTime3 - dateTime2).Days;
				int days3 = (dateTime4 - dateTime3).Days;
				text = ((soNgayNC > GetSoNgay(hanNopToKhai, dateTime4)) ? (days + "*" + strColRow + "*0.05/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + days3 + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days - days2 - days3) + "*" + strColRow + "*0.03/100") : ((soNgayNC <= GetSoNgay(hanNopToKhai, dateTime4) && soNgayNC > GetSoNgay(hanNopToKhai, dateTime3)) ? (days + "*" + strColRow + "*0.05/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + (soNgayNC - days - days2) + "*" + strColRow + "*0.05/100") : ((soNgayNC > GetSoNgay(hanNopToKhai, dateTime3) || soNgayNC <= GetSoNgay(hanNopToKhai, dateTime2)) ? (soNgayNC + "*" + strColRow + "*0.05/100") : (days + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days) + "*" + strColRow + "*0.07/100"))));
			}
			else if (DateTime.Compare(hanNopToKhai, t) >= 0 && DateTime.Compare(hanNopToKhai, dateTime3) <= 0)
			{
				int days;
				int days2;
				if (DateTime.Compare(hanNopToKhai, dateTime2) > 0)
				{
					days = 0;
					days2 = GetSoNgay(hanNopToKhai, dateTime3);
				}
				else
				{
					days = GetSoNgay(hanNopToKhai, dateTime2) + 1;
					days2 = GetSoNgay(dateTime2, dateTime3);
				}
				int days3 = GetSoNgay(dateTime3, dateTime4);
				text = ((soNgayNC > GetSoNgay(hanNopToKhai, dateTime4)) ? (days + "*" + strColRow + "*0.07/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + days3 + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days - days2 - days3) + "*" + strColRow + "*0.03/100") : ((soNgayNC <= GetSoNgay(hanNopToKhai, dateTime4) && soNgayNC > GetSoNgay(hanNopToKhai, dateTime3)) ? (days + "*" + strColRow + "*0.07/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + (soNgayNC - days - days2) + "*" + strColRow + "*0.05/100") : ((soNgayNC > GetSoNgay(hanNopToKhai, dateTime3) || soNgayNC <= GetSoNgay(hanNopToKhai, dateTime2)) ? (soNgayNC + "*" + strColRow + "*0.05/100") : ((soNgayNC > 90) ? (days + "*" + strColRow + "*0.07/100 + " + (soNgayNC - days) + "*" + strColRow + "*0.07/100") : (days + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days) + "*" + strColRow + "*0.05/100")))));
			}
			else if (DateTime.Compare(hanNopToKhai, dateTime3) > 0 && DateTime.Compare(hanNopToKhai, dateTime4) <= 0)
			{
				int days3 = GetSoNgay(hanNopToKhai, dateTime4);
				text = ((soNgayNC <= GetSoNgay(hanNopToKhai, dateTime4)) ? (soNgayNC + "*" + strColRow + "*0.05/100") : (days3 + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days3) + "*" + strColRow + "*0.03/100"));
			}
			else
			{
				text = soNgayNC + "*" + strColRow + "*0.03/100";
			}
			if (text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				return "IF(" + text + ">0,ROUND(" + text + ",2),0)";
			}
			return "IF(" + text + ">0,ROUND(" + text + ",0),0)";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static int GetSoNgay(DateTime fromDate, DateTime toDate)
	{
		try
		{
			return (toDate - fromDate).Days;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	private bool XuLyCheckboxCt21(FpSpread winFpSpread, bool isBindData, HTKKCMCToKhaiInfo ToKhaiInfo, string mstNNT, Dictionary<string, int> DicSheet)
	{
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Invalid comparison between Unknown and I4
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("I") - 1;
			int columnIndex = HTKKFile.ColumnLettersToIndex("J") - 1;
			int columnIndex2 = HTKKFile.ColumnLettersToIndex("L") - 1;
			bool flag = (bool)winFpSpread.Sheets[0].Cells[CellId_ct21[1], CellId_ct21[0]].Value;
			int rowIndex = 26;
			int rowIndex2 = 27;
			int rowIndex3 = 29;
			int rowIndex4 = 31;
			int rowIndex5 = 32;
			int rowIndex6 = 33;
			int rowIndex7 = 34;
			winFpSpread.Sheets[0].Protect = true;
			if (bool.Parse(flag.ToString()))
			{
				DialogResult val = (DialogResult)6;
				if (!isBindData)
				{
					val = MessageBox.Show("Các số liệu đã có trong các chỉ tiêu phát sinh trong kỳ sẽ bị xóa bằng 0. Bạn có đồng ý không?", "Thong bao", (MessageBoxButtons)4);
				}
				if ((int)val == 6)
				{
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex, columnIndex);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex, columnIndex, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex, columnIndex2);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex, columnIndex2, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex2, columnIndex2);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex2, columnIndex2, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex3, columnIndex);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex3, columnIndex, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex4, columnIndex);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex4, columnIndex, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex5, columnIndex);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex5, columnIndex, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex5, columnIndex2);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex5, columnIndex2, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex6, columnIndex);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex6, columnIndex, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex6, columnIndex2);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex6, columnIndex2, valueChange: true);
					ResetValueCell_Zero(winFpSpread.Sheets[0], rowIndex7, columnIndex);
					ChangeLockedCell(winFpSpread.Sheets[0], rowIndex7, columnIndex, valueChange: true);
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
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex, columnIndex, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex, columnIndex2, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex2, columnIndex2, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex3, columnIndex, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex4, columnIndex, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex5, columnIndex, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex5, columnIndex2, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex6, columnIndex, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex6, columnIndex2, valueChange: false);
				ChangeLockedCell(winFpSpread.Sheets[0], rowIndex7, columnIndex, valueChange: false);
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
			int row = 44;
			int row2 = 46;
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

	public void ResetValueCell_Zero(SheetView sView, int rowIndex, int columnIndex)
	{
		try
		{
			sView.Cells[rowIndex, columnIndex].Value = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ChangeLockedCell(SheetView sView, int rowIndex, int columnIndex, bool valueChange)
	{
		try
		{
			sView.Cells[rowIndex, columnIndex].Locked = valueChange;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
								string valueInXml = GeneralFunctions.GetValueInXml(text + "\\" + fileInfo.Name, "L_48");
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
			if (array.Length >= 3 && maNNKD == array[1])
			{
				result = array[2];
				break;
			}
		}
		return result;
	}

	public void LayThemThongTinTinhHuyen_PL01_6(SheetView sView, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("M");
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int colTenTinh = HTKKFile.ColumnLettersToIndex("R");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("S");
			int colTenHuyen = HTKKFile.ColumnLettersToIndex("T");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("U");
			string text = sView.Cells[iRow - 1, num - 1].Text;
			string text2 = sView.Cells[iRow - 1, num2 - 1].Text;
			if (!string.IsNullOrEmpty(text) && text != "0")
			{
				text = text.Replace(".", "");
				GeneralFunctions.ThayDoiTinhTheoCQTDangBang2(text, text2, sView, colTenTinh, colMaTinh, colTenHuyen, colMaHuyen, iRow, awaySet: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LayThemThongTinTinhHuyen_PL_01_7(SheetView sView, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("J");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int colTenTinh = HTKKFile.ColumnLettersToIndex("N");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("O");
			int colTenHuyen = HTKKFile.ColumnLettersToIndex("P");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("Q");
			string text = sView.Cells[iRow - 1, num - 1].Text;
			string text2 = sView.Cells[iRow - 1, num2 - 1].Text;
			if (!string.IsNullOrEmpty(text))
			{
				GeneralFunctions.ThayDoiTinhTheoCQTDangBang2(text, text2, sView, colTenTinh, colMaTinh, colTenHuyen, colMaHuyen, iRow, awaySet: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LayThemThongTinTinhHuyen_PL_01_2_TD(SheetView sView, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AH");
			int num2 = HTKKFile.ColumnLettersToIndex("AK");
			int colTenTinh = HTKKFile.ColumnLettersToIndex("AN");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("AO");
			int colTenHuyen = HTKKFile.ColumnLettersToIndex("AP");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("AQ");
			string text = sView.Cells[iRow - 1, num - 1].Text;
			string text2 = sView.Cells[iRow - 1, num2 - 1].Text;
			if (!string.IsNullOrEmpty(text) && text != "0")
			{
				text = text.Replace(".", "");
				GeneralFunctions.ThayDoiTinhTheoCQTDangBang2(text, text2, sView, colTenTinh, colMaTinh, colTenHuyen, colMaHuyen, iRow, awaySet: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LayThemThongTinCQTChiCuc_PL_01_2_TD(SheetView sView, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AH");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AK");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("AJ");
			string text = sView.Cells[iRow - 1, num - 1].Text;
			if (!string.IsNullOrEmpty(text) && text != "0")
			{
				text = text.Replace(".", "");
				GeneralFunctions.GetCQTChiCucTuCQTCuc(text, sView, colTenCQTChiCuc, colMaCQTChiCuc, iRow);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetLaiChiTieu07_PL01_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo != null)
			{
				int[] array = GeneralFunctions.ArrayCellId("J_31");
				int[] array2 = GeneralFunctions.ArrayCellId("J_32");
				int[] array3 = GeneralFunctions.ArrayCellId("J_35");
				string pathCt07 = "PL01_6_GTGT/ct07";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt07);
				if (iCol == array[0] + 1 && iRow >= array2[1] + 1 && iRow <= array3[1] + 1)
				{
					sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = sheetView.Cells[array[1], array[0]].Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetLaiCt15_17_PL01_6(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (sInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("J");
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("V");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			string pathCt06 = "PL01_6_GTGT/ct06";
			int[] array = GeneralFunctions.ArrayCellId("V_5");
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt06);
			double num5 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			double num6 = Convert.ToDouble(sView.Cells[array[1], array[0]].Value);
			if (num5 >= num6)
			{
				for (int num7 = 0; num7 < hTKKCMCTableInfo.RowsValues.Count; num7++)
				{
					sView.Cells[dataLocationTemplateRowStart + num7 - 1, num - 1].Value = sView.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Value;
					sView.Cells[dataLocationTemplateRowStart + num7 - 1, num2 - 1].Value = 0;
				}
			}
			else
			{
				for (int num7 = 0; num7 < hTKKCMCTableInfo.RowsValues.Count; num7++)
				{
					sView.Cells[dataLocationTemplateRowStart + num7 - 1, num - 1].Value = 0;
					sView.Cells[dataLocationTemplateRowStart + num7 - 1, num2 - 1].Value = sView.Cells[dataLocationTemplateRowStart + num7 - 1, num4 - 1].Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiCt06_PL_01_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRowTK, int iColTK)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo != null)
			{
				int[] array = GeneralFunctions.ArrayCellId("L_28");
				int[] array2 = GeneralFunctions.ArrayCellId("L_44");
				string pathCt06_PL01_6 = "PL01_6_GTGT/ct06";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCt06_PL01_6);
				if (iColTK == array[0] + 1 && iRowTK >= array[1] + 1 && iRowTK <= array2[1] + 1)
				{
					SetLaiCt15_17_PL01_6(sView, hTKKCMCSheetInfo);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckPhuLucAnHien(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		sheetView.Cells[cellPL_01_1TD[1], cellPL_01_1TD[0]].Value = ((ToKhaiInfo.SheetInfos[4] != null) ? 1 : 0);
		sheetView.Cells[cellPL_01_2TD[1], cellPL_01_2TD[0]].Value = ((ToKhaiInfo.SheetInfos[5] != null) ? 1 : 0);
	}
}
