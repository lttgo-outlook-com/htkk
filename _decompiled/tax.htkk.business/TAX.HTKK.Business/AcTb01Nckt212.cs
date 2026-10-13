using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
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

public class AcTb01Nckt212
{
	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

	private List<string> listHD = new List<string>();

	private string pathNgayThongBao = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayBCao";

	private string pathTenDvCq = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DonViChuQuan/ten";

	private string pathMstDvCq = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DonViChuQuan/mst";

	private string pathNguoiKy = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private string pathCQT = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/tenCQTTiepNhan";

	private int idxCol_MaLoaiHD => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_TenLoaiHD => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_MauSoHD => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_KyHieuHD => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_SoLuong => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("R") - 1;

	private int idxCol_NgayBDSD => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_TenDN => HTKKFile.ColumnLettersToIndex("W") - 1;

	private int idxCol_MSTDN => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_HDSo => HTKKFile.ColumnLettersToIndex("AD") - 1;

	private int idxCol_HDNgay => HTKKFile.ColumnLettersToIndex("AE") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayThongBao));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCQT));
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].ResetFormula();
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = PeriodInfo.kyKKhai;
				sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = TTinChung.tenCQTNoiNop;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 2 && idxCol == 1)
			{
				ShowBangKe("Bangke_01TBAC.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow < dataLocationTemplateRowStart || iRow > num || iCol - 1 != idxCol_MaLoaiHD)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, idxCol_MaLoaiHD].Text;
			if (HTKKFormater.isNullString(text))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].ResetValue();
				sheetView.Cells[(int)iRow - 1, idxCol_TenLoaiHD].ResetValue();
				return;
			}
			sheetView.Cells[(int)iRow - 1, idxCol_TenLoaiHD].Text = dicDMHoaDon[text][2];
			if (text.Equals("TT120"))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].ResetValue();
			}
			else
			{
				sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow >= dataLocationTemplateRowStart && iRow <= num && iCol - 1 == idxCol_MauSoHD)
			{
				string text = sheetView.Cells[(int)iRow - 1, idxCol_MaLoaiHD].Text;
				string text2 = "";
				if (!string.IsNullOrEmpty(text))
				{
					text2 = dicDMHoaDon[text][0];
				}
				if (text2.Equals("0") || text2.Equals("4"))
				{
					TextCellType textCellType = new TextCellType();
					textCellType.MaxLength = 13;
					TextCellType cellType = textCellType;
					sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].CellType = cellType;
				}
				else
				{
					TextCellType textCellType2 = new TextCellType();
					textCellType2.MaxLength = 20;
					TextCellType cellType = textCellType2;
					sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].CellType = cellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			string text = sheetView.Cells[(int)iRow - 1, num3 - 1].Text;
			string text2 = "";
			string text3 = "";
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_TuSo || iCol - 1 == idxCol_DenSo)
			{
				text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text2);
				string text4 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
				if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text5))
				{
					double num4 = ConvertHelpers.ConvertToDouble(text4);
					double num5 = ConvertHelpers.ConvertToDouble(text5);
					if (num4 == 0.0 && num5 == 0.0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Text = "0";
					}
					else if (num4 == 0.0 && num5 != 0.0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Text = (num5 - num4).ToString();
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Text = (num5 - num4 + 1.0).ToString();
					}
				}
			}
			if (iCol - 1 == idxCol_MauSoHD)
			{
				if (!string.IsNullOrEmpty(text))
				{
					text3 = dicDMHoaDon[text][0];
				}
				if (!text3.Equals("3"))
				{
					string text6 = sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].Text = text6.ToUpper();
				}
			}
			if (iCol - 1 == idxCol_KyHieuHD)
			{
				string text7 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuHD].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuHD].Text = GeneralFunctions.FormatKyHieu(text7);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LeaveCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, long iNewRow, long iNewCol)
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

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num && idxRow <= num2 && (idxCol == idxCol_TuSo || idxCol == idxCol_DenSo))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			checkTKChinh(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void checkTKChinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[25];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayThongBao));
			string obj = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTenDvCq));
			string text = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathMstDvCq));
			string text2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiKy));
			string value2 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, num, "P", "R", "B", "D", "J", "L", "", isMST: false, 1);
			for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
			{
				array[0] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_MaLoaiHD);
				array[1] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_TenLoaiHD);
				array[2] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_MauSoHD);
				array[3] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_KyHieuHD);
				array[4] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_SoLuong);
				array[5] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_TuSo);
				array[6] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_DenSo);
				array[7] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_NgayBDSD);
				array[8] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_TenDN);
				array[9] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_MSTDN);
				array[10] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_HDSo);
				array[11] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_HDNgay);
				string strLoai = "";
				if (!string.IsNullOrEmpty(array[0]))
				{
					strLoai = dicDMHoaDon[array[0]][0];
				}
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MaLoaiHD);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TenLoaiHD);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MaLoaiHD);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MauSoHD);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_KyHieuHD);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TuSo);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_DenSo);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_NgayBDSD);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TenDN);
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MSTDN);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MaLoaiHD);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TenLoaiHD);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MauSoHD);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_KyHieuHD);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TuSo);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_DenSo);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TenDN);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MSTDN);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_HDSo);
				GeneralFunctions.strMauSoHD_01GTKT = array[0];
				if (HTKKFormater.isNullString(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MaLoaiHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				if (HTKKFormater.isNullString(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TenLoaiHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				if (HTKKFormater.isNullString(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				else
				{
					string text3 = GeneralFunctions.CheckMauSoHD(array[2], strLoai, array[0]);
					if (text3.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					else if (text3.Equals("2"))
					{
						if (array[3].Length > 0)
						{
							if (array[3].Substring(array[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienHDDT(array[2], strLoai, array[0], array[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
									value = 1;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
								value = 2;
							}
						}
					}
					else if (array[3].Length > 0 && array[3].Substring(array[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(array[2], strLoai, array[0], array[3]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
						value = 1;
					}
				}
				if (HTKKFormater.isNullString(array[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_KyHieuHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				else
				{
					string text4 = GeneralFunctions.CheckSoHD(array[3], strLoai);
					if (text4.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_KyHieuHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						value = 2;
					}
					else if (text4.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_KyHieuHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(array[1]))
				{
					if (string.IsNullOrEmpty(array[7]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_NgayBDSD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
						value = 2;
					}
					else
					{
						DateTime dateTime = ConvertHelpers.ConvertToDateTime(array[7], "dd/MM/yyyy");
						DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(obj, "dd/MM/yyyy");
						if ((dateTime - dateTime2).TotalDays < 2.0 && dateTime != DateTime.MinValue)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_NgayBDSD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
							value = 2;
						}
						else if (dateTime != DateTime.MinValue)
						{
							HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_HDNgay);
						}
					}
				}
				if (!string.IsNullOrEmpty(array[3]) && array[3][array[3].Length - 1].ToString().Equals("P"))
				{
					if (string.IsNullOrEmpty(array[8]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TenDN, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
						value = 2;
					}
					if (string.IsNullOrEmpty(array[9]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MSTDN, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThue(array[9]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MSTDN, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
						value = 2;
					}
				}
				else if (!string.IsNullOrEmpty(array[9]) && !ValidateBase.CheckMaSoThue(array[9]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MSTDN, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (HTKKFormater.isNullString(array[5]) && HTKKFormater.isNullString(array[6]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
					continue;
				}
				if (HTKKFormater.isNullString(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_DenSo);
					value = 2;
					continue;
				}
				if (HTKKFormater.isNullString(array[6]))
				{
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TuSo);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
					continue;
				}
				double num3 = ConvertHelpers.ConvertToDouble(array[5]);
				double num4 = ConvertHelpers.ConvertToDouble(array[6]);
				if (num3 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
					value = 2;
				}
				else
				{
					if (array2 == null || array2.Length <= 1)
					{
						continue;
					}
					for (int num5 = 0; num5 < array2.Length; num5++)
					{
						if (num + num2 == ConvertHelpers.ConvertToInt(array2[num5]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21")));
							HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21")));
							value = 2;
							break;
						}
					}
				}
			}
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
			}
			string mst = TTinChung.mst;
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
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMHoaDon(HTKKPeriodInfo PeriodInfo, string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0073: Expected O, but got Unknown
		try
		{
			if (dicDMHoaDonTmp.Count > 0)
			{
				return dicDMHoaDonTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			listHD = new List<string>();
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string text = array[1];
				if (!(DateTime.ParseExact(PeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.CurrentCulture) > new DateTime(2014, 6, 1)) || !text.Equals("06HDXK"))
				{
					dictionary2.Add(key, text);
					dictionary.Add(key, array);
					listHD.Add(value);
				}
			}
			dicDMHoaDon = dictionary;
			dicDMHoaDonTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num3 - 1].Text;
				bool flag = true;
				for (int j = 0; j < listHD.Count; j++)
				{
					string input = listHD[j];
					string[] array = Regex.Split(input, "###");
					if (text.IndexOf(array[1], 0) == 0)
					{
						sheetView.Cells[num + i, num2 - 1].Text = array[1];
						flag = false;
						break;
					}
				}
				if (flag && !string.IsNullOrEmpty(text))
				{
					sheetView.Cells[num + i, num2 - 1].Text = "TT120";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
