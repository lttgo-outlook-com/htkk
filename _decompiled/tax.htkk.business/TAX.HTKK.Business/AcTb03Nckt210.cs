using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class AcTb03Nckt210
{
	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

	private List<string> listHD = new List<string>();

	private string cellId_NgayMat = "O_8";

	private string cellId_GioMat = "L_8";

	private string cellId_PhutMat = "N_8";

	private string pathPhuongPhapHuy = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/phuongPhapHuy";

	private string pathNguoiLapBieu = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiLapBieu";

	private string pathNgayThongBao = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayBCao";

	private string pathNguoiDaiDien = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private int[] CellId_NgayMat => GeneralFunctions.ArrayCellId(cellId_NgayMat);

	private int[] CellId_GioMat => GeneralFunctions.ArrayCellId(cellId_GioMat);

	private int[] CellId_PhutMat => GeneralFunctions.ArrayCellId(cellId_PhutMat);

	private int idxCol_MaLoaiHD => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_TenLoaiHD => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_MauSoHD => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_KyHieuHD => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_SoLuong => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("N");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				sView.Cells[1, num4 - 1].Text = TTinChung.tenCQTNoiNop;
			}
			int num5 = 8;
			if (HTKKFormater.isNullString(sView.Cells[num5 - 1, num - 1].Text) && HTKKFormater.isNullString(sView.Cells[num5 - 1, num2 - 1].Text) && HTKKFormater.isNullString(sView.Cells[num5 - 1, num3 - 1].Text))
			{
				sView.Cells[num5 - 1, num - 1].Text = DateTime.Now.Hour.ToString();
				sView.Cells[num5 - 1, num2 - 1].Text = DateTime.Now.Minute.ToString();
				sView.Cells[num5 - 1, num3 - 1].Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
				int num6 = HTKKFile.ColumnLettersToIndex("L");
				int num7 = HTKKFile.ColumnLettersToIndex("N");
				int num8 = HTKKFile.ColumnLettersToIndex("O");
				string text = sView.Cells[num5 - 1, num8 - 1].Text;
				string text2 = sView.Cells[num5 - 1, num6 - 1].Text;
				string text3 = sView.Cells[num5 - 1, num7 - 1].Text;
				string strInput = "00";
				if (text.IndexOf("/") > 0 && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
				{
					string[] array = text.Split(new char[1] { '/' });
					sView.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text2, 2) + ":" + GeneralFunctions.GetStrByLengh(text3, 2) + ":" + GeneralFunctions.GetStrByLengh(strInput, 2) + "Z";
				}
				else
				{
					sView.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = "";
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
			int num3 = HTKKFile.ColumnLettersToIndex("L");
			int num4 = HTKKFile.ColumnLettersToIndex("N");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = 8;
			if (iRow == 8 && (iCol == num3 || iCol == num4))
			{
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.GetStrByLengh(sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text, 2);
			}
			if (iRow == 8 && (iCol == num3 || iCol == num4 || iCol == num5))
			{
				string text = sheetView.Cells[num6 - 1, num5 - 1].Text;
				string text2 = sheetView.Cells[num6 - 1, num3 - 1].Text;
				string text3 = sheetView.Cells[num6 - 1, num4 - 1].Text;
				string strInput = "00";
				if (text.IndexOf("/") > 0 && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
				{
					string[] array = text.Split(new char[1] { '/' });
					string text4 = text2;
					sheetView.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text2, 2) + ":" + GeneralFunctions.GetStrByLengh(text3, 2) + ":" + GeneralFunctions.GetStrByLengh(strInput, 2) + "Z";
				}
				else
				{
					sheetView.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = "";
				}
			}
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_MauSoHD)
			{
				string text5 = sheetView.Cells[(int)iRow - 1, idxCol_MaLoaiHD].Text;
				string text6 = "";
				if (!string.IsNullOrEmpty(text5))
				{
					text6 = dicDMHoaDon[text5][0];
				}
				if (!text6.Equals("3"))
				{
					string text7 = sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].Text = text7.ToUpper();
				}
			}
			else if (iCol - 1 == idxCol_KyHieuHD)
			{
				string text8 = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuHD].Text;
				sheetView.Cells[(int)iRow - 1, idxCol_KyHieuHD].Text = GeneralFunctions.FormatKyHieu(text8);
			}
			else if (iCol - 1 == idxCol_TuSo)
			{
				string text9 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
				if (!HTKKFormater.isNullString(text9))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text = GeneralFunctions.FormatSoHD(text9);
				}
			}
			else if (iCol - 1 == idxCol_DenSo)
			{
				string text9 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
				if (!HTKKFormater.isNullString(text9))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text = GeneralFunctions.FormatSoHD(text9);
				}
			}
			if (iCol - 1 != idxCol_TuSo && iCol - 1 != idxCol_DenSo)
			{
				return;
			}
			string text10 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
			string text11 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
			if (!string.IsNullOrEmpty(text10) && !string.IsNullOrEmpty(text11))
			{
				int num7 = ConvertHelpers.ConvertToInt(text10);
				int num8 = ConvertHelpers.ConvertToInt(text11);
				if (num7 == 0 && num8 == 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = "0";
				}
				else if (num7 == 0 && num8 != 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = (num8 - num7).ToString();
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = (num8 - num7 + 1).ToString();
				}
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
			string text2 = "";
			if (!string.IsNullOrEmpty(text))
			{
				text2 = dicDMHoaDon[text][0];
			}
			if (text2.Equals("0") || text2.Equals("4"))
			{
				TextCellType textCellType = new TextCellType();
				textCellType.WordWrap = true;
				TextCellType textCellType2 = textCellType;
				textCellType2.MaxLength = 13;
				sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].CellType = textCellType2;
			}
			else
			{
				TextCellType textCellType3 = new TextCellType();
				textCellType3.WordWrap = true;
				TextCellType textCellType2 = textCellType3;
				textCellType2.MaxLength = 20;
				sheetView.Cells[(int)iRow - 1, idxCol_MauSoHD].CellType = textCellType2;
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
			int iwarningColor = ((levelWarning.Count > 0) ? levelWarning[0] : 0);
			checkTKChinh(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iwarningColor);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void checkTKChinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iwarningColor)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[7];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayThongBao));
			string text = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathPhuongPhapHuy));
			string value = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiLapBieu));
			string value2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiDaiDien));
			string value3 = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
			for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
			{
				array[0] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_MaLoaiHD);
				array[1] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_TenLoaiHD);
				array[2] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_MauSoHD);
				array[3] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_KyHieuHD);
				array[4] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_TuSo);
				array[5] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_DenSo);
				array[6] = ConvertHelpers.ConvertCellValueToText(sheetView, num + num2, idxCol_SoLuong);
				string[] array2 = GeneralFunctions.checkTrung(winFpSpread, num, "V", "Y", "B", "E", "O", "S", "", isMST: false, 1);
				string strLoai = "";
				string text2 = "";
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
					continue;
				}
				GeneralFunctions.strMauSoHD_01GTKT = array[0];
				if (HTKKFormater.isNullString(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MaLoaiHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MaLoaiHD);
				}
				if (HTKKFormater.isNullString(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TenLoaiHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TenLoaiHD);
				}
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_MauSoHD);
				if (HTKKFormater.isNullString(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else
				{
					string text3 = GeneralFunctions.CheckMauSoHD(array[2], strLoai, array[0]);
					if (text3.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
						iwarningColor = 2;
					}
					else if (text3.Equals("2"))
					{
						if (array[3].Length > 0)
						{
							if (array[3].Substring(array[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienHDDT(array[2], strLoai, array[0], array[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
									if (iwarningColor != 2)
									{
										iwarningColor = 1;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
								iwarningColor = 2;
							}
						}
					}
					else if (array[3].Length > 0 && array[3].Substring(array[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(array[2], strLoai, array[0], array[3]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num + num2, idxCol_MauSoHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
						if (iwarningColor != 2)
						{
							iwarningColor = 1;
						}
					}
				}
				if (HTKKFormater.isNullString(array[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_KyHieuHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else
				{
					string text4 = GeneralFunctions.CheckSoHD(array[3], strLoai);
					if (text4.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_KyHieuHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
						iwarningColor = 2;
					}
					else if (text4.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_KyHieuHD, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						iwarningColor = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_KyHieuHD);
					}
				}
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TuSo);
				HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_DenSo);
				if (HTKKFormater.isNullString(array[4]) && HTKKFormater.isNullString(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else if (HTKKFormater.isNullString(array[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_DenSo);
					iwarningColor = 2;
				}
				else if (HTKKFormater.isNullString(array[5]))
				{
					HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TuSo);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else
				{
					double num3 = ConvertHelpers.ConvertToDouble(array[4]);
					double num4 = ConvertHelpers.ConvertToDouble(array[5]);
					if ((num3 > num4 && num4 > 0.0) || (num3 == 0.0 && num4 == 0.0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						iwarningColor = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_TuSo);
						HTKKFile.RemoveValidate(sheetView, num + num2, idxCol_DenSo);
					}
				}
				if (array2 == null || array2.Length <= 1)
				{
					continue;
				}
				for (int num5 = 0; num5 < array2.Length; num5++)
				{
					if (num + num2 == ConvertHelpers.ConvertToInt(array2[num5]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num2, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						iwarningColor = 2;
						break;
					}
				}
			}
			if (string.IsNullOrEmpty(value3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
			}
			if (string.IsNullOrEmpty(value))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			}
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
			}
			DateTime dtDate;
			if (string.IsNullOrEmpty(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate);
				if (dtDate != DateTime.MinValue)
				{
					if (DateTime.Compare(dtDate, DateTime.Now) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
						iwarningColor = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					}
				}
			}
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, GeneralFunctions.ArrayCellId("O_8")[1] + 1, GeneralFunctions.ArrayCellId("O_8")[0] + 1, "14");
			if (string.IsNullOrEmpty(valueCustom))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("O_8")[1], GeneralFunctions.ArrayCellId("O_8")[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				iwarningColor = 2;
			}
			else
			{
				HTKKFormater.IsDateTime(valueCustom, "dd/MM/yyyy", out var dtDate2);
				HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate);
				if (dtDate2 != DateTime.MinValue)
				{
					if (DateTime.Compare(dtDate2, DateTime.Now) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("O_8")[1], GeneralFunctions.ArrayCellId("O_8")[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
						iwarningColor = 2;
					}
					else if (DateTime.Compare(dtDate2, DateTime.ParseExact("01/01/2011", "dd/MM/yyyy", CultureInfo.CurrentCulture)) < 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("O_8")[1], GeneralFunctions.ArrayCellId("O_8")[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
						iwarningColor = 2;
					}
					else if (DateTime.Compare(dtDate2, dtDate) > 0 && dtDate != DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("O_8")[1], GeneralFunctions.ArrayCellId("O_8")[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28")));
						iwarningColor = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, GeneralFunctions.ArrayCellId("O_8")[1], GeneralFunctions.ArrayCellId("O_8")[0]);
					}
				}
			}
			switch (iwarningColor)
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

	public Dictionary<string, string> GetDMHoaDon(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Expected O, but got Unknown
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
				string value2 = array[1];
				dictionary2.Add(key, value2);
				dictionary.Add(key, array);
				listHD.Add(value);
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

	public void CellFormatImport(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			if (sheetIndex == 0)
			{
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			List<HTKKCMCCellInfo> staticCells = hTKKCMCSheetInfo.StaticCells;
			string s = string.Empty;
			string value = string.Empty;
			string value2 = string.Empty;
			string value3 = string.Empty;
			foreach (HTKKCMCCellInfo item in staticCells)
			{
				if (item.Path.Contains("thoiGian"))
				{
					s = item.Value.Substring(0, 10);
					value = item.Value.Substring(11, 2);
					value2 = item.Value.Substring(14, 2);
					value3 = item.Value;
					break;
				}
			}
			DateTime dateTime = DateTime.ParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture);
			winFpSpread.Sheets[0].Cells[CellId_NgayMat[1], CellId_NgayMat[0]].Value = dateTime.ToString("dd/MM/yyyy");
			winFpSpread.Sheets[0].Cells[CellId_NgayMat[1] + 1, CellId_NgayMat[0]].Value = value3;
			winFpSpread.Sheets[0].Cells[CellId_GioMat[1], CellId_GioMat[0]].Value = value;
			winFpSpread.Sheets[0].Cells[CellId_PhutMat[1], CellId_PhutMat[0]].Value = value2;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
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
