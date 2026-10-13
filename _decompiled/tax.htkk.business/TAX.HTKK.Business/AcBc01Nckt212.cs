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

public class AcBc01Nckt212
{
	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

	private List<string> listHD = new List<string>();

	private string pathNgayThongBao = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayLapBCao";

	private string pathNguoiKy = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private int idxCol_MaLoaiHD => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_TenLoaiHD => HTKKFile.ColumnLettersToIndex("I") - 1;

	private int idxCol_MauSoHD => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_KyHieuHD => HTKKFile.ColumnLettersToIndex("K") - 1;

	private int idxCol_SoLuong => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("M") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayThongBao));
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].ResetFormula();
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = DateTime.Now.ToString("dd/MM/yyyy");
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("H");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			if (iCol != num)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, idxCol_MaLoaiHD].Text;
			string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
			if (HTKKFormater.isNullString(text))
			{
				sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
				sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
				return;
			}
			sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("H");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			if (iCol != num)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, idxCol_MaLoaiHD].Text;
			string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
			if (HTKKFormater.isNullString(text))
			{
				sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
				sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
				return;
			}
			sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "";
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_TuSo || iCol - 1 == idxCol_DenSo)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text2 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
				if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3))
				{
					int num3 = ConvertHelpers.ConvertToInt(text2);
					int num4 = ConvertHelpers.ConvertToInt(text3);
					if (num3 == 0 && num4 == 0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = "0";
					}
					else if (num3 == 0 && num4 != 0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = (num4 - num3).ToString();
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = (num4 - num3 + 1).ToString();
					}
				}
			}
			if (iCol - 1 == idxCol_MauSoHD)
			{
				string text4 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("H") - 1].Text;
				string text5 = "";
				if (!string.IsNullOrEmpty(text4))
				{
					text5 = dicDMHoaDon[text4][0];
				}
				if (!text5.Equals("3"))
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
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[11, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			string[] array = new string[12];
			List<int> list = new List<int>();
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, dataLocationTemplateRowStart - 1, "L", "M", "B", "I", "J", "K", "C", isMST: true, 1);
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int num = HTKKFile.ColumnLettersToIndex("H");
				int num2 = HTKKFile.ColumnLettersToIndex("C");
				int num3 = HTKKFile.ColumnLettersToIndex("D");
				int num4 = HTKKFile.ColumnLettersToIndex("E");
				int num5 = HTKKFile.ColumnLettersToIndex("F");
				int num6 = HTKKFile.ColumnLettersToIndex("G");
				int num7 = HTKKFile.ColumnLettersToIndex("I");
				int num8 = HTKKFile.ColumnLettersToIndex("J");
				int num9 = HTKKFile.ColumnLettersToIndex("K");
				int num10 = HTKKFile.ColumnLettersToIndex("L");
				int num11 = HTKKFile.ColumnLettersToIndex("M");
				int num12 = HTKKFile.ColumnLettersToIndex("Q");
				array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value);
				array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value);
				array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
				array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value);
				array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Value);
				array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Text);
				array[6] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num7 - 1].Value);
				array[7] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num8 - 1].Value);
				array[8] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num9 - 1].Value);
				array[9] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num10 - 1].Value);
				array[10] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Value);
				array[11] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num12 - 1].Value);
				string strLoai = "";
				string text10 = "";
				bool flag = false;
				if (!string.IsNullOrEmpty(array[0]))
				{
					strLoai = dicDMHoaDon[array[0]][0];
				}
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num12 - 1);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1);
				if (!HTKKFormater.isNullString(array[0]) || !HTKKFormater.isNullString(array[6]) || !HTKKFormater.isNullString(array[7]) || !HTKKFormater.isNullString(array[8]) || !HTKKFormater.isNullString(array[9]) || !HTKKFormater.isNullString(array[10]))
				{
					flag = true;
				}
				GeneralFunctions.strMauSoHD_01GTKT = array[0];
				if (HTKKFormater.isNullString(array[0]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, text7);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[6]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, text7);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[9]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1, text7);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[10]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1, text7);
					list.Add(1);
				}
				if (!HTKKFormater.isNullString(array[9]) && !HTKKFormater.isNullString(array[10]))
				{
					if (ConvertHelpers.ConvertToInt(array[10]) > 0 && ConvertHelpers.ConvertToInt(array[9]) > ConvertHelpers.ConvertToInt(array[10]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1, text6);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1, text6);
						list.Add(1);
					}
					else if (array2 != null && array2.Length > 1)
					{
						for (int j = 0; j < array2.Length; j++)
						{
							if (dataLocationTemplateRowStart + i - 1 == ConvertHelpers.ConvertToInt(array2[j]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num11 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
								list.Add(1);
								break;
							}
						}
					}
				}
				if (array[9].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, idxCol_TuSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				if (array[10].ToIntOrDefault(-1) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, idxCol_DenSo, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				if (HTKKFormater.isNullString(array[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, text7);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text7);
					list.Add(1);
				}
				else
				{
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(array[5], "dd/MM/yyyy");
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
					DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(PeriodInfo.DenNgay, "dd/MM/yyyy");
					if (((dateTime - dateTime3).TotalDays > 0.0 || (dateTime2 - dateTime).TotalDays > 0.0) && dateTime != DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text8);
						list.Add(0);
					}
				}
				if (HTKKFormater.isNullString(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text7);
					list.Add(1);
				}
				else if (!ValidateBase.CheckMaSoThue(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text9);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[7]))
				{
					if (flag)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text7);
						list.Add(1);
					}
				}
				else
				{
					text10 = GeneralFunctions.CheckMauSoHD(array[7], strLoai, array[0]);
					if (text10.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text);
						list.Add(1);
					}
					else if (text10.Equals("2"))
					{
						if (array[8].Length > 0)
						{
							if (array[8].Substring(array[8].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienHDDT(array[7], strLoai, array[0], array[8]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text3);
									list.Add(0);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text2);
								list.Add(1);
							}
						}
					}
					else if (array[8].Length > 0 && array[8].Substring(array[8].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(array[7], strLoai, array[0], array[8]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text3);
						list.Add(0);
					}
				}
				if (HTKKFormater.isNullString(array[8]))
				{
					if (flag)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text7);
						list.Add(1);
					}
					continue;
				}
				text10 = GeneralFunctions.CheckSoHD(array[8], strLoai);
				if (text10.Equals("1"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text4);
					list.Add(1);
				}
				else if (text10.Equals("2"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text5);
					list.Add(1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayThongBao));
			string text11 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			if (text11 != "")
			{
				DateTime dateTime4 = DateTime.ParseExact(text11.Replace("/", ""), "ddMMyyyy", CultureInfo.CurrentCulture);
				if (dateTime4 > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					list.Add(1);
				}
				else if (dateTime4 < DateTime.ParseExact(PeriodInfo.TuNgay.Replace("/", ""), "ddMMyyyy", CultureInfo.CurrentCulture))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21")));
					list.Add(1);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
				list.Add(1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiKy));
			string value2 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13")));
				list.Add(1);
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
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
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMHoaDon(string pathFile, HTKKPeriodInfo PeriodInfo)
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
				string text = array[1];
				string value2 = array[1];
				if (text.IndexOf("06HDXK") <= -1 || (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) <= 2014 && (!(PeriodInfo.Quy == "3") || ConvertHelpers.ConvertToInt(PeriodInfo.Nam) != 2014)) || !(PeriodInfo.KieuKy == "Q"))
				{
					dictionary2.Add(text, value2);
					dictionary.Add(text, array);
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 4 && idxCol == 1)
			{
				ShowBangKe("Bangke_01BCAC.xls");
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("H");
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
				string text2 = sheetView.Cells[num + i, idxCol_TuSo].Text;
				string text3 = sheetView.Cells[num + i, idxCol_DenSo].Text;
				if (text2.Equals("0000000") && text3.Equals("0000000"))
				{
					sheetView.Cells[num + i, idxCol_TuSo].Text = "";
					sheetView.Cells[num + i, idxCol_DenSo].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/HoaDon");
			setValueTuSoDenSo(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/HoaDon/ChiTietHoaDon/tuSo");
			setValueTuSoDenSo(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/HoaDon/ChiTietHoaDon/denSo");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void setValueTuSoDenSo(XmlNode xmlNode, string path)
	{
		//IL_0036: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Expected O, but got Unknown
		try
		{
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(path);
			if (val == null || val.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				if (val2.InnerText.Trim().Equals(""))
				{
					val2.InnerText = "0000000";
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
