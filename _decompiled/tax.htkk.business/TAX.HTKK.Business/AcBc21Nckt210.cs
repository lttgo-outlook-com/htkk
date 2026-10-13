using System;
using System.Collections.Generic;
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

public class AcBc21Nckt210
{
	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMGhiChu = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMGhiChuTmp = new Dictionary<string, string>();

	private List<string> listHD = new List<string>();

	private string pathNguoiKy = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiDaiDien";

	private string pathNgayLap = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayLapBCao";

	private string pathLyDoMat = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/lyDo";

	private string cellId_NgayMat = "O_8";

	private string cellId_GioMat = "K_8";

	private string cellId_PhutMat = "M_8";

	private int[] CellId_NgayMat => GeneralFunctions.ArrayCellId(cellId_NgayMat);

	private int[] CellId_GioMat => GeneralFunctions.ArrayCellId(cellId_GioMat);

	private int[] CellId_PhutMat => GeneralFunctions.ArrayCellId(cellId_PhutMat);

	private int iCol_MaLoaiHD => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int iCol_TenLoaiHD => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int iCol_MauSoHD => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int iCol_KyHieuHD => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int iCol_SoLuong => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int iCol_TuSo => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int iCol_DenSo => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int iCol_LienHoaDon => HTKKFile.ColumnLettersToIndex("AB") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("K");
			int num2 = HTKKFile.ColumnLettersToIndex("M");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = 8;
			if (HTKKFormater.isNullString(sView.Cells[num4 - 1, num - 1].Text) && HTKKFormater.isNullString(sView.Cells[num4 - 1, num2 - 1].Text) && HTKKFormater.isNullString(sView.Cells[num4 - 1, num3 - 1].Text))
			{
				sView.Cells[9, HTKKFile.ColumnLettersToIndex("S") - 1].Text = DateTime.Now.Second.ToString();
				sView.Cells[num4 - 1, num - 1].Text = DateTime.Now.Hour.ToString();
				sView.Cells[num4 - 1, num2 - 1].Text = DateTime.Now.Minute.ToString();
				sView.Cells[num4 - 1, num3 - 1].Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
				int num5 = HTKKFile.ColumnLettersToIndex("K");
				int num6 = HTKKFile.ColumnLettersToIndex("M");
				int num7 = HTKKFile.ColumnLettersToIndex("O");
				int num8 = HTKKFile.ColumnLettersToIndex("S");
				string text = sView.Cells[num4 - 1, num7 - 1].Text;
				string text2 = sView.Cells[num4 - 1, num5 - 1].Text;
				string text3 = sView.Cells[num4 - 1, num6 - 1].Text;
				string text4 = sView.Cells[9, num8 - 1].Text;
				if (text.IndexOf("/") > 0 && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
				{
					string[] array = text.Split(new char[1] { '/' });
					sView.Cells[9, num5 - 1].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text2, 2) + ":" + GeneralFunctions.GetStrByLengh(text3, 2) + ":" + GeneralFunctions.GetStrByLengh(text4, 2) + "Z";
				}
				else
				{
					sView.Cells[9, num5 - 1].Text = "";
				}
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
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AL");
			if (iCol == num)
			{
				string text = sheetView.Cells[(int)iRow - 1, iCol_MaLoaiHD].Text;
				string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
					if (text.Equals("TT120"))
					{
						sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].ResetValue();
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].Text = text;
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
						sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].CellType = textCellType2;
					}
					else
					{
						TextCellType textCellType3 = new TextCellType();
						textCellType3.WordWrap = true;
						TextCellType textCellType2 = textCellType3;
						textCellType2.MaxLength = 20;
						sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].CellType = textCellType2;
					}
				}
			}
			if (iCol == num4)
			{
				string[] array = dicDMGhiChu[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length >= 2)
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = array[0];
				}
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
			if (idxSheet != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (idxRow >= num && idxRow <= num2)
			{
				if (idxCol == iCol_TuSo || idxCol == iCol_DenSo)
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
				if (idxCol == iCol_LienHoaDon)
				{
					GeneralFunctions.CompareKeycode(e, 2);
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
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = HTKKFile.ColumnLettersToIndex("S");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("Y");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AB");
			int num11 = HTKKFile.ColumnLettersToIndex("AC");
			string text2 = sheetView.Cells[(int)iRow, num7 - 1].Text;
			string text3 = sheetView.Cells[(int)iRow, num8 - 1].Text;
			int num12 = HTKKFile.ColumnLettersToIndex("K");
			int num13 = HTKKFile.ColumnLettersToIndex("M");
			int num14 = HTKKFile.ColumnLettersToIndex("O");
			int num15 = HTKKFile.ColumnLettersToIndex("S");
			int num16 = 8;
			if (iRow + 1 == 8 && (iCol == num12 || iCol == num13))
			{
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.GetStrByLengh(sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text, 2);
			}
			if (iRow == 8 && (iCol == num12 || iCol == num13 || iCol == num14))
			{
				string text4 = sheetView.Cells[num16 - 1, num14 - 1].Text;
				string text5 = sheetView.Cells[num16 - 1, num12 - 1].Text;
				string text6 = sheetView.Cells[num16 - 1, num13 - 1].Text;
				string text7 = sheetView.Cells[9, num15 - 1].Text;
				if (text4.IndexOf("/") > 0 && !HTKKFormater.isNullString(text5) && !HTKKFormater.isNullString(text6))
				{
					string[] array = text4.Split(new char[1] { '/' });
					string text8 = text5;
					sheetView.Cells[9, num12 - 1].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text5, 2) + ":" + GeneralFunctions.GetStrByLengh(text6, 2) + ":" + GeneralFunctions.GetStrByLengh(text7, 2) + "Z";
				}
				else
				{
					sheetView.Cells[9, num12 - 1].Text = "";
				}
			}
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == iCol_TuSo || iCol - 1 == iCol_DenSo)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text9 = sheetView.Cells[(int)iRow - 1, iCol_TuSo].Text;
				string text10 = sheetView.Cells[(int)iRow - 1, iCol_DenSo].Text;
				if (!string.IsNullOrEmpty(text9) && !string.IsNullOrEmpty(text10))
				{
					int num17 = ConvertHelpers.ConvertToInt(text9);
					int num18 = ConvertHelpers.ConvertToInt(text10);
					if (num17 == 0 && num18 == 0)
					{
						sheetView.Cells[(int)iRow - 1, iCol_SoLuong].Value = "0";
					}
					else if (num17 == 0 && num18 != 0)
					{
						sheetView.Cells[(int)iRow - 1, iCol_SoLuong].Value = (num18 - num17).ToString();
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, iCol_SoLuong].Value = (num18 - num17 + 1).ToString();
					}
				}
			}
			if (iCol - 1 == iCol_MauSoHD)
			{
				string text11 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text;
				string text12 = "";
				if (!string.IsNullOrEmpty(text11))
				{
					text12 = dicDMHoaDon[text11][0];
				}
				if (!text12.Equals("3"))
				{
					string text13 = sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].Text;
					sheetView.Cells[(int)iRow - 1, iCol_MauSoHD].Text = text13.ToUpper();
				}
			}
			if (iCol - 1 == iCol_KyHieuHD)
			{
				string text14 = sheetView.Cells[(int)iRow - 1, iCol_KyHieuHD].Text;
				sheetView.Cells[(int)iRow - 1, iCol_KyHieuHD].Text = GeneralFunctions.FormatKyHieu(text14);
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[13, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			List<int> list = new List<int>();
			string[] array = new string[9];
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, dataLocationTemplateRowStart - 1, "V", "Y", "B", "E", "O", "S", "", isMST: false, 1);
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int num = HTKKFile.ColumnLettersToIndex("C");
				int num2 = HTKKFile.ColumnLettersToIndex("E");
				int num3 = HTKKFile.ColumnLettersToIndex("O");
				int num4 = HTKKFile.ColumnLettersToIndex("S");
				int num5 = HTKKFile.ColumnLettersToIndex("V");
				int num6 = HTKKFile.ColumnLettersToIndex("Y");
				int num7 = HTKKFile.ColumnLettersToIndex("AA");
				int num8 = HTKKFile.ColumnLettersToIndex("AB");
				int num9 = HTKKFile.ColumnLettersToIndex("AC");
				array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value);
				array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value);
				array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
				array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value);
				array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Value);
				array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Value);
				array[6] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num7 - 1].Value);
				array[7] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num8 - 1].Value);
				array[8] = ConvertHelpers.ConvertToString(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num9 - 1].Value);
				string text14 = "";
				string text15 = "";
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
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
				if (!HTKKFormater.isNullString(array[0]))
				{
					text15 = dicDMHoaDon[array[0]][0];
					GeneralFunctions.strMauSoHD_01GTKT = array[0];
				}
				int num10 = 0;
				if (array[2].Length > array[0].Length + 1 && text15.Equals("0"))
				{
					num10 = ConvertHelpers.ConvertToInt(array[2].Substring(array[0].Length, 1));
				}
				if (!string.IsNullOrEmpty(array[7]))
				{
					string[] array3 = array[7].Split(new char[1] { ';' });
					if (array3 != null)
					{
						for (int j = 0; j < array3.Length; j++)
						{
							if (array3[j].Equals("") || ConvertHelpers.ConvertToInt(array3[j]) > 9 || ConvertHelpers.ConvertToInt(array3[j]) < 0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, HTKKFile.GetMessage("0387"));
								list.Add(1);
							}
							else if (ConvertHelpers.ConvertToInt(array3[j]) > num10 && text15.Equals("0"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, text13);
								list.Add(1);
							}
						}
					}
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, text6);
					list.Add(1);
				}
				if (string.IsNullOrEmpty(array[8]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, text7);
					list.Add(1);
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, text6);
					list.Add(1);
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text6);
					list.Add(1);
				}
				else
				{
					text14 = GeneralFunctions.CheckMauSoHD(array[2], text15, array[0]);
					if (text14.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text);
						list.Add(1);
					}
					else if (text14.Equals("2"))
					{
						if (array[3].Length > 0)
						{
							if (array[3].Substring(array[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienHDDT(array[2], text15, array[0], array[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
									list.Add(0);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, text2);
								list.Add(1);
							}
						}
					}
					else if (array[3].Length > 0 && array[3].Substring(array[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(array[2], text15, array[0], array[3]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
						list.Add(0);
					}
				}
				if (string.IsNullOrEmpty(array[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text6);
					list.Add(1);
				}
				else
				{
					text14 = GeneralFunctions.CheckSoHD(array[3], text15);
					if (text14.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text3);
						list.Add(1);
					}
					else if (text14.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, text4);
						list.Add(1);
					}
				}
				if (HTKKFormater.isNullString(array[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, text6);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text6);
					list.Add(1);
				}
				if (HTKKFormater.isNullString(array[4]) || HTKKFormater.isNullString(array[5]))
				{
					continue;
				}
				if (ConvertHelpers.ConvertToInt(array[5]) > 0 && ConvertHelpers.ConvertToInt(array[4]) > ConvertHelpers.ConvertToInt(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, text5);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text5);
					list.Add(1);
				}
				else
				{
					if (array2 == null || array2.Length <= 1)
					{
						continue;
					}
					for (int k = 0; k < array2.Length; k++)
					{
						if (dataLocationTemplateRowStart + i - 1 == ConvertHelpers.ConvertToInt(array2[k]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, text8);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, text8);
							list.Add(1);
							break;
						}
					}
				}
			}
			bool success = false;
			bool success2 = false;
			DateTime dateTime = HTKKFormater.ConvertToDateTime(sheetView.Cells[7, HTKKFile.ColumnLettersToIndex("O") - 1].Text, out success);
			DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("V") - 1].Text, "dd/MM/yyyy");
			string formula = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("C") - 1].Formula;
			string[] array4 = GeneralFunctions.GetCellValild(formula).Split(new char[1] { '#' });
			string text16 = sheetView.Cells[7, HTKKFile.ColumnLettersToIndex("O") - 1].Text;
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayLap));
			string text17 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			DateTime dateTime3 = HTKKFormater.ConvertToDateTime(text17, out success2);
			if (string.IsNullOrEmpty(text17))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text6);
				value = 2;
			}
			else if (!success2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0071"));
				value = 2;
			}
			else if ((dateTime3 - dateTime2).TotalDays < 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, ConvertHelpers.ConvertToInt(array4[1]) - 1, HTKKFile.ColumnLettersToIndex(array4[2]) - 1, text10);
				value = 2;
			}
			else if ((dateTime3 - DateTime.Now.Date).TotalDays > 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, ConvertHelpers.ConvertToInt(array4[1]) - 1, HTKKFile.ColumnLettersToIndex(array4[2]) - 1, text12);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, ConvertHelpers.ConvertToInt(array4[1]) - 1, HTKKFile.ColumnLettersToIndex(array4[2]) - 1);
			}
			if (string.IsNullOrEmpty(text16))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 7, HTKKFile.ColumnLettersToIndex("O") - 1, text6);
				value = 2;
			}
			else if (!success)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 7, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage("0071"));
				value = 2;
			}
			else if ((dateTime - dateTime3).TotalDays > 0.0 && success2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 7, HTKKFile.ColumnLettersToIndex("O") - 1, text9);
				value = 2;
			}
			else if ((dateTime - dateTime2).TotalDays < 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 7, HTKKFile.ColumnLettersToIndex("O") - 1, text11);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 7, HTKKFile.ColumnLettersToIndex("O") - 1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathLyDoMat));
			string value2 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiKy));
			string value3 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
			if (string.IsNullOrEmpty(value3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
			}
			for (int i = 0; i < list.Count; i++)
			{
				if (list[i] == 1)
				{
					value = 2;
					break;
				}
			}
			GeneralFunctions.SetTabStripColor(sheetView, 0, value);
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

	public Dictionary<string, string> GetDMGhiChu(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Expected O, but got Unknown
		try
		{
			if (dicDMGhiChuTmp.Count > 0)
			{
				return dicDMGhiChuTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			pathFile += "\\Catalogue_ghi_chu.xml";
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[0];
				string text = array[1];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
			}
			dicDMGhiChu = dictionary;
			dicDMGhiChuTmp = dictionary2;
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

	public void ExportHeaderXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode parentNode = xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue").ParentNode;
			XmlNode val = parentNode.SelectSingleNode("DLyThue");
			if (val != null)
			{
				parentNode.RemoveChild(val);
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
				if (item.Path.Contains("ngayMatChayHongHD"))
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
			winFpSpread.Sheets[0].Cells[CellId_GioMat[1], CellId_GioMat[0]].Value = value;
			winFpSpread.Sheets[0].Cells[CellId_GioMat[1] + 2, CellId_GioMat[0]].Value = value3;
			winFpSpread.Sheets[0].Cells[CellId_PhutMat[1], CellId_PhutMat[0]].Value = value2;
			string text = "";
			string empty = string.Empty;
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AL");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = winFpSpread.Sheets[0].Cells[num + i, num3 - 1].Text;
				winFpSpread.Sheets[0].Cells[num + i, num2 - 1].ResetText();
				bool flag = true;
				for (int j = 0; j < listHD.Count; j++)
				{
					string input = listHD[j];
					string[] array = Regex.Split(input, "###");
					if (text.IndexOf(array[1], 0) == 0)
					{
						winFpSpread.Sheets[0].Cells[num + i, num2 - 1].Text = array[1];
						flag = false;
						break;
					}
				}
				if (flag && !string.IsNullOrEmpty(text))
				{
					winFpSpread.Sheets[0].Cells[num + i, num2 - 1].Text = "TT120";
				}
				empty = winFpSpread.Sheets[0].Cells[num + i, num5 - 1].Text;
				if (!string.IsNullOrEmpty(empty) && dicDMGhiChuTmp.ContainsKey(empty))
				{
					winFpSpread.Sheets[0].Cells[num + i, num4 - 1].Text = dicDMGhiChuTmp[empty];
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
