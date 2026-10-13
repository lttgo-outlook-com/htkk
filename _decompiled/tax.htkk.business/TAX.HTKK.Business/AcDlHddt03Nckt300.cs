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
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class AcDlHddt03Nckt300
{
	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicThueSuatTmp = new Dictionary<string, string>();

	private int idxCellThueSuat = 1;

	private int idxCellGiaTriThueSuat = 6;

	private int idxCellMaThueSuat = 4;

	private int idxCellMaThueSuatKxXml = 5;

	private string pathCellDiaDanh = "TKhai/DLTKhai/TTChung/DDanh";

	private string pathCellNgayLap = "TKhai/DLTKhai/TTChung/NLap";

	private int irowForcus = -1;

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_I => HTKKFile.ColumnLettersToIndex("I") - 1;

	private int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_K => HTKKFile.ColumnLettersToIndex("K") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_M => HTKKFile.ColumnLettersToIndex("M") - 1;

	private int idxCol_O => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_T => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_U => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxCol_Y => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int rowIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.RowIndex;
				int colIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.ColIndex;
				int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.RowIndex;
				int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.ColIndex;
				int rowIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.RowIndex;
				int colIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.ColIndex;
				Dictionary<string, string> dictionary = DMThueSuat();
				string text = ConvertHelpers.ConvertToString(sView.Cells[rowIndex2 - 1, colIndex2 - 1].Value);
				string text2 = sView.Cells[rowIndex - 1, colIndex - 1].Text;
				string text3 = "";
				if (text2 != null && !text2.Equals("") && dictionary.Count > 0)
				{
					text3 = dictionary[text2];
				}
				sView.Cells[rowIndex - 1, colIndex3 - 1].Text = text3;
				if (text3.Equals("KHAC"))
				{
					sView.Cells[rowIndex2 - 1, colIndex2 - 1].Locked = false;
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.MinimumValue = 0.0;
					numberCellType.MaximumValue = 99.0;
					numberCellType.DecimalPlaces = 2;
					sView.Cells[rowIndex2 - 1, colIndex2 - 1].CellType = numberCellType;
				}
				else
				{
					sView.Cells[rowIndex2 - 1, colIndex2 - 1].Locked = true;
					TextCellType cellType = new TextCellType();
					sView.Cells[rowIndex2 - 1, colIndex2 - 1].CellType = cellType;
				}
			}
			int[] array = GeneralFunctions.ArrayCellId("J_5");
			sView.Cells[array[1], array[0]].Text = PeriodInfo.LanXuatBan;
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
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("AF");
			int num5 = HTKKFile.ColumnLettersToIndex("AG");
			SetLaiDoRong(sheetView, (int)iRow);
			Dictionary<string, string> dictionary = DMThueSuat();
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num6 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num7 = num6 + hTKKCMCTableInfo.RowsValues.Count - 1;
				int rowIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.RowIndex;
				int colIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.ColIndex;
				int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.RowIndex;
				int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.ColIndex;
				int rowIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.RowIndex;
				int colIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.ColIndex;
				string text = "";
				string text2 = "";
				string text3 = "";
				text3 = ConvertHelpers.ConvertToString(sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Value);
				text2 = sheetView.Cells[rowIndex - 1, colIndex - 1].Text;
				text = dictionary[text2];
				if (iRow == rowIndex && iCol == colIndex)
				{
					sheetView.Cells[rowIndex - 1, num4 - 1].Text = text;
					if (text.Equals("KHAC"))
					{
						sheetView.Cells[(int)iRow - 1, num2 - 1].Locked = false;
						NumberCellType numberCellType = new NumberCellType();
						numberCellType.MinimumValue = 0.0;
						numberCellType.MaximumValue = 99.0;
						numberCellType.DecimalPlaces = 2;
						sheetView.Cells[(int)iRow - 1, num2 - 1].CellType = numberCellType;
						sheetView.Cells[(int)iRow - 1, num2 - 1].Value = 0;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num2 - 1].Locked = true;
						TextCellType cellType = new TextCellType();
						sheetView.Cells[(int)iRow - 1, num2 - 1].CellType = cellType;
						sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
					}
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						string text4 = "";
						if (text.Equals("0%") || text.Equals("KCT") || text.Equals("KKKNT"))
						{
							text4 = "0*T" + (num6 + j + 1);
						}
						else if (text.Equals("5%"))
						{
							text4 = "ROUND(5/100*T" + (num6 + j + 1) + ",0)";
						}
						else if (text.Equals("8%"))
						{
							text4 = "ROUND(8/100*T" + (num6 + j + 1) + ",0)";
						}
						else if (text.Equals("10%"))
						{
							text4 = "ROUND(10/100*T" + (num6 + j + 1) + ",0)";
						}
						else if (text.Equals("KHAC"))
						{
							if ("".Equals(text3))
							{
								text3 = "0";
							}
							text4 = "ROUND(" + text3 + "/100*T" + (num6 + j + 1) + ",0)";
						}
						else
						{
							text4 = "0*T" + (num6 + j + 1);
						}
						sheetView.Cells[num6 + j, idxCol_U].Formula = text4;
					}
				}
				if (iRow == rowIndex2 && iCol == colIndex2)
				{
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						string text4 = "";
						if ("".Equals(text3))
						{
							text3 = "0";
						}
						text4 = "ROUND(" + text3 + "/100*T" + (num6 + j + 1) + ",0)";
						sheetView.Cells[num6 + j, idxCol_U].Formula = text4;
					}
				}
				if (iRow >= num6 + 1 && iRow <= num7 + 1 && iCol == idxCol_T + 1)
				{
					string text4 = "";
					if (!"".Equals(text2))
					{
						if (text.Equals("0%") || text.Equals("KCT") || text.Equals("KKKNT"))
						{
							text4 = "0*T" + iRow;
						}
						else if (text.Equals("5%"))
						{
							text4 = "ROUND(5/100*T" + iRow + ",0)";
						}
						else if (text.Equals("8%"))
						{
							text4 = "ROUND(8/100*T" + iRow + ",0)";
						}
						else if (text.Equals("10%"))
						{
							text4 = "ROUND(10/100*T" + iRow + ",0)";
						}
						else if (text.Equals("KHAC"))
						{
							if ("".Equals(text3))
							{
								text3 = "0";
							}
							text4 = "ROUND(" + text3 + "/100*T" + iRow + ",0)";
						}
						else
						{
							text4 = "0*T" + iRow;
						}
					}
					else
					{
						text4 = "0*T" + iRow;
					}
					sheetView.Cells[(int)iRow - 1, idxCol_U].Formula = text4;
				}
				if (iRow - 1 >= num6 && iRow - 1 <= num7 && iCol == num)
				{
					string text5 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
					string text6 = "";
					string[] array = dicDMHoaDon[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
					if (array.Length >= 3)
					{
						text6 = array[0];
						sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[1];
					}
					else if (array.Length == 2)
					{
						sheetView.Cells[(int)iRow - 1, num2 - 1].Text = "";
					}
					if (text6.Equals("5"))
					{
						TextCellType textCellType = new TextCellType();
						textCellType.WordWrap = true;
						TextCellType textCellType2 = textCellType;
						textCellType2.MaxLength = 12;
						sheetView.Cells[(int)iRow - 1, num2 - 1].CellType = textCellType2;
					}
					else
					{
						TextCellType textCellType3 = new TextCellType();
						textCellType3.WordWrap = true;
						TextCellType textCellType2 = textCellType3;
						textCellType2.MaxLength = 11;
						sheetView.Cells[(int)iRow - 1, num2 - 1].CellType = textCellType2;
					}
				}
			}
			sheetView.Recalculate();
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
			SetLaiDoRong(sheetView, (int)iRow);
			string text = "";
			string[] array = new string[1];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("T");
			int num5 = HTKKFile.ColumnLettersToIndex("U");
			string text2 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
			string text3 = "";
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num6 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num7 = num6 + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (iRow - 1 >= num6 && iRow - 1 <= num7)
				{
					if (!string.IsNullOrEmpty(text2))
					{
						text3 = dicDMHoaDon[text2][0];
					}
					if (iCol == num3)
					{
						text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
						sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieu(text);
					}
				}
			}
			if (iCol == num4 || iCol == num5 || iCol == num2)
			{
				tinhGiaTriDongTong(winFpSpread, ToKhaiInfo);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("K") - 1;
			if (idxSheet != 0)
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num3 = num2 + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num2 && idxRow <= num3 && idxCol == num)
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
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				int rowIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.RowIndex;
				int colIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.ColIndex;
				int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.RowIndex;
				int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.ColIndex;
				int rowIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.RowIndex;
				int colIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.ColIndex;
				string text10 = sheetView.Cells[rowIndex - 1, colIndex - 1].Text;
				string text11 = sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Text;
				string text12 = sheetView.Cells[rowIndex3 - 1, colIndex3 - 1].Text;
				if (string.IsNullOrEmpty(text12))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex3 - 1, colIndex3 - 1, text);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, rowIndex3 - 1, colIndex3 - 1);
				}
				HTKKFile.RemoveValidate(sheetView, rowIndex - 1, colIndex - 1);
				if (text11.Equals("KHAC") && ConvertHelpers.ConvertToDouble(text10) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex - 1, colIndex - 1, text);
					value = 2;
				}
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_C].Text);
					array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_I].Text);
					array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_J].Text);
					array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_K].Text);
					array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_L].Text);
					array[6] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_M].Text);
					array[7] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_O].Text);
					array[8] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_T].Value);
					array[9] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_U].Value);
					array[10] = ConvertHelpers.ConvertToString(sheetView.Cells[num + j, idxCol_Y].Text);
					bool flag = GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array);
					string text13 = "";
					string text14 = "";
					string strTemp = "";
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_C);
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_I);
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_J);
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_K);
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_L);
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_T);
					HTKKFile.RemoveValidate(sheetView, num + j, idxCol_O);
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_C, text, j + 1);
						value = 2;
					}
					if (string.IsNullOrEmpty(array[2]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_I, text, j + 1);
						value = 2;
					}
					else
					{
						if (!string.IsNullOrEmpty(array[1]))
						{
							text13 = dicDMHoaDon[array[1]][0];
							strTemp = dicDMHoaDon[array[1]][1];
						}
						if (!text13.Equals("5"))
						{
							text14 = GeneralFunctions.CheckMauSoHD(array[2], text13, strTemp);
							if (text14.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_I, text2, j + 1);
								value = 2;
							}
							else if (text14.Equals("2"))
							{
								if (array[3].Length > 0)
								{
									if (array[3].Substring(array[3].Length - 1).Equals("E"))
									{
										if (GeneralFunctions.CheckSoLienHDDT(array[2], text13, strTemp, array[3]).Equals("3"))
										{
											HTKKFile.SetNodeBackGroundYellow(sheetView, num + j, idxCol_I, text7);
										}
									}
									else
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_I, text6, j + 1);
										value = 2;
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_I, text6, j + 1);
									value = 2;
								}
							}
							else if (array[3].Length > 0 && array[3].Substring(array[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienHDDT(array[2], text13, strTemp, array[3]).Equals("3"))
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, num + j, idxCol_I, text7);
							}
						}
					}
					if (string.IsNullOrEmpty(array[3]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_J, text, j + 1);
						value = 2;
					}
					else
					{
						if (!string.IsNullOrEmpty(array[1]))
						{
							text13 = dicDMHoaDon[array[1]][0];
						}
						if (!text13.Equals("5"))
						{
							text14 = GeneralFunctions.CheckSoHD(array[3], text13);
							if (text14.Equals("1"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_J, text3, j + 1);
								value = 2;
							}
							else if (text14.Equals("2"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_J, text4, j + 1);
								value = 2;
							}
						}
					}
					if (string.IsNullOrEmpty(array[4]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_K, text, j + 1);
						value = 2;
					}
					DateTime dtDate;
					if (string.IsNullOrEmpty(array[5]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_L, text, j + 1);
						value = 2;
					}
					else if (HTKKFormater.IsDateTime(array[5], "dd/MM/yyyy", out dtDate) && dtDate >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
					{
						if (dtDate > DateTime.Now.Date)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_L, text8, j + 1);
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_L, text9, j + 1);
						value = 2;
					}
					if (!string.IsNullOrEmpty(array[7]) && !ValidateBase.CheckMaSoThue(array[7].Trim()))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxCol_O, text5, j + 1);
						value = 2;
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellDiaDanh));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayLap));
			string text15 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
			string text16 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
			if (string.IsNullOrEmpty(text15))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			}
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			DateTime dtDate2;
			if (string.IsNullOrEmpty(text16))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text);
				value = 2;
			}
			else if (!HTKKFormater.IsDateTime(text16, "dd/MM/yyyy", out dtDate2) || !(dtDate2 >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture)))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text9);
				value = 2;
			}
			if (value == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			else if (value == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
			}
			else if (value == 0)
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMHoaDonDT(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
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
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[2];
				string text = array[2];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AH");
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			Dictionary<string, string> source = DMThueSuat();
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int rowIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.RowIndex;
				int colIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.ColIndex;
				int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.RowIndex;
				int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.ColIndex;
				int rowIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.RowIndex;
				int colIndex3 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuat].CellTemplate.ColIndex;
				int rowIndex4 = hTKKCMCTableInfo.Header.StaticCells[0].CellTemplate.RowIndex;
				string maThueSuat = sheetView.Cells[rowIndex4 - 1, colIndex3 - 1].Text.Trim();
				KeyValuePair<string, string> keyValuePair = source.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maThueSuat);
				if (!string.IsNullOrEmpty(keyValuePair.Key))
				{
					sheetView.Cells[rowIndex4 - 1, colIndex3 - 1].Value = maThueSuat;
					sheetView.Cells[rowIndex4 - 1, colIndex - 1].Value = keyValuePair.Key;
				}
				if (maThueSuat.Equals("KHAC"))
				{
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Locked = false;
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.MinimumValue = 0.0;
					numberCellType.MaximumValue = 99.0;
					numberCellType.DecimalPlaces = 2;
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].CellType = numberCellType;
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Formula = "AH" + (rowIndex2 + 1);
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Value = sheetView.Cells[rowIndex2, num - 1].Value;
				}
				else
				{
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Locked = true;
					TextCellType cellType = new TextCellType();
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].CellType = cellType;
				}
				string text = ConvertHelpers.ConvertToString(sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Value);
				string text2 = sheetView.Cells[rowIndex - 1, colIndex - 1].Text;
			}
			sheetView.Recalculate();
			tinhGiaTriDongTong(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int rowIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.RowIndex;
				int colIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.ColIndex;
				int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuatKxXml].CellTemplate.RowIndex;
				int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellMaThueSuatKxXml].CellTemplate.ColIndex;
				Dictionary<string, string> dictionary = DMMaThueSuat();
				string text = sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Text;
				string text2 = "";
				if (text != null && text.IndexOf("KHAC") >= 0)
				{
					text2 = dictionary["KHAC"];
					sheetView.Cells[rowIndex - 1, colIndex - 1].Text = text2;
				}
				else
				{
					text2 = dictionary[text];
					sheetView.Cells[rowIndex - 1, colIndex - 1].Text = text2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void reloadCacheTuSoDenSo(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow)
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

	private void tinhGiaTriDongTong(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			double num = 0.0;
			double num2 = 0.0;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num3 = HTKKFile.ColumnLettersToIndex("T");
			int num4 = HTKKFile.ColumnLettersToIndex("U");
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("TKhai/DLTKhai/NDTKhai/TgDThu"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("TKhai/DLTKhai/NDTKhai/TgThue"));
			for (int num5 = 0; num5 < hTKKCMCSheetInfo.Tables.Count; num5++)
			{
				double num6 = 0.0;
				double num7 = 0.0;
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[num5];
				int num8 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int num9 = 0; num9 < hTKKCMCTableInfo.RowsValues.Count; num9++)
				{
					num6 += ConvertHelpers.ConvertToDouble(sheetView.Cells[num8 + num9, idxCol_T].Value);
					num7 += ConvertHelpers.ConvertToDouble(sheetView.Cells[num8 + num9, idxCol_U].Value);
				}
				num += num6;
				num2 += num7;
				sheetView.Cells[hTKKCMCTableInfo.Header.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[2].CellTemplate.ColIndex - 1].Value = num6;
				sheetView.Cells[hTKKCMCTableInfo.Header.StaticCells[3].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[3].CellTemplate.ColIndex - 1].Value = num7;
			}
			sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
			sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private Dictionary<string, string> DMThueSuat()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_thue_suat_HDDT.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[2];
				string value2 = array[1];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private Dictionary<string, string> DMMaThueSuat()
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_thue_suat_HDDT.xml";
			XmlDocument val = new XmlDocument();
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[2];
				dictionary.Add(key, value2);
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int rowIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.RowIndex;
				int colIndex = hTKKCMCTableInfo.Header.StaticCells[idxCellThueSuat].CellTemplate.ColIndex;
				int rowIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.RowIndex;
				int colIndex2 = hTKKCMCTableInfo.Header.StaticCells[idxCellGiaTriThueSuat].CellTemplate.ColIndex;
				string text = sheetView.Cells[rowIndex - 1, colIndex - 1].Text;
				if (string.IsNullOrEmpty(text))
				{
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Locked = true;
					TextCellType cellType = new TextCellType();
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].CellType = cellType;
					sheetView.Cells[rowIndex2 - 1, colIndex2 - 1].Text = "";
				}
			}
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
			if (idxRow == 3 && idxCol == 2)
			{
				ShowBangKe("Bangke_03_DL_HDDT.xls");
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

	public void SetLaiDoRong(SheetView sView, int iRow)
	{
		if (irowForcus != iRow)
		{
			irowForcus = iRow;
			string text = sView.Cells[iRow - 1, idxCol_C].Text;
			string text2 = "";
			string[] array = dicDMHoaDon[text];
			if (array.Length >= 3)
			{
				text2 = array[0];
			}
			if (text2.Equals("5"))
			{
				TextCellType textCellType = new TextCellType();
				textCellType.WordWrap = true;
				TextCellType textCellType2 = textCellType;
				textCellType2.MaxLength = 12;
				sView.Cells[iRow - 1, idxCol_I].CellType = textCellType2;
			}
			else
			{
				TextCellType textCellType3 = new TextCellType();
				textCellType3.WordWrap = true;
				TextCellType textCellType2 = textCellType3;
				textCellType2.MaxLength = 11;
				sView.Cells[iRow - 1, idxCol_I].CellType = textCellType2;
			}
		}
	}
}
