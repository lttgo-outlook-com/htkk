using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tts01Nckt216
{
	private Dictionary<string, string[]> dicDMHopDong = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMHopDongTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMMucDich = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMMucDichTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMTTHopDong = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMTTHopDongTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicFormulaTK = new Dictionary<string, string>();

	private Dictionary<string, double[]> dicTongCT20a = new Dictionary<string, double[]>();

	private Dictionary<string, double[]> dicTongCT20 = new Dictionary<string, double[]>();

	private Dictionary<string, string> dicHDXoa = new Dictionary<string, string>();

	private bool isFirtLoad = true;

	private int levelCall = 0;

	private DateTime dateNamDau;

	private DateTime dateNamCuoi;

	private bool isLienTiep = false;

	private bool isDataVB6 = false;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 47;
			int num2 = 49;
			int num3 = 90;
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			int num5 = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("G");
			int num7 = HTKKFile.ColumnLettersToIndex("H");
			int num8 = HTKKFile.ColumnLettersToIndex("I");
			int num9 = HTKKFile.ColumnLettersToIndex("J");
			int num10 = HTKKFile.ColumnLettersToIndex("K");
			int num11 = HTKKFile.ColumnLettersToIndex("N");
			int num12 = HTKKFile.ColumnLettersToIndex("P");
			int num13 = HTKKFile.ColumnLettersToIndex("S");
			int num14 = HTKKFile.ColumnLettersToIndex("T");
			int num15 = HTKKFile.ColumnLettersToIndex("U");
			int num16 = HTKKFile.ColumnLettersToIndex("V");
			int num17 = HTKKFile.ColumnLettersToIndex("W");
			int num18 = HTKKFile.ColumnLettersToIndex("X");
			int num19 = HTKKFile.ColumnLettersToIndex("Y");
			int num20 = HTKKFile.ColumnLettersToIndex("Z");
			int num21 = HTKKFile.ColumnLettersToIndex("AA");
			int num22 = HTKKFile.ColumnLettersToIndex("AB");
			int num23 = HTKKFile.ColumnLettersToIndex("AC");
			int num24 = HTKKFile.ColumnLettersToIndex("AD");
			int num25 = HTKKFile.ColumnLettersToIndex("AE");
			int num26 = HTKKFile.ColumnLettersToIndex("AF");
			int num27 = HTKKFile.ColumnLettersToIndex("AG");
			int num28 = HTKKFile.ColumnLettersToIndex("AI");
			int num29 = HTKKFile.ColumnLettersToIndex("AJ");
			string text = sheetView.Cells[53, HTKKFile.ColumnLettersToIndex("U") - 1].Text;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[0].StaticCells.Count; i++)
			{
				int rowIndex = ToKhaiInfo.SheetInfos[0].StaticCells[i].CellTemplate.RowIndex;
				if (rowIndex >= 54 && rowIndex <= 62)
				{
					string formula = sheetView.Cells[rowIndex - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Formula;
					HTKKFile.AddOrUpdateDicStr(dicFormulaTK, "U_" + rowIndex, formula);
				}
			}
			int num30 = HTKKFile.ColumnLettersToIndex("E");
			HTKKFile.AddOrUpdateDicStr(dicFormulaTK, "E_49", "IF(L90=\"1\",\"\",TRIM('###'!AJ22))");
			winFpSpread.Sheets[0].Cells[48, num30 - 1].Formula = "";
			if (PeriodInfo.KieuKy == "Y")
			{
				winFpSpread.Sheets[0].Cells[num3 - 1, num4 - 1].Text = "";
				winFpSpread.Sheets[0].Cells[num3 - 1, num4 - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[num3 - 1, num5 - 1].Text = "";
				winFpSpread.Sheets[0].Cells[num3 - 1, num5 - 1].Formula = "";
				Row row = sheetView.Rows[num - 1];
				row.Visible = false;
				Row row2 = sheetView.Rows[num2 - 1];
				row2.Visible = false;
			}
			if (hTKKCMCSheetInfo != null && ConvertHelpers.ConvertToDouble1(text) == 0.0)
			{
				string sheetName = sheetView2.SheetName;
				object value = sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("B") - 1].Value;
				double num31 = 0.0;
				double num32 = 0.0;
				if ((bool)value || PeriodInfo.KieuKy == "Y")
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					for (int i = 0; i < rowsValues.Count; i++)
					{
						string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num23 - 1].Text;
						string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num25 - 1].Text;
						num31 += ConvertHelpers.ConvertToDouble1(text2);
						num32 += ConvertHelpers.ConvertToDouble1(text3);
					}
				}
				else
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					for (int i = 0; i < rowsValues.Count; i++)
					{
						string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num23 - 1].Text;
						string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num25 - 1].Text;
						string text4 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num19 - 1].Text;
						string text5 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num20 - 1].Text;
						DateTime dateTime = ConvertHelpers.ConvertToDateTime(text4, "dd/MM/yyyy");
						DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text5, "dd/MM/yyyy");
						DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
						DateTime dateTime4 = ConvertHelpers.ConvertToDateTime(PeriodInfo.DenNgay, "dd/MM/yyyy");
						if ((dateTime - dateTime3).TotalDays >= 0.0 && (dateTime2 - dateTime3).TotalDays >= 0.0 && (dateTime - dateTime4).TotalDays <= 0.0 && (dateTime2 - dateTime4).TotalDays <= 0.0)
						{
							num31 += ConvertHelpers.ConvertToDouble1(text2);
							num32 += ConvertHelpers.ConvertToDouble1(text3);
						}
					}
				}
				sheetView.Cells[53, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num31;
			}
			string nam = PeriodInfo.Nam;
			string tuNgay = PeriodInfo.TuNgay;
			DateTime dateTime5 = ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy");
			DateTime dateTime6 = new DateTime(2018, 8, 1);
			string kieuKy = PeriodInfo.KieuKy;
			int column = GeneralFunctions.ArrayCellId("U_58")[0];
			int row3 = GeneralFunctions.ArrayCellId("U_58")[1];
			int column2 = GeneralFunctions.ArrayCellId("U_59")[0];
			int row4 = GeneralFunctions.ArrayCellId("U_59")[1];
			if (kieuKy.Equals("Y"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018)
				{
					sheetView.Cells[row3, column].Value = "0";
					sheetView.Cells[row3, column].Locked = true;
					sheetView.Cells[row4, column2].Value = "0";
					sheetView.Cells[row4, column2].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && dateTime5 >= dateTime6))
			{
				sheetView.Cells[row3, column].Value = "0";
				sheetView.Cells[row3, column].Locked = true;
				sheetView.Cells[row4, column2].Value = "0";
				sheetView.Cells[row4, column2].Locked = true;
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
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCCellInfo> staticCells = hTKKCMCSheetInfo.StaticCells;
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("Y");
			int num4 = HTKKFile.ColumnLettersToIndex("Z");
			int num5 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol = HTKKFile.ColumnLettersToIndex("AC");
			int num6 = HTKKFile.ColumnLettersToIndex("AD");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				if (!hTKKCMCSheetInfo.IsReadDataFile && isFirtLoad)
				{
					sheetView.Cells[dataLocationTemplateRowStart - 1, num - 1].Text = "0.0";
				}
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text.Replace("/", "");
					string text2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text.Replace("/", "");
					string text3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text4 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					if (!HTKKFormater.isNullString(text3))
					{
						string key = text + text2;
						string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "NumberCellType");
						double[] newValue = new double[4]
						{
							ConvertHelpers.ConvertToDouble(text4),
							ConvertHelpers.ConvertToDouble(text3),
							ConvertHelpers.ConvertToDouble(valueCustom),
							1.0
						};
						AddOrUpdateDicArr(dicTongCT20a, key, newValue);
						AddOrUpdateDicArr(dicTongCT20, key, newValue);
					}
				}
			}
			foreach (HTKKCMCCellInfo item in staticCells)
			{
				if (item.Value == "VB6")
				{
					isDataVB6 = true;
					break;
				}
			}
			if (isDataVB6)
			{
				FinishImportXML(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = 49;
			string value = "";
			int num3 = 54;
			int num4 = 55;
			int num5 = 56;
			int num6 = 57;
			int num7 = 58;
			int num8 = 59;
			int num9 = 60;
			int num10 = 61;
			int num11 = 62;
			int num12 = 73;
			int num13 = 75;
			switch (actionType)
			{
			case 2:
			{
				winFpSpread.Sheets[0].Cells[48, num - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[48, num - 1].Text = "";
				int[] array = GeneralFunctions.ArrayCellId("U_54");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				array = GeneralFunctions.ArrayCellId("U_55");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				array = GeneralFunctions.ArrayCellId("U_56");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				array = GeneralFunctions.ArrayCellId("U_57");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				array = GeneralFunctions.ArrayCellId("U_59");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				array = GeneralFunctions.ArrayCellId("U_62");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				break;
			}
			case 1:
			{
				int[] array = GeneralFunctions.ArrayCellId("B_45");
				if (ConvertHelpers.ConvertToBool(winFpSpread.Sheets[0].Cells[array[1], array[0]].Value))
				{
					dicFormulaTK.TryGetValue("E_49", out value);
					winFpSpread.Sheets[0].Cells[48, num - 1].Formula = value.Replace("###", winFpSpread.Sheets[1].SheetName);
				}
				array = GeneralFunctions.ArrayCellId("U_54");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "AE54";
				array = GeneralFunctions.ArrayCellId("U_56");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "AE55";
				string formula = "ROUND(U" + num4 + "*0.05,0)";
				array = GeneralFunctions.ArrayCellId("U_56");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
				array = GeneralFunctions.ArrayCellId("U_57");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
				string formula2 = "IF(U" + num4 + "<>0,ROUND(U" + num6 + "*(U" + num7 + "/U" + num4 + ")*0.5,0),0)";
				array = GeneralFunctions.ArrayCellId("U_59");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula2;
				array = GeneralFunctions.ArrayCellId("U_62");
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "IF(U57-U59+U61>0,U57-U59+U61,0)";
				break;
			}
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
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("H");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("N");
			int num7 = HTKKFile.ColumnLettersToIndex("P");
			int num8 = HTKKFile.ColumnLettersToIndex("S");
			int num9 = HTKKFile.ColumnLettersToIndex("T");
			int num10 = HTKKFile.ColumnLettersToIndex("U");
			int num11 = HTKKFile.ColumnLettersToIndex("V");
			int num12 = HTKKFile.ColumnLettersToIndex("W");
			int num13 = HTKKFile.ColumnLettersToIndex("X");
			int num14 = HTKKFile.ColumnLettersToIndex("Y");
			int num15 = HTKKFile.ColumnLettersToIndex("Z");
			int num16 = HTKKFile.ColumnLettersToIndex("AA");
			int num17 = HTKKFile.ColumnLettersToIndex("AB");
			int num18 = HTKKFile.ColumnLettersToIndex("AC");
			int num19 = HTKKFile.ColumnLettersToIndex("AD");
			int num20 = HTKKFile.ColumnLettersToIndex("AE");
			int num21 = HTKKFile.ColumnLettersToIndex("AF");
			int num22 = HTKKFile.ColumnLettersToIndex("AG");
			int num23 = HTKKFile.ColumnLettersToIndex("AI");
			int num24 = HTKKFile.ColumnLettersToIndex("AJ");
			if (iCol == num)
			{
				string[] value = null;
				dicDMHopDong.TryGetValue(sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim(), out value);
				if (value != null && value.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num - 2].Text = value[0];
				}
				string[] value2 = null;
				dicDMTTHopDong.TryGetValue(sheetView.Cells[(int)iRow - 1, num23 - 1].Text.Trim(), out value2);
				if (value2 != null && value2[0] == "01")
				{
					string text = sheetView.Cells[(int)iRow - 1, num14 - 1].Text.Replace("/", "");
					string text2 = sheetView.Cells[(int)iRow - 1, num - 2].Text;
					sheetView.Cells[(int)iRow - 1, num24 - 1].Text = TTinChung.maCQTNoiNop + "_" + TTinChung.mst + "000_" + text + "_" + text2;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num24 - 1].Text = "";
				}
				CellChange_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, (int)iRow, (int)iCol);
			}
			if (iCol == num7)
			{
				string[] value = dicDMMucDich[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (value.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num7 - 2].Text = value[0];
				}
			}
			if (iCol == num23)
			{
				string[] value = dicDMTTHopDong[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (value.Length == 2)
				{
					sheetView.Cells[(int)iRow - 1, num23 - 2].Text = value[0];
				}
				if (value[0] == "01")
				{
					string text = sheetView.Cells[(int)iRow - 1, num14 - 1].Text.Replace("/", "");
					string text2 = sheetView.Cells[(int)iRow - 1, num - 2].Text;
					sheetView.Cells[(int)iRow - 1, num24 - 1].Text = TTinChung.maCQTNoiNop + "_" + TTinChung.mst + "000_" + text + "_" + text2;
					int num25 = 44;
					int num26 = HTKKFile.ColumnLettersToIndex("B") - 1;
					int[] array = GeneralFunctions.ArrayCellId("E_49");
					int[] array2 = GeneralFunctions.ArrayCellId("AJ_22");
					if (winFpSpread.Sheets[1].Visible)
					{
						string value3 = "";
						dicFormulaTK.TryGetValue("E_49", out value3);
						winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = value3.Replace("###", winFpSpread.Sheets[1].SheetName);
					}
					else
					{
						winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = "";
						winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = "";
					}
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num24 - 1].Text = "";
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num27 = HTKKFile.ColumnLettersToIndex("G");
			int num28 = HTKKFile.ColumnLettersToIndex("B");
			if (iCol != num27)
			{
				return;
			}
			string text3 = "";
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					object value4 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num27 - 1].Value;
					if (j == 0)
					{
						if ((bool)value4)
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].HorizontalAlignment = CellHorizontalAlignment.Center;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].Text = "1";
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].Text = "0.0";
						}
						continue;
					}
					string text4;
					if ((bool)value4)
					{
						text4 = sheetView.Cells[dataLocationTemplateRowStart + j - 2, num28 - 1].Text;
						if (text4.IndexOf(".") > 0)
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].Text = (ConvertHelpers.ConvertToInt(text4.Split(new char[1] { '.' })[0]) + 1).ToString();
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].HorizontalAlignment = CellHorizontalAlignment.Center;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].Text = (ConvertHelpers.ConvertToInt(text4) + 1).ToString();
						}
						continue;
					}
					text4 = sheetView.Cells[dataLocationTemplateRowStart + j - 2, num28 - 1].Text;
					if (text4.IndexOf(".") > 0)
					{
						sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
						sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].Text = text4.Split(new char[1] { '.' })[1] + "." + (ConvertHelpers.ConvertToInt(text4.Split(new char[1] { '.' })[1]) + 1);
					}
					else
					{
						sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
						sheetView.Cells[dataLocationTemplateRowStart + j - 1, num28 - 1].Text = ConvertHelpers.ConvertToInt(text4) + 1 + ".0";
					}
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
			int num = HTKKFile.ColumnLettersToIndex("I");
			int num2 = HTKKFile.ColumnLettersToIndex("U");
			int num3 = 54;
			int num4 = 55;
			int num5 = 56;
			int num6 = 57;
			int num7 = 58;
			int num8 = 59;
			int num9 = 60;
			int num10 = 61;
			int num11 = 62;
			int num12 = 73;
			int num13 = 75;
			if (iCol == num2 && (iRow == num4 || iRow == num6 || iRow == num7))
			{
				string formula = "IF(U" + num4 + "<>0,ROUND(U" + num6 + "*(U" + num7 + "/U" + num4 + ")*0.5,0),0)";
				sheetView.Cells[num8 - 1, num2 - 1].Formula = formula;
			}
			if (iCol == num2 && iRow == num4)
			{
				string formula2 = "ROUND(U" + num4 + "*0.05,0)";
				sheetView.Cells[num5 - 1, num2 - 1].Formula = formula2;
				sheetView.Cells[num6 - 1, num2 - 1].Formula = formula2;
			}
			if (iCol == num2 && iRow != num3 && iRow != num9 && iRow >= 54 && iRow <= 62)
			{
				sheetView.Cells[num11 - 1, num2 - 1].Formula = "IF(U57-U59+U61>0,U57-U59+U61,0)";
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
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("H");
			int num5 = HTKKFile.ColumnLettersToIndex("I");
			int num6 = HTKKFile.ColumnLettersToIndex("J");
			int num7 = HTKKFile.ColumnLettersToIndex("K");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("P");
			int num10 = HTKKFile.ColumnLettersToIndex("S");
			int num11 = HTKKFile.ColumnLettersToIndex("T");
			int num12 = HTKKFile.ColumnLettersToIndex("U");
			int num13 = HTKKFile.ColumnLettersToIndex("V");
			int num14 = HTKKFile.ColumnLettersToIndex("W");
			int num15 = HTKKFile.ColumnLettersToIndex("X");
			int num16 = HTKKFile.ColumnLettersToIndex("Y");
			int num17 = HTKKFile.ColumnLettersToIndex("Z");
			int num18 = HTKKFile.ColumnLettersToIndex("AA");
			int num19 = HTKKFile.ColumnLettersToIndex("AB");
			int num20 = HTKKFile.ColumnLettersToIndex("AC");
			int num21 = HTKKFile.ColumnLettersToIndex("AD");
			int num22 = HTKKFile.ColumnLettersToIndex("AE");
			int num23 = HTKKFile.ColumnLettersToIndex("AF");
			int num24 = HTKKFile.ColumnLettersToIndex("AG");
			int num25 = HTKKFile.ColumnLettersToIndex("AI");
			int num26 = HTKKFile.ColumnLettersToIndex("AJ");
			int num27 = HTKKFile.ColumnLettersToIndex("AL");
			string text = sheetView.Cells[(int)iRow - 1, num16 - 1].Text;
			string text2 = sheetView.Cells[(int)iRow - 1, num17 - 1].Text;
			string text3 = sheetView.Cells[(int)iRow - 1, num18 - 1].Text;
			string text4 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text;
			string text5 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num27, "0");
			double soThang = 0.0;
			string key = text5 + text.Replace("/", "") + text2.Replace("/", "");
			if (iCol == num20 && !HTKKFormater.isNullString(text3))
			{
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num20, "NumberCellType");
				double[] newValue = new double[4]
				{
					ConvertHelpers.ConvertToDouble(text5),
					ConvertHelpers.ConvertToDouble(text3),
					ConvertHelpers.ConvertToDouble(valueCustom2),
					1.0
				};
				AddOrUpdateDicArr(dicTongCT20a, key, newValue);
			}
			if (iCol == num21 && !HTKKFormater.isNullString(text3))
			{
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num21, "NumberCellType");
				double[] newValue2 = new double[4]
				{
					ConvertHelpers.ConvertToDouble(text5),
					ConvertHelpers.ConvertToDouble(text3),
					ConvertHelpers.ConvertToDouble(valueCustom2),
					1.0
				};
				AddOrUpdateDicArr(dicTongCT20, key, newValue2);
			}
			if (iCol == num16)
			{
				string[] value = null;
				dicDMTTHopDong.TryGetValue(sheetView.Cells[(int)iRow - 1, num25 - 1].Text.Trim(), out value);
				if (value != null)
				{
					if (value[0] == "01")
					{
						string text6 = sheetView.Cells[(int)iRow - 1, num16 - 1].Text.Replace("/", "");
						string text7 = sheetView.Cells[(int)iRow - 1, num2 - 2].Text;
						sheetView.Cells[(int)iRow - 1, num26 - 1].Text = TTinChung.maCQTNoiNop + "_" + TTinChung.mst + "000_" + text6 + "_" + text7;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num26].Text = "";
					}
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num26].Text = "";
				}
				string text8 = text.Substring(text.LastIndexOf("/") + 1);
				string obj = text2.Substring(text2.LastIndexOf("/") + 1);
				if (ConvertHelpers.ConvertToInt(obj) > 0 && ConvertHelpers.ConvertToInt(obj) - ConvertHelpers.ConvertToInt(text8) == 0)
				{
					sheetView.Cells[(int)iRow - 1, num18 - 1].Text = text8;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num18 - 1].Text = "";
				}
				if (!HTKKFormater.isNullString(valueCustom))
				{
					HTKKFile.AddOrUpdateDicStr(dicHDXoa, valueCustom, "1");
				}
			}
			else if (iCol == num17)
			{
				string text8 = text.Substring(text.LastIndexOf("/") + 1);
				string obj = text2.Substring(text2.LastIndexOf("/") + 1);
				if (ConvertHelpers.ConvertToInt(obj) - ConvertHelpers.ConvertToInt(text8) == 0)
				{
					sheetView.Cells[(int)iRow - 1, num18 - 1].Text = text8;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num18 - 1].Text = "";
				}
				if (!HTKKFormater.isNullString(valueCustom))
				{
					HTKKFile.AddOrUpdateDicStr(dicHDXoa, valueCustom, "1");
				}
			}
			if (iCol == num16 || iCol == num17 || iCol == num19 || iCol == num3 || iCol == num2)
			{
				string text9 = text.Substring(text.LastIndexOf("/") + 1);
				string text10 = text2.Substring(text2.LastIndexOf("/") + 1);
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy");
				dateNamDau = dateTime;
				dateNamCuoi = dateTime2;
				if (valueCustom.IndexOf("_") > -1)
				{
					string[] array = valueCustom.Split(new char[1] { '_' });
					for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
					{
						if (GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + i, num27, "0").StartsWith(array[0] + "."))
						{
							dateNamDau = ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num16 - 1].Text, "dd/MM/yyyy");
							dateNamCuoi = ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num17 - 1].Text, "dd/MM/yyyy");
							break;
						}
					}
				}
				double num28 = 0.0;
				double num29 = 0.0;
				double totalDays = (dateTime2 - dateTime).TotalDays;
				DateTime dateTime3 = dateTime.AddDays(totalDays);
				double totalDays2 = (dateTime2 - dateTime3).TotalDays;
				num28 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num19, "NumberCellType"));
				num29 = getValueCot20a(dateTime, dateTime2, dateNamDau, dateNamCuoi, num28, valueCustom, out soThang);
				GetDicCT20a(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
				Dictionary<string, double> dictionary = new Dictionary<string, double>();
				double[] value3;
				double value4;
				foreach (KeyValuePair<string, double[]> item in dicTongCT20a)
				{
					bool flag = false;
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						string text11 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text;
						string text12 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text;
						string text13 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
						string value2 = text13 + text11.Replace("/", "") + text12.Replace("/", "");
						if (item.Key.IndexOf(value2) > -1)
						{
							flag = true;
							break;
						}
					}
					if (flag)
					{
						value3 = item.Value;
						if (value3.Length == 4 && value3[3] != 1.0)
						{
							string key2 = value3[0] + "#" + value3[1];
							dictionary.TryGetValue(key2, out value4);
							AddOrUpdateDicStrDouble(dictionary, key2, value4 + value3[2]);
						}
					}
				}
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string text13 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
					double num30 = ConvertHelpers.ConvertToDouble(text13);
					string text14 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Text;
					string text11 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text;
					string text12 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text;
					string value2 = text13 + text11.Replace("/", "") + text12.Replace("/", "");
					string key2 = num30 + "#" + text14;
					dictionary.TryGetValue(key2, out value4);
					dicTongCT20a.TryGetValue(value2, out value3);
					if ((value3 == null || ConvertHelpers.ConvertToDouble(value3[3]) != 1.0) && value4 > 0.0)
					{
						sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num20 - 1].Value = value4;
					}
				}
				Dictionary<string, double> dictionary2 = new Dictionary<string, double>();
				double[] value5;
				double value6;
				foreach (KeyValuePair<string, double[]> item2 in dicTongCT20a)
				{
					bool flag = false;
					for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
					{
						string text11 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text;
						string text12 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text;
						string text13 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
						string value2 = text11.Replace("/", "") + text12.Replace("/", "");
						if (item2.Key.IndexOf(value2) > -1)
						{
							flag = true;
							break;
						}
					}
					if (flag)
					{
						value5 = item2.Value;
						if (value5.Length == 4 && value5[3] != 1.0)
						{
							string key2 = value5[1].ToString();
							dictionary2.TryGetValue(key2, out value6);
							AddOrUpdateDicStrDouble(dictionary2, key2, value6 + value5[2]);
						}
					}
				}
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string text13 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
					double num30 = ConvertHelpers.ConvertToDouble(text13);
					string text14 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Text;
					string text11 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text;
					string text12 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text;
					string value2 = text13 + text11.Replace("/", "") + text12.Replace("/", "");
					string key2 = text14;
					dictionary2.TryGetValue(key2, out value6);
					dicTongCT20.TryGetValue(value2, out value5);
					if ((value5 == null || ConvertHelpers.ConvertToDouble(value5[3]) != 1.0) && value6 > 0.0)
					{
						sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num21 - 1].Value = value6;
					}
				}
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AK") - 1].Value = soThang;
				if (valueCustom.ToUpper() == "X")
				{
					dateNamDau = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
					dateNamCuoi = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy");
				}
				double num31 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num21, "NumberCellType"));
				if (valueCustom.IndexOf("_") > -1)
				{
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = 0;
					string text15 = "IF(AD###<=100000000,0,AB###*AK###)";
					sheetView.Cells[(int)iRow - 1, num22 - 1].Formula = text15.Replace("###", iRow.ToString());
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Formula = text15.Replace("###", iRow.ToString());
					if (num31 < 100000000.0)
					{
						sheetView.Cells[(int)iRow - 1, num22 - 1].Value = 0;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = 0;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num22 - 1].Value = Math.Round(num29, 0);
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[(int)iRow - 1, num22 - 1].Value;
					}
					string[] array2 = valueCustom.Split(new char[1] { '_' });
					int num32 = (int)iRow - ConvertHelpers.ConvertToInt(array2[1]);
					string[] array3 = null;
					if (!HTKKFormater.isNullString(sheetView.Cells[num32 - 1, num27 - 1].Text) && sheetView.Cells[num32 - 1, num27 - 1].Value.ToString().IndexOf(".") > 0)
					{
						array3 = sheetView.Cells[num32 - 1, num27 - 1].Value.ToString().Split(new char[1] { '.' });
					}
					if (array3 != null)
					{
						int num33 = ConvertHelpers.ConvertToInt(array3[1]);
						HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
						HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
						HTKKCMCDataLocation location = hTKKCMCTableInfo2.Location;
						int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
						sheetView.Cells[num32 - 1, num22 - 1].Value = 0;
						double num34 = 0.0;
						while (num33 > 0)
						{
							num34 += ConvertHelpers.ConvertToDouble(sheetView.Cells[num32 + num33 - 1, num22 - 1].Value);
							sheetView.Cells[num32 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[num32 - 1, num22 - 1].Value;
							num33--;
						}
						sheetView.Cells[num32 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = Math.Round(num34, 0);
						sheetView.Cells[num32 - 1, num22 - 1].Value = Math.Round(num34, 0);
					}
				}
				else if (valueCustom.IndexOf(".") > -1)
				{
					dateNamDau = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
					dateNamCuoi = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = (dateTime2.Year - dateTime.Year) * 12 + dateTime2.Month - dateTime.Month + ((Math.Abs(dateTime2.Day - dateTime.Day) >= 15) ? 1 : 0);
					sheetView.Cells[(int)iRow - 1, num14 - 1].Value = num28 * soThang;
					sheetView.Cells[(int)iRow - 1, num22 - 1].Formula = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Formula = "";
					if (iCol == num19)
					{
						string[] array4 = valueCustom.Split(new char[1] { '.' });
						int num35 = ConvertHelpers.ConvertToInt(array4[1]);
						HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
						HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo2.Tables[0];
						HTKKCMCDataLocation location2 = hTKKCMCTableInfo3.Location;
						int dataLocationTemplateRowStart3 = location2.DataLocationTemplateRowStart;
						while (num35 > 0)
						{
							sheetView.Cells[(int)iRow + num35 - 1, num14 - 1].Value = sheetView.Cells[(int)iRow - 1, num14 - 1].Value;
							sheetView.Cells[(int)iRow + num35 - 1, num19 - 1].Value = sheetView.Cells[(int)iRow - 1, num19 - 1].Value;
							CellChange_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow + num35, num17);
							num35--;
						}
					}
					string[] array2 = valueCustom.Split(new char[1] { '.' });
					int num33 = ConvertHelpers.ConvertToInt(array2[1]);
					HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
					HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
					HTKKCMCDataLocation location = hTKKCMCTableInfo2.Location;
					int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
					sheetView.Cells[(int)iRow - 1, num22 - 1].Value = 0;
					double num34 = 0.0;
					while (num33 > 0)
					{
						num34 += ConvertHelpers.ConvertToDouble(sheetView.Cells[(int)iRow + num33 - 1, num22 - 1].Value);
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[(int)iRow - 1, num22 - 1].Value;
						num33--;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = Math.Round(num34, 0);
					sheetView.Cells[(int)iRow - 1, num22 - 1].Value = Math.Round(num34, 0);
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = (dateTime2.Year - dateTime.Year) * 12 + dateTime2.Month - dateTime.Month + ((Math.Abs(dateTime2.Day - dateTime.Day) >= 15) ? 1 : 0);
					if (!ConvertHelpers.ConvertToBool(sheetView.Cells[(int)iRow - 1, num3 - 1].Value))
					{
						sheetView.Cells[(int)iRow - 1, num14 - 1].Value = num28 * soThang;
					}
					sheetView.Cells[(int)iRow - 1, num22 - 1].Formula = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Formula = "";
					if (num31 < 100000000.0)
					{
						sheetView.Cells[(int)iRow - 1, num22 - 1].Value = 0;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = 0;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num22 - 1].Value = Math.Round(num29, 0);
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = Math.Round(num29, 0);
					}
				}
				if (HTKKFormater.isNullString(text3) || Convert.ToBoolean(text4))
				{
					sheetView.Cells[(int)iRow - 1, num20 - 1].Value = 0;
					num29 = 0.0;
				}
				if (!ConvertHelpers.ConvertToBool(sheetView.Cells[(int)iRow - 1, num3 - 1].Value))
				{
					DateTime dateTime4 = new DateTime(1990, 1, 1);
					DateTime dateTime5 = new DateTime(1990, 1, 1);
					DateTime dateTime6 = new DateTime(1990, 1, 1);
					double num36 = 0.0;
					int num37 = 0;
					string text16 = "";
					string text17 = "";
					HTKKCMCTableInfo hTKKCMCTableInfo4 = ToKhaiInfo.SheetInfos[1].Tables[0];
					int dataLocationTemplateRowStart4 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart;
					int num38 = 0;
					int num39 = 0;
					string text18 = "";
					isLienTiep = true;
					Dictionary<int, int> dictionary3 = new Dictionary<int, int>();
					for (int j = num38; j < hTKKCMCTableInfo4.RowsValues.Count; j++)
					{
						num39 = j;
						if (Convert.ToBoolean(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num3 - 1].Value))
						{
							num38++;
							dictionary3.Add(num38, num39);
						}
					}
					foreach (KeyValuePair<int, int> item3 in dictionary3)
					{
						int value7 = 0;
						dictionary3.TryGetValue(item3.Key + 1, out value7);
						for (int j = item3.Value; j < ((value7 > 0) ? value7 : hTKKCMCTableInfo4.RowsValues.Count); j++)
						{
							if (Convert.ToBoolean(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num3 - 1].Value))
							{
								dateTime4 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num16 - 1].Text, "dd/MM/yyyy");
								dateTime5 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num17 - 1].Text, "dd/MM/yyyy");
							}
						}
						num36 = 0.0;
						if (dateTime4.Year > 1990 && dateTime5.Year > 1990)
						{
							for (int j = item3.Value; j < ((value7 > 0) ? value7 : hTKKCMCTableInfo4.RowsValues.Count); j++)
							{
								string text19 = sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num27 - 1].Text;
								if ((HTKKFormater.isNullString(text19) || (!HTKKFormater.isNullString(text19) && text19.IndexOf("_") > 0)) && !Convert.ToBoolean(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num3 - 1].Value) && ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num16 - 1].Text, "dd/MM/yyyy").Year >= dateTime4.Year && ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num17 - 1].Text, "dd/MM/yyyy").Year <= dateTime5.Year)
								{
									num36 += ConvertHelpers.ConvertToDouble(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, HTKKFile.ColumnLettersToIndex("AK") - 1].Value) * ConvertHelpers.ConvertToDouble(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num19 - 1].Value);
								}
							}
						}
						for (int j = item3.Value; j < ((value7 > 0) ? value7 : hTKKCMCTableInfo4.RowsValues.Count); j++)
						{
							if (ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num16 - 1].Text, "dd/MM/yyyy").Year >= dateTime4.Year && ConvertHelpers.ConvertToDateTime(sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num17 - 1].Text, "dd/MM/yyyy").Year <= dateTime5.Year)
							{
								sheetView.Cells[dataLocationTemplateRowStart4 + j - 1, num14 - 1].Value = Math.Round(num36, 0);
							}
						}
					}
				}
			}
			if (iCol == num21)
			{
				string text20 = "IF(LEN(AL###)>0,IF(OR(MID(AL###,2,1)=\".\",MID(AL###,3,1)=\".\"),AS###,IF(AC###<=100000000,0,AB###*AK###)),IF(LEN(AA###)>0,IF(AD###<=100000000,0,AB###*AK###),0))";
				text20 = text20.Replace("###", iRow.ToString());
				sheetView.Cells[(int)iRow - 1, num22 - 1].Formula = text20;
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, HTKKFile.ColumnLettersToIndex("AK"), "NumberCellType");
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num19, "NumberCellType");
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num21, "NumberCellType");
				if (valueCustom.IndexOf("_") > -1)
				{
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = 0;
					sheetView.Cells[(int)iRow - 1, num22 - 1].Formula = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Formula = "";
					if (ConvertHelpers.ConvertToDouble(valueCustom5) < 100000000.0)
					{
						sheetView.Cells[(int)iRow - 1, num22 - 1].Value = 0;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = 0;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num22 - 1].Value = Math.Round(ConvertHelpers.ConvertToDouble(valueCustom3) * ConvertHelpers.ConvertToDouble(valueCustom4), 0);
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[(int)iRow - 1, num22 - 1].Value;
					}
					string[] array2 = valueCustom.Split(new char[1] { '_' });
					int num32 = (int)iRow - ConvertHelpers.ConvertToInt(array2[1]);
					string[] array3 = null;
					if (!HTKKFormater.isNullString(sheetView.Cells[num32 - 1, num27 - 1].Text) && sheetView.Cells[num32 - 1, num27 - 1].Value.ToString().IndexOf(".") > 0)
					{
						array3 = sheetView.Cells[num32 - 1, num27 - 1].Value.ToString().Split(new char[1] { '.' });
					}
					if (array3 != null)
					{
						int num33 = ConvertHelpers.ConvertToInt(array3[1]);
						HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
						HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
						HTKKCMCDataLocation location = hTKKCMCTableInfo2.Location;
						int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
						sheetView.Cells[num32 - 1, num22 - 1].Value = 0;
						double num34 = 0.0;
						while (num33 > 0)
						{
							num34 += ConvertHelpers.ConvertToDouble(sheetView.Cells[num32 + num33 - 1, num22 - 1].Value);
							sheetView.Cells[num32 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[num32 - 1, num22 - 1].Value;
							num33--;
						}
						sheetView.Cells[num32 - 1, num22 - 1].Formula = "";
						sheetView.Cells[num32 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Formula = "";
						sheetView.Cells[num32 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = Math.Round(num34, 0);
						sheetView.Cells[num32 - 1, num22 - 1].Value = Math.Round(num34, 0);
					}
				}
			}
			if (HTKKFormater.isNullString(valueCustom) && HTKKFormater.isNullString(sheetView.Cells[(int)iRow - 1, num18 - 1].Text))
			{
				sheetView.Cells[(int)iRow - 1, num22 - 1].Value = 0;
			}
			if (PeriodInfo.KieuKy.Equals("Q"))
			{
				int[] array5 = GeneralFunctions.ArrayCellId("B_47");
				bool flag2 = ConvertHelpers.ConvertToBool(winFpSpread.Sheets[0].Cells[array5[1], array5[0]].Value);
				double num40 = 0.0;
				double num41 = 0.0;
				Dictionary<string, double> dictionary4 = new Dictionary<string, double>();
				Dictionary<string, double> dictionary5 = new Dictionary<string, double>();
				bool flag3 = false;
				bool flag4 = false;
				bool flag5 = false;
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime(PeriodInfo.DenNgay, "dd/MM/yyyy");
				DateTime dateTime9 = dateNamDau;
				DateTime dateTime10 = dateNamCuoi;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num27, "0");
					string text21 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text14 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Text;
					DateTime dateTime11 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text, "dd/MM/yyyy");
					DateTime dateTime12 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text, "dd/MM/yyyy");
					double num42 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num20, "NumberCellType"));
					double valueCot = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num19, "NumberCellType"));
					double newValue3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num22, "NumberCellType"));
					double soThang2 = 0.0;
					double value8;
					if (!HTKKFormater.isNullString(valueCustom6) && valueCustom6.IndexOf(".") > -1)
					{
						dateTime9 = dateTime11;
						dateTime10 = dateTime12;
						value8 = 0.0;
						dictionary4.TryGetValue(text21, out value8);
						if (value8 == 0.0)
						{
							flag4 = false;
						}
					}
					if ((!HTKKFormater.isNullString(valueCustom6) && valueCustom6.IndexOf("_") == -1) || HTKKFormater.isNullString(text14) || (!HTKKFormater.isNullString(text14) && !text14.Equals(dateTime7.ToString("yyyy"))))
					{
						continue;
					}
					if (DateTime.Compare(dateTime11, dateTime7) <= 0 && DateTime.Compare(dateTime8, dateTime12) <= 0)
					{
						if (!flag4)
						{
							double valueCot20a = getValueCot20a(dateTime7, dateTime8, dateTime9, dateTime10, valueCot, valueCustom6, out soThang2);
							AddOrUpdateDicStrDouble(dictionary4, text21, valueCot20a);
							flag4 = true;
						}
					}
					else if (DateTime.Compare(dateTime11, dateTime7) == 0 && DateTime.Compare(dateTime8, dateTime12) < 0 && !flag4)
					{
						double valueCot20a = getValueCot20a(dateTime11, dateTime8, dateTime9, dateTime10, valueCot, valueCustom6, out soThang2);
						value8 = 0.0;
						dictionary4.TryGetValue(text21, out value8);
						AddOrUpdateDicStrDouble(dictionary4, text21, value8 + num42);
					}
					else if (DateTime.Compare(dateTime11, dateTime7) < 0 && DateTime.Compare(dateTime8, dateTime12) == 0 && !flag4)
					{
						double valueCot20a = getValueCot20a(dateTime7, dateTime12, dateTime9, dateTime10, valueCot, valueCustom6, out soThang2);
						value8 = 0.0;
						dictionary4.TryGetValue(text21, out value8);
						AddOrUpdateDicStrDouble(dictionary4, text21, value8 + num42);
					}
					else if (DateTime.Compare(dateTime11, dateTime7) < 0 && DateTime.Compare(dateTime8, dateTime12) < 0 && !flag4)
					{
						double valueCot20a = getValueCot20a(dateTime7, dateTime8, dateTime9, dateTime10, valueCot, valueCustom6, out soThang2);
						value8 = 0.0;
						dictionary4.TryGetValue(text21, out value8);
						AddOrUpdateDicStrDouble(dictionary4, text21, value8 + num42);
					}
					dictionary5.TryGetValue(text21, out value8);
					if (DateTime.Compare(dateTime11, dateTime7) == 0 && DateTime.Compare(dateTime8, dateTime12) == 0 && value8 == 0.0)
					{
						value8 = 0.0;
						dictionary5.TryGetValue(text21, out value8);
						AddOrUpdateDicStrDouble(dictionary5, text21, newValue3);
					}
				}
				foreach (KeyValuePair<string, double> item4 in dictionary4)
				{
					num40 = item4.Value + num40;
				}
				foreach (KeyValuePair<string, double> item5 in dictionary5)
				{
					num41 = item5.Value + num41;
				}
				string text22 = "AE_54";
				string text23 = "U_54";
				string[] array6 = text23.Split(new char[1] { '_' });
				string[] array7 = text22.Split(new char[1] { '_' });
				string value9 = "";
				dicFormulaTK.TryGetValue(text23, out value9);
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Value = num40;
				if (!flag2)
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[0]) - 1].Formula = "AE54";
				}
				text22 = "AE_55";
				text23 = "U_55";
				array6 = text23.Split(new char[1] { '_' });
				array7 = text22.Split(new char[1] { '_' });
				value9 = "";
				dicFormulaTK.TryGetValue(text23, out value9);
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Value = num41;
				if (!flag2)
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[0]) - 1].Formula = "AE55";
				}
				dictionary4 = new Dictionary<string, double>();
				dictionary5 = new Dictionary<string, double>();
				num40 = 0.0;
				num41 = 0.0;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num27, "0");
					string text21 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text14 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Text;
					DateTime dateTime11 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text, "dd/MM/yyyy");
					DateTime dateTime12 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text, "dd/MM/yyyy");
					double num42 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num20, "NumberCellType"));
					double valueCot = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num19, "NumberCellType"));
					double newValue3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num22, "NumberCellType"));
					double soThang2 = 0.0;
					if (!HTKKFormater.isNullString(text14))
					{
						AddOrUpdateDicStrDouble(dictionary4, text14, num42);
						if (HTKKFormater.isNullString(valueCustom6))
						{
							num41 += newValue3;
						}
					}
					else
					{
						num41 += newValue3;
					}
				}
				foreach (KeyValuePair<string, double> item6 in dictionary4)
				{
					num40 = item6.Value + num40;
				}
				dictionary4 = new Dictionary<string, double>();
				dictionary5 = new Dictionary<string, double>();
				text22 = "AF_54";
				text23 = "U_54";
				array6 = text23.Split(new char[1] { '_' });
				array7 = text22.Split(new char[1] { '_' });
				value9 = "";
				dicFormulaTK.TryGetValue(text23, out value9);
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Value = num40;
				if (flag2)
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[0]) - 1].Formula = "AF54";
				}
				text22 = "AF_55";
				text23 = "U_55";
				array6 = text23.Split(new char[1] { '_' });
				array7 = text22.Split(new char[1] { '_' });
				value9 = "";
				dicFormulaTK.TryGetValue(text23, out value9);
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Value = num41;
				if (flag2)
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[0]) - 1].Formula = "AF55";
				}
			}
			else
			{
				Dictionary<string, double> dictionary4 = new Dictionary<string, double>();
				Dictionary<string, double> dictionary5 = new Dictionary<string, double>();
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime(PeriodInfo.DenNgay, "dd/MM/yyyy");
				double num40 = 0.0;
				double num41 = 0.0;
				string nam = PeriodInfo.Nam;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num27, "0");
					string text21 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text14 = sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Text;
					DateTime dateTime11 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Text, "dd/MM/yyyy");
					DateTime dateTime12 = ConvertHelpers.ConvertToDateTime(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text, "dd/MM/yyyy");
					double num42 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num20, "NumberCellType"));
					double valueCot = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num19, "NumberCellType"));
					double newValue3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + j, num22, "NumberCellType"));
					double soThang2 = 0.0;
					if (!HTKKFormater.isNullString(text14) && (HTKKFormater.isNullString(text14) || text14.Equals(nam)))
					{
						if (!HTKKFormater.isNullString(text14))
						{
							AddOrUpdateDicStrDouble(dictionary4, text21, num42);
						}
						num41 = newValue3 + num41;
					}
				}
				foreach (KeyValuePair<string, double> item7 in dictionary4)
				{
					num40 = item7.Value + num40;
				}
				string text22 = "AE_54";
				string text23 = "U_54";
				string[] array6 = text23.Split(new char[1] { '_' });
				string[] array7 = text22.Split(new char[1] { '_' });
				string value9 = "";
				dicFormulaTK.TryGetValue(text23, out value9);
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Value = num40;
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[0]) - 1].Formula = "AE54";
				text22 = "AE_55";
				text23 = "U_55";
				array6 = text23.Split(new char[1] { '_' });
				array7 = text22.Split(new char[1] { '_' });
				value9 = "";
				dicFormulaTK.TryGetValue(text23, out value9);
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Formula = "";
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[0]) - 1].Value = num41;
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[0]) - 1].Formula = "AE55";
			}
			string text24 = "U_56";
			string[] array8 = text24.Split(new char[1] { '_' });
			winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array8[1]) - 1, HTKKFile.ColumnLettersToIndex(array8[0]) - 1].Formula = "ROUND(U55*0.05,0)";
			text24 = "U_57";
			array8 = text24.Split(new char[1] { '_' });
			winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array8[1]) - 1, HTKKFile.ColumnLettersToIndex(array8[0]) - 1].Formula = "ROUND(U55*0.05,0)";
			text24 = "U_59";
			array8 = text24.Split(new char[1] { '_' });
			winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array8[1]) - 1, HTKKFile.ColumnLettersToIndex(array8[0]) - 1].Formula = "IF(U55<>0,ROUND(U57*(U58/U55)*0.5,0),0)";
			text24 = "U_62";
			array8 = text24.Split(new char[1] { '_' });
			winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array8[1]) - 1, HTKKFile.ColumnLettersToIndex(array8[0]) - 1].Formula = "IF(U57-U59+U61>0,U57-U59+U61,0)";
			int[] array9 = GeneralFunctions.ArrayCellId("E_49");
			int[] array10 = GeneralFunctions.ArrayCellId("AJ_22");
			if (ConvertHelpers.ConvertToBool(winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("B") - 1].Value) && winFpSpread.Sheets[1].Visible)
			{
				string value10 = "";
				dicFormulaTK.TryGetValue("E_49", out value10);
				winFpSpread.Sheets[0].Cells[array9[1], array9[0]].Formula = value10.Replace("###", winFpSpread.Sheets[1].SheetName);
			}
			else
			{
				winFpSpread.Sheets[0].Cells[array9[1], array9[0]].Formula = "";
			}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = 47;
			if (iCol == num && iRow != num2)
			{
			}
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
		if (ToKhaiInfo.SheetInfos[1] != null)
		{
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		}
		if (ToKhaiInfo.SheetInfos[0] != null)
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		}
		bool result = true;
		if (winFpSpread.TabStrip[0].BackColor != Color.White && winFpSpread.TabStrip[1].BackColor != Color.White)
		{
			result = false;
		}
		foreach (KeyValuePair<int, int> item in levelWarning)
		{
			int key = item.Key;
			int value = item.Value;
			if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[key].BackColor = Color.Pink;
			}
			else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
			{
				result = false;
				winFpSpread.TabStrip[key].BackColor = Color.Yellow;
			}
			else
			{
				result = false;
				winFpSpread.TabStrip[key].BackColor = Color.White;
			}
		}
		return result;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = 45;
			int num3 = 47;
			int num4 = 0;
			HTKKFile.RemoveValidate(sheetView, 22, HTKKFile.ColumnLettersToIndex("J") - 1);
			HTKKFile.RemoveValidate(sheetView, 24, HTKKFile.ColumnLettersToIndex("J") - 1);
			HTKKFile.RemoveValidate(sheetView, 26, HTKKFile.ColumnLettersToIndex("K") - 1);
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, 23, HTKKFile.ColumnLettersToIndex("J"), "0");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, 25, HTKKFile.ColumnLettersToIndex("J"), "0");
			string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, 27, HTKKFile.ColumnLettersToIndex("F"), "0");
			string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, 29, HTKKFile.ColumnLettersToIndex("F"), "0");
			if (!HTKKFormater.isNullString(valueCustom))
			{
				if (HTKKFormater.isNullString(valueCustom2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 24, HTKKFile.ColumnLettersToIndex("J") - 1, text3);
					num4 = 2;
				}
				if (HTKKFormater.isNullString(valueCustom3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 26, HTKKFile.ColumnLettersToIndex("F") - 1, text3);
					num4 = 2;
				}
			}
			else if (!HTKKFormater.isNullString(valueCustom2))
			{
				if (HTKKFormater.isNullString(valueCustom))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 22, HTKKFile.ColumnLettersToIndex("J") - 1, text3);
					num4 = 2;
				}
				if (HTKKFormater.isNullString(valueCustom3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 26, HTKKFile.ColumnLettersToIndex("F") - 1, text3);
					num4 = 2;
				}
			}
			else if (!HTKKFormater.isNullString(valueCustom3))
			{
				if (HTKKFormater.isNullString(valueCustom))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 22, HTKKFile.ColumnLettersToIndex("J") - 1, text3);
					num4 = 2;
				}
				if (HTKKFormater.isNullString(valueCustom2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 24, HTKKFile.ColumnLettersToIndex("J") - 1, text3);
					num4 = 2;
				}
			}
			if (valueCustom4.Length > 0)
			{
				if (HTKKFormater.isNullString(valueCustom))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 22, HTKKFile.ColumnLettersToIndex("J") - 1, text3);
					num4 = 2;
				}
				if (HTKKFormater.isNullString(valueCustom2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 24, HTKKFile.ColumnLettersToIndex("J") - 1, text3);
					num4 = 2;
				}
				if (HTKKFormater.isNullString(valueCustom3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 26, HTKKFile.ColumnLettersToIndex("F") - 1, text3);
					num4 = 2;
				}
			}
			if (valueCustom2.Length > 0 && !ValidateBase.CheckMaSoThue(valueCustom2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 24, HTKKFile.ColumnLettersToIndex("J") - 1, text2);
				num4 = 2;
			}
			HTKKFile.RemoveValidate(sheetView, 53, HTKKFile.ColumnLettersToIndex("U") - 1);
			HTKKFile.RemoveValidate(sheetView, 54, HTKKFile.ColumnLettersToIndex("U") - 1);
			HTKKFile.RemoveValidate(sheetView, 57, HTKKFile.ColumnLettersToIndex("U") - 1);
			string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, 54, HTKKFile.ColumnLettersToIndex("U"), "NumberCellType");
			string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, 55, HTKKFile.ColumnLettersToIndex("U"), "NumberCellType");
			string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, 58, HTKKFile.ColumnLettersToIndex("U"), "NumberCellType");
			string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, 59, HTKKFile.ColumnLettersToIndex("U"), "NumberCellType");
			if (ConvertHelpers.ConvertToDouble(valueCustom6) > ConvertHelpers.ConvertToDouble(valueCustom5))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 54, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28")));
				num4 = 2;
			}
			if (ConvertHelpers.ConvertToDouble(valueCustom7) > ConvertHelpers.ConvertToDouble(valueCustom6))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 57, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
				num4 = 2;
			}
			HTKKFile.RemoveValidate(sheetView, num2 - 1, num - 1);
			HTKKFile.RemoveValidate(sheetView, num3 - 1, num - 1);
			string valueCustom9 = GeneralFunctions.GetValueCustom(sheetView, num2, num, "2");
			string valueCustom10 = GeneralFunctions.GetValueCustom(sheetView, num3, num, "2");
			if (Convert.ToBoolean(valueCustom9) && !winFpSpread.Sheets[1].Visible)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num2 - 1, num - 1, text);
				num4 = 2;
			}
			if (Convert.ToBoolean(valueCustom10) && !winFpSpread.Sheets[1].Visible)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num3 - 1, num - 1, text);
				num4 = 2;
			}
			int[] array = GeneralFunctions.ArrayCellId("E_49");
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			string valueCustom11 = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "0");
			if (!HTKKFormater.isNullString(valueCustom11))
			{
				if (valueCustom11.IndexOf("_") > 0)
				{
					string[] array2 = valueCustom11.Split(new char[1] { '_' });
					if (array2.Length != 4)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						num4 = 2;
					}
					else
					{
						if (HTKKFormater.isNullString(array2[0]) || (!HTKKFormater.isNullString(array2[0]) && !array2[0].Equals(TTinChung.maCQTNoiNop)))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							num4 = 2;
						}
						if (HTKKFormater.isNullString(array2[1]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							num4 = 2;
						}
						else if ((TTinChung.mst.Length >= array2[1].Length && !array2[1].Equals(TTinChung.mst)) || (TTinChung.mst.Length < array2[1].Length && !array2[1].Substring(0, TTinChung.mst.Length).Equals(TTinChung.mst)))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							num4 = 2;
						}
						if (HTKKFormater.isNullString(array2[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							num4 = 2;
						}
						else
						{
							DateTime dtDate = DateTime.Now;
							string strDate = HTKKFormater.toDateDDMMYYYY(array2[2], isMess: false);
							if (!HTKKFormater.IsDateTime(strDate, "dd/MM/yyyy", out dtDate))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
								num4 = 2;
							}
						}
						if (HTKKFormater.isNullString(array2[3]) || (!HTKKFormater.isNullString(array2[3]) && ConvertHelpers.ConvertToInt(array2[3]) == 0))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							num4 = 2;
						}
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
					num4 = 2;
				}
			}
			string nam = PeriodInfo.Nam;
			string tuNgay = PeriodInfo.TuNgay;
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy");
			DateTime dateTime2 = new DateTime(2018, 8, 1);
			string kieuKy = PeriodInfo.KieuKy;
			if (kieuKy.Equals("Y"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018 && ConvertHelpers.ConvertToDouble(valueCustom7) > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 57, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
					num4 = 2;
				}
			}
			else if ((ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && dateTime >= dateTime2)) && ConvertHelpers.ConvertToDouble(valueCustom7) > 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 57, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
				num4 = 2;
			}
			switch (num4)
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		//IL_1a18: Unknown result type (might be due to invalid IL or missing references)
		//IL_191e: Unknown result type (might be due to invalid IL or missing references)
		int num = 0;
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<int> list = new List<int>();
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			bool flag = false;
			string text7 = "";
			DateTime dateTime = new DateTime(1990, 10, 10);
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("H");
			int num6 = HTKKFile.ColumnLettersToIndex("I");
			int num7 = HTKKFile.ColumnLettersToIndex("J");
			int num8 = HTKKFile.ColumnLettersToIndex("K");
			int num9 = HTKKFile.ColumnLettersToIndex("N");
			int num10 = HTKKFile.ColumnLettersToIndex("P");
			int iCol = HTKKFile.ColumnLettersToIndex("S");
			int num11 = HTKKFile.ColumnLettersToIndex("T");
			int num12 = HTKKFile.ColumnLettersToIndex("U");
			int num13 = HTKKFile.ColumnLettersToIndex("V");
			int iCol2 = HTKKFile.ColumnLettersToIndex("W");
			int iCol3 = HTKKFile.ColumnLettersToIndex("X");
			int num14 = HTKKFile.ColumnLettersToIndex("Y");
			int num15 = HTKKFile.ColumnLettersToIndex("Z");
			int num16 = HTKKFile.ColumnLettersToIndex("AA");
			int num17 = HTKKFile.ColumnLettersToIndex("AB");
			int iCol4 = HTKKFile.ColumnLettersToIndex("AC");
			int iCol5 = HTKKFile.ColumnLettersToIndex("AD");
			int iCol6 = HTKKFile.ColumnLettersToIndex("AE");
			int iCol7 = HTKKFile.ColumnLettersToIndex("AF");
			int iCol8 = HTKKFile.ColumnLettersToIndex("AG");
			int num18 = HTKKFile.ColumnLettersToIndex("AI");
			int num19 = HTKKFile.ColumnLettersToIndex("AJ");
			int num20 = HTKKFile.ColumnLettersToIndex("AL");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
				int num21 = 0;
				GetRowDelete(winFpSpread, ToKhaiInfo);
				sheetView.SetActiveCell(0, sheetView.ActiveColumn.Index);
				foreach (KeyValuePair<string, string> item in dicHDXoa)
				{
					string[] array = new string[2] { "0", "0" };
					if (item.Key.IndexOf(".") == -1 && item.Key.IndexOf("_") == -1)
					{
						continue;
					}
					if (item.Key.IndexOf(".") > -1)
					{
						array = item.Key.Split(new char[1] { '.' });
					}
					if (item.Key.IndexOf("_") > -1)
					{
						array = item.Key.Split(new char[1] { '_' });
					}
					bool flag2 = false;
					while (true)
					{
						int j;
						for (j = 0; j < rowsValues.Count; j++)
						{
							string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num20, "0");
							if (HTKKFormater.isNullString(valueCustom) || !valueCustom.StartsWith(array[0]))
							{
								continue;
							}
							if (valueCustom.IndexOf(".") > -1)
							{
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num20 - 1].Text = "";
								flag2 = true;
							}
							else if (j > 0)
							{
								if (flag2)
								{
									goto IL_04c8;
								}
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num20 - 1].Text = "";
							}
						}
						break;
						IL_04c8:
						hTKKCMCSheetInfo.Tables[i].RowsValues.RemoveAt(dataLocationTemplateRowStart + j - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart);
						RemoveDataRows(sheetView, dataLocationTemplateRowStart + j - 1, 1, hTKKCMCSheetInfo);
					}
				}
				dicHDXoa = new Dictionary<string, string>();
				bool flag3 = false;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text8 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text.Trim();
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num3, "0");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num4, "2");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num5, "0");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num6, "2");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num7, "2");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num8, "0");
					string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num9, "14");
					string valueCustom9 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num10, "0");
					string valueCustom10 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "2");
					string valueCustom11 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num11, "0");
					string valueCustom12 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num12, "NumberCellType");
					string valueCustom13 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num13, "NumberCellType");
					string valueCustom14 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol2, "NumberCellType");
					string valueCustom15 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol3, "2");
					string valueCustom16 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num14, "14");
					string valueCustom17 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num15, "14");
					string valueCustom18 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num16, "NumberCellType");
					string valueCustom19 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num17, "NumberCellType");
					string valueCustom20 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol4, "NumberCellType");
					string valueCustom21 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol5, "NumberCellType");
					string valueCustom22 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol6, "NumberCellType");
					string valueCustom23 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol7, "NumberCellType");
					string valueCustom24 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol8, "NumberCellType");
					string valueCustom25 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num18, "0");
					string valueCustom26 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num19, "0");
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num20, "0");
					Dictionary<int, string[]> dictionary3 = new Dictionary<int, string[]>();
					string obj = valueCustom16.Substring(valueCustom16.LastIndexOf("/") + 1);
					string obj2 = valueCustom17.Substring(valueCustom17.LastIndexOf("/") + 1);
					if (!HTKKFormater.isNullString(valueCustom17) && !HTKKFormater.isNullString(valueCustom16) && ConvertHelpers.ConvertToInt(obj2) - ConvertHelpers.ConvertToInt(obj) > 0)
					{
						if (!HTKKFormater.isNullString(valueCustom) && valueCustom.IndexOf(".") > 0)
						{
							string text9 = "";
							string text10 = valueCustom.Split(new char[1] { '.' })[0];
							int num22 = 1;
							bool flag4 = false;
							while (rowsValues.Count != 1 && num22 <= rowsValues.Count)
							{
								text9 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j + num22, num20, "0");
								if (!HTKKFormater.isNullString(text9) && text9.IndexOf(text10 + "_") >= 0)
								{
									flag4 = true;
									break;
								}
								num22++;
								if ((HTKKFormater.isNullString(text9) || text9.IndexOf(".") < 0) && !HTKKFormater.isNullString(text9))
								{
									break;
								}
							}
							if (!flag4)
							{
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num20 - 1].Text = "";
								valueCustom = "";
							}
						}
						dictionary3 = splitKy(valueCustom16, valueCustom17);
						if (dictionary3.Count > 1 && HTKKFormater.isNullString(valueCustom) && !Convert.ToBoolean(valueCustom3))
						{
							int num23 = 1;
							num21++;
							int num24 = 0;
							foreach (KeyValuePair<int, string[]> item2 in dictionary3)
							{
								int num25 = dataLocationTemplateRowStart + j;
								num24++;
								if (num25 < hTKKCMCTableInfo.Location.DataLocationTemplateRowStart)
								{
									num25 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
								}
								else if (num25 > hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
								{
									num25 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
								}
								InsertTableRow(sheetView, num25, -1, num25, -1, -1, hTKKCMCSheetInfo, 0);
								string[] value = item2.Value;
								sheetView.Cells[dataLocationTemplateRowStart + j + num23 - 1, num14 - 1].Text = value[0];
								sheetView.Cells[dataLocationTemplateRowStart + j + num23 - 1, num15 - 1].Text = value[1];
								sheetView.Cells[dataLocationTemplateRowStart + j + num23 - 1, num16 - 1].Text = value[0].Substring(value[0].LastIndexOf("/") + 1);
								sheetView.Cells[dataLocationTemplateRowStart + j + num23 - 1, 1].Text = "";
								sheetView.Cells[dataLocationTemplateRowStart + j + num23 - 1, num20 - 1].Text = num21 + "_" + num24;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, num20 - 1].Text = num21 + "." + num24;
								CellChange_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, dataLocationTemplateRowStart + j + num23, num15);
							}
						}
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num16 - 1);
					if (HTKKFormater.isNullString(valueCustom2) && HTKKFormater.isNullString(valueCustom4) && HTKKFormater.isNullString(valueCustom7) && HTKKFormater.isNullString(valueCustom8) && HTKKFormater.isNullString(valueCustom9) && HTKKFormater.isNullString(valueCustom11) && HTKKFormater.isNullString(valueCustom16) && HTKKFormater.isNullString(valueCustom17) && ConvertHelpers.ConvertToDouble(valueCustom12) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom13) == 0.0 && ConvertHelpers.ConvertToDouble(valueCustom14) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num16 - 1, text);
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num13 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1);
					bool success = false;
					bool success2 = false;
					DateTime dateTime2 = DateTime.MinValue;
					DateTime dateTime3 = DateTime.MinValue;
					if (ConvertHelpers.ConvertToBool(valueCustom3))
					{
						flag = true;
						if (HTKKFormater.isNullString(valueCustom4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text);
							list.Add(1);
						}
						if (!Convert.ToBoolean(valueCustom5) && !Convert.ToBoolean(valueCustom6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text);
							list.Add(1);
						}
						if (HTKKFormater.isNullString(valueCustom8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, text);
							list.Add(1);
						}
						else
						{
							bool success3 = false;
							DateTime dateTime4 = HTKKFormater.ConvertToDateTime(valueCustom8, out success3);
							if (!success3)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, HTKKFile.GetMessage("0071"));
								list.Add(1);
							}
						}
						if (HTKKFormater.isNullString(valueCustom11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text);
							list.Add(1);
						}
						if (!HTKKFormater.isNullString(valueCustom16))
						{
							dateTime2 = HTKKFormater.ConvertToDateTime(valueCustom16, out success);
							if (!success)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, HTKKFile.GetMessage("0071"));
								list.Add(1);
							}
						}
						if (!HTKKFormater.isNullString(valueCustom17))
						{
							dateTime3 = HTKKFormater.ConvertToDateTime(valueCustom17, out success2);
							if (!success2)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, HTKKFile.GetMessage("0071"));
								list.Add(1);
							}
						}
						if (!HTKKFormater.isNullString(valueCustom16) && !HTKKFormater.isNullString(valueCustom17) && success && success2 && (dateTime3 - dateTime2).TotalDays < 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, text2);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, text2);
							list.Add(1);
						}
					}
					else
					{
						if (!HTKKFormater.isNullString(valueCustom8))
						{
							bool success3 = false;
							DateTime dateTime4 = HTKKFormater.ConvertToDateTime(valueCustom8, out success3);
							if (!success3)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, HTKKFile.GetMessage("0071"));
								list.Add(1);
							}
						}
						if (HTKKFormater.isNullString(valueCustom16))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, text);
							list.Add(1);
						}
						else
						{
							dateTime2 = HTKKFormater.ConvertToDateTime(valueCustom16, out success);
							if (!success)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, HTKKFile.GetMessage("0071"));
								list.Add(1);
							}
						}
						if (HTKKFormater.isNullString(valueCustom17))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, text);
							list.Add(1);
						}
						else
						{
							dateTime3 = HTKKFormater.ConvertToDateTime(valueCustom17, out success2);
							if (!success2)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, HTKKFile.GetMessage("0071"));
								list.Add(1);
							}
						}
						if (!HTKKFormater.isNullString(valueCustom16) || !HTKKFormater.isNullString(valueCustom17))
						{
							text7 = text7 + ";" + valueCustom16.Replace("/", "") + valueCustom17.Replace("/", "");
							dateTime2 = ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy");
							dateTime3 = ConvertHelpers.ConvertToDateTime(valueCustom17, "dd/MM/yyyy");
							if (success && success2 && (dateTime3 - dateTime2).TotalDays < 0.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, text2);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, text2);
								list.Add(1);
							}
							string value2 = PeriodInfo.TuNgay.Replace("/", "") + PeriodInfo.DenNgay.Replace("/", "");
							if (text7.IndexOf(value2) == -1 && PeriodInfo.KieuKy == "Q")
							{
								flag3 = true;
							}
						}
						if (j > 0 && GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j - 1, num20, "0").IndexOf(".") < 0)
						{
							string text11 = sheetView.Cells[dataLocationTemplateRowStart + j - 2, num2 - 1].Text.Trim();
							DateTime dateTime5 = dateTime.AddDays(1.0);
							if ((dateTime5 - ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy")).TotalDays != 0.0 && text11 == text8 && !Convert.ToBoolean(sheetView.Cells[dataLocationTemplateRowStart + j - 2, num4 - 1].Value))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 2, num15 - 1, text6);
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, text6);
								list.Add(1);
							}
						}
						dateTime = ConvertHelpers.ConvertToDateTime(valueCustom17, "dd/MM/yyyy");
						if (ConvertHelpers.ConvertToDouble(valueCustom19) == 0.0 && (HTKKFormater.isNullString(valueCustom) || (!HTKKFormater.isNullString(valueCustom) && valueCustom.IndexOf(".") > -1)))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1, text);
							list.Add(1);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1);
						}
					}
					if (HTKKFormater.isNullString(valueCustom2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
						list.Add(1);
					}
					else
					{
						string value3 = "";
						dictionary.TryGetValue(text8, out value3);
						if (value3 == "1")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text5);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						}
					}
					if (!HTKKFormater.isNullString(text8))
					{
						string value3 = "";
						dictionary.TryGetValue(text8, out value3);
						if (ConvertHelpers.ConvertToBool(valueCustom15))
						{
							HTKKFile.AddOrUpdateDicStr(dictionary, text8, "1");
						}
						else if (HTKKFormater.isNullString(value3))
						{
							HTKKFile.AddOrUpdateDicStr(dictionary, text8, "0");
						}
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
					if (ConvertHelpers.ConvertToBool(valueCustom5) && ConvertHelpers.ConvertToBool(valueCustom3))
					{
						if (HTKKFormater.isNullString(valueCustom9))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text);
							list.Add(1);
						}
						if (ConvertHelpers.ConvertToDouble(valueCustom12) == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text);
							list.Add(1);
						}
					}
					if (HTKKFormater.isNullString(valueCustom25))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1, text);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1);
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1);
					if (HTKKFormater.isNullString(valueCustom26))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, text);
						list.Add(1);
					}
					else if (valueCustom26.IndexOf("_") > 0)
					{
						string[] array2 = valueCustom26.Split(new char[1] { '_' });
						if (array2.Length != 4)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							list.Add(1);
							continue;
						}
						if (HTKKFormater.isNullString(array2[0]) || (!HTKKFormater.isNullString(array2[0]) && !array2[0].Equals(TTinChung.maCQTNoiNop)))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							list.Add(1);
						}
						if (HTKKFormater.isNullString(array2[1]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							list.Add(1);
						}
						else if ((TTinChung.mst.Length >= array2[1].Length && !array2[1].Equals(TTinChung.mst)) || (TTinChung.mst.Length < array2[1].Length && !array2[1].Substring(0, TTinChung.mst.Length).Equals(TTinChung.mst)))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							list.Add(1);
						}
						if (HTKKFormater.isNullString(array2[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							list.Add(1);
						}
						else
						{
							DateTime dtDate = DateTime.Now;
							string strDate = HTKKFormater.toDateDDMMYYYY(array2[2], isMess: false);
							if (!HTKKFormater.IsDateTime(strDate, "dd/MM/yyyy", out dtDate))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
								list.Add(1);
							}
						}
						if (HTKKFormater.isNullString(array2[3]) || (!HTKKFormater.isNullString(array2[3]) && ConvertHelpers.ConvertToInt(array2[3]) == 0))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
							list.Add(1);
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						list.Add(1);
					}
				}
				if (flag3)
				{
					frmMessage frmMessage = new frmMessage("0373", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
			}
			if (dictionary.Count > 1 && PeriodInfo.KieuKy == "Q")
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					for (int j = 0; j < rowsValues.Count; j++)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
					}
				}
				num = 2;
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0369", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				winFpSpread.SuspendLayout();
			}
			else
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					if (!flag)
					{
						for (int j = 0; j < rowsValues.Count; j++)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text4);
						}
						list.Add(1);
						continue;
					}
					for (int j = 0; j < rowsValues.Count; j++)
					{
						if (!(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].BackColor == Color.Pink))
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						}
					}
					num = 0;
				}
			}
			isFirtLoad = false;
			ButtonClick_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 1L, 1L);
			dicHDXoa = new Dictionary<string, string>();
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					num = 2;
					break;
				}
			}
			switch (num)
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

	public Dictionary<string, string> GetDMHopDong(string pathFile)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			string text = pathFile + "\\Catalogue_DM_Hop_dong.xml";
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

	private Dictionary<int, string[]> splitKy(string tuNgay, string denNgay)
	{
		try
		{
			Dictionary<int, string[]> dictionary = new Dictionary<int, string[]>();
			string[] array = tuNgay.Split(new char[1] { '/' });
			string[] array2 = denNgay.Split(new char[1] { '/' });
			int num = ConvertHelpers.ConvertToInt(array2[2]);
			int num2 = num - ConvertHelpers.ConvertToInt(array[2]);
			string[] array3 = new string[2];
			while (num2 > 0)
			{
				array3 = new string[2];
				if (num == ConvertHelpers.ConvertToInt(array2[2]))
				{
					array3[0] = "01/01/" + num;
					array3[1] = denNgay;
					dictionary.Add(num2, array3);
					num2--;
					num--;
				}
				else
				{
					array3[0] = "01/01/" + num;
					array3[1] = "31/12/" + num;
					dictionary.Add(num2, array3);
					num2--;
					num--;
				}
			}
			dictionary.Add(num2, new string[2]
			{
				tuNgay,
				"31/12/" + num
			});
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private Dictionary<int, string[]> splitKy1(string tuNgay, string denNgay)
	{
		try
		{
			Dictionary<int, string[]> dictionary = new Dictionary<int, string[]>();
			string[] array = tuNgay.Split(new char[1] { '/' });
			string[] array2 = denNgay.Split(new char[1] { '/' });
			int num = 0;
			int num2 = ConvertHelpers.ConvertToInt(array[2]);
			string[] array3 = new string[2];
			while (ConvertHelpers.ConvertToInt(array2[2]) - num2 > 0)
			{
				array3 = new string[2];
				if (num == 0)
				{
					array3[0] = tuNgay;
					array3[1] = "31/12/" + num2;
					dictionary.Add(num, array3);
					num++;
					num2++;
				}
				else
				{
					array3[0] = "01/01/" + num2;
					array3[1] = "31/12/" + num2;
					dictionary.Add(num, array3);
					num++;
					num2++;
				}
			}
			dictionary.Add(num, new string[2]
			{
				"01/01/" + num2,
				denNgay
			});
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void InsertTableRow(SheetView sView, int frRow, int frCol, int toRow, int toCol, int colCount, HTKKCMCSheetInfo sInfo, int iCurTable)
	{
		sView.Rows.Add(frRow, 1);
		sView.CopyRange(frRow - 1, frCol, toRow, toCol, 1, colCount, dataOnly: false);
		sInfo.IncrementStaticCell(frRow, 1);
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > frRow)
			{
				table.Increment(1);
			}
		}
		sInfo.Tables[iCurTable].addNewDefaultRow(frRow - sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + 1);
	}

	public static void RemoveDataRows(SheetView sView, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (rowCount <= 0)
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
		catch (Exception)
		{
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			if (idxRow == 46 && idxCol == HTKKFile.ColumnLettersToIndex("B") - 1)
			{
				bool flag = ConvertHelpers.ConvertToBool(winFpSpread.Sheets[0].Cells[(int)idxRow, (int)idxCol].Value);
				winFpSpread.Sheets[0].SetActiveCell((int)idxRow, (int)idxCol - 1);
				string text = "U_54";
				string[] array = text.Split(new char[1] { '_' });
				if (!flag)
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array[1]) - 1, HTKKFile.ColumnLettersToIndex(array[0]) - 1].Formula = "AE54";
				}
				else
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array[1]) - 1, HTKKFile.ColumnLettersToIndex(array[0]) - 1].Formula = "AF54";
				}
				string text2 = "U_55";
				string[] array2 = text2.Split(new char[1] { '_' });
				if (!flag)
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[0]) - 1].Formula = "AE55";
				}
				else
				{
					winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[0]) - 1].Formula = "AF55";
				}
				string text3 = "U_56";
				string[] array3 = text3.Split(new char[1] { '_' });
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array3[1]) - 1, HTKKFile.ColumnLettersToIndex(array3[0]) - 1].Formula = "ROUND(U55*0.05,0)";
				text3 = "U_57";
				array3 = text3.Split(new char[1] { '_' });
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array3[1]) - 1, HTKKFile.ColumnLettersToIndex(array3[0]) - 1].Formula = "ROUND(U55*0.05,0)";
				text3 = "U_59";
				array3 = text3.Split(new char[1] { '_' });
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array3[1]) - 1, HTKKFile.ColumnLettersToIndex(array3[0]) - 1].Formula = "IF(U55<>0,ROUND(U57*(U58/U55)*0.5,0),0)";
				text3 = "U_62";
				array3 = text3.Split(new char[1] { '_' });
				winFpSpread.Sheets[0].Cells[ConvertHelpers.ConvertToInt(array3[1]) - 1, HTKKFile.ColumnLettersToIndex(array3[0]) - 1].Formula = "IF(U57-U59+U61>0,U57-U59+U61,0)";
			}
			if (idxRow == 44 && idxCol == HTKKFile.ColumnLettersToIndex("B") - 1)
			{
				int[] array4 = GeneralFunctions.ArrayCellId("E_49");
				int[] array5 = GeneralFunctions.ArrayCellId("AJ_22");
				if (ConvertHelpers.ConvertToBool(winFpSpread.Sheets[0].Cells[(int)idxRow, (int)idxCol].Value) && winFpSpread.Sheets[1].Visible)
				{
					string value = "";
					dicFormulaTK.TryGetValue("E_49", out value);
					winFpSpread.Sheets[0].Cells[array4[1], array4[0]].Formula = value.Replace("###", winFpSpread.Sheets[1].SheetName);
				}
				else
				{
					winFpSpread.Sheets[0].Cells[array4[1], array4[0]].Formula = "";
					winFpSpread.Sheets[0].Cells[array4[1], array4[0]].Value = "";
				}
			}
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
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("AL");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			int num5 = HTKKFile.ColumnLettersToIndex("I");
			string text = "";
			if (idxCol == num4 - 1 || idxCol == num5 - 1)
			{
				object value = sheetView.Cells[(int)idxRow, num4 - 1].Value;
				object value2 = sheetView.Cells[(int)idxRow, num5 - 1].Value;
				if ((bool)value && (bool)value && idxCol == num4 - 1)
				{
					sheetView.Cells[(int)idxRow, num5 - 1].Value = false;
				}
				if ((bool)value && (bool)value && idxCol == num5 - 1)
				{
					sheetView.Cells[(int)idxRow, num4 - 1].Value = false;
				}
				sheetView.SetActiveCell((int)idxRow, (int)idxCol);
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				if (!hTKKCMCSheetInfo.IsReadDataFile && isFirtLoad)
				{
					sheetView.Cells[dataLocationTemplateRowStart - 1, num2 - 1].Text = "0.0";
				}
				for (int j = 0; j < rowsValues.Count; j++)
				{
					object value3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Value;
					if (HTKKFormater.isNullString(value3.ToString()))
					{
						continue;
					}
					if (j == 0)
					{
						if (ConvertHelpers.ConvertToBool(value3))
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].HorizontalAlignment = CellHorizontalAlignment.Center;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = "1";
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = "0.0";
						}
						if (!HTKKFormater.isNullString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text))
						{
							text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
						}
						continue;
					}
					string text2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					if ((bool)value3)
					{
						if (text.IndexOf(".") > 0)
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].HorizontalAlignment = CellHorizontalAlignment.Center;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = (ConvertHelpers.ConvertToInt(text.Split(new char[1] { '.' })[0]) + 1).ToString();
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].HorizontalAlignment = CellHorizontalAlignment.Center;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = (ConvertHelpers.ConvertToInt(text) + 1).ToString();
						}
					}
					else
					{
						if (text2.Trim().IndexOf("_") > -1)
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = "";
							continue;
						}
						object value4 = sheetView.Cells[dataLocationTemplateRowStart + j - 2, num - 1].Value;
						if (text.IndexOf(".") > 0)
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = text.Split(new char[1] { '.' })[0] + "." + (ConvertHelpers.ConvertToInt(text.Split(new char[1] { '.' })[1]) + 1);
						}
						else
						{
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].HorizontalAlignment = CellHorizontalAlignment.Left;
							sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text = ConvertHelpers.ConvertToInt(text) + ".1";
						}
					}
					if (!HTKKFormater.isNullString(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text))
					{
						text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					}
				}
			}
			isFirtLoad = false;
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
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("H");
			int num5 = HTKKFile.ColumnLettersToIndex("I");
			int num6 = HTKKFile.ColumnLettersToIndex("J");
			int num7 = HTKKFile.ColumnLettersToIndex("K");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("P");
			int num10 = HTKKFile.ColumnLettersToIndex("S");
			int num11 = HTKKFile.ColumnLettersToIndex("T");
			int num12 = HTKKFile.ColumnLettersToIndex("U");
			int num13 = HTKKFile.ColumnLettersToIndex("V");
			int num14 = HTKKFile.ColumnLettersToIndex("W");
			int num15 = HTKKFile.ColumnLettersToIndex("X");
			int num16 = HTKKFile.ColumnLettersToIndex("Y");
			int num17 = HTKKFile.ColumnLettersToIndex("Z");
			int num18 = HTKKFile.ColumnLettersToIndex("AA");
			int num19 = HTKKFile.ColumnLettersToIndex("AB");
			int num20 = HTKKFile.ColumnLettersToIndex("AC");
			int num21 = HTKKFile.ColumnLettersToIndex("AD");
			int num22 = HTKKFile.ColumnLettersToIndex("AE");
			int num23 = HTKKFile.ColumnLettersToIndex("AF");
			int num24 = HTKKFile.ColumnLettersToIndex("AG");
			int num25 = HTKKFile.ColumnLettersToIndex("AI");
			int num26 = HTKKFile.ColumnLettersToIndex("AJ");
			int iCol = HTKKFile.ColumnLettersToIndex("AL");
			dicHDXoa = new Dictionary<string, string>();
			GetRowDelete(winFpSpread, ToKhaiInfo);
			foreach (KeyValuePair<string, string> item in dicHDXoa)
			{
				string[] array = new string[2] { "0", "0" };
				if (item.Key.IndexOf(".") == -1 && item.Key.IndexOf("_") == -1)
				{
					continue;
				}
				if (item.Key.IndexOf(".") > -1)
				{
					array = item.Key.Split(new char[1] { '.' });
				}
				if (item.Key.IndexOf("_") > -1)
				{
					array = item.Key.Split(new char[1] { '_' });
				}
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, iCol, "0");
					if (!HTKKFormater.isNullString(valueCustom) && valueCustom.StartsWith(array[0]) && valueCustom.IndexOf(".") > -1)
					{
						sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num22 - 1].Formula = "";
						sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num22 - 1].Value = 0;
					}
				}
			}
			dicHDXoa = new Dictionary<string, string>();
			ButtonClick_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 1L, 1L);
			return true;
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
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<int> list = new List<int>();
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			DateTime dateTime = new DateTime(1990, 10, 10);
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CQTGhiThongTin/maCQTQLyKThu"));
			if (hTKKCMCCellInfo != null && !HTKKFormater.isNullString(hTKKCMCCellInfo.Value))
			{
				string value = hTKKCMCCellInfo.Value;
				string text = "";
				string text2 = "";
				text = value.Substring(0, 3) + "00";
				foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc)
				{
					if (item.Value == text && item.Key.IndexOf(" - ") == -1)
					{
						text2 = item.Key;
						break;
					}
				}
				HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("null_3"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("null_2"));
				winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = text;
				winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = text2;
			}
			int num = HTKKFile.ColumnLettersToIndex("E");
			int iCol = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int iCol2 = HTKKFile.ColumnLettersToIndex("H");
			int iCol3 = HTKKFile.ColumnLettersToIndex("I");
			int iCol4 = HTKKFile.ColumnLettersToIndex("J");
			int iCol5 = HTKKFile.ColumnLettersToIndex("K");
			int iCol6 = HTKKFile.ColumnLettersToIndex("N");
			int iCol7 = HTKKFile.ColumnLettersToIndex("P");
			int iCol8 = HTKKFile.ColumnLettersToIndex("S");
			int iCol9 = HTKKFile.ColumnLettersToIndex("T");
			int iCol10 = HTKKFile.ColumnLettersToIndex("U");
			int iCol11 = HTKKFile.ColumnLettersToIndex("V");
			int iCol12 = HTKKFile.ColumnLettersToIndex("W");
			int iCol13 = HTKKFile.ColumnLettersToIndex("X");
			int iCol14 = HTKKFile.ColumnLettersToIndex("Y");
			int iCol15 = HTKKFile.ColumnLettersToIndex("Z");
			int iCol16 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol17 = HTKKFile.ColumnLettersToIndex("AB");
			int iCol18 = HTKKFile.ColumnLettersToIndex("AC");
			int iCol19 = HTKKFile.ColumnLettersToIndex("AD");
			int num3 = HTKKFile.ColumnLettersToIndex("AE");
			int iCol20 = HTKKFile.ColumnLettersToIndex("AF");
			int iCol21 = HTKKFile.ColumnLettersToIndex("AG");
			int iCol22 = HTKKFile.ColumnLettersToIndex("AI");
			int iCol23 = HTKKFile.ColumnLettersToIndex("AJ");
			int num4 = HTKKFile.ColumnLettersToIndex("AL");
			for (int num5 = 0; num5 < hTKKCMCSheetInfo.Tables.Count; num5++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[num5];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
				int num6 = 0;
				int num7 = 0;
				int num8 = 0;
				num8 = dataLocationTemplateRowStart;
				DateTime t = DateTime.Now;
				DateTime t2 = DateTime.Now;
				double soThang = 0.0;
				for (int num9 = 0; num9 < rowsValues.Count; num9++)
				{
					string text3 = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num - 1].Text.Trim();
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol, "0");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, num2, "2");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol2, "0");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol3, "2");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol4, "2");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol5, "0");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol6, "14");
					string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol7, "0");
					string valueCustom9 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol8, "2");
					string valueCustom10 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol9, "0");
					string valueCustom11 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol10, "NumberCellType");
					string valueCustom12 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol11, "NumberCellType");
					string valueCustom13 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol12, "NumberCellType");
					string valueCustom14 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol13, "2");
					string valueCustom15 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol14, "14");
					string valueCustom16 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol15, "14");
					string valueCustom17 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol16, "NumberCellType");
					string valueCustom18 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol17, "NumberCellType");
					string valueCustom19 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol18, "NumberCellType");
					string valueCustom20 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol19, "NumberCellType");
					string valueCustom21 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, num3, "NumberCellType");
					string valueCustom22 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol20, "NumberCellType");
					string valueCustom23 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol21, "NumberCellType");
					string valueCustom24 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol22, "0");
					string valueCustom25 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol23, "0");
					string valueCustom26 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, num4, "0");
					if (HTKKFormater.isNullString(valueCustom17))
					{
						if (!ConvertHelpers.ConvertToBool(valueCustom2))
						{
							num8 = dataLocationTemplateRowStart + num9;
							num6++;
							num7 = 0;
							t = ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy");
							t2 = ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy");
						}
					}
					else if (DateTime.Compare(ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy"), t) >= 0 && DateTime.Compare(ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy"), t2) <= 0 && num9 > 0 && num8 >= dataLocationTemplateRowStart && num8 < dataLocationTemplateRowStart + num9)
					{
						num7++;
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num4 - 1].Text = num6 + "_" + num7;
						sheetView.Cells[num8 - 1, num4 - 1].Text = num6 + "." + num7;
					}
					string text4 = valueCustom15.Replace("/", "") + valueCustom16.Replace("/", "");
					DateTime dateCol = ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy");
					DateTime dateCol2 = ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy");
					if (!isDataVB6)
					{
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = getValueCot20a(dateCol, dateCol2, t, t2, ConvertHelpers.ConvertToDouble(valueCustom18), sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num4 - 1].Text, out soThang);
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, HTKKFile.ColumnLettersToIndex("AK") - 1].Value = soThang;
					}
					else
					{
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Value;
					}
				}
				dicTongCT20a = new Dictionary<string, double[]>();
				dicTongCT20 = new Dictionary<string, double[]>();
				for (int num9 = 0; num9 < rowsValues.Count; num9++)
				{
					string text3 = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num - 1].Text.Trim();
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol, "0");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, num2, "2");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol2, "0");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol3, "2");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol4, "2");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol5, "0");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol6, "14");
					string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol7, "0");
					string valueCustom9 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol8, "2");
					string valueCustom10 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol9, "0");
					string valueCustom11 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol10, "NumberCellType");
					string valueCustom12 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol11, "NumberCellType");
					string valueCustom13 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol12, "NumberCellType");
					string valueCustom14 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol13, "2");
					string valueCustom15 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol14, "14");
					string valueCustom16 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol15, "14");
					string valueCustom17 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol16, "NumberCellType");
					string valueCustom18 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol17, "NumberCellType");
					string valueCustom19 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol18, "NumberCellType");
					string valueCustom20 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol19, "NumberCellType");
					string valueCustom21 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, num3, "NumberCellType");
					string valueCustom22 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol20, "NumberCellType");
					string valueCustom23 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol21, "NumberCellType");
					string valueCustom24 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol22, "0");
					string valueCustom25 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, iCol23, "0");
					string valueCustom26 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + num9, num4, "0");
					if (!HTKKFormater.isNullString(valueCustom26) && valueCustom26.IndexOf(".") > 0)
					{
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Formula = "";
						string[] array = valueCustom26.Split(new char[1] { '.' });
						int num10 = ConvertHelpers.ConvertToInt(array[1]);
						HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
						HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[0];
						HTKKCMCDataLocation location2 = hTKKCMCTableInfo2.Location;
						int dataLocationTemplateRowStart2 = location2.DataLocationTemplateRowStart;
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Value = 0;
						double num11 = 0.0;
						while (num10 > 0)
						{
							num11 += ConvertHelpers.ConvertToDouble(sheetView.Cells[dataLocationTemplateRowStart + num9 + num10 - 1, num3 - 1].Value);
							sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Value;
							num10--;
						}
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, HTKKFile.ColumnLettersToIndex("AS") - 1].Value = Math.Round(num11, 0);
						sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Value = Math.Round(num11, 0);
					}
					string text4 = text3 + valueCustom15.Replace("/", "") + valueCustom16.Replace("/", "");
					if (!HTKKFormater.isNullString(valueCustom17))
					{
						double valueCot20a = getValueCot20a(ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy"), ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy"), t, t2, ConvertHelpers.ConvertToDouble(valueCustom18), valueCustom26, out soThang);
						double[] newValue = new double[4]
						{
							ConvertHelpers.ConvertToDouble(text3),
							ConvertHelpers.ConvertToDouble(valueCustom17),
							valueCot20a,
							0.0
						};
						AddOrUpdateDicArr(dicTongCT20a, text4, newValue);
					}
				}
				ButtonClick_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, dataLocationTemplateRowStart - 1, num2 - 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static double getValueCot20a(DateTime dateCol17, DateTime dateCol18, DateTime dateNamDau, DateTime dateNamCuoi, double valueCot19, string valueDongTuSinh, out double soThang)
	{
		double num = 0.0;
		soThang = 0.0;
		try
		{
			string obj = dateCol17.ToString("yyyy");
			string obj2 = dateCol18.ToString("yyyy");
			if (dateCol17.Day == 1 && dateCol18.Day == DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month))
			{
				num = (soThang = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month + 1));
				num *= valueCot19;
			}
			else if (dateCol17.Day != 1 && dateCol18.Day != DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month) && dateCol17.Day - dateCol18.Day == 1)
			{
				num = (soThang = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month));
				num *= valueCot19;
			}
			else if (ConvertHelpers.ConvertToInt(obj2) - ConvertHelpers.ConvertToInt(obj) == 0 && ConvertHelpers.ConvertToInt(dateNamDau.Day) - ConvertHelpers.ConvertToInt(dateNamCuoi.Day) > 0 && !HTKKFormater.isNullString(valueDongTuSinh))
			{
				if (dateCol17.Day != 1 && dateCol18.Day == DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month))
				{
					num = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month);
					num = (soThang = num + (double)(dateCol18.Day - dateCol17.Day + 1) / ((double)DateTime.DaysInMonth(dateNamDau.Year, dateNamDau.Month) - (double)dateNamDau.Day + (double)dateNamCuoi.Day + 1.0));
					num *= valueCot19;
				}
				else if (dateCol17.Day == 1 && dateCol18.Day != DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month))
				{
					num = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month);
					num = (soThang = num + (double)dateCol18.Day / ((double)(DateTime.DaysInMonth(dateNamDau.Year, dateNamDau.Month) - dateNamDau.Day + 1) + (double)dateNamCuoi.Day));
					num *= valueCot19;
				}
				else if (valueDongTuSinh.IndexOf(".") > -1)
				{
					num = (soThang = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month + 1));
					num *= valueCot19;
				}
			}
			else if (dateCol17.Day != 1 && dateCol18.Day != DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month) && dateCol17.Day - dateCol18.Day > 1)
			{
				num = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month);
				num = (soThang = num - (double)(dateCol17.Day - dateCol18.Day - 1) / (double)DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month));
				num *= valueCot19;
			}
			else if (dateCol17.Day != 1 && dateCol18.Day != DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month) && dateCol17.Day - dateCol18.Day < 1)
			{
				num = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month);
				num = (soThang = num + (double)(dateCol18.Day - dateCol17.Day + 1) / (double)DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month));
				num *= valueCot19;
			}
			else if (dateCol17.Day != 1 && dateCol18.Day == DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month) && dateCol17.Day - dateCol18.Day < 1)
			{
				num = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month);
				num = (soThang = num + (double)(dateCol18.Day - dateCol17.Day + 1) / (double)DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month));
				num *= valueCot19;
			}
			else if (dateCol17.Day == 1 && dateCol18.Day != DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month) && dateCol17.Day - dateCol18.Day < 1)
			{
				num = (soThang = (double)((dateCol18.Year - dateCol17.Year) * 12) + (double)(dateCol18.Month - dateCol17.Month) + ConvertHelpers.ConvertToDouble1(dateCol18.Day) / ConvertHelpers.ConvertToDouble1(DateTime.DaysInMonth(dateCol18.Year, dateCol18.Month)));
				num *= valueCot19;
			}
			else
			{
				num = (soThang = (dateCol18.Year - dateCol17.Year) * 12 + dateCol18.Month - dateCol17.Month + 1);
				num *= valueCot19;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKHeaderInfo header)
	{
		try
		{
			if (PeriodInfo.KieuKy.Equals("Y"))
			{
				return;
			}
			string tuNgay = PeriodInfo.TuNgay;
			if (!HTKKFormater.isNullString(tuNgay) && tuNgay.IndexOf("/") > 0)
			{
				string[] array = tuNgay.Split(new char[1] { '/' });
				string ngayBatDauTC = "01/01";
				List<string> quyByNamTC = GeneralFunctions.GetQuyByNamTC(ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[2]), ngayBatDauTC);
				if (quyByNamTC != null)
				{
					xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai").InnerText = SecurityElement.Escape(quyByNamTC[0] + "/" + quyByNamTC[1]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void AddOrUpdateDicArr(Dictionary<string, double[]> dic, string key, double[] newValue)
	{
		try
		{
			if (dic.TryGetValue(key, out var _))
			{
				dic[key] = newValue;
			}
			else
			{
				dic.Add(key, newValue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void AddOrUpdateDicStrDouble(Dictionary<string, double> dic, string key, double newValue)
	{
		try
		{
			if (dic.TryGetValue(key, out var _))
			{
				dic[key] = newValue;
			}
			else
			{
				dic.Add(key, newValue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void GetRowDelete(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AL");
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text;
				if (HTKKFormater.isNullString(text))
				{
					continue;
				}
				if (text.IndexOf(".") > -1)
				{
					bool flag = false;
					int num2 = ConvertHelpers.ConvertToInt(text.Split(new char[1] { '.' })[1]);
					int num3 = 0;
					for (int j = i + 1; j < dataLocationTemplateRowStart; j++)
					{
						string text2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
						if (HTKKFormater.isNullString(text2))
						{
							break;
						}
						if (text2.StartsWith(text.Split(new char[1] { '.' })[0]))
						{
							num3++;
						}
					}
					if (num2 != num3)
					{
						HTKKFile.AddOrUpdateDicStr(dicHDXoa, text, "1");
					}
				}
				else
				{
					if (text.IndexOf("_") <= -1)
					{
						continue;
					}
					bool flag2 = false;
					for (int j = 0; j < i; j++)
					{
						string text3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
						if (!HTKKFormater.isNullString(text3) && text3.StartsWith(text.Split(new char[1] { '_' })[0]))
						{
							flag2 = true;
							break;
						}
					}
					if (!flag2)
					{
						HTKKFile.AddOrUpdateDicStr(dicHDXoa, text, "1");
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void GetDicCT20a(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<int> list = new List<int>();
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			DateTime dateTime = new DateTime(1990, 10, 10);
			int num = HTKKFile.ColumnLettersToIndex("E");
			int iCol = HTKKFile.ColumnLettersToIndex("F");
			int iCol2 = HTKKFile.ColumnLettersToIndex("G");
			int iCol3 = HTKKFile.ColumnLettersToIndex("H");
			int iCol4 = HTKKFile.ColumnLettersToIndex("I");
			int iCol5 = HTKKFile.ColumnLettersToIndex("J");
			int iCol6 = HTKKFile.ColumnLettersToIndex("K");
			int iCol7 = HTKKFile.ColumnLettersToIndex("N");
			int iCol8 = HTKKFile.ColumnLettersToIndex("P");
			int iCol9 = HTKKFile.ColumnLettersToIndex("S");
			int iCol10 = HTKKFile.ColumnLettersToIndex("T");
			int iCol11 = HTKKFile.ColumnLettersToIndex("U");
			int iCol12 = HTKKFile.ColumnLettersToIndex("V");
			int iCol13 = HTKKFile.ColumnLettersToIndex("W");
			int iCol14 = HTKKFile.ColumnLettersToIndex("X");
			int iCol15 = HTKKFile.ColumnLettersToIndex("Y");
			int iCol16 = HTKKFile.ColumnLettersToIndex("Z");
			int iCol17 = HTKKFile.ColumnLettersToIndex("AA");
			int iCol18 = HTKKFile.ColumnLettersToIndex("AB");
			int iCol19 = HTKKFile.ColumnLettersToIndex("AC");
			int iCol20 = HTKKFile.ColumnLettersToIndex("AD");
			int iCol21 = HTKKFile.ColumnLettersToIndex("AE");
			int iCol22 = HTKKFile.ColumnLettersToIndex("AF");
			int iCol23 = HTKKFile.ColumnLettersToIndex("AG");
			int iCol24 = HTKKFile.ColumnLettersToIndex("AI");
			int iCol25 = HTKKFile.ColumnLettersToIndex("AJ");
			int iCol26 = HTKKFile.ColumnLettersToIndex("AL");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
				int num2 = 0;
				num2 = dataLocationTemplateRowStart;
				DateTime t = DateTime.Now;
				DateTime t2 = DateTime.Now;
				double soThang = 0.0;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "0");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol2, "2");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol3, "0");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol4, "2");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol5, "2");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol6, "0");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol7, "14");
					string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol8, "0");
					string valueCustom9 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol9, "2");
					string valueCustom10 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol10, "0");
					string valueCustom11 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol11, "NumberCellType");
					string valueCustom12 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol12, "NumberCellType");
					string valueCustom13 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol13, "NumberCellType");
					string valueCustom14 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol14, "2");
					string valueCustom15 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol15, "14");
					string valueCustom16 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol16, "14");
					string valueCustom17 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol17, "NumberCellType");
					string valueCustom18 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol18, "NumberCellType");
					string valueCustom19 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol19, "NumberCellType");
					string valueCustom20 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol20, "NumberCellType");
					string valueCustom21 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol21, "NumberCellType");
					string valueCustom22 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol22, "NumberCellType");
					string valueCustom23 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol23, "NumberCellType");
					string valueCustom24 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol24, "0");
					string valueCustom25 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol25, "0");
					string valueCustom26 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol26, "0");
					if (HTKKFormater.isNullString(valueCustom17))
					{
						if (!ConvertHelpers.ConvertToBool(valueCustom2))
						{
							t = ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy");
							t2 = ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy");
						}
					}
					else if (DateTime.Compare(ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy"), t) >= 0 && DateTime.Compare(ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy"), t2) <= 0)
					{
					}
					string text2 = valueCustom15.Replace("/", "") + valueCustom16.Replace("/", "");
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy");
					DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy");
				}
				dicTongCT20a = new Dictionary<string, double[]>();
				dicTongCT20 = new Dictionary<string, double[]>();
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol, "0");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol2, "2");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol3, "0");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol4, "2");
					string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol5, "2");
					string valueCustom6 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol6, "0");
					string valueCustom7 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol7, "14");
					string valueCustom8 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol8, "0");
					string valueCustom9 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol9, "2");
					string valueCustom10 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol10, "0");
					string valueCustom11 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol11, "NumberCellType");
					string valueCustom12 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol12, "NumberCellType");
					string valueCustom13 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol13, "NumberCellType");
					string valueCustom14 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol14, "2");
					string valueCustom15 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol15, "14");
					string valueCustom16 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol16, "14");
					string valueCustom17 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol17, "NumberCellType");
					string valueCustom18 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol18, "NumberCellType");
					string valueCustom19 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol19, "NumberCellType");
					string valueCustom20 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol20, "NumberCellType");
					string valueCustom21 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol21, "NumberCellType");
					string valueCustom22 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol22, "NumberCellType");
					string valueCustom23 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol23, "NumberCellType");
					string valueCustom24 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol24, "0");
					string valueCustom25 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol25, "0");
					string valueCustom26 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, iCol26, "0");
					string text2 = text + valueCustom15.Replace("/", "") + valueCustom16.Replace("/", "");
					if (!HTKKFormater.isNullString(valueCustom17))
					{
						double valueCot20a = getValueCot20a(ConvertHelpers.ConvertToDateTime(valueCustom15, "dd/MM/yyyy"), ConvertHelpers.ConvertToDateTime(valueCustom16, "dd/MM/yyyy"), t, t2, ConvertHelpers.ConvertToDouble(valueCustom18), valueCustom26, out soThang);
						double[] newValue = new double[4]
						{
							ConvertHelpers.ConvertToDouble(text),
							ConvertHelpers.ConvertToDouble(valueCustom17),
							valueCot20a,
							0.0
						};
						AddOrUpdateDicArr(dicTongCT20a, text2, newValue);
					}
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
