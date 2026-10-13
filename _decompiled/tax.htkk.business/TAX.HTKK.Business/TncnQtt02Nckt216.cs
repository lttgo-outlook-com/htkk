using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnQtt02Nckt216
{
	private static double mucGiamTruNPT = 3600000.0;

	private string strNam = "";

	public string cellId_n4 = "N_4";

	public string cellId_j4 = "J_4";

	public string cellId_p4 = "P_4";

	public string cellId_r4 = "R_4";

	public string cellId_ct3 = "E_8";

	public string cellId_ct4 = "E_10";

	public string cellId_ct5 = "E_12";

	public string cellId_ct6 = "E_14";

	public string cellId_ct22 = "M_42";

	public string cellId_ct23 = "M_43";

	public string cellId_ct24 = "M_44";

	public string cellId_ct25 = "M_45";

	public string cellId_ct26 = "M_46";

	public string cellId_ct27 = "M_47";

	public string cellId_ct28 = "M_48";

	public string cellId_ct29 = "M_49";

	public string cellId_ct30 = "M_50";

	public string cellId_ct31 = "M_51";

	public string cellId_ct32 = "M_52";

	public string cellId_ct33 = "M_53";

	public string cellId_ct34 = "M_54";

	public string cellId_ct35 = "M_55";

	public string cellId_ct36 = "M_56";

	public string cellId_ct37 = "M_57";

	public string cellId_ct38 = "M_58";

	public string cellId_ct39 = "M_59";

	public string cellId_ct40 = "M_60";

	public string cellId_ct41 = "M_62";

	public string cellId_ct42 = "M_63";

	public string cellId_ct43 = "M_64";

	public string cellId_ct44 = "M_65";

	public string cellId_ct45 = "M_66";

	public string cellId_ct46 = "M_67";

	public string cellId_ct47 = "M_68";

	public string cellId_ct48 = "M_69";

	public string cellId_ct49 = "M_70";

	public int[] CellId_j4 => GeneralFunctions.ArrayCellId(cellId_j4);

	public int[] CellId_n4 => GeneralFunctions.ArrayCellId(cellId_n4);

	public int[] CellId_p4 => GeneralFunctions.ArrayCellId(cellId_p4);

	public int[] CellId_r4 => GeneralFunctions.ArrayCellId(cellId_r4);

	public int[] CellId_ct3 => GeneralFunctions.ArrayCellId(cellId_ct3);

	public int[] CellId_ct4 => GeneralFunctions.ArrayCellId(cellId_ct4);

	public int[] CellId_ct5 => GeneralFunctions.ArrayCellId(cellId_ct5);

	public int[] CellId_ct6 => GeneralFunctions.ArrayCellId(cellId_ct6);

	public int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	public int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	public int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	public int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	public int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	public int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	public int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	public int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	public int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	public int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	public int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	public int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	public int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	public int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	public int[] CellId_ct36 => GeneralFunctions.ArrayCellId(cellId_ct36);

	public int[] CellId_ct37 => GeneralFunctions.ArrayCellId(cellId_ct37);

	public int[] CellId_ct38 => GeneralFunctions.ArrayCellId(cellId_ct38);

	public int[] CellId_ct39 => GeneralFunctions.ArrayCellId(cellId_ct39);

	public int[] CellId_ct40 => GeneralFunctions.ArrayCellId(cellId_ct40);

	public int[] CellId_ct41 => GeneralFunctions.ArrayCellId(cellId_ct41);

	public int[] CellId_ct42 => GeneralFunctions.ArrayCellId(cellId_ct42);

	public int[] CellId_ct43 => GeneralFunctions.ArrayCellId(cellId_ct43);

	public int[] CellId_ct44 => GeneralFunctions.ArrayCellId(cellId_ct44);

	public int[] CellId_ct45 => GeneralFunctions.ArrayCellId(cellId_ct45);

	public int[] CellId_ct46 => GeneralFunctions.ArrayCellId(cellId_ct46);

	public int[] CellId_ct47 => GeneralFunctions.ArrayCellId(cellId_ct47);

	public int[] CellId_ct48 => GeneralFunctions.ArrayCellId(cellId_ct48);

	public int[] CellId_ct49 => GeneralFunctions.ArrayCellId(cellId_ct49);

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string text = sheetView.Cells[CellId_n4[1], CellId_n4[0]].Text;
			string text2 = sheetView.Cells[CellId_p4[1], CellId_p4[0]].Text;
			int soThang = GeneralFunctions.GetSoThang(text.Trim().Replace("/", ""), text2.Trim().Replace("/", ""));
			sheetView.Cells[CellId_r4[1], CellId_r4[0]].Value = soThang;
			string text3 = sheetView.Cells[CellId_j4[1], CellId_j4[0]].Text;
			if (ConvertHelpers.ConvertToInt(text3) > 2019)
			{
				mucGiamTruNPT = 4400000.0;
			}
			else
			{
				mucGiamTruNPT = 3600000.0;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			string nam = PeriodInfo.Nam;
			string tuThang = PeriodInfo.TuThang;
			string denThang = PeriodInfo.DenThang;
			strNam = nam;
			string[] array = tuThang.Split(new char[1] { '/' });
			string[] array2 = denThang.Split(new char[1] { '/' });
			int column = GeneralFunctions.ArrayCellId("M_44")[0];
			int row = GeneralFunctions.ArrayCellId("M_44")[1];
			if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(array[0]) >= 8))
			{
				sView.Cells[row, column].Value = "0";
				sView.Cells[row, column].Locked = true;
			}
			int soThang = GeneralFunctions.GetSoThang(tuThang.Trim().Replace("/", ""), denThang.Trim().Replace("/", ""));
			if (ConvertHelpers.ConvertToInt(nam) == 2019 && ConvertHelpers.ConvertToInt(array2[1]) > 2019)
			{
				int num = 12 - ConvertHelpers.ConvertToInt(array[0]) + 1;
				int num2 = soThang - num;
				double num3 = num * 9000000 + num2 * 11000000;
				sView.Cells[CellId_ct29[1], CellId_ct29[0]].Value = num3;
			}
			else if (ConvertHelpers.ConvertToInt(nam) > 2019)
			{
				double num3 = soThang * 11000000;
				sView.Cells[CellId_ct29[1], CellId_ct29[0]].Value = num3;
			}
			else
			{
				double num3 = soThang * 9000000;
				sView.Cells[CellId_ct29[1], CellId_ct29[0]].Value = num3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			List<HTKKCMCRowValue> rowsValues = ToKhaiInfo.SheetInfos[1].Tables[1].RowsValues;
			int column = GeneralFunctions.ArrayCellId("K_33")[0];
			int column2 = GeneralFunctions.ArrayCellId("M_33")[0];
			int column3 = GeneralFunctions.ArrayCellId("O_33")[0];
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int row = ToKhaiInfo.SheetInfos[1].Tables[1].Location.DataLocationTemplateRowStart + i - 1;
				if (ConvertHelpers.ConvertToString(sView.Cells[row, HTKKFile.ColumnLettersToIndex("I") - 1].Value) == "Khác")
				{
					sView.Cells[row, column].Locked = true;
					sView.Cells[row, column2].Locked = true;
					sView.Cells[row, column3].Locked = true;
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				CellChange_1(sView.FpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, HTKKFile.ColumnLettersToIndex("S") - 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string nam = PeriodInfo.Nam;
			SheetView sheetView = winFpSpread.Sheets[1];
			int row = (int)iRow - 1;
			int num = (int)iCol - 1;
			int column = HTKKFile.ColumnLettersToIndex("Y") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("Z") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("Q") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("S") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("Q") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("Q") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("T") - 1;
			int num7 = HTKKFile.ColumnLettersToIndex("U") - 1;
			if (num == num2 || num == num3)
			{
				DateTime result = default(DateTime);
				DateTime.TryParseExact("01/" + sheetView.Cells[row, num2].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
				DateTime result2 = default(DateTime);
				DateTime.TryParseExact("01/" + sheetView.Cells[row, num3].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result2);
				if (!result2.Equals(DateTime.MinValue))
				{
					int num8 = (result2.Year - result.Year) * 12 + (result2.Month - result.Month) + 1;
					double num9 = 0.0;
					if (result.Year > 2019)
					{
						mucGiamTruNPT = 4400000.0;
					}
					else
					{
						mucGiamTruNPT = 3600000.0;
					}
					if (nam.Equals("2019") && result2.Year > result.Year)
					{
						int num10 = 0;
						if (result.Year == 2019)
						{
							num10 = 12 - result.Month + 1;
						}
						int num11 = num8 - num10;
						num9 = num10 * 3600000 + num11 * 4400000;
					}
					else
					{
						num9 = (double)num8 * mucGiamTruNPT;
					}
					sheetView.Cells[row, column].Value = num8;
					sheetView.Cells[row, column2].Value = num9;
				}
			}
			if (num == num6 || num == num7)
			{
				DateTime result = default(DateTime);
				DateTime.TryParseExact("01/" + sheetView.Cells[row, num6].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
				DateTime result2 = default(DateTime);
				DateTime.TryParseExact("01/" + sheetView.Cells[row, num7].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result2);
				if (!result2.Equals(DateTime.MinValue))
				{
					int num8 = (result2.Year - result.Year) * 12 + (result2.Month - result.Month) + 1;
					double num9 = 0.0;
					if (result.Year > 2019)
					{
						mucGiamTruNPT = 4400000.0;
					}
					else
					{
						mucGiamTruNPT = 3600000.0;
					}
					if (nam.Equals("2019") && result2.Year > result.Year)
					{
						int num10 = 0;
						if (result.Year == 2019)
						{
							num10 = 12 - result.Month + 1;
						}
						int num11 = num8 - num10;
						num9 = num10 * 3600000 + num11 * 4400000;
					}
					else
					{
						num9 = (double)num8 * mucGiamTruNPT;
					}
					sheetView.Cells[row, column].Value = num8;
					sheetView.Cells[row, column2].Value = num9;
				}
			}
			string[] array = new string[1] { "I_33" };
			int num12 = 0;
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + "_" + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			if (text.Split(new char[1] { '_' })[0] == array[0].Split(new char[1] { '_' })[0] && int.Parse(text.Split(new char[1] { '_' })[1]) >= 33)
			{
				int num13 = GeneralFunctions.ArrayCellId(text)[1];
				int num14 = GeneralFunctions.ArrayCellId("K_33")[0];
				int num15 = GeneralFunctions.ArrayCellId("M_33")[0];
				int num16 = GeneralFunctions.ArrayCellId("O_33")[0];
				if (GeneralFunctions.GetValue(sheetView, text) == "Việt Nam")
				{
					sheetView.Cells[num13, num14].Locked = false;
					sheetView.Cells[num13, num15].Locked = false;
					sheetView.Cells[num13, num16].Locked = false;
					sheetView.Cells[num13, num14].ResetValue();
					sheetView.Cells[num13, num15].ResetValue();
					sheetView.Cells[num13, num16].ResetValue();
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num13, num14);
					HTKKFile.RemoveValidate(sheetView, num13, num15);
					HTKKFile.RemoveValidate(sheetView, num13, num16);
					sheetView.Cells[num13, num14].Value = "Khác";
					sheetView.Cells[num13, num15].Value = "Khác";
					sheetView.Cells[num13, num16].Value = "Khác";
					sheetView.Cells[num13, num14].Locked = true;
					sheetView.Cells[num13, num15].Locked = true;
					sheetView.Cells[num13, num16].Locked = true;
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
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL2_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string text = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sView, "F_20"));
			string text2 = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sView, "O_20"));
			string obj = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sView, "M_68"));
			if (ConvertHelpers.ConvertToDouble(obj) > 0.0)
			{
				if (text.Trim().Length == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, 19, HTKKFile.ColumnLettersToIndex("F") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_64")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround1(sView, 19, HTKKFile.ColumnLettersToIndex("F") - 1);
				}
				if (text2.Trim().Length == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, 19, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_64")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, 19, HTKKFile.ColumnLettersToIndex("O") - 1);
				}
			}
			else
			{
				HTKKFile.RemoveValidateBackGround1(sView, 19, HTKKFile.ColumnLettersToIndex("F") - 1);
				HTKKFile.RemoveValidate(sView, 19, HTKKFile.ColumnLettersToIndex("O") - 1);
			}
			string obj2 = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sView, "M_44"));
			string nam = PeriodInfo.Nam;
			string tuThang = PeriodInfo.TuThang;
			string[] array = tuThang.Split(new char[1] { '/' });
			if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(array[0]) >= 8))
			{
				if (ConvertHelpers.ConvertToInt(obj2) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, 43, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_69")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, 43, HTKKFile.ColumnLettersToIndex("M") - 1);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, 43, HTKKFile.ColumnLettersToIndex("M") - 1);
			}
			string text3 = "43~44~45~46~51~52~53~57~58~59~60~64~68~69";
			string[] array2 = text3.Split(new char[1] { '~' });
			string text4 = "0";
			bool flag = true;
			if (array2.Length > 0)
			{
				for (int i = 0; i < array2.Length; i++)
				{
					text4 = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sView, "M_" + array2[i]));
					if (ConvertHelpers.ConvertToDouble(text4) > 0.0)
					{
						flag = false;
						break;
					}
				}
			}
			if (flag)
			{
				for (int i = 0; i < array2.Length; i++)
				{
					HTKKFile.SetNodeBackGroundPink(sView, ConvertHelpers.ConvertToInt(array2[i]) - 1, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_70")));
				}
				value = 2;
			}
			else
			{
				for (int i = 0; i < array2.Length; i++)
				{
					HTKKFile.RemoveValidate(sView, ConvertHelpers.ConvertToInt(array2[i]) - 1, HTKKFile.ColumnLettersToIndex("M") - 1);
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

	private void checkPL2_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int num2 = rowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			int num3 = rowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int num4 = rowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int num5 = rowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num6 = rowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int num7 = rowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int num8 = rowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int num9 = rowsInfo.Cells[7].CellTemplate.ColIndex - 1;
			int column = rowsInfo.Cells[8].CellTemplate.ColIndex - 1;
			int column2 = rowsInfo.Cells[9].CellTemplate.ColIndex - 1;
			string[] array = new string[10];
			string text = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sheetView, cellId_ct3));
			string text2 = ConvertHelpers.ConvertToString(GeneralFunctions.GetObjectValue(sheetView, cellId_ct5));
			HTKKFile.RemoveValidate(sheetView, CellId_ct5[1], CellId_ct5[0]);
			if (!string.IsNullOrEmpty(text2))
			{
				if (text2.Trim().Equals(text.Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct5[1], CellId_ct5[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct5[1], CellId_ct5[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				else
				{
					int foundRowIndex = 0;
					int foundColumnIndex = 0;
					string text3 = winFpSpread.Search(1, text2, caseSensitive: false, exactMatch: true, alternateSearch: false, useWildcards: false, num, HTKKFile.ColumnLettersToIndex("F") - 1, num - 1 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("F") - 1, ref foundRowIndex, ref foundColumnIndex);
					if (foundRowIndex > -1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct5[1], CellId_ct5[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
						value = 2;
					}
				}
			}
			string text4 = (string)GeneralFunctions.GetObjectValue(sheetView, cellId_ct6);
			if (!string.IsNullOrEmpty(text4) && text4.Contains(" "))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct6[1], CellId_ct6[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct6[1], CellId_ct6[0]);
			}
			DateTime result3;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num + i, num2].Text;
				array[1] = sheetView.Cells[num + i, num3].Text;
				array[2] = sheetView.Cells[num + i, num4].Text;
				array[3] = sheetView.Cells[num + i, num5].Text;
				array[4] = sheetView.Cells[num + i, num6].Text;
				array[5] = sheetView.Cells[num + i, num7].Text;
				array[6] = sheetView.Cells[num + i, num8].Text;
				array[7] = sheetView.Cells[num + i, num9].Text;
				array[8] = sheetView.Cells[num + i, column].Text;
				array[9] = sheetView.Cells[num + i, column2].Text;
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num2);
					HTKKFile.RemoveValidate(sheetView, num + i, num3);
					HTKKFile.RemoveValidate(sheetView, num + i, num5);
					HTKKFile.RemoveValidate(sheetView, num + i, num4);
					HTKKFile.RemoveValidate(sheetView, num + i, num6);
					HTKKFile.RemoveValidate(sheetView, num + i, num7);
					HTKKFile.RemoveValidate(sheetView, num + i, num8);
					HTKKFile.RemoveValidate(sheetView, num + i, num9);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num2);
				}
				string text5 = sheetView.Cells[num + i, num3].Text;
				string text6 = sheetView.Cells[num + i, num8].Text;
				string text7 = sheetView.Cells[num + i, num9].Text;
				DateTime result = default(DateTime);
				DateTime result2 = default(DateTime);
				bool flag = false;
				if (string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				else if (!DateTime.TryParseExact(text6, "MM/yyyy", null, DateTimeStyles.None, out result))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
					value = 2;
				}
				else if (result > DateTime.ParseExact(PeriodInfo.DenThang, "MM/yyyy", CultureInfo.CurrentCulture) || result < DateTime.ParseExact(PeriodInfo.TuThang, "MM/yyyy", CultureInfo.CurrentCulture))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num8);
					flag = true;
				}
				if (string.IsNullOrEmpty(text7))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				else if (!DateTime.TryParseExact(text7, "MM/yyyy", null, DateTimeStyles.None, out result2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
					value = 2;
				}
				else if (result2 < result)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
					value = 2;
				}
				else if (result2 > DateTime.ParseExact(PeriodInfo.DenThang, "MM/yyyy", CultureInfo.CurrentCulture))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
					value = 2;
				}
				else
				{
					if (flag)
					{
						HTKKFile.RemoveValidate(sheetView, num + i, num8);
					}
					HTKKFile.RemoveValidate(sheetView, num + i, num9);
				}
				if (!string.IsNullOrEmpty(text5))
				{
					if (!DateTime.TryParseExact(text5, "dd/MM/yyyy", null, DateTimeStyles.None, out result3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
					else if (result3 > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
						value = 2;
					}
					else
					{
						DateTime dateTime = new DateTime(result.Year, result.Month, DateTime.DaysInMonth(result.Year, result.Month));
						if (result3 > dateTime)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, num3);
						}
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[2]) && string.IsNullOrEmpty(array[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num6, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num4);
					HTKKFile.RemoveValidate(sheetView, num + i, num6);
				}
				if (!string.IsNullOrEmpty(array[2]))
				{
					if (!ValidateBase.CheckMaSoThue(array[2]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
						value = 2;
					}
					else if (array[2].Equals(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
						value = 2;
					}
					else
					{
						int foundRowIndex = 0;
						int foundColumnIndex = 0;
						string text3 = winFpSpread.Search(1, array[2], caseSensitive: false, exactMatch: true, alternateSearch: false, useWildcards: false, num, HTKKFile.ColumnLettersToIndex("F") - 1, num - 1 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("F") - 1, ref foundRowIndex, ref foundColumnIndex);
						string text8 = winFpSpread.Search(1, array[2], caseSensitive: false, exactMatch: true, alternateSearch: false, useWildcards: false, foundRowIndex + 1, HTKKFile.ColumnLettersToIndex("F") - 1, num - 1 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("F") - 1, ref foundRowIndex, ref foundColumnIndex);
						if (foundRowIndex > -1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, num4);
						}
					}
				}
				if (!string.IsNullOrEmpty(array[4]))
				{
					if (array[4].Contains(" "))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num6, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
						value = 2;
					}
					else
					{
						int foundRowIndex = 0;
						int foundColumnIndex = 0;
						string text3 = winFpSpread.Search(1, array[4], caseSensitive: false, exactMatch: true, alternateSearch: false, useWildcards: false, num, HTKKFile.ColumnLettersToIndex("M") - 1, num - 1 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("M") - 1, ref foundRowIndex, ref foundColumnIndex);
						string text8 = winFpSpread.Search(1, array[4], caseSensitive: false, exactMatch: true, alternateSearch: false, useWildcards: false, foundRowIndex + 1, HTKKFile.ColumnLettersToIndex("M") - 1, num - 1 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("M") - 1, ref foundRowIndex, ref foundColumnIndex);
						if (foundRowIndex > -1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num6, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, num6);
						}
					}
				}
				if (string.IsNullOrEmpty(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num7, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num7);
				}
			}
			hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			rowsInfo = hTKKCMCTableInfo.RowsInfo;
			num2 = rowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			num3 = rowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int column3 = rowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int column4 = rowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num10 = rowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int num11 = rowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int num12 = rowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int num13 = rowsInfo.Cells[7].CellTemplate.ColIndex - 1;
			int num14 = rowsInfo.Cells[8].CellTemplate.ColIndex - 1;
			int num15 = rowsInfo.Cells[9].CellTemplate.ColIndex - 1;
			num8 = rowsInfo.Cells[10].CellTemplate.ColIndex - 1;
			num9 = rowsInfo.Cells[11].CellTemplate.ColIndex - 1;
			int column5 = rowsInfo.Cells[12].CellTemplate.ColIndex - 1;
			int column6 = rowsInfo.Cells[13].CellTemplate.ColIndex - 1;
			int column7 = rowsInfo.Cells[14].CellTemplate.ColIndex - 1;
			int column8 = rowsInfo.Cells[15].CellTemplate.ColIndex - 1;
			column = rowsInfo.Cells[16].CellTemplate.ColIndex - 1;
			column2 = rowsInfo.Cells[17].CellTemplate.ColIndex - 1;
			array = new string[18];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num + i, num2].Text;
				array[1] = sheetView.Cells[num + i, num3].Text;
				array[2] = sheetView.Cells[num + i, column3].Text;
				array[3] = sheetView.Cells[num + i, column4].Text;
				array[4] = sheetView.Cells[num + i, num10].Text;
				array[5] = sheetView.Cells[num + i, num11].Text;
				array[6] = sheetView.Cells[num + i, num12].Text;
				array[7] = sheetView.Cells[num + i, num13].Text;
				array[8] = sheetView.Cells[num + i, num14].Text;
				array[9] = sheetView.Cells[num + i, num15].Text;
				array[10] = sheetView.Cells[num + i, num8].Text;
				array[11] = sheetView.Cells[num + i, num9].Text;
				array[12] = sheetView.Cells[num + i, column5].Text;
				array[13] = sheetView.Cells[num + i, column6].Text;
				array[14] = sheetView.Cells[num + i, column7].Text;
				array[15] = sheetView.Cells[num + i, column8].Text;
				array[16] = sheetView.Cells[num + i, column].Text;
				array[17] = sheetView.Cells[num + i, column2].Text;
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num2);
					HTKKFile.RemoveValidate(sheetView, num + i, num3);
					HTKKFile.RemoveValidate(sheetView, num + i, num10);
					HTKKFile.RemoveValidate(sheetView, num + i, num11);
					HTKKFile.RemoveValidate(sheetView, num + i, num12);
					HTKKFile.RemoveValidate(sheetView, num + i, num13);
					HTKKFile.RemoveValidate(sheetView, num + i, num14);
					HTKKFile.RemoveValidate(sheetView, num + i, num15);
					HTKKFile.RemoveValidate(sheetView, num + i, num8);
					HTKKFile.RemoveValidate(sheetView, num + i, num9);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num2);
				}
				string text5 = sheetView.Cells[num + i, num3].Text;
				string text6 = sheetView.Cells[num + i, num8].Text;
				string text7 = sheetView.Cells[num + i, num9].Text;
				DateTime result = default(DateTime);
				DateTime result2 = default(DateTime);
				bool flag = false;
				if (string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
					value = 2;
				}
				else if (!DateTime.TryParseExact(text6, "MM/yyyy", null, DateTimeStyles.None, out result))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
					value = 2;
				}
				else if (result > DateTime.ParseExact(PeriodInfo.DenThang, "MM/yyyy", CultureInfo.CurrentCulture) || result < DateTime.ParseExact(PeriodInfo.TuThang, "MM/yyyy", CultureInfo.CurrentCulture))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num8);
				}
				if (string.IsNullOrEmpty(text7))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				else if (!DateTime.TryParseExact(text7, "MM/yyyy", null, DateTimeStyles.None, out result2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
					value = 2;
				}
				else if (result2 < result)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
					value = 2;
				}
				else if (result2 > DateTime.ParseExact(PeriodInfo.DenThang, "MM/yyyy", CultureInfo.CurrentCulture))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
					value = 2;
				}
				else
				{
					if (flag)
					{
						HTKKFile.RemoveValidate(sheetView, num + i, num8);
					}
					HTKKFile.RemoveValidate(sheetView, num + i, num9);
				}
				if (!string.IsNullOrEmpty(text5))
				{
					if (!DateTime.TryParseExact(text5, "dd/MM/yyyy", null, DateTimeStyles.None, out result3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
					else if (result3 > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")));
						value = 2;
					}
					else if (DateTime.ParseExact(result3.ToString("MM/yyyy"), "MM/yyyy", CultureInfo.InvariantCulture) > DateTime.ParseExact(result.ToString("MM/yyyy"), "MM/yyyy", CultureInfo.InvariantCulture) && !result.Equals(DateTime.MinValue))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, num3);
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
					value = 2;
				}
				if (string.IsNullOrEmpty(array[9]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num15, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, num15);
				}
				HTKKFile.RemoveValidate(sheetView, num + i, num11);
				HTKKFile.RemoveValidate(sheetView, num + i, num12);
				HTKKFile.RemoveValidate(sheetView, num + i, num13);
				if (array[12].Equals("01"))
				{
					if (string.IsNullOrEmpty(array[5]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num11, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						value = 2;
					}
					else if (array[13].Equals("999"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num11, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
					if (string.IsNullOrEmpty(array[6]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num12, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						value = 2;
					}
					else if (array[14].Equals("99999"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num12, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
					if (string.IsNullOrEmpty(array[7]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num13, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
						value = 2;
					}
					else if (array[15].Equals("9999999"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num13, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
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
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			bool success = false;
			bool success2 = false;
			DateTime dateTime = default(DateTime);
			DateTime dateTime2 = default(DateTime);
			int column = HTKKFile.ColumnLettersToIndex("Y") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("Z") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				int column3 = HTKKFile.ColumnLettersToIndex("Q") - 1;
				int column4 = HTKKFile.ColumnLettersToIndex("S") - 1;
				dateTime = HTKKFormater.ConvertToDateTime("01/" + sheetView.Cells[i + num, column3].Text, out success);
				dateTime2 = HTKKFormater.ConvertToDateTime("01/" + sheetView.Cells[i + num, column4].Text, out success2);
				if (!success || !success2)
				{
					continue;
				}
				int num2 = (dateTime2.Year - dateTime.Year) * 12 + (dateTime2.Month - dateTime.Month) + 1;
				double num3 = 0.0;
				if (dateTime.Year > 2019)
				{
					mucGiamTruNPT = 4400000.0;
				}
				else
				{
					mucGiamTruNPT = 3600000.0;
				}
				string text = strNam;
				if (text.Equals("2019") && dateTime2.Year > dateTime.Year)
				{
					int num4 = 0;
					if (dateTime.Year == 2019)
					{
						num4 = 12 - dateTime.Month + 1;
					}
					int num5 = num2 - num4;
					num3 = num4 * 3600000 + num5 * 4400000;
				}
				else
				{
					num3 = (double)num2 * mucGiamTruNPT;
				}
				sheetView.Cells[i + num, column].Value = num2;
				sheetView.Cells[i + num, column2].Value = num3;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			int num6 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			success = false;
			success2 = false;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				int column5 = HTKKFile.ColumnLettersToIndex("T") - 1;
				int column6 = HTKKFile.ColumnLettersToIndex("U") - 1;
				dateTime = HTKKFormater.ConvertToDateTime("01/" + sheetView.Cells[i + num6, column5].Text, out success);
				dateTime2 = HTKKFormater.ConvertToDateTime("01/" + sheetView.Cells[i + num6, column6].Text, out success2);
				if (!success || !success2)
				{
					continue;
				}
				int num2 = (dateTime2.Year - dateTime.Year) * 12 + (dateTime2.Month - dateTime.Month) + 1;
				double num3 = 0.0;
				if (dateTime.Year > 2019)
				{
					mucGiamTruNPT = 4400000.0;
				}
				else
				{
					mucGiamTruNPT = 3600000.0;
				}
				string text = strNam;
				if (text.Equals("2019") && dateTime2.Year > dateTime.Year)
				{
					int num4 = 0;
					if (dateTime.Year == 2019)
					{
						num4 = 12 - dateTime.Month + 1;
					}
					int num5 = num2 - num4;
					num3 = num4 * 3600000 + num5 * 4400000;
				}
				else
				{
					num3 = (double)num2 * mucGiamTruNPT;
				}
				sheetView.Cells[i + num6, column].Value = num2;
				sheetView.Cells[i + num6, column2].Value = num3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
