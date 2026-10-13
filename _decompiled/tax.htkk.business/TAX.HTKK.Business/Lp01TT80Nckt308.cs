using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Lp01TT80Nckt308
{
	private string DEFAULT_LOAI_TIEN_KHAC = "0,00";

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private int col2_Tb1 = HTKKFile.ColumnLettersToIndex("C");

	private int col3_Tb1 = HTKKFile.ColumnLettersToIndex("N");

	private int col4_Tb1 = HTKKFile.ColumnLettersToIndex("O");

	private int col3Ma_Tb1 = HTKKFile.ColumnLettersToIndex("V");

	private int colMaTm_Tb1 = HTKKFile.ColumnLettersToIndex("K");

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int iCol = HTKKFile.ColumnLettersToIndex("O");
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, iCol, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
				string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, iCol, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = ToKhaiInfo.SheetInfos[0].Tables[2];
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo3.RowsValues[i].RowValue[4];
				string text = hTKKCMCTableInfo3.RowsValues[i].RowValue[6];
				HTKKCMCValidate validate = hTKKCMCTableInfo3.RowsInfo.Cells[4].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, iCol, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			int num = HTKKFile.ColumnLettersToIndex("V");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			CapNhatLaiCellType(sView, ToKhaiInfo);
			int num3 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int num4 = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			switch (num3)
			{
			case 2023:
				if (num4 >= 12)
				{
					break;
				}
				goto default;
			default:
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[2];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				sView.Rows[dataLocationTemplateRowStart - 3].Visible = false;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sView.Rows[dataLocationTemplateRowStart + i - 1].Visible = false;
				}
				break;
			}
			case 2024:
			case 2025:
			case 2026:
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Invalid comparison between Unknown and I4
		try
		{
			if ((int)GeneralFunctions.KEY_PRESS == 116)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			if (iCol == col2_Tb1 || iCol == col3_Tb1)
			{
				List<int> list = new List<int>();
				list.Add(col2_Tb1);
				list.Add(col3_Tb1);
				List<string> list2 = new List<string>();
				list2.Add(sheetView.Cells[(int)iRow - 1, col2_Tb1 - 1].Text.Trim());
				list2.Add(sheetView.Cells[(int)iRow - 1, col3_Tb1 - 1].Text.Trim());
				List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
				if (rowIndexInSViewByCompares.Count > 1)
				{
					string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, col2_Tb1 - 1, message);
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, col3_Tb1 - 1, message);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, col2_Tb1 - 1);
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, col3_Tb1 - 1);
				}
			}
			if (iCol == col3_Tb1)
			{
				string value = sheetView.Cells[(int)iRow - 1, col4_Tb1 - 1].Value.ToString().Trim();
				string text = sheetView.Cells[(int)iRow - 1, col3Ma_Tb1 - 1].Value.ToString().Trim();
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				string defaultValue = ((text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col4_Tb1, value, validate, defaultValue);
			}
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
			{
				CongTheoLoaiTien(winFpSpread, ToKhaiInfo);
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
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
			return true;
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int[] cell = GeneralFunctions.ArrayCellId("E_22");
			string message = HTKKFile.GetMessage(sheetView, cell);
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			int num2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("N") - 1;
			int column = HTKKFile.ColumnLettersToIndex("AC") - 1;
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			string[,] array = new string[hTKKCMCTableInfo.RowsValues.Count, 3];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num + i, num2);
				HTKKFile.RemoveValidate(sheetView, num + i, num3);
				HTKKFile.RemoveValidate(sheetView, num + i, col2_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col3_Tb1 - 1);
				string text = sheetView.Cells[num + i, num2].Text;
				string text2 = sheetView.Cells[num + i, num3].Text;
				string text3 = sheetView.Cells[num + i, col3Ma_Tb1 - 1].Text;
				string text4 = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, col4_Tb1 - 1].Value);
				string text5 = sheetView.Cells[num + i, colMaTm_Tb1 - 1].Text;
				string item = text5 + "~" + text3;
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, message);
					value = 2;
				}
				List<int> list3 = new List<int>();
				list3.Add(col2_Tb1);
				list3.Add(col3_Tb1);
				List<string> list4 = new List<string>();
				list4.Add(sheetView.Cells[num + i, col2_Tb1 - 1].Text.Trim());
				list4.Add(sheetView.Cells[num + i, col3_Tb1 - 1].Text.Trim());
				List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list3, list4);
				if (rowIndexInSViewByCompares.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message2);
				}
				if (!list.Contains(text5))
				{
					list.Add(text5);
				}
				if (!list2.Contains(item))
				{
					list2.Add(item);
				}
				array[i, 0] = text5;
				array[i, 1] = text3;
				array[i, 2] = text4;
			}
			int num4 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int num5 = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (num4 >= 2023 && num4 <= 2026 && (num4 != 2023 || num5 >= 12))
			{
				Dictionary<string, double> dictionary = new Dictionary<string, double>();
				for (int i = 0; i < list2.Count; i++)
				{
					string[] array2 = list2[i].Split(new char[1] { '~' });
					double num6 = 0.0;
					if (array2 != null && array2.Length == 2)
					{
						for (int j = 0; j < array.GetLength(0); j++)
						{
							if (array2[0] == array[j, 0] && array2[1] == array[j, 1])
							{
								num6 += ConvertHelpers.ConvertToDouble(array[j, 2]);
							}
						}
					}
					dictionary.Add(list2[i], num6);
				}
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[2];
				int num7 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
				List<string> list5 = new List<string>();
				for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
				{
					HTKKFile.RemoveValidate(sheetView, num7 + i, col2_Tb1 - 1);
					HTKKFile.RemoveValidate(sheetView, num7 + i, col3_Tb1 - 1);
					HTKKFile.RemoveValidate(sheetView, num7 + i, col4_Tb1 - 1);
					string text = sheetView.Cells[num7 + i, col2_Tb1 - 1].Text;
					string text2 = sheetView.Cells[num7 + i, col3_Tb1 - 1].Text;
					string text3 = sheetView.Cells[num7 + i, col3Ma_Tb1 - 1].Text;
					string text6 = sheetView.Cells[num7 + i, column].Text;
					string text5 = sheetView.Cells[num7 + i, colMaTm_Tb1 - 1].Text;
					double num8 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num7 + i, col4_Tb1 - 1].Value);
					string item2 = text + text2;
					if (!list5.Contains(item2))
					{
						list5.Add(item2);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col2_Tb1 - 1, message3);
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col3_Tb1 - 1, message3);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text) && string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col3_Tb1 - 1, message);
						value = 2;
					}
					if (string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col2_Tb1 - 1, message);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text5))
					{
						if (!list.Contains(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col2_Tb1 - 1, message4);
							value = 2;
						}
						string key = text5 + "~" + text3;
						if (dictionary.ContainsKey(key))
						{
							double num9 = dictionary[key];
							if (num8 > num9)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col4_Tb1 - 1, message5);
								value = 2;
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col4_Tb1 - 1, message5);
							value = 2;
						}
					}
					if (text6 == "63012853" && text3 == "VND" && num8 % 115000.0 != 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + i, col4_Tb1 - 1, message6);
						value = 2;
					}
				}
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

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
		if ((int)GeneralFunctions.KEY_PRESS == 117)
		{
			CongTheoLoaiTien(winFpSpread, ToKhaiInfo);
		}
		return true;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		CapNhatLaiCellType(winFpSpread.Sheets[0], ToKhaiInfo);
	}

	public void CapNhatLaiCellType(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		int num = HTKKFile.ColumnLettersToIndex("V");
		int num2 = HTKKFile.ColumnLettersToIndex("O");
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string text = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value);
			string text2 = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value);
			HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
			string defaultValue = ((text2 == "" || text2 == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, num2, text.ToString(), validate, defaultValue);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
		{
			string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
			string text2 = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart2 + i - 1, num - 1].Value);
			HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, num2, text, validate, (text2 == "" || text2 == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo3 = ToKhaiInfo.SheetInfos[0].Tables[2];
		int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
		{
			string text = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart3 + i - 1, num2 - 1].Value);
			string text2 = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart3 + i - 1, num - 1].Value);
			HTKKCMCValidate validate = hTKKCMCTableInfo3.RowsInfo.Cells[4].Validate;
			string defaultValue = ((text2 == "" || text2 == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, num2, text.ToString(), validate, defaultValue);
		}
		GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[0], 0);
	}

	public void CongTheoLoaiTien(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int colTenFrom = HTKKFile.ColumnLettersToIndex("N");
		int num = HTKKFile.ColumnLettersToIndex("O");
		int num2 = HTKKFile.ColumnLettersToIndex("V");
		int colTenTo = HTKKFile.ColumnLettersToIndex("C");
		int num3 = HTKKFile.ColumnLettersToIndex("O");
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		dictionary.Add(num, num);
		GeneralFunctions.CopyPaseTableToTable(hTKKCMCSheetInfo.Tables[0], hTKKCMCSheetInfo.Tables[1], sheetView, sheetView, dictionary, num2, num2, colTenFrom, colTenTo, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
	}
}
