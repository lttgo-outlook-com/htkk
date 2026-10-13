using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ph01TT80Nckt275
{
	private string DEFAULT_LOAI_TIEN_KHAC = "0,00";

	private int col2_Tb1 = HTKKFile.ColumnLettersToIndex("C");

	private int col3_Tb1 = HTKKFile.ColumnLettersToIndex("K");

	private int col4_Tb1 = HTKKFile.ColumnLettersToIndex("L");

	private int col5_Tb1 = HTKKFile.ColumnLettersToIndex("O");

	private int col6_Tb1 = HTKKFile.ColumnLettersToIndex("Q");

	private int col7_Tb1 = HTKKFile.ColumnLettersToIndex("T");

	private int col3Ma_Tb1 = HTKKFile.ColumnLettersToIndex("AE");

	private int colMaTm_Tb1 = HTKKFile.ColumnLettersToIndex("I");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
				string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
				string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[7];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
				HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col6_Tb1, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[2];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo2.RowsValues[i].RowValue[4];
				string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[4].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
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
			CapNhatLaiCellType(sView, ToKhaiInfo);
			int num = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			switch (num)
			{
			case 2023:
				if (num2 >= 12)
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
			if ((int)GeneralFunctions.KEY_PRESS != 116)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				int num = HTKKFile.ColumnLettersToIndex("AB");
				int num2 = HTKKFile.ColumnLettersToIndex("AE");
				int num3 = HTKKFile.ColumnLettersToIndex("C");
				int num4 = HTKKFile.ColumnLettersToIndex("T");
				int num5 = HTKKFile.ColumnLettersToIndex("AE");
				SheetView sheetView = winFpSpread.Sheets[0];
				double num6 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col4_Tb1 - 1].Value);
				double num7 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col5_Tb1 - 1].Value);
				double num8 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col6_Tb1 - 1].Value);
				double num9 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col7_Tb1 - 1].Value);
				double num10 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, num - 1].Value);
				if (iCol == col3_Tb1)
				{
					string text = sheetView.Cells[(int)iRow - 1, col3_Tb1 - 1].Value.ToString().Trim();
					HTKKCMCValidate validate = hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[4].Validate;
					HTKKCMCValidate validate2 = hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[6].Validate;
					HTKKCMCValidate validate3 = hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[7].Validate;
					GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col4_Tb1, num6.ToString(), validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col6_Tb1, num8.ToString(), validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col7_Tb1, num9.ToString(), validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[2];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				if ((iCol == col4_Tb1 || iCol == col5_Tb1) && iRow < dataLocationTemplateRowStart)
				{
					sheetView.Cells[(int)iRow - 1, col7_Tb1 - 1].Value = num10;
				}
				if (iRow < dataLocationTemplateRowStart)
				{
					TongTheoLoaiTien(winFpSpread, ToKhaiInfo);
				}
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[0]);
				GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, ToKhaiInfo.SheetInfos[0], 0);
			}
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
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			List<string> list3 = new List<string>();
			string[,] array = new string[hTKKCMCTableInfo.RowsValues.Count, 3];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num + i, col2_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col3_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col5_Tb1 - 1);
				string text = sheetView.Cells[num + i, col2_Tb1 - 1].Text;
				string text2 = sheetView.Cells[num + i, col3_Tb1 - 1].Text;
				string text3 = sheetView.Cells[num + i, col3Ma_Tb1 - 1].Text;
				string text4 = sheetView.Cells[num + i, col5_Tb1 - 1].Text;
				string text5 = ConvertHelpers.ConvertToString(sheetView.Cells[num + i, col4_Tb1 - 1].Value);
				string text6 = sheetView.Cells[num + i, colMaTm_Tb1 - 1].Text;
				string item = text6 + "~" + text3;
				string item2 = text + text2 + text4;
				if (!list.Contains(item2))
				{
					list.Add(item2);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col5_Tb1 - 1, message2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message3);
					value = 2;
				}
				if (!list2.Contains(text6))
				{
					list2.Add(text6);
				}
				if (!list3.Contains(item))
				{
					list3.Add(item);
				}
				array[i, 0] = text6;
				array[i, 1] = text3;
				array[i, 2] = text5;
			}
			int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int num3 = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (num2 >= 2023 && num2 <= 2026 && (num2 != 2023 || num3 >= 12))
			{
				Dictionary<string, double> dictionary = new Dictionary<string, double>();
				for (int i = 0; i < list3.Count; i++)
				{
					string[] array2 = list3[i].Split(new char[1] { '~' });
					double num4 = 0.0;
					if (array2 != null && array2.Length == 2)
					{
						for (int j = 0; j < array.GetLength(0); j++)
						{
							if (array2[0] == array[j, 0] && array2[1] == array[j, 1])
							{
								num4 += ConvertHelpers.ConvertToDouble(array[j, 2]);
							}
						}
					}
					dictionary.Add(list3[i], num4);
				}
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[2];
				int num5 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
				List<string> list4 = new List<string>();
				for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, col2_Tb1 - 1);
					HTKKFile.RemoveValidate(sheetView, num5 + i, col3_Tb1 - 1);
					HTKKFile.RemoveValidate(sheetView, num5 + i, col4_Tb1 - 1);
					string text = sheetView.Cells[num5 + i, col2_Tb1 - 1].Text;
					string text2 = sheetView.Cells[num5 + i, col3_Tb1 - 1].Text;
					string text3 = sheetView.Cells[num5 + i, col3Ma_Tb1 - 1].Text;
					string text6 = sheetView.Cells[num5 + i, colMaTm_Tb1 - 1].Text;
					double num6 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num5 + i, col4_Tb1 - 1].Value);
					string item2 = text + text2;
					if (!list4.Contains(item2))
					{
						list4.Add(item2);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col2_Tb1 - 1, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col3_Tb1 - 1, message4);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text) && string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col3_Tb1 - 1, message);
						value = 2;
					}
					if (string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col2_Tb1 - 1, message);
						value = 2;
					}
					if (string.IsNullOrEmpty(text6))
					{
						continue;
					}
					if (!list2.Contains(text6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col2_Tb1 - 1, message5);
						value = 2;
					}
					string key = text6 + "~" + text3;
					if (dictionary.ContainsKey(key))
					{
						double num7 = dictionary[key];
						if (num6 > num7)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col4_Tb1 - 1, message6);
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, col4_Tb1 - 1, message6);
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
			TongTheoLoaiTien(winFpSpread, ToKhaiInfo);
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

	public void TongTheoLoaiTien(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int num = HTKKFile.ColumnLettersToIndex("AB");
		int colMaFrom = HTKKFile.ColumnLettersToIndex("AE");
		int colTenTo = HTKKFile.ColumnLettersToIndex("C");
		int num2 = HTKKFile.ColumnLettersToIndex("T");
		int colMaTo = HTKKFile.ColumnLettersToIndex("AE");
		SheetView sheetView = winFpSpread.Sheets[0];
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		dictionary.Add(col7_Tb1, col7_Tb1);
		GeneralFunctions.CopyPaseTableToTable(hTKKCMCSheetInfo.Tables[0], hTKKCMCSheetInfo.Tables[1], sheetView, sheetView, dictionary, colMaFrom, colMaTo, col3_Tb1, colTenTo, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
	}

	public void CapNhatLaiCellType(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
			string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
			string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[7];
			string text = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart + i - 1, col3Ma_Tb1 - 1].Value);
			HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
			HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
			HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col6_Tb1, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
		{
			string value3 = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
			string text = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart2 + i - 1, col3Ma_Tb1 - 1].Value);
			HTKKCMCValidate validate4 = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col7_Tb1, value3, validate4, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo3 = ToKhaiInfo.SheetInfos[0].Tables[2];
		int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo3.RowsValues[i].RowValue[4];
			string text = ConvertHelpers.ConvertToString(sView.Cells[dataLocationTemplateRowStart3 + i - 1, col3Ma_Tb1 - 1].Value);
			HTKKCMCValidate validate = hTKKCMCTableInfo3.RowsInfo.Cells[4].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[0], 0);
	}
}
