using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TdGtgt01Nckt201
{
	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			checkPL1_2TD(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
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
			throw ex;
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("U"))
			{
				double num = 0.0;
				bool flag = false;
				int num2 = -1;
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					if (iRow == hTKKCMCTableInfo.Header.StaticCells[5].CellTemplate.RowIndex)
					{
						num2 = i;
						flag = true;
					}
					num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
				}
				if (flag)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_2_TD_GTGT/tongCong"));
					winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2];
					int num3 = ((ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count - 1 > num2) ? (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2 + 1].Location.DataLocationTemplateRowStart - 1) : (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].RowsValues.Count));
					for (int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num4 < num3; num4++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num6 * (num5 / 100.0);
					}
				}
			}
			else if (iCol == HTKKFile.ColumnLettersToIndex("R"))
			{
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					double num7 = 0.0;
					int num2 = -1;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					int num3 = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].RowsValues.Count;
					for (int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num4 < num3; num4++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						num7 += num5;
						if (iRow == num4)
						{
							num2 = i;
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num6 * (num5 / 100.0);
						}
					}
					if (num2 > -1)
					{
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].Header.StaticCells[4].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Value = num7;
					}
				}
			}
			int num8 = HTKKFile.ColumnLettersToIndex("P");
			int num9 = HTKKFile.ColumnLettersToIndex("AJ");
			if (iCol == num8 || iCol == num9)
			{
				LayThemThongTinCQTChiCuc(winFpSpread.Sheets[1], (int)iRow);
				LayThemThongTinTinhHuyen_PL_01_2_TD(winFpSpread.Sheets[1], (int)iRow);
			}
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sView = winFpSpread.Sheets[1];
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

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
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

	public void checkPL1_2TD(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
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
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
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
				int num = 0;
				for (int i = 0; i < tables.Count; i++)
				{
					hTKKCMCStaticOfTable = tables[i].Header;
					hTKKCMCRowInfo = tables[i].RowsInfo;
					num = tables[i].Location.DataLocationTemplateRowStart - 1;
					array[1] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1].Value);
					array[2] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1].Value);
					array[4] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.ColIndex - 1].Value);
					array[5] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1].Value);
					bool flag = GeneralFunctions.HaveToValidateHeaderRow(header, array);
					bool flag2 = false;
					double num2 = 0.0;
					for (int j = 0; j < tables[i].RowsValues.Count; j++)
					{
						array2[1] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Value);
						array2[2] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1].Value);
						array2[3] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1].Value);
						array2[4] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1].Value);
						array2[5] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1].Value);
						array2[6] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[6].CellTemplate.ColIndex - 1].Value);
						bool visible = sView.Cells[num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Row.Visible;
						if (GeneralFunctions.HaveToValidateRow(hTKKCMCRowInfo, array2) && visible)
						{
							flag2 = true;
						}
						else
						{
							if (!flag)
							{
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidateBackGround(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
								continue;
							}
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidateBackGround(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
						if (!string.IsNullOrEmpty(array2[2]))
						{
							if (!ValidateBase.CheckMaSoThue(array2[2]))
							{
								HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
						}
						if (string.IsNullOrEmpty(array2[3]))
						{
							HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
						}
						num2 += ConvertHelpers.ConvertToDouble(array2[4]);
						if (ConvertHelpers.ConvertToDouble(array2[5]) == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidateBackGround(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
					}
					if (!flag2 && !flag)
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
						continue;
					}
					if (num2 != 100.0)
					{
						for (int j = 0; j < tables[i].RowsValues.Count; j++)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
							value = 2;
						}
					}
					else
					{
						for (int j = 0; j < tables[i].RowsValues.Count; j++)
						{
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
						}
					}
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
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
							HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
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
					if (ConvertHelpers.ConvertToDouble(array[5]) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
					}
				}
			}
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				string maCQTCapCuc = headerInfo.maCQTCapCuc;
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				int colCheck = HTKKFile.ColumnLettersToIndex("AH");
				int num3 = HTKKFile.ColumnLettersToIndex("P");
				string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30"));
				int rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare2(sView, dataLocationTemplateRowStart, colCheck, maCQTCapCuc);
				if (rowIndexInSViewByCompare < 0)
				{
					for (int k = 0; k < table.RowsValues.Count; k++)
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + k - 1, num3 - 1, message);
					}
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

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					LayThemThongTinCQTChiCuc(sView, dataLocationTemplateRowStart + j);
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
			int index = 1;
			SheetView sView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					LayThemThongTinCQTChiCuc(sView, dataLocationTemplateRowStart + j);
					LayThemThongTinTinhHuyen_PL_01_2_TD(sView, dataLocationTemplateRowStart + j);
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
					GeneralFunctions.UpdateFromSViewToSInfo(sView, dataLocationTemplateRowStart + j, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
				}
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

	public void LayThemThongTinCQTChiCuc(SheetView sView, int iRow)
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
}
