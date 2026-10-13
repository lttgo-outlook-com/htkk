using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn03Nckt202
{
	private string cellId_ct10 = "N_11";

	private string cellId_ct11 = "N_13";

	private string cellId_ct22d = "V_57";

	private string cellId_ct23d = "V_58";

	private string cellId_ct24d = "V_59";

	private string cellId_ct25d = "V_60";

	private int[] CellId_ct10 => GeneralFunctions.ArrayCellId(cellId_ct10);

	private int[] CellId_ct11 => GeneralFunctions.ArrayCellId(cellId_ct11);

	private int[] CellId_ct22d => GeneralFunctions.ArrayCellId(cellId_ct22d);

	private int[] CellId_ct23d => GeneralFunctions.ArrayCellId(cellId_ct23d);

	private int[] CellId_ct24d => GeneralFunctions.ArrayCellId(cellId_ct24d);

	private int[] CellId_ct25d => GeneralFunctions.ArrayCellId(cellId_ct25d);

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.KieuKy == "D")
			{
				int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
				int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				for (int i = num; i <= num2; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Locked = true;
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
			int[] cellChangeID = new int[2]
			{
				(int)iCol - 1,
				(int)iRow - 1
			};
			List<int[]> list = new List<int[]>();
			list.Add(GeneralFunctions.ArrayCellId("BK_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("BU_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("BX_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("CG_" + iRow));
			List<int[]> list2 = list;
			XuLyCellChange(winFpSpread.Sheets[0], ToKhaiInfo, cellChangeID, list2[0], list2[1], list2[2], list2[3]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string value = GeneralFunctions.GetValue(winFpSpread.Sheets[0], "CG_26");
			int[] array = GeneralFunctions.ArrayCellId("CF_26");
			string text = MoneyHelpers.MoneyToText(value);
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Text = text;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.KieuKy == "D")
			{
				int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
				int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				for (int i = num; i <= num2; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("BE") - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void XuLyCellChange(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, int[] cellChangeID, int[] ct6, int[] ct7, int[] ct8, int[] ct9)
	{
		try
		{
			string formula = string.Format("IF(BK{0}*(BU{0}/100)-BX{0}<0,0,ROUND(BK{0}*(BU{0}/100)-BX{0},0))", ConvertHelpers.ConvertToInt(cellChangeID[1]) + 1);
			if (ct6[0] == cellChangeID[0] && ct6[1] == cellChangeID[1])
			{
				sView.Cells[ct9[1], ct9[0]].Formula = formula;
			}
			else if (ct7[0] == cellChangeID[0] && ct7[1] == cellChangeID[1])
			{
				sView.Cells[ct9[1], ct9[0]].Formula = formula;
			}
			else if (ct8[0] == cellChangeID[0] && ct8[1] == cellChangeID[1])
			{
				sView.Cells[ct9[1], ct9[0]].Formula = formula;
			}
			else if (ct9[0] == cellChangeID[0] && ct9[1] == cellChangeID[1])
			{
				sView.Cells[ct9[1], ct9[0]].Formula = "";
			}
			if (ConvertHelpers.ConvertToDouble(sView.Cells[ct6[1], ct6[0]].Value) < 0.0)
			{
				sView.Cells[ct8[1], ct8[0]].ResetFormula();
				sView.Cells[ct9[1], ct9[0]].ResetFormula();
				sView.Cells[ct8[1], ct8[0]].Value = 0;
				sView.Cells[ct9[1], ct9[0]].Value = 0;
				sView.Cells[ct8[1], ct8[0]].Locked = true;
				sView.Cells[ct9[1], ct9[0]].Locked = true;
			}
			else
			{
				sView.Cells[ct8[1], ct8[0]].Locked = false;
				sView.Cells[ct9[1], ct9[0]].Locked = false;
			}
			string value = GeneralFunctions.GetValue(sView, "CG_26");
			int[] array = GeneralFunctions.ArrayCellId("CF_26");
			string text = MoneyHelpers.MoneyToText(value);
			sView.Cells[array[1], array[0]].Text = text;
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
			checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[9];
			int num2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("AD") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("AL") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("AU") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("BE") - 1;
			int num7 = HTKKFile.ColumnLettersToIndex("BK") - 1;
			int num8 = HTKKFile.ColumnLettersToIndex("BU") - 1;
			int num9 = HTKKFile.ColumnLettersToIndex("BX") - 1;
			int num10 = HTKKFile.ColumnLettersToIndex("CG") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num2].Value);
				array[1] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num3].Value);
				array[2] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num4].Value);
				array[3] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num5].Value);
				array[4] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num6].Value);
				array[5] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num7].Value);
				array[6] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num8].Value);
				array[7] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num9].Value);
				array[8] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num10].Value);
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
					HTKKFile.RemoveValidate(sView, num + i, num3);
					HTKKFile.RemoveValidate(sView, num + i, num6);
					HTKKFile.RemoveValidate(sView, num + i, num10);
					HTKKFile.RemoveValidate(sView, num + i, num5);
					HTKKFile.RemoveValidate(sView, num + i, num4);
					HTKKFile.RemoveValidate(sView, num + i, num6);
					HTKKFile.RemoveValidate(sView, num + i, num7);
					HTKKFile.RemoveValidate(sView, num + i, num8);
					HTKKFile.RemoveValidate(sView, num + i, num9);
					continue;
				}
				HTKKFile.RemoveValidate(sView, num + i, num2);
				HTKKFile.RemoveValidate(sView, num + i, num3);
				HTKKFile.RemoveValidate(sView, num + i, num6);
				HTKKFile.RemoveValidate(sView, num + i, num10);
				HTKKFile.RemoveValidate(sView, num + i, num5);
				HTKKFile.RemoveValidate(sView, num + i, num4);
				HTKKFile.RemoveValidate(sView, num + i, num6);
				HTKKFile.RemoveValidate(sView, num + i, num7);
				HTKKFile.RemoveValidate(sView, num + i, num8);
				HTKKFile.RemoveValidate(sView, num + i, num10);
				HTKKFile.RemoveValidate(sView, num + i, num9);
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
				}
				if (!string.IsNullOrEmpty(array[1]) & !ValidateBase.CheckMaSoThue(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num3);
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num4);
				}
				if (string.IsNullOrEmpty(array[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(array[4].Replace("/", ""), "ddMMyyyy");
					if (!ValidateBase.ValidateFullDate(array[4], "ddMMyyyy"))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					else if (dateTime > DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture) || dateTime < DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, num + i, num6);
					}
				}
				if (ConvertHelpers.ConvertToDouble(array[8]) != 0.0 && ConvertHelpers.ConvertToDouble(array[8]) > ConvertHelpers.ConvertToDouble(array[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_22")));
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDouble(array[8]) != Math.Round(ConvertHelpers.ConvertToDouble(array[5]) * ConvertHelpers.ConvertToDouble(array[6]) / 100.0, 0) - ConvertHelpers.ConvertToDouble(array[7]))
				{
					if (ConvertHelpers.ConvertToDouble(array[8]) == 0.0 && Math.Round(ConvertHelpers.ConvertToDouble(array[5]) * ConvertHelpers.ConvertToDouble(array[6]) / 100.0, 0) - ConvertHelpers.ConvertToDouble(array[7]) < 0.0)
					{
						HTKKFile.RemoveValidate(sView, num + i, num10);
					}
					else
					{
						HTKKFile.SetNodeBackGroundYellow(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num10);
				}
				if (Math.Round(ConvertHelpers.ConvertToDouble(array[5]) * ConvertHelpers.ConvertToDouble(array[6]) / 100.0, 0) - ConvertHelpers.ConvertToDouble(array[7]) < 0.0 && ConvertHelpers.ConvertToDouble(array[7]) != 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num9, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num9);
				}
				double num11 = Convert.ToDouble(array[5]);
				double num12 = Convert.ToDouble(array[6]) / 100.0;
				double num13 = Convert.ToDouble(array[7]);
				if (num13 > num11 * num12)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				if (!string.IsNullOrEmpty(array[2].Replace("0", "")) && !string.IsNullOrEmpty(array[4].Replace("0", "")))
				{
					continue;
				}
				if (string.IsNullOrEmpty(array[3].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, num5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[5].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, num7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[6].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, num8, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[7].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, num9, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[8].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
					if (value == 0)
					{
						value = 1;
					}
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
			throw ex;
		}
	}
}
