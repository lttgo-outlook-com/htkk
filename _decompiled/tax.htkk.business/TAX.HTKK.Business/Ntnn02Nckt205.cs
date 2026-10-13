using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn02Nckt205
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

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			int num = HTKKFile.ColumnLettersToIndex("BG");
			int num2 = HTKKFile.ColumnLettersToIndex("BY");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			int num = HTKKFile.ColumnLettersToIndex("BG");
			int num2 = HTKKFile.ColumnLettersToIndex("BY");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			string formula3 = "BK32";
			int num = HTKKFile.ColumnLettersToIndex("BG");
			int num2 = HTKKFile.ColumnLettersToIndex("BY");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			int[] array3 = GeneralFunctions.ArrayCellId("AF_32");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
			winFpSpread.Sheets[0].Cells[array3[1], array3[0]].Formula = formula3;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Refresh_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int actionType)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			string formula3 = "BK32";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			int[] array3 = GeneralFunctions.ArrayCellId("AF_32");
			if (actionType == 2 && !winFpSpread.Sheets[2].Visible)
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[array3[1], array3[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array3[1], array3[0]].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
				winFpSpread.Sheets[0].Cells[array3[1], array3[0]].Formula = formula3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			if (iCol == num)
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			}
			if (iCol == num2)
			{
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK32";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_32");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Refresh_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int actionType)
	{
		try
		{
			string formula = "BK30";
			string formula2 = "BK31";
			string formula3 = "BK32";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AF_30");
			int[] array2 = GeneralFunctions.ArrayCellId("AF_31");
			int[] array3 = GeneralFunctions.ArrayCellId("AF_32");
			if (actionType == 2 && !winFpSpread.Sheets[1].Visible)
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[array3[1], array3[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array3[1], array3[0]].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
				winFpSpread.Sheets[0].Cells[array3[1], array3[0]].Formula = formula3;
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
			checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			checkPL2_1(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			checkPL2_2(winFpSpread.Sheets[2], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
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

	private void checkPL2_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo tinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			if (!sView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[1];
			int num2 = GeneralFunctions.ArrayCellId("O_23")[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, num2].Text;
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
				}
				else if (!string.IsNullOrEmpty(array[0]) & !ValidateBase.CheckMaSoThue(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
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
			throw ex;
		}
	}

	private void checkPL2_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo tinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			if (!sView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[11];
			int column = HTKKFile.ColumnLettersToIndex("E") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("L") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("R") - 1;
			int column3 = HTKKFile.ColumnLettersToIndex("Y") - 1;
			int column4 = HTKKFile.ColumnLettersToIndex("AG") - 1;
			int column5 = HTKKFile.ColumnLettersToIndex("AM") - 1;
			int column6 = HTKKFile.ColumnLettersToIndex("AR") - 1;
			int column7 = HTKKFile.ColumnLettersToIndex("AW") - 1;
			int column8 = HTKKFile.ColumnLettersToIndex("BD") - 1;
			int column9 = HTKKFile.ColumnLettersToIndex("BL") - 1;
			int column10 = HTKKFile.ColumnLettersToIndex("BS") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sView.Cells[num + i, column].Text;
				array[1] = sView.Cells[num + i, num2].Text;
				array[2] = sView.Cells[num + i, column2].Text;
				array[3] = sView.Cells[num + i, column3].Text;
				array[4] = sView.Cells[num + i, column4].Text;
				array[5] = sView.Cells[num + i, column5].Text;
				array[6] = sView.Cells[num + i, column6].Text;
				array[7] = sView.Cells[num + i, column7].Text;
				array[8] = sView.Cells[num + i, column8].Text;
				array[9] = sView.Cells[num + i, column9].Text;
				array[10] = sView.Cells[num + i, column10].Text;
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
				}
				else if (!string.IsNullOrEmpty(array[1]) & !ValidateBase.CheckMaSoThue(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[2].BackColor = Color.White;
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
