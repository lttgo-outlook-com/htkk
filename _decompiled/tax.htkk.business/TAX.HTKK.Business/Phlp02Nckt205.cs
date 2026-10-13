using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Phlp02Nckt205
{
	private string fmlCol7 = "IF(H{0}-N{0}<0,0,ROUND(H{0}-N{0},0))";

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_E => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_H => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_Q => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_T => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_N => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_X => HTKKFile.ColumnLettersToIndex("X") - 1;

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow >= dataLocationTemplateRowStart && iRow <= num && (iCol - 1 == idxCol_H || iCol - 1 == idxCol_L))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_Q].Formula = string.Format(fmlCol7, iRow.ToString());
			}
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
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow >= dataLocationTemplateRowStart && iRow <= num && iCol - 1 == idxCol_Q)
			{
				sheetView.Cells[(int)iRow - 1, idxCol_Q].ResetFormula();
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
			SheetView sheetView = winFpSpread.Sheets[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				num = hTKKCMCTableInfo.RowsValues.Count;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_C].Value);
					array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_E].Value);
					array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_H].Value);
					array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_L].Value);
					array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_N].Value);
					array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_Q].Value);
					array[6] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_T].Value);
					array[7] = ConvertHelpers.ConvertToString(sheetView.Cells[num2 + j, idxCol_X].Value);
					if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
					{
						HTKKFile.RemoveValidate(sheetView, num2 + j, idxCol_C);
					}
					else if (string.IsNullOrEmpty(array[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + j, idxCol_C, message);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num2 + j, idxCol_C);
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
			throw ex;
		}
	}
}
