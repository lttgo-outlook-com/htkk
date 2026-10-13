using System;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt95Nckt200
{
	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = 0;
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("D") - 1].Value))
			{
				if (string.IsNullOrEmpty(sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 10, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 10, HTKKFile.ColumnLettersToIndex("T") - 1);
				}
			}
			if (num == 2 && winFpSpread.TabStrip[0].BackColor == Color.Pink)
			{
				return false;
			}
			if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.Yellow)
			{
				if (num == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[0].BackColor = Color.Pink;
				}
				else if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
				{
					winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				}
				else if (num == 0)
				{
					winFpSpread.TabStrip[0].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return true;
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_11");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = !Convert.ToBoolean(valueCustom);
				if (!Convert.ToBoolean(valueCustom))
				{
					sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_11");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			sheetView.Cells[5, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = !Convert.ToBoolean(valueCustom);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
