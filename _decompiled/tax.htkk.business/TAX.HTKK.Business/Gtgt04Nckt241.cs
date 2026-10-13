using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt04Nckt241
{
	public void AutoBindData_0(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) >= 2021)
			{
				if (PeriodInfo.ThuHo.Equals("TH"))
				{
					sView.Cells[4, HTKKFile.ColumnLettersToIndex("K") - 1].Value = true;
					sView.Cells[42, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = "0";
					sView.Cells[42, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = true;
					sView.Cells[44, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = "0";
					sView.Cells[44, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = true;
					sView.Cells[45, HTKKFile.ColumnLettersToIndex("V") - 1].Text = "10%";
					sView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q46*10/100,0)";
					sView.Cells[45, HTKKFile.ColumnLettersToIndex("X") - 1].Text = "[29]=[28]x10%";
				}
				else
				{
					CheckBoxCellType cellType = new CheckBoxCellType();
					sView.Cells[4, HTKKFile.ColumnLettersToIndex("K") - 1].CellType = cellType;
					sView.Cells[4, HTKKFile.ColumnLettersToIndex("K") - 1].Value = false;
					sView.Cells[42, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = false;
					sView.Cells[44, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = false;
					sView.Cells[45, HTKKFile.ColumnLettersToIndex("V") - 1].Text = "2%";
					sView.Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q46*2/100,0)";
					sView.Cells[45, HTKKFile.ColumnLettersToIndex("X") - 1].Text = "[29]=[28]x2%";
				}
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
			bool result = true;
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
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else
				{
					winFpSpread.TabStrip[key].BackColor = Color.White;
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

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}
}
