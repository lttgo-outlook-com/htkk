using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class QtCnv02TT57Nckt271
{
	private int[] cellMst = GeneralFunctions.ArrayCellId("I_11");

	private int[] cellChuyenNhuongVon = GeneralFunctions.ArrayCellId("L_5");

	private int[] cellChuyenNhuongMuaCP = GeneralFunctions.ArrayCellId("X_5");

	private int[] cellChuyenNhuongQuyenGV = GeneralFunctions.ArrayCellId("T_7");

	private int[] cellCt11 = GeneralFunctions.ArrayCellId("U_52");

	private int[] cellCt12 = GeneralFunctions.ArrayCellId("U_53");

	private int[] cellCt13 = GeneralFunctions.ArrayCellId("U_54");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellMst[1], cellMst[0]].Value = PeriodInfo.MaHSDuAnDT;
			bool flag = Convert.ToBoolean(sView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value);
			bool flag2 = Convert.ToBoolean(sView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value);
			bool flag3 = Convert.ToBoolean(sView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value);
			if (flag)
			{
				sView.Cells[cellCt11[1], cellCt11[0]].Locked = false;
			}
			if (flag2)
			{
				sView.Cells[cellCt12[1], cellCt12[0]].Locked = false;
			}
			if (flag3)
			{
				sView.Cells[cellCt13[1], cellCt13[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]);
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]);
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]);
			int num = 0;
			if (flag)
			{
				num++;
			}
			if (flag2)
			{
				num++;
			}
			if (flag3)
			{
				num++;
			}
			if (num != 1)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cellChuyenNhuongVon[1], cellChuyenNhuongVon[0], text);
				HTKKFile.SetNodeBackGroundPink(sheetView, cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0], text);
				HTKKFile.SetNodeBackGroundPink(sheetView, cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0], text);
				value = 2;
			}
			switch (value)
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
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			if (idxRow == cellChuyenNhuongVon[1] && idxCol == cellChuyenNhuongVon[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value))
				{
					sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value = false;
					sheetView.Cells[cellCt12[1], cellCt12[0]].Value = 0;
					sheetView.Cells[cellCt13[1], cellCt13[0]].Value = 0;
				}
			}
			else if (idxRow == cellChuyenNhuongMuaCP[1] && idxCol == cellChuyenNhuongMuaCP[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value))
				{
					sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value = false;
					sheetView.Cells[cellCt11[1], cellCt11[0]].Value = 0;
					sheetView.Cells[cellCt13[1], cellCt13[0]].Value = 0;
				}
			}
			else if (idxRow == cellChuyenNhuongQuyenGV[1] && idxCol == cellChuyenNhuongQuyenGV[0] && Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value))
			{
				sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value = false;
				sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value = false;
				sheetView.Cells[cellCt11[1], cellCt11[0]].Value = 0;
				sheetView.Cells[cellCt12[1], cellCt12[0]].Value = 0;
			}
			flag = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value);
			flag2 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value);
			flag3 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value);
			sheetView.Cells[cellCt11[1], cellCt11[0]].Locked = true;
			sheetView.Cells[cellCt12[1], cellCt12[0]].Locked = true;
			sheetView.Cells[cellCt13[1], cellCt13[0]].Locked = true;
			if (flag)
			{
				sheetView.Cells[cellCt11[1], cellCt11[0]].Locked = false;
			}
			if (flag2)
			{
				sheetView.Cells[cellCt12[1], cellCt12[0]].Locked = false;
			}
			if (flag3)
			{
				sheetView.Cells[cellCt13[1], cellCt13[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
