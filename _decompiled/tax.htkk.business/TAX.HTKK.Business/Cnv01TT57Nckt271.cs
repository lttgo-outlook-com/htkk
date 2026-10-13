using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Cnv01TT57Nckt271
{
	private int[] cellMst = GeneralFunctions.ArrayCellId("I_13");

	private int[] cellChuyenNhuongVon = GeneralFunctions.ArrayCellId("L_5");

	private int[] cellChuyenNhuongVonThoaThuan = GeneralFunctions.ArrayCellId("X_5");

	private int[] cellChuyenNhuongVonKhac = GeneralFunctions.ArrayCellId("X_7");

	private int[] cellChuyenNhuongMuaCP = GeneralFunctions.ArrayCellId("L_9");

	private int[] cellChuyenNhuongQuyenGV = GeneralFunctions.ArrayCellId("X_9");

	private int[] cellCt09 = GeneralFunctions.ArrayCellId("U_52");

	private int[] cellCt10 = GeneralFunctions.ArrayCellId("U_53");

	private int[] cellCt11 = GeneralFunctions.ArrayCellId("U_54");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellMst[1], cellMst[0]].Value = PeriodInfo.MaHSDuAnDT;
			bool flag = Convert.ToBoolean(sView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value);
			if (flag)
			{
				sView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value = false;
				sView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value = false;
				sView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Locked = false;
				sView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Locked = false;
			}
			bool flag2 = Convert.ToBoolean(sView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value);
			bool flag3 = Convert.ToBoolean(sView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value);
			if (flag)
			{
				sView.Cells[cellCt09[1], cellCt09[0]].Locked = false;
			}
			if (flag2)
			{
				sView.Cells[cellCt10[1], cellCt10[0]].Locked = false;
			}
			if (flag3)
			{
				sView.Cells[cellCt11[1], cellCt11[0]].Locked = false;
			}
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Locked = true;
				sView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Locked = true;
				sView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Locked = true;
				sView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Locked = true;
				sView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Locked = true;
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
			bool flag4 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value);
			bool flag5 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]);
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]);
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]);
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]);
			HTKKFile.RemoveValidate(sheetView, cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]);
			if (flag && ((flag4 && flag5) || (!flag4 && !flag5)))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0], text2);
				HTKKFile.SetNodeBackGroundPink(sheetView, cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0], text2);
				value = 2;
			}
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
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Locked = false;
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Locked = false;
					sheetView.Cells[cellCt10[1], cellCt10[0]].Value = 0;
					sheetView.Cells[cellCt11[1], cellCt11[0]].Value = 0;
				}
				else
				{
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Locked = true;
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Locked = true;
				}
			}
			else if (idxRow == cellChuyenNhuongMuaCP[1] && idxCol == cellChuyenNhuongMuaCP[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value))
				{
					sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value = false;
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Locked = true;
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Locked = true;
					sheetView.Cells[cellCt09[1], cellCt09[0]].Value = 0;
					sheetView.Cells[cellCt11[1], cellCt11[0]].Value = 0;
				}
			}
			else if (idxRow == cellChuyenNhuongQuyenGV[1] && idxCol == cellChuyenNhuongQuyenGV[0] && Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value))
			{
				sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value = false;
				sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value = false;
				sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value = false;
				sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value = false;
				sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Locked = true;
				sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Locked = true;
				sheetView.Cells[cellCt09[1], cellCt09[0]].Value = 0;
				sheetView.Cells[cellCt10[1], cellCt10[0]].Value = 0;
			}
			if (idxRow == cellChuyenNhuongVonThoaThuan[1] && idxCol == cellChuyenNhuongVonThoaThuan[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value))
				{
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value = false;
				}
				else
				{
					sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value = true;
				}
			}
			if (idxRow == cellChuyenNhuongVonKhac[1] && idxCol == cellChuyenNhuongVonKhac[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVonKhac[1], cellChuyenNhuongVonKhac[0]].Value))
				{
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value = false;
				}
				else
				{
					sheetView.Cells[cellChuyenNhuongVonThoaThuan[1], cellChuyenNhuongVonThoaThuan[0]].Value = true;
				}
			}
			flag = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongVon[1], cellChuyenNhuongVon[0]].Value);
			flag2 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongMuaCP[1], cellChuyenNhuongMuaCP[0]].Value);
			flag3 = Convert.ToBoolean(sheetView.Cells[cellChuyenNhuongQuyenGV[1], cellChuyenNhuongQuyenGV[0]].Value);
			sheetView.Cells[cellCt09[1], cellCt09[0]].Locked = true;
			sheetView.Cells[cellCt10[1], cellCt10[0]].Locked = true;
			sheetView.Cells[cellCt11[1], cellCt11[0]].Locked = true;
			if (flag)
			{
				sheetView.Cells[cellCt09[1], cellCt09[0]].Locked = false;
			}
			if (flag2)
			{
				sheetView.Cells[cellCt10[1], cellCt10[0]].Locked = false;
			}
			if (flag3)
			{
				sheetView.Cells[cellCt11[1], cellCt11[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
