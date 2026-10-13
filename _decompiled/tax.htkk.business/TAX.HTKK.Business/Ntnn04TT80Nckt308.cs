using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn04TT80Nckt308
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("BE_10");
			int[] array2 = GeneralFunctions.ArrayCellId("BD_10");
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[array[1], array[0]].Text = PeriodInfo.TenQuyetToan;
				sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaQuyetToan;
			}
			int[] array3 = GeneralFunctions.ArrayCellId("N_23");
			int[] array4 = GeneralFunctions.ArrayCellId("AH_23");
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[array3[1], array3[0]].Locked = true;
				sView.Cells[array4[1], array4[0]].Locked = true;
			}
			string[] array5 = PeriodInfo.DenNgay.Split(new char[1] { '/' });
			if (array5.Length == 3 && PeriodInfo.Nam != array5[2])
			{
				PeriodInfo.Nam = array5[2];
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			string formula = "BD32";
			string formula2 = "BD33";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AB_32");
			int[] array2 = GeneralFunctions.ArrayCellId("AB_33");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string formula = "BD32";
			string formula2 = "BD33";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AB_32");
			int[] array2 = GeneralFunctions.ArrayCellId("AB_33");
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

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string formula = "BD32";
			string formula2 = "BD33";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AB_32");
			int[] array2 = GeneralFunctions.ArrayCellId("AB_33");
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
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
			string formula = "BD32";
			string formula2 = "BD33";
			int num = HTKKFile.ColumnLettersToIndex("BD");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int[] array = GeneralFunctions.ArrayCellId("AB_32");
			int[] array2 = GeneralFunctions.ArrayCellId("AB_33");
			if (actionType == 2)
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].ResetFormula();
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Formula = formula;
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Formula = formula2;
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
			checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			checkPL4_1(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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

	private void checkPL4_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
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
			string text = "";
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			string text6 = "";
			string text7 = "";
			string text8 = "";
			double num2 = 0.0;
			double num3 = 0.0;
			int num4 = GeneralFunctions.ArrayCellId("E_23")[0];
			int num5 = GeneralFunctions.ArrayCellId("L_23")[0];
			int num6 = GeneralFunctions.ArrayCellId("R_23")[0];
			int num7 = GeneralFunctions.ArrayCellId("Y_23")[0];
			int num8 = GeneralFunctions.ArrayCellId("AG_23")[0];
			int num9 = GeneralFunctions.ArrayCellId("AM_23")[0];
			int num10 = GeneralFunctions.ArrayCellId("AR_23")[0];
			int num11 = GeneralFunctions.ArrayCellId("AW_23")[0];
			int num12 = GeneralFunctions.ArrayCellId("BD_23")[0];
			int num13 = GeneralFunctions.ArrayCellId("BS_23")[0];
			int col = GeneralFunctions.ArrayCellId("CC_23")[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sView.Cells[num + i, num4].Text;
				text2 = sView.Cells[num + i, num5].Text;
				text3 = sView.Cells[num + i, num6].Text;
				text4 = sView.Cells[num + i, num7].Text;
				text5 = sView.Cells[num + i, num8].Text;
				text6 = sView.Cells[num + i, num9].Text;
				text7 = sView.Cells[num + i, num10].Text;
				text8 = sView.Cells[num + i, num11].Text;
				num2 = ConvertHelpers.ConvertToDouble(sView.Cells[num + i, num12].Value);
				num3 = ConvertHelpers.ConvertToDouble(sView.Cells[num + i, num13].Value);
				string valueCheckTrung = ConvertHelpers.ConvertCellValueToText(sView, num + i, col);
				HTKKFile.RemoveValidate(sView, num + i, num4);
				HTKKFile.RemoveValidate(sView, num + i, num5);
				HTKKFile.RemoveValidate(sView, num + i, num6);
				HTKKFile.RemoveValidate(sView, num + i, num7);
				HTKKFile.RemoveValidate(sView, num + i, num8);
				HTKKFile.RemoveValidate(sView, num + i, num9);
				HTKKFile.RemoveValidate(sView, num + i, num10);
				HTKKFile.RemoveValidate(sView, num + i, num11);
				HTKKFile.RemoveValidate(sView, num + i, num12);
				HTKKFile.RemoveValidate(sView, num + i, num13);
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (!string.IsNullOrEmpty(text4) && !ValidateBase.CheckMaSoThue(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text5))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num8, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num9, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(text7, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text7))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else if (dateTime == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				else if (dateTime.CompareTo(DateTime.Today) >= 0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30")));
					value = 2;
				}
				if (string.IsNullOrEmpty(text8))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text8, "dd/MM/yyyy");
					if (dateTime2 == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
						value = 2;
					}
					else if (dateTime2.CompareTo(DateTime.Today) >= 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30")));
						value = 2;
					}
					else if (dateTime2.CompareTo(dateTime) < 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_31")));
						value = 2;
					}
				}
				if (num2 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num12, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (num3 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num13, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				int num14 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[12].Equals(valueCheckTrung.Trim()));
				if (num14 > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")), i + 1);
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")), i + 1);
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num8, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")), i + 1);
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num9, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")), i + 1);
					value = 2;
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")), i + 1);
					value = 2;
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")), i + 1);
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
			throw ex;
		}
	}
}
