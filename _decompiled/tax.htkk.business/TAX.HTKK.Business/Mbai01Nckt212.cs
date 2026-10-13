using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Mbai01Nckt212
{
	private string cellId_CT22_ct4 = "P_45";

	private string cellId_CT22_ct5 = "U_45";

	private string cellId_CT22_ct6 = "Y_45";

	private int[] CellId_CT22_ct4 => GeneralFunctions.ArrayCellId(cellId_CT22_ct4);

	private int[] CellId_CT22_ct5 => GeneralFunctions.ArrayCellId(cellId_CT22_ct5);

	private int[] CellId_CT22_ct6 => GeneralFunctions.ArrayCellId(cellId_CT22_ct6);

	private int idxCT23_ColC => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCT23_ColP => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCT23_ColU => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxCT23_ColY => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			XuLyCT22_ct5(sheetView);
			int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
			int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
			for (int i = num; i <= num2; i++)
			{
				double num3 = Convert.ToDouble(sheetView.Cells[i, idxCT23_ColP].Value);
				if (num3 > 0.0)
				{
					sheetView.Cells[i, idxCT23_ColU].Value = 4;
				}
				else
				{
					sheetView.Cells[i, idxCT23_ColU].Value = 0;
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iRow - 1 == CellId_CT22_ct4[1] && iCol - 1 == CellId_CT22_ct4[0])
			{
				XuLyCT22_ct5(sheetView);
				XuLyCT22_ct6(sheetView);
			}
			int num = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1;
			if (iCol - 1 == idxCT23_ColP && iRow - 1 >= num && iRow - 1 <= num2)
			{
				double num3 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, idxCT23_ColP].Value);
				if (num3 > 0.0)
				{
					sheetView.Cells[(int)iRow - 1, idxCT23_ColU].Value = 4;
					sheetView.Cells[(int)iRow - 1, idxCT23_ColY].Value = 1000000;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCT23_ColU].Value = 0;
					sheetView.Cells[(int)iRow - 1, idxCT23_ColY].Value = 0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		if (winFpSpread.TabStrip[0].BackColor != Color.White)
		{
			return false;
		}
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			int num2 = Convert.ToInt32(sheetView.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value);
			double num3 = Convert.ToDouble(sheetView.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value);
			double num4 = 0.0;
			switch (num2)
			{
			case 1:
				num4 = 3000000.0;
				break;
			case 2:
				num4 = 2000000.0;
				break;
			case 3:
				num4 = 1500000.0;
				break;
			case 4:
				num4 = 1000000.0;
				break;
			}
			if (num3 != num4 && num3 != num4 / 2.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_CT22_ct6[1], CellId_CT22_ct6[0], message2);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_CT22_ct6[1], CellId_CT22_ct6[0]);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num5 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num5 + i, idxCT23_ColC].Text;
				array[1] = sheetView.Cells[num5 + i, idxCT23_ColP].Text;
				array[3] = sheetView.Cells[num5 + i, idxCT23_ColY].Value.ToString();
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColC);
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColY);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, idxCT23_ColC, message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColC);
				}
				double num6 = ConvertHelpers.ConvertToDouble(array[3]);
				if (num6 != 1000000.0 && num6 != 500000.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, idxCT23_ColY, message2);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColY);
				}
			}
			switch (num)
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

	private void XuLyCT22_ct5(SheetView sViewTK)
	{
		try
		{
			double num = Convert.ToDouble(sViewTK.Cells[CellId_CT22_ct4[1], CellId_CT22_ct4[0]].Value);
			double num2 = 10000000000.0;
			double num3 = 5000000000.0;
			double num4 = 2000000000.0;
			if (num > num2)
			{
				sViewTK.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value = 1;
			}
			else if (num > num3)
			{
				sViewTK.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value = 2;
			}
			else if (num > num4)
			{
				sViewTK.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value = 3;
			}
			else if (num > 0.0)
			{
				sViewTK.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value = 4;
			}
			else
			{
				sViewTK.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value = 0;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void XuLyCT22_ct6(SheetView sViewTK)
	{
		try
		{
			switch (Convert.ToInt32(sViewTK.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value))
			{
			case 1:
				sViewTK.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value = 3000000;
				break;
			case 2:
				sViewTK.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value = 2000000;
				break;
			case 3:
				sViewTK.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value = 1500000;
				break;
			case 4:
				sViewTK.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value = 1000000;
				break;
			default:
				sViewTK.Cells[CellId_CT22_ct6[1], CellId_CT22_ct6[0]].Value = 0;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
