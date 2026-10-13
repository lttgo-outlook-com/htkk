using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BkKtTncnNckt213
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ChungTu/tongSoTienThue"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathnull_01"));
			string text = MoneyHelpers.MoneyToText(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value.ToStringOrDefault());
			sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num && idxRow <= num2 && idxCol == HTKKFile.ColumnLettersToIndex("I") - 1)
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sViewTK = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iCol == HTKKFile.ColumnLettersToIndex("U") && iRow >= dataLocationTemplateRowStart && iRow <= num)
			{
				SoTienBangChu(sViewTK, hTKKCMCSheetInfo);
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Text = winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Text.ToUpper();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void SoTienBangChu(SheetView sViewTK, HTKKCMCSheetInfo sInfoTK)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ChungTu/tongSoTienThue"));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfoTK.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathnull_01"));
		double num = Convert.ToDouble(sViewTK.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
		if (num >= 0.0)
		{
			string text = MoneyHelpers.MoneyToText(num.ToStringOrDefault());
			sViewTK.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = text;
		}
		else
		{
			sViewTK.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = "";
		}
	}

	public void Validate_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[7];
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("L") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("P") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("T") - 1;
			int column = HTKKFile.ColumnLettersToIndex("U") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("W") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num2].Value);
				array[1] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num3].Value);
				array[2] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num4].Value);
				array[3] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num5].Value);
				array[4] = ConvertHelpers.ConvertToString(sView.Cells[num + i, num6].Value);
				array[5] = ConvertHelpers.ConvertToString(sView.Cells[num + i, column].Value);
				array[6] = ConvertHelpers.ConvertToString(sView.Cells[num + i, column2].Value);
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, num2);
					HTKKFile.RemoveValidate(sView, num + i, num3);
					HTKKFile.RemoveValidate(sView, num + i, num4);
					HTKKFile.RemoveValidate(sView, num + i, num5);
					HTKKFile.RemoveValidate(sView, num + i, num6);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")), i + 1);
					value = 2;
				}
				else
				{
					string text = GeneralFunctions.CheckSoChungTu(array[0]);
					if (text == "1")
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
						value = 2;
					}
					else if (text == "2")
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, num2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_16")), i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, num + i, num2);
					}
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")), i + 1);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num3);
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")), i + 1);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num4);
				}
				if (string.IsNullOrEmpty(array[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")), i + 1);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num5);
				}
				if (string.IsNullOrEmpty(array[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, num6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")), i + 1);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, num6);
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_0(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					break;
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeRowIndex == 1 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_KhauTruThueTNCN.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SoTienBangChu(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
