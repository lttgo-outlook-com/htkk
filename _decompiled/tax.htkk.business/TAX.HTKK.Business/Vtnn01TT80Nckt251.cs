using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Vtnn01TT80Nckt251
{
	private int[] cellTenQuyetToan = GeneralFunctions.ArrayCellId("H_5");

	private int[] cellMaQuyetToan = GeneralFunctions.ArrayCellId("AK_5");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[cellTenQuyetToan[1], cellTenQuyetToan[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[cellMaQuyetToan[1], cellMaQuyetToan[0]].Value = PeriodInfo.MaQuyetToan;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		GeneralFunctions.setButton(sView, 5, 2, "Tải bảng kê");
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		GeneralFunctions.setButton(sView, 5, 2, "Tải bảng kê");
	}

	public void AutoBindData_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		GeneralFunctions.setButton(sView, 5, 2, "Tải bảng kê");
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
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
			setKey(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
			tinhSum(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
			setStt(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
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
			setKey(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
			tinhSum(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
			setStt(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			setKey(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
			tinhSum(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
			setStt(ToKhaiInfo, winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheetIndex);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bangke_01_1_VTNN.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bangke_01_2_VTNN.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bangke_01_3_VTNN.xls");
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
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int[] cell = GeneralFunctions.ArrayCellId("E_22");
			string message = HTKKFile.GetMessage(sView, cell);
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24"));
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
		}
	}

	private void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int index = 1;
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[index];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			int num4 = HTKKFile.ColumnLettersToIndex("I");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("P");
			int num8 = HTKKFile.ColumnLettersToIndex("S");
			int num9 = HTKKFile.ColumnLettersToIndex("V");
			if (hTKKCMCSheetInfo != null)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num10 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct04"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct06"));
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
					string text3 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
					string text4 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					if (string.IsNullOrEmpty(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (string.IsNullOrEmpty(text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					for (int num11 = 0; num11 < hTKKCMCTableInfo.RowsValues.Count; num11++)
					{
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, num10 + num11, num9 - 1);
						string text5 = sheetView.Cells[num10 + num11, num - 1].Text;
						string text6 = sheetView.Cells[num10 + num11, num2 - 1].Text;
						string text7 = sheetView.Cells[num10 + num11, num3 - 1].Text;
						string text8 = sheetView.Cells[num10 + num11, num4 - 1].Text;
						string text9 = sheetView.Cells[num10 + num11, num5 - 1].Text;
						string text10 = sheetView.Cells[num10 + num11, num6 - 1].Text;
						string text11 = sheetView.Cells[num10 + num11, num7 - 1].Text;
						double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num10 + num11, num8 - 1].Value);
						double num13 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num10 + num11, num9 - 1].Value);
						if (string.IsNullOrEmpty(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num2 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num3 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num4 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text9))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num5 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num6 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num7 - 1, text);
							value = 2;
						}
						else if (!ValidateBase.CheckMaSoThue(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num7 - 1, text2);
							value = 2;
						}
						if (num12 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num8 - 1, text);
							value = 2;
						}
						if (num13 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num10 + num11, num9 - 1, text);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[index].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int index = 2;
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[index];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			int num4 = HTKKFile.ColumnLettersToIndex("I");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("P");
			int num8 = HTKKFile.ColumnLettersToIndex("S");
			int num9 = HTKKFile.ColumnLettersToIndex("V");
			int num10 = HTKKFile.ColumnLettersToIndex("X");
			int num11 = HTKKFile.ColumnLettersToIndex("Z");
			if (hTKKCMCSheetInfo != null)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num12 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct04"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct06"));
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
					string text3 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
					string text4 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					if (string.IsNullOrEmpty(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (string.IsNullOrEmpty(text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					for (int num13 = 0; num13 < hTKKCMCTableInfo.RowsValues.Count; num13++)
					{
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num9 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num10 - 1);
						HTKKFile.RemoveValidate(sheetView, num12 + num13, num11 - 1);
						string text5 = sheetView.Cells[num12 + num13, num - 1].Text;
						string text6 = sheetView.Cells[num12 + num13, num2 - 1].Text;
						string text7 = sheetView.Cells[num12 + num13, num3 - 1].Text;
						string text8 = sheetView.Cells[num12 + num13, num4 - 1].Text;
						string text9 = sheetView.Cells[num12 + num13, num5 - 1].Text;
						string text10 = sheetView.Cells[num12 + num13, num6 - 1].Text;
						string text11 = sheetView.Cells[num12 + num13, num7 - 1].Text;
						double num14 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num12 + num13, num8 - 1].Value);
						double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num12 + num13, num9 - 1].Value);
						double num16 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num12 + num13, num10 - 1].Value);
						double num17 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num12 + num13, num11 - 1].Value);
						if (string.IsNullOrEmpty(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num2 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num3 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num4 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text9))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num5 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num6 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num7 - 1, text);
							value = 2;
						}
						else if (!ValidateBase.CheckMaSoThue(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num7 - 1, text2);
							value = 2;
						}
						if (num14 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num8 - 1, text);
							value = 2;
						}
						if (num15 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num9 - 1, text);
							value = 2;
						}
						if (num16 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num10 - 1, text);
							value = 2;
						}
						if (num17 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num12 + num13, num11 - 1, text);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[index].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void Validate_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int index = 3;
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[index];
			string text = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = sheetView.FpSpread.Sheets[sheetView.FpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			int num4 = HTKKFile.ColumnLettersToIndex("I");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("N");
			int num7 = HTKKFile.ColumnLettersToIndex("Q");
			int num8 = HTKKFile.ColumnLettersToIndex("T");
			if (hTKKCMCSheetInfo != null)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					int num9 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct04"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("ct06"));
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
					string text3 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
					string text4 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
					if (string.IsNullOrEmpty(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					if (string.IsNullOrEmpty(text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, text);
						value = 2;
					}
					for (int num10 = 0; num10 < hTKKCMCTableInfo.RowsValues.Count; num10++)
					{
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, num9 + num10, num8 - 1);
						string text5 = sheetView.Cells[num9 + num10, num - 1].Text;
						string text6 = sheetView.Cells[num9 + num10, num2 - 1].Text;
						string text7 = sheetView.Cells[num9 + num10, num3 - 1].Text;
						string text8 = sheetView.Cells[num9 + num10, num4 - 1].Text;
						string text9 = sheetView.Cells[num9 + num10, num5 - 1].Text;
						string text10 = sheetView.Cells[num9 + num10, num6 - 1].Text;
						double num11 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num9 + num10, num7 - 1].Value);
						double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num9 + num10, num8 - 1].Value);
						if (string.IsNullOrEmpty(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num2 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num3 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num4 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text9))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num5 - 1, text);
							value = 2;
						}
						if (string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num6 - 1, text);
							value = 2;
						}
						else if (!ValidateBase.CheckMaSoThue(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num6 - 1, text2);
							value = 2;
						}
						if (num11 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num7 - 1, text);
							value = 2;
						}
						if (num12 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + num10, num8 - 1, text);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[index].BackColor = Color.White;
				break;
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[activeSheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			int num = HTKKFile.ColumnLettersToIndex("AD");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				object value = sheetView.Cells[num2 - 1, num - 1].Value;
				if (value != null)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct04");
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct05_ten");
					HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct06");
					HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct05_ma");
					string[] array = value.ToString().Split(new string[1] { "___" }, StringSplitOptions.None);
					if (array.Length > 3)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = array[0];
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = array[1];
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = array[2];
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = array[3];
					}
				}
			}
			tinhSum(ToKhaiInfo, sheetView, activeSheetIndex);
			setStt(ToKhaiInfo, sheetView, activeSheetIndex);
		}
		catch (Exception)
		{
		}
	}

	public void tinhSum(HTKKCMCToKhaiInfo ToKhaiInfo, SheetView sView, int sheetIndex)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "tongCot7");
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "tongCot8");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "tongCot9");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "tongCot10");
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "tongCot11");
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCTableInfo.Footer.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "tongCot12");
				double num = 0.0;
				double num2 = 0.0;
				double num3 = 0.0;
				double num4 = 0.0;
				double num5 = 0.0;
				double num6 = 0.0;
				for (int num7 = 0; num7 < hTKKCMCTableInfo.RowsValues.Count; num7++)
				{
					if (hTKKCMCCellInfo != null)
					{
						num += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num7 - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
					}
					if (hTKKCMCCellInfo2 != null)
					{
						num2 += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num7 - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
					}
					if (hTKKCMCCellInfo3 != null)
					{
						num3 += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num7 - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
					}
					if (hTKKCMCCellInfo4 != null)
					{
						num4 += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num7 - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
					}
					if (hTKKCMCCellInfo5 != null)
					{
						num5 += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num7 - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value);
					}
					if (hTKKCMCCellInfo6 != null)
					{
						num6 += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num7 - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value);
					}
				}
				if (hTKKCMCCellInfo != null)
				{
					sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
				}
				if (hTKKCMCCellInfo2 != null)
				{
					sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
				}
				if (hTKKCMCCellInfo3 != null)
				{
					sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = num3;
				}
				if (hTKKCMCCellInfo4 != null)
				{
					sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = num4;
				}
				if (hTKKCMCCellInfo5 != null)
				{
					sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = num5;
				}
				if (hTKKCMCCellInfo6 != null)
				{
					sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = num6;
				}
			}
		}
		catch (Exception)
		{
		}
	}

	public void setKey(HTKKCMCToKhaiInfo ToKhaiInfo, SheetView sView, int sheetIndex)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			int num = HTKKFile.ColumnLettersToIndex("AD");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct04");
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct05_ten");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct06");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "ct05_ma");
				string text = "";
				text += sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
				text += sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
				text += sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
				text += sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
				sView.Cells[num2 - 1, num - 1].Value = text;
			}
		}
		catch (Exception)
		{
		}
	}

	public void setStt(HTKKCMCToKhaiInfo ToKhaiInfo, SheetView sView, int sheetIndex)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			int num = HTKKFile.ColumnLettersToIndex("B");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					sView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Value = j + 1;
				}
			}
		}
		catch (Exception)
		{
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sView = winFpSpread.Sheets[activeSheetIndex];
			tinhSum(ToKhaiInfo, sView, activeSheetIndex);
			setStt(ToKhaiInfo, sView, activeSheetIndex);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sView = winFpSpread.Sheets[activeSheetIndex];
			tinhSum(ToKhaiInfo, sView, activeSheetIndex);
			setStt(ToKhaiInfo, sView, activeSheetIndex);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sView = winFpSpread.Sheets[activeSheetIndex];
			tinhSum(ToKhaiInfo, sView, activeSheetIndex);
			setStt(ToKhaiInfo, sView, activeSheetIndex);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
