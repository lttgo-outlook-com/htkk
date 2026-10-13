using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn04TT80Nckt251
{
	private string cellId_ct16 = "BH_48";

	private string cellId_ct16_ma = "BS_48";

	private string cellId_ct22 = "BH_54";

	private string cellId_ct23 = "BH_55";

	private string cellId_ct24 = "BH_56";

	private string cellId_ct25 = "BH_57";

	private int[] CellId_ct16 => GeneralFunctions.ArrayCellId(cellId_ct16);

	private int[] CellId_ct16_ma => GeneralFunctions.ArrayCellId(cellId_ct16_ma);

	private int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("CR_15");
			int[] array2 = GeneralFunctions.ArrayCellId("S_27");
			int[] array3 = GeneralFunctions.ArrayCellId("CR_27");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			DateTime dateTime = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime value = DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			double num = Math.Round((double)dateTime.Subtract(value).Days / 30.4375, 3);
			sView.Cells[array[1], array[0]].Value = num;
			if (PeriodInfo.KieuKy == "D")
			{
				sView.Cells[CellId_ct22[1], CellId_ct22[0]].Locked = true;
				sView.Cells[CellId_ct23[1], CellId_ct23[0]].Locked = true;
				sView.Cells[CellId_ct24[1], CellId_ct24[0]].Formula = "";
				sView.Cells[CellId_ct25[1], CellId_ct25[0]].Formula = "";
			}
			if (PeriodInfo.KieuKy == "Y" && !string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[array2[1], array2[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[array3[1], array3[0]].Value = PeriodInfo.MaQuyetToan;
			}
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("BH_44");
			int[] array2 = GeneralFunctions.ArrayCellId("BH_45");
			int[] array3 = GeneralFunctions.ArrayCellId("BH_51");
			int[] array4 = GeneralFunctions.ArrayCellId("BH_52");
			double num = ConvertHelpers.ConvertToDouble(sheetView.Cells[array[1], array[0]].Value);
			double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
			double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
			double num4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array4[1], array4[0]].Value);
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			if (PeriodInfo.KieuKy == "Y" && hTKKCMCSheetInfo != null)
			{
				string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
				string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
				string tuNgay = PeriodInfo.TuNgay;
				string denNgay = PeriodInfo.DenNgay;
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy");
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy");
				int num5 = (dateTime2.Year - dateTime.Year) * 12 + dateTime2.Month - dateTime.Month + 1;
				if (num5 == 0)
				{
					num5 = 1;
				}
				if (num5 > 12)
				{
					if ((num + num2) / (double)num5 * 12.0 > 200000000000.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text);
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text);
						value = 2;
					}
				}
				else if (num + num2 > 200000000000.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text);
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text);
					value = 2;
				}
				if (num4 <= num3)
				{
					HTKKFile.RemoveValidate(sheetView, array4[1], array4[0]);
					int[] array5 = GeneralFunctions.ArrayCellId("E_26");
					double num6 = ConvertHelpers.ConvertToDouble(sheetView2.Cells[array5[1], array5[0]].Value);
					if (num4 != num6)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, array4[1], array4[0], text2);
						if (value != 2)
						{
							value = 1;
						}
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
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string text = sheetView.Cells[CellId_ct16_ma[1], CellId_ct16_ma[0]].Text;
			sheetView.Cells[CellId_ct16[1], CellId_ct16[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
