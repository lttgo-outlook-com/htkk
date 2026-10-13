using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Hknn01TT80Nckt308
{
	private string pathCellTuNgay_PL01_1 = "ct01_tu";

	private string pathCellDenNgay_PL01_1 = "ct01_den";

	private string pathCellTuNgay_PL01_2 = "ct01_tu";

	private string pathCellDenNgay_PL01_2 = "ct01_den";

	private int[] cell_PL01_1_TK = GeneralFunctions.ArrayCellId("AD_22");

	private int[] cell_PL01_2_TK = GeneralFunctions.ArrayCellId("AE_22");

	private HTKKCMCToKhaiInfo _toKhaiInfo = null;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_toKhaiInfo = ToKhaiInfo;
			winFpSpread.Sheets[0].Cells[cell_PL01_1_TK[1], cell_PL01_1_TK[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			winFpSpread.Sheets[0].Cells[cell_PL01_2_TK[1], cell_PL01_2_TK[0]].Value = ((ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				SheetView sView = winFpSpread.Sheets[1];
				setCt01Default(sView, hTKKCMCSheetInfo.Tables, pathCellTuNgay_PL01_1, pathCellDenNgay_PL01_1, PeriodInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo != null)
			{
				SheetView sView = winFpSpread.Sheets[2];
				setCt01Default(sView, hTKKCMCSheetInfo.Tables, pathCellTuNgay_PL01_2, pathCellDenNgay_PL01_2, PeriodInfo);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPhuLuc_01_1(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPhuLuc_01_2(winFpSpread.Sheets[2], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_36"));
			int[] array = GeneralFunctions.ArrayCellId("L_86");
			int[] array2 = GeneralFunctions.ArrayCellId("L_87");
			int colSum = HTKKFile.ColumnLettersToIndex("O");
			int colSum2 = HTKKFile.ColumnLettersToIndex("R");
			int colSum3 = HTKKFile.ColumnLettersToIndex("S");
			int colSum4 = HTKKFile.ColumnLettersToIndex("V");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[2];
			HTKKFile.RemoveValidate(sView, array[1], array[0]);
			HTKKFile.RemoveValidate(sView, array2[1], array2[0]);
			double num = 0.0;
			double num2 = 0.0;
			if (hTKKCMCSheetInfo != null)
			{
				num = GeneralFunctions.CalculateSumByGroup(sView.FpSpread.Sheets[1], hTKKCMCSheetInfo.Tables, colSum, isView: false);
				num2 = GeneralFunctions.CalculateSumByGroup(sView.FpSpread.Sheets[1], hTKKCMCSheetInfo.Tables, colSum2, isView: false);
			}
			double num3 = 0.0;
			double num4 = 0.0;
			if (hTKKCMCSheetInfo2 != null)
			{
				num3 = GeneralFunctions.CalculateSumByGroup(sView.FpSpread.Sheets[2], hTKKCMCSheetInfo2.Tables, colSum3, isView: false);
				num4 = GeneralFunctions.CalculateSumByGroup(sView.FpSpread.Sheets[2], hTKKCMCSheetInfo2.Tables, colSum4, isView: false);
			}
			double num5 = Convert.ToDouble(sView.Cells[array[1], array[0]].Value);
			double num6 = Convert.ToDouble(sView.Cells[array2[1], array2[0]].Value);
			int num7 = Convert.ToInt32(sView.Cells[cell_PL01_1_TK[1], cell_PL01_1_TK[0]].Value);
			int num8 = Convert.ToInt32(sView.Cells[cell_PL01_2_TK[1], cell_PL01_2_TK[0]].Value);
			if ((num7 == 1 || num8 == 1) && num5 != num2 + num4)
			{
				HTKKFile.SetNodeBackGroundYellow(sView, array[1], array[0], message);
				value = 1;
			}
			if ((num7 == 1 || num8 == 1) && num6 != num + num3)
			{
				HTKKFile.SetNodeBackGroundYellow(sView, array2[1], array2[0], message2);
				value = 1;
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
		}
	}

	private void checkPhuLuc_01_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_34"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_38"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_39"));
			int num = HTKKFile.ColumnLettersToIndex("C");
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCellTuNgay_PL01_1);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCellDenNgay_PL01_1);
				HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "ct02");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "ct03_ten");
				HTKKCMCCellInfo hTKKCMCCellInfo5 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "ct04");
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
				string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				string text2 = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				DateTime dateTime = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime2 = DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime3 = (string.IsNullOrEmpty(text) ? DateTime.MinValue : DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture));
				DateTime dateTime4 = (string.IsNullOrEmpty(text2) ? DateTime.MinValue : DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture));
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text) && (dateTime3 > dateTime2 || dateTime3 < dateTime))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message2);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && (dateTime4 > dateTime2 || dateTime4 < dateTime))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message3);
					value = 2;
				}
				string value2 = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				string value3 = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				string value4 = sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num2 = 0; num2 < table.RowsValues.Count; num2++)
				{
					string text3 = sView.Cells[dataLocationTemplateRowStart + num2 - 1, num - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text3))
					{
						DateTime dateTime5 = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						if (dateTime5 > dateTime4 || dateTime5 < dateTime3)
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + num2 - 1, num - 1, message);
							value = 2;
						}
					}
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
		}
	}

	private void checkPhuLuc_01_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("C");
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_34"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_38"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_39"));
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCellTuNgay_PL01_2);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCellDenNgay_PL01_2);
				HTKKCMCCellInfo hTKKCMCCellInfo3 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "ct02");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "ct03_ten");
				HTKKCMCCellInfo hTKKCMCCellInfo5 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "ct04");
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
				HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
				string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				string text2 = sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				DateTime dateTime = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime2 = DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime dateTime3 = (string.IsNullOrEmpty(text) ? DateTime.MinValue : DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture));
				DateTime dateTime4 = (string.IsNullOrEmpty(text2) ? DateTime.MinValue : DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture));
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text) && (dateTime3 > dateTime2 || dateTime3 < dateTime))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message2);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && (dateTime4 > dateTime2 || dateTime4 < dateTime))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message3);
					value = 2;
				}
				string value2 = sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				string value3 = sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim();
				string value4 = sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text.Trim();
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, message4);
					value = 2;
				}
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num2 = 0; num2 < table.RowsValues.Count; num2++)
				{
					string text3 = sView.Cells[dataLocationTemplateRowStart + num2 - 1, num - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text3))
					{
						DateTime dateTime5 = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						if (dateTime5 > dateTime4 || dateTime5 < dateTime3)
						{
							HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + num2 - 1, num - 1, message);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void setCt01Default(SheetView sView, List<HTKKCMCTableInfo> tables, string pathCellTuNgay, string pathCellDenNgay, HTKKPeriodInfo PeriodInfo)
	{
		foreach (HTKKCMCTableInfo table in tables)
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCellTuNgay);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathCellDenNgay);
			if (string.IsNullOrEmpty(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text))
			{
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = PeriodInfo.TuNgay;
			}
			if (string.IsNullOrEmpty(sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text))
			{
				sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = PeriodInfo.DenNgay;
			}
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return false;
			}
			SheetView sView = winFpSpread.Sheets[1];
			setCt01Default(sView, hTKKCMCSheetInfo.Tables, pathCellTuNgay_PL01_1, pathCellDenNgay_PL01_1, PeriodInfo);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return false;
			}
			SheetView sView = winFpSpread.Sheets[2];
			setCt01Default(sView, hTKKCMCSheetInfo.Tables, pathCellTuNgay_PL01_2, pathCellDenNgay_PL01_2, PeriodInfo);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cell_PL01_1_TK[1], cell_PL01_1_TK[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cell_PL01_1_TK[1], cell_PL01_1_TK[0]].Value = ((_toKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cell_PL01_2_TK[1], cell_PL01_2_TK[0]].Value = ((ToKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_2(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cell_PL01_2_TK[1], cell_PL01_2_TK[0]].Value = ((_toKhaiInfo.SheetInfos[2] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
