using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Phlpng01TT80Nckt250
{
	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int iCol = HTKKFile.ColumnLettersToIndex("O");
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, iCol, value, validate, (text == "" || text == "VND") ? "0" : "0,000000");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
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
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("R");
			int key = HTKKFile.ColumnLettersToIndex("T");
			int key2 = HTKKFile.ColumnLettersToIndex("W");
			int num4 = HTKKFile.ColumnLettersToIndex("AE");
			int num5 = HTKKFile.ColumnLettersToIndex("C");
			int num6 = HTKKFile.ColumnLettersToIndex("O");
			int value = HTKKFile.ColumnLettersToIndex("T");
			int value2 = HTKKFile.ColumnLettersToIndex("W");
			SheetView sheetView = winFpSpread.Sheets[0];
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(num2, num6);
			dictionary.Add(key, value);
			dictionary.Add(key2, value2);
			GeneralFunctions.CopyPaseTableToTable(hTKKCMCSheetInfo.Tables[0], hTKKCMCSheetInfo.Tables[1], sheetView, sheetView, dictionary, num, num5, num, num5, "", hTKKCMCSheetInfo);
			if (iCol == num)
			{
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
					string text = hTKKCMCTableInfo.RowsValues[i].RowValue[3];
					HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
					GeneralFunctions.InsertNumberCell(sheetView, dataLocationTemplateRowStart + i, num2, value3, validate, (text == "" || text == "VND") ? "0" : "0,000000");
				}
			}
			int dataLocationTemplateRowStart2 = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[1].RowsValues.Count; i++)
			{
				sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num6 - 1].CellType = sheetView.Cells[0, 0].CellType;
				object value4 = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num6 - 1].Value;
				string text = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, num5 - 1].Text.Trim();
				HTKKCMCValidate validate = hTKKCMCSheetInfo.Tables[1].RowsInfo.Cells[2].Validate;
				GeneralFunctions.InsertNumberCell(sheetView, dataLocationTemplateRowStart2 + i, num6, value4.ToString(), validate, (text == "" || text == "VND") ? "0" : "0,000000");
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
