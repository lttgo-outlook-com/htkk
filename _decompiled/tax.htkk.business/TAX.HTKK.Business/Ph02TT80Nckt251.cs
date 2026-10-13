using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ph02TT80Nckt251
{
	private string DEFAULT_LOAI_TIEN_KHAC = "0,00";

	private int col2_Tb1 = HTKKFile.ColumnLettersToIndex("C");

	private int col3_Tb1 = HTKKFile.ColumnLettersToIndex("J");

	private int col4_Tb1 = HTKKFile.ColumnLettersToIndex("L");

	private int col5_Tb1 = HTKKFile.ColumnLettersToIndex("M");

	private int col6_Tb1 = HTKKFile.ColumnLettersToIndex("P");

	private int col7_Tb1 = HTKKFile.ColumnLettersToIndex("S");

	private int col8_Tb1 = HTKKFile.ColumnLettersToIndex("V");

	private int col9_Tb1 = HTKKFile.ColumnLettersToIndex("AB");

	private int col3Ma_Tb1 = HTKKFile.ColumnLettersToIndex("AH");

	private int col7Formula_Tb1 = HTKKFile.ColumnLettersToIndex("AE");

	private int colMaTien_Tb2 = HTKKFile.ColumnLettersToIndex("AH");

	private int colTien_Tb2 = HTKKFile.ColumnLettersToIndex("C");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
				string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
				string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[7];
				string value4 = hTKKCMCTableInfo.RowsValues[i].RowValue[8];
				string value5 = hTKKCMCTableInfo.RowsValues[i].RowValue[9];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[11];
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
				HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
				HTKKCMCValidate validate4 = hTKKCMCTableInfo.RowsInfo.Cells[8].Validate;
				HTKKCMCValidate validate5 = hTKKCMCTableInfo.RowsInfo.Cells[9].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col6_Tb1, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col8_Tb1, value4, validate4, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col9_Tb1, value5, validate5, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value5 = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate6 = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col9_Tb1, value5, validate6, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
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
			_ToKhaiInfo = ToKhaiInfo;
			CapNhatLaiCellType(sView, ToKhaiInfo);
			int[] array = GeneralFunctions.ArrayCellId("D_36");
			int[] array2 = GeneralFunctions.ArrayCellId("E_36");
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[array[1], array[0]].Text = PeriodInfo.TenQuyetToan;
				sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaQuyetToan;
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == col4_Tb1 || iCol == col5_Tb1)
			{
				sheetView.Cells[(int)iRow - 1, col7_Tb1 - 1].Value = sheetView.Cells[(int)iRow - 1, col7Formula_Tb1 - 1].Value;
			}
			TongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
			CapNhatLaiCellType(winFpSpread.Sheets[0], ToKhaiInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[0]);
			GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, ToKhaiInfo.SheetInfos[0], 0);
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int[] cell = GeneralFunctions.ArrayCellId("E_22");
			string message = HTKKFile.GetMessage(sheetView, cell);
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			List<string> list = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num + i, col2_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col3_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col5_Tb1 - 1);
				string text = sheetView.Cells[num + i, col2_Tb1 - 1].Text;
				string text2 = sheetView.Cells[num + i, col3_Tb1 - 1].Text;
				string text3 = sheetView.Cells[num + i, col5_Tb1 - 1].Text;
				string text4 = sheetView.Cells[num + i, col3Ma_Tb1 - 1].Text;
				string item = text + text2 + text3;
				if (!list.Contains(item))
				{
					list.Add(item);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col5_Tb1 - 1, message2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message3);
					value = 2;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TongTheoLoaiTien(winFpSpread, ToKhaiInfo);
		return true;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
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
			TongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		CapNhatLaiCellType(winFpSpread.Sheets[0], ToKhaiInfo);
	}

	public void TongTheoLoaiTien(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(col9_Tb1, col9_Tb1);
			GeneralFunctions.CopyPaseTableToTable(hTKKCMCSheetInfo.Tables[0], hTKKCMCSheetInfo.Tables[1], sheetView, sheetView, dictionary, col3Ma_Tb1, colMaTien_Tb2, col3_Tb1, colTien_Tb2, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CapNhatLaiCellType(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart + i - 1, col4_Tb1 - 1].Value.ToString().Trim();
				string value2 = sView.Cells[dataLocationTemplateRowStart + i - 1, col6_Tb1 - 1].Value.ToString().Trim();
				string value3 = sView.Cells[dataLocationTemplateRowStart + i - 1, col7_Tb1 - 1].Value.ToString().Trim();
				string value4 = sView.Cells[dataLocationTemplateRowStart + i - 1, col8_Tb1 - 1].Value.ToString().Trim();
				string value5 = sView.Cells[dataLocationTemplateRowStart + i - 1, col9_Tb1 - 1].Value.ToString().Trim();
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, col3Ma_Tb1 - 1].Value.ToString().Trim();
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
				HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
				HTKKCMCValidate validate4 = hTKKCMCTableInfo.RowsInfo.Cells[8].Validate;
				HTKKCMCValidate validate5 = hTKKCMCTableInfo.RowsInfo.Cells[9].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col6_Tb1, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col8_Tb1, value4, validate4, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col9_Tb1, value5, validate5, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value5 = sView.Cells[dataLocationTemplateRowStart + i - 1, col9_Tb1 - 1].Value.ToString().Trim();
				object value6 = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col3Ma_Tb1 - 1].Value;
				if (value6 != null)
				{
					HTKKCMCValidate validate6 = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col9_Tb1, value5, validate6, (value6.ToString().Trim() == "" || value6.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[0], 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
