using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn07TT89Nckt309
{
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = HTKKFile.ColumnLettersToIndex("AI");
			int num2 = HTKKFile.ColumnLettersToIndex("AN");
			int num3 = HTKKFile.ColumnLettersToIndex("BH");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num4 = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count;
			if (iRow >= dataLocationTemplateRowStart && iRow < num4 && iCol == num)
			{
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = sheetView.Cells[(int)iRow - 1, num3 - 1].Value;
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
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo.RowsInfo.Cells;
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKCMCRowValue hTKKCMCRowValue = rowsValues[i];
				hTKKCMCRowValue.RowValue = hTKKCMCRowValue.RowValue.GetRange(0, cells.Count);
				hTKKCMCRowValue.RowValue.Take(hTKKCMCTableInfo.RowsInfo.Cells.Count).ToList();
				if (!GeneralFunctions.checkHaveNewDataOfTable(hTKKCMCTableInfo.RowsInfo, hTKKCMCRowValue.RowValue))
				{
					continue;
				}
				string value2 = rowsValues[i].RowValue[0];
				string value3 = rowsValues[i].RowValue[1];
				string value4 = rowsValues[i].RowValue[2];
				string value5 = rowsValues[i].RowValue[3];
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num3 - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num - 1);
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart2 + i - 1, num4 - 1);
				if (string.IsNullOrEmpty(value3))
				{
					if (value5.ToStringOrDefault() == "VN")
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1, message);
						value = 2;
					}
				}
				else
				{
					if (!ValidateBase.CheckMaSoThue(value3))
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1, message2);
						value = 2;
					}
					if (string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num3 - 1, message);
						value = 2;
					}
				}
				if (rowsValues.Where((HTKKCMCRowValue x) => x.RowValue[0] == value2 && x.RowValue[1] == value3 && x.RowValue[3] == value5).Count() > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num2 - 1, message3);
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num - 1, message3);
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart2 + i - 1, num4 - 1, message3);
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 8 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_07_TNDN_TT89.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}
}
