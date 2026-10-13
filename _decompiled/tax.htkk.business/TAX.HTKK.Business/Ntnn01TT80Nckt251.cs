using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn01TT80Nckt251
{
	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.KieuKy == "D")
			{
				int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
				int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				for (int i = num; i <= num2; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AE") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AE") - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.KieuKy == "D")
			{
				int num = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart - 1;
				int num2 = num + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				for (int i = num; i <= num2; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AE") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AE") - 1].Locked = true;
				}
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
			if (idxRow == 5 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01NTNN_TT80.xls");
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
		{
			return false;
		}
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int colunmIndexByColumName = HTKKFormater.GetColunmIndexByColumName("C");
			int colunmIndexByColumName2 = HTKKFormater.GetColunmIndexByColumName("L");
			int colunmIndexByColumName3 = HTKKFormater.GetColunmIndexByColumName("M");
			int colunmIndexByColumName4 = HTKKFormater.GetColunmIndexByColumName("S");
			int colunmIndexByColumName5 = HTKKFormater.GetColunmIndexByColumName("W");
			int colunmIndexByColumName6 = HTKKFormater.GetColunmIndexByColumName("Y");
			int colunmIndexByColumName7 = HTKKFormater.GetColunmIndexByColumName("AE");
			int colunmIndexByColumName8 = HTKKFormater.GetColunmIndexByColumName("AJ");
			int colunmIndexByColumName9 = HTKKFormater.GetColunmIndexByColumName("AN");
			int colunmIndexByColumName10 = HTKKFormater.GetColunmIndexByColumName("AN");
			int colunmIndexByColumName11 = HTKKFormater.GetColunmIndexByColumName("AV");
			int colunmIndexByColumName12 = HTKKFormater.GetColunmIndexByColumName("AZ");
			int colunmIndexByColumName13 = HTKKFormater.GetColunmIndexByColumName("BD");
			int colunmIndexByColumName14 = HTKKFormater.GetColunmIndexByColumName("BD");
			int colunmIndexByColumName15 = HTKKFormater.GetColunmIndexByColumName("BH");
			int colunmIndexByColumName16 = HTKKFormater.GetColunmIndexByColumName("BN");
			int colunmIndexByColumName17 = HTKKFormater.GetColunmIndexByColumName("BR");
			int colunmIndexByColumName18 = HTKKFormater.GetColunmIndexByColumName("CA");
			int colunmIndexByColumName19 = HTKKFormater.GetColunmIndexByColumName("Z");
			int colunmIndexByColumName20 = HTKKFormater.GetColunmIndexByColumName("BX");
			string[] arrARowValue = new string[19];
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				arrARowValue[0] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName);
				arrARowValue[1] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName2);
				arrARowValue[2] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName3);
				arrARowValue[3] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName4);
				arrARowValue[4] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName5);
				arrARowValue[5] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName6);
				arrARowValue[6] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName7);
				arrARowValue[7] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName8);
				arrARowValue[8] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName9);
				arrARowValue[9] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName11);
				arrARowValue[10] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName12);
				arrARowValue[11] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName13);
				arrARowValue[12] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName15);
				arrARowValue[13] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName16);
				arrARowValue[14] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName17);
				arrARowValue[15] = "";
				arrARowValue[18] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName18);
				string value2 = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName20);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName2);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName3);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName4);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName5);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName7);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName8);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName12);
				if (string.IsNullOrEmpty(arrARowValue[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				if (string.IsNullOrEmpty(arrARowValue[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				if (!string.IsNullOrEmpty(arrARowValue[2]) && !ValidateBase.CheckMaSoThue(arrARowValue[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				if (string.IsNullOrEmpty(arrARowValue[3]))
				{
				}
				if (!string.IsNullOrEmpty(arrARowValue[4]))
				{
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(arrARowValue[4], "dd/MM/yyyy");
					if (dateTime == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_33")));
						value = 2;
					}
					else if (dateTime.CompareTo(DateTime.Today) >= 0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(arrARowValue[6]))
				{
					DateTime dateTime2 = DateTime.ParseExact(arrARowValue[6], "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (dateTime2 > DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
						value = 2;
					}
					else if (dateTime2 < DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
						value = 2;
					}
				}
				int num2 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[18].Equals(arrARowValue[18].Trim()));
				if (num2 > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")), i + 1);
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")), i + 1);
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")), i + 1);
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")), i + 1);
					value = 2;
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("BX");
			int num3 = HTKKFile.ColumnLettersToIndex("AN");
			int num4 = HTKKFile.ColumnLettersToIndex("BD");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num + i - 1, num2 - 1].Text;
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
}
