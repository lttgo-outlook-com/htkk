using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tncn06TT89Nckt309
{
	private int[] cellTemPhuLuc_tk = GeneralFunctions.ArrayCellId("AB_37");

	private int[] cellKhaiThay_tk = GeneralFunctions.ArrayCellId("F_17");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellTemPhuLuc_tk[1], cellTemPhuLuc_tk[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
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
			if (iRow == cellKhaiThay_tk[1] + 1)
			{
				GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
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
			int num = HTKKFile.ColumnLettersToIndex("AP");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			int num4 = HTKKFile.ColumnLettersToIndex("M");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[2];
			HTKKCMCTableInfo hTKKCMCTableInfo4 = hTKKCMCSheetInfo.Tables[3];
			HTKKCMCTableInfo hTKKCMCTableInfo5 = hTKKCMCSheetInfo.Tables[5];
			HTKKCMCTableInfo hTKKCMCTableInfo6 = hTKKCMCSheetInfo.Tables[6];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart4 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart5 = hTKKCMCTableInfo5.Location.DataLocationTemplateRowStart;
			int dataLocationTemplateRowStart6 = hTKKCMCTableInfo6.Location.DataLocationTemplateRowStart;
			if (iCol == num2 && iRow >= dataLocationTemplateRowStart && iRow < dataLocationTemplateRowStart2)
			{
				sheetView.Cells[(int)iRow - 1, num3 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
			if (iCol == num4 && iRow >= dataLocationTemplateRowStart2 && iRow < dataLocationTemplateRowStart3)
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
			if ((iCol == num4 && iRow >= dataLocationTemplateRowStart4 && iRow < dataLocationTemplateRowStart5) || (iCol == num4 && iRow >= dataLocationTemplateRowStart6))
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
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
			int index = 2;
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int num = HTKKFile.ColumnLettersToIndex("AC");
			int colHide = HTKKFile.ColumnLettersToIndex("AD");
			int colHide2 = HTKKFile.ColumnLettersToIndex("AE");
			int num2 = HTKKFile.ColumnLettersToIndex("AF");
			double num3 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, num - 1].Value);
			if (num3 > 0.0)
			{
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide2, locked: true, deleteData: true);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num2, locked: true, deleteData: true);
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = num3;
			}
			else
			{
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, colHide2, locked: false, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, (int)iRow, num2, locked: false, deleteData: false);
			}
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
			if (activeRowIndex == 6 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_06_1_TNCN_TT89.xls");
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
			if (activeRowIndex == 6 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_06_2_TNCN_TT89.xls");
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
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					break;
				default:
					winFpSpread.TabStrip[key].BackColor = Color.White;
					break;
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int index = 1;
			SheetView sView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int num = 0;
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_90"));
			int colCheck = HTKKFile.ColumnLettersToIndex("AO");
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				if (table.TableName == "table_1")
				{
					int num2 = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sView, table, colCheck, message, isRemoveValidate: false, HTKKFile.ColumnLettersToIndex("M"));
					if (num2 > 0)
					{
						num = num2;
					}
				}
			}
			switch (num)
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int index = 2;
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_115"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_113"));
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("I");
			int num5 = HTKKFile.ColumnLettersToIndex("J");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text;
				bool flag = Convert.ToBoolean(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Value);
				string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text;
				string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Text;
				if (string.IsNullOrEmpty(text))
				{
					continue;
				}
				if (flag)
				{
					if (!ValidateBase.CheckMaSoThueCCCD(text) || text.Length != 10)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, message, i + 1);
						num = 2;
					}
				}
				else if (text.Length != 12)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, message2, i + 1);
					num = 2;
				}
			}
			switch (num)
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

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc_tk[1], cellTemPhuLuc_tk[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
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
			winFpSpread.Sheets[0].Cells[cellTemPhuLuc_tk[1], cellTemPhuLuc_tk[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			GeneralFunctions.FillDataMstCu(xmlNode);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
