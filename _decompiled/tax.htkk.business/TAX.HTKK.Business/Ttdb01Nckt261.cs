using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ttdb01Nckt261
{
	private Dictionary<string, string[]> dicDMThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDichVuThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDichVuThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private string nameTable_01_TTDB_I = "01_TTDB_I";

	private string nameTable_01_TTDB_II = "01_TTDB_II";

	private string nameTable_01_TTDB_III_A = "01_TTDB_III_A";

	private string nameTable_01_TTDB_III_B = "01_TTDB_III_B";

	private string nameTable_01_TTDB_III_C = "01_TTDB_III_C";

	private string maCaSiNo = "20200";

	private string cellId_TieuMuc = "L_14";

	private string cellId_N34 = "N_34";

	private string formatFormulaCt6 = "V{0}/(1+AD{0}/100)";

	private string formatFormulaCt9 = "ROUND(AA{0}*(AD{0}/100),0)-AG{0}";

	private string fmlCt10_PL1 = "H{0}-J{0}";

	private int[] CellId_TieuMuc => GeneralFunctions.ArrayCellId(cellId_TieuMuc);

	private int[] CellId_N34 => GeneralFunctions.ArrayCellId(cellId_N34);

	private int idxCol_F => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_G => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_P => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_S => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_AA => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_AD => HTKKFile.ColumnLettersToIndex("AD") - 1;

	private int idxCol_AG => HTKKFile.ColumnLettersToIndex("AG") - 1;

	private int idxCol_AN => HTKKFile.ColumnLettersToIndex("AN") - 1;

	private int idxCol_C_PL1 => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_D_PL1 => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_E_PL1 => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_F_PL1 => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_G_PL1 => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_H_PL1 => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_I_PL1 => HTKKFile.ColumnLettersToIndex("I") - 1;

	private int idxCol_J_PL1 => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_K_PL1 => HTKKFile.ColumnLettersToIndex("K") - 1;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.Cells[CellId_TieuMuc[1], CellId_TieuMuc[0]].Text = PeriodInfo.TieuMucHachToanVal;
			bool isCheckN = ConvertHelpers.ConvertToBool(sheetView.Cells[CellId_N34[1], CellId_N34[0]].Value);
			ClearAndLockCellTK(sheetView, hTKKCMCSheetInfo, isBindData: true, isCheckN);
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					for (int k = 0; k < hTKKCMCTableInfo.RowsInfo.Cells.Count; k++)
					{
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells[k];
						if (hTKKCMCTableCellInfo.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex("G"))
						{
							if (!HTKKFormater.isNullString(hTKKCMCTableInfo.RowsValues[j].RowValue[k]))
							{
								sheetView.Cells[num + j, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = false;
							}
							else
							{
								sheetView.Cells[num + j, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_I);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_II);
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_III_A);
			HTKKCMCTableInfo hTKKCMCTableInfo4 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_III_B);
			HTKKCMCTableInfo hTKKCMCTableInfo5 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_III_C);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num2 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			int num3 = dataLocationTemplateRowStart3 + hTKKCMCTableInfo3.RowsValues.Count - 1;
			int dataLocationTemplateRowStart4 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart;
			int num4 = dataLocationTemplateRowStart4 + hTKKCMCTableInfo4.RowsValues.Count - 1;
			int dataLocationTemplateRowStart5 = hTKKCMCTableInfo5.Location.DataLocationTemplateRowStart;
			int num5 = dataLocationTemplateRowStart5 + hTKKCMCTableInfo5.RowsValues.Count - 1;
			if (iCol - 1 != idxCol_G)
			{
				return;
			}
			if ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart3 && iRow <= num3) || (iRow >= dataLocationTemplateRowStart4 && iRow <= num4) || (iRow >= dataLocationTemplateRowStart5 && iRow <= num5))
			{
				string text = sheetView.Cells[(int)iRow - 1, idxCol_G].Text;
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_F].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_P].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_P].Locked = true;
					if (iRow >= dataLocationTemplateRowStart && iRow <= num)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_AD].Value = 0;
					}
					return;
				}
				string[] array = dicDMThueDB[text];
				sheetView.Cells[(int)iRow - 1, idxCol_G].Text = text;
				sheetView.Cells[(int)iRow - 1, idxCol_F].Text = array[1];
				sheetView.Cells[(int)iRow - 1, idxCol_P].Locked = false;
				sheetView.Cells[(int)iRow - 1, idxCol_P].Value = array[3];
				if (iRow >= dataLocationTemplateRowStart && iRow <= num)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_AD].Value = ConvertHelpers.ConvertToDouble(array[4]);
				}
			}
			else if (iRow >= dataLocationTemplateRowStart2 && iRow <= num2)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, idxCol_G].Text;
				if (HTKKFormater.isNullString(text2))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_F].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_AD].Value = 0;
				}
				else
				{
					string[] array = dicDichVuThueDB[text2];
					sheetView.Cells[(int)iRow - 1, idxCol_F].Value = array[1];
					sheetView.Cells[(int)iRow - 1, idxCol_AD].Value = ConvertHelpers.ConvertToDouble(array[4]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "01_1_TTDB_I");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iCol - 1 == idxCol_K_PL1)
			{
				sheetView.Cells[(int)iRow - 1, idxCol_K_PL1].ResetFormula();
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_I);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable_01_TTDB_II);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int num2 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
			if (!Convert.ToBoolean(sheetView.Cells[CellId_N34[1], CellId_N34[0]].Value))
			{
				string text = formatFormulaCt9;
				if (iCol == idxCol_V + 1 && ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2)))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_AA].Formula = string.Format(formatFormulaCt6, iRow);
					string newValue = sheetView.Cells[(int)iRow - 1, idxCol_AA].Text.Replace(".", "").Replace(",", ".");
					text = text.Replace("AA{0}", newValue);
					sheetView.Cells[(int)iRow - 1, idxCol_AN].Formula = string.Format(text, iRow);
				}
				else if (iCol == idxCol_AA + 1 && ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2)))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_AA].ResetFormula();
					string newValue = sheetView.Cells[(int)iRow - 1, idxCol_AA].Text.Replace(".", "").Replace(",", ".");
					text = text.Replace("AA{0}", newValue);
					sheetView.Cells[(int)iRow - 1, idxCol_AN].Formula = string.Format(text, iRow);
				}
				else if (iCol == idxCol_AG + 1 && ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2)))
				{
					string newValue = sheetView.Cells[(int)iRow - 1, idxCol_AA].Text.Replace(".", "").Replace(",", ".");
					text = text.Replace("AA{0}", newValue);
					sheetView.Cells[(int)iRow - 1, idxCol_AN].Formula = string.Format(text, iRow);
				}
			}
			else if (iCol == idxCol_AG + 1 && ((iRow >= dataLocationTemplateRowStart && iRow <= num) || (iRow >= dataLocationTemplateRowStart2 && iRow <= num2)))
			{
				string text = formatFormulaCt9;
				string newValue = sheetView.Cells[(int)iRow - 1, idxCol_AA].Text.Replace(".", "").Replace(",", ".");
				text = text.Replace("AA{0}", newValue);
				sheetView.Cells[(int)iRow - 1, idxCol_AN].Formula = string.Format(text, iRow);
			}
			int num3 = 0;
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			if (iCol == num4)
			{
				string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43"));
				string text2 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim();
				if (text2.StartsWith("-----") || text2.StartsWith("---"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num4 - 1, message);
					num3 = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num4 - 1);
				}
				if (num3 == 2)
				{
					winFpSpread.TabStrip[0].BackColor = Color.Pink;
				}
				else
				{
					winFpSpread.TabStrip[0].BackColor = Color.White;
				}
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
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "01_1_TTDB_I");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if ((iCol - 1 == idxCol_H_PL1 || iCol - 1 == idxCol_J_PL1) && iRow <= num)
			{
				sheetView.Cells[(int)iRow - 1, idxCol_K_PL1].Formula = string.Format(fmlCt10_PL1, iRow.ToString());
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo sInfoTK = ToKhaiInfo.SheetInfos[0];
			if (activeRowIndex == 2 && activeColumnIndex == 1)
			{
				ShowBangKe("Bangkekhautru_01TTDB.xls");
			}
			if (activeSheetIndex != 0 || activeRowIndex != CellId_N34[1] || activeColumnIndex != CellId_N34[0])
			{
				return;
			}
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_N34[1], CellId_N34[0]].Value);
			if (flag)
			{
				frmMessage frmMessage = new frmMessage("0394", "5", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				if (frmMessage.Yes)
				{
					ClearAndLockCellTK(sheetView, sInfoTK, isBindData: false, flag);
					if (ToKhaiInfo.SheetInfos[1] != null)
					{
						SheetView sViewPL = winFpSpread.Sheets[1];
						ClearCellNumberPL1(sViewPL, ToKhaiInfo.SheetInfos[1]);
					}
				}
				else
				{
					sheetView.Cells[CellId_N34[1], CellId_N34[0]].Value = false;
				}
			}
			else
			{
				ClearAndLockCellTK(sheetView, sInfoTK, isBindData: false, flag);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeRowIndex == 2 && activeColumnIndex == 1)
			{
				ShowBangKe("Bangkekhautru_01TTDB.xls");
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
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
			{
				result = true;
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					continue;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						result = false;
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
						continue;
					}
					break;
				}
				result = false;
				winFpSpread.TabStrip[key].BackColor = Color.White;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			bool flag = false;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43"));
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_N34[1], CellId_N34[0]].Value);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo obj) => obj.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietThueTTDB/tongThueKTru"));
			HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				object value = sheetView.Cells[14, HTKKFile.ColumnLettersToIndex("AI") - 1].Value;
				object value2 = sheetView.Cells[15, HTKKFile.ColumnLettersToIndex("AI") - 1].Value;
				if (value.ToNumber() > 0m && value2.ToNumber() <= 0m)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
					num = 2;
				}
				else if (value.ToNumber() > value2.ToNumber())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
					num = 2;
				}
			}
			for (int num2 = 0; num2 < hTKKCMCSheetInfo.Tables.Count; num2++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[num2];
				int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int num4 = 0; num4 < hTKKCMCTableInfo.RowsValues.Count; num4++)
				{
					if (hTKKCMCSheetInfo.Tables[num2].TableName == nameTable_01_TTDB_I || hTKKCMCSheetInfo.Tables[num2].TableName == nameTable_01_TTDB_II)
					{
						array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_G].Value);
						array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_P].Value);
						array[8] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_S].Value);
						array[9] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_V].Value);
						array[10] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_AA].Value);
						array[11] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_AD].Value);
						array[12] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_AG].Value);
						array[14] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_AN].Value);
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_V);
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_AA);
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_AG);
							continue;
						}
						if (string.IsNullOrEmpty(array[3]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_G, message5);
							num = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_G);
						}
						if (hTKKCMCSheetInfo.Tables[num2].TableName == nameTable_01_TTDB_I && !string.IsNullOrEmpty(array[3]))
						{
							string hhdv = array[3] + array[5];
							if (hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[3] + p.RowValue[5] == hhdv) > 1)
							{
								XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
								HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_G, xMLMessage.GetMessage("0261"));
								num = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_G);
							}
						}
						if (!flag2)
						{
							if (!string.IsNullOrEmpty(array[3]) && array[9] == hTKKCMCTableInfo.RowsInfo.Cells[9].DefaultValue)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_V, message);
								num = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_V);
							}
							if (!string.IsNullOrEmpty(array[3]) && array[10] == "0")
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_AA, message2);
								num = 2;
								flag = true;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_AA);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_V);
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_AA);
						}
						double num5 = Convert.ToDouble(array[14]);
						if (num5 < 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_AG, message4);
							num = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_AG);
						}
						double num6 = Convert.ToDouble(array[9]);
						double num7 = Convert.ToDouble(array[10]);
						if (!string.IsNullOrEmpty(array[3]) && num7 >= num6 && !flag2 && !flag)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_AA, message3);
							num = 2;
						}
						else if (!flag)
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_AA);
						}
					}
					else
					{
						array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_G].Value);
						array[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_P].Value);
						array[8] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_S].Value);
						array[9] = ConvertHelpers.ConvertToString(sheetView.Cells[num3 + num4, idxCol_V].Value);
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_G);
							continue;
						}
						if (string.IsNullOrEmpty(array[3]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_G, message5);
							num = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_G);
						}
						if (!string.IsNullOrEmpty(array[3]))
						{
							string hhdv2 = array[3] + array[5];
							if (hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[3] + p.RowValue[5] == hhdv2) > 1)
							{
								XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
								HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_G, xMLMessage.GetMessage("0261"));
								num = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num3 + num4, idxCol_G);
							}
						}
					}
					string text = sheetView.Cells[num3 + num4, idxCol_G].Text.Trim();
					if (text.StartsWith("-----"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num3 + num4, idxCol_G, message6);
						num = 2;
					}
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
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			string message = HTKKFile.GetMessage("0071");
			string message2 = HTKKFile.GetMessage("0112");
			string message3 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_21"));
			string message4 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_24"));
			string message5 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_25"));
			string message6 = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_26"));
			HTKKCMCTableInfo hTKKCMCTableInfo = null;
			if (hTKKCMCSheetInfo2 != null)
			{
				hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "01_1_TTDB_I");
			}
			if (hTKKCMCTableInfo != null)
			{
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
				{
					array[1] = sheetView2.Cells[num2 + num3, idxCol_C_PL1].Text;
					array[2] = sheetView2.Cells[num2 + num3, idxCol_D_PL1].Text;
					array[3] = sheetView2.Cells[num2 + num3, idxCol_E_PL1].Text;
					array[4] = sheetView2.Cells[num2 + num3, idxCol_F_PL1].Text;
					array[5] = ConvertHelpers.ConvertToString(sheetView2.Cells[num2 + num3, idxCol_G_PL1].Value);
					array[6] = ConvertHelpers.ConvertToString(sheetView2.Cells[num2 + num3, idxCol_H_PL1].Value);
					array[8] = ConvertHelpers.ConvertToString(sheetView2.Cells[num2 + num3, idxCol_J_PL1].Value);
					array[9] = ConvertHelpers.ConvertToString(sheetView2.Cells[num2 + num3, idxCol_K_PL1].Value);
					if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array))
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_E_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_F_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_G_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_H_PL1);
						continue;
					}
					string value = array[1].Replace(",", "").Replace(".", "").Replace("0", "") + array[2].Replace(",", "").Replace(".", "").Replace("0", "") + array[3].Replace(",", "").Replace(".", "").Replace("0", "") + array[4].Replace(",", "").Replace(".", "").Replace("0", "") + array[5].Replace(",", "").Replace(".", "").Replace("0", "") + array[6].Replace(",", "").Replace(".", "").Replace("0", "") + array[8].Replace(",", "").Replace(".", "").Replace("0", "") + array[9].Replace(",", "").Replace(".", "").Replace("0", "");
					if (string.IsNullOrEmpty(value))
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_C_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_D_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_E_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_F_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_G_PL1);
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_H_PL1);
						break;
					}
					if (HTKKFormater.isNullString(array[1]) && num3 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_C_PL1, message3);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_C_PL1);
					}
					if (HTKKFormater.isNullString(array[2]) && num3 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_D_PL1, message3);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_D_PL1);
					}
					if (!string.IsNullOrEmpty(array[3]))
					{
						if (DateTime.TryParseExact(array[3], "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
						{
							int num4 = 12 * ConvertHelpers.ConvertToInt(PeriodInfo.Nam) + ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
							int num5 = 12 * result.Year + result.Month;
							if (num5 > num4)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_E_PL1, message2);
								num = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_E_PL1);
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_E_PL1, message);
							num = 2;
						}
					}
					else if (num3 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_E_PL1, message3);
						num = 2;
					}
					if (string.IsNullOrEmpty(array[4]) && num3 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_F_PL1, message3);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_F_PL1);
					}
					if ((string.IsNullOrEmpty(array[5]) || Convert.ToDouble(array[5]) == 0.0) && num3 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_G_PL1, message3);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_G_PL1);
					}
					if ((string.IsNullOrEmpty(array[6]) || Convert.ToDouble(array[6]) == 0.0) && num3 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, num2 + num3, idxCol_H_PL1, message3);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView2, num2 + num3, idxCol_H_PL1);
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo2?.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "Pluc_01_1_TTDB/TTinThueNLMVao/tong_ct10");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo2?.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "Pluc_01_1_TTDB/TTinThueNLDuocKTru/tong_ct7");
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo?.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BangChiTietThueTTDB/tongThueKTru");
			double num6 = ((hTKKCMCCellInfo != null) ? Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value) : 0.0);
			double num7 = ((hTKKCMCCellInfo2 != null) ? Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value) : 0.0);
			double num8 = ((hTKKCMCCellInfo3 != null) ? Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value) : 0.0);
			if (hTKKCMCCellInfo2 != null)
			{
				if (num7 > num6)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView2, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message4);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround(sheetView2, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				}
			}
			switch (num)
			{
			case 2:
				sheetView2.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sheetView2.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sheetView2.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			SheetView sheetView = winFpSpread.Sheets[2];
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("L");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, GeneralFunctions.ArrayCellId("J_21")[1] + 1, GeneralFunctions.ArrayCellId("J_21")[0] + 1, "14");
			if (string.IsNullOrEmpty(valueCustom))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("J_21")[1], GeneralFunctions.ArrayCellId("j_21")[0], text2);
				value = 2;
			}
			else
			{
				HTKKFormater.IsDateTime(valueCustom, "dd/MM/yyyy", out var dtDate);
				HTKKFormater.IsDateTime(valueCustom, "dd/MM/yyyy", out dtDate);
				if (DateTime.Compare(dtDate, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, GeneralFunctions.ArrayCellId("J_21")[1], GeneralFunctions.ArrayCellId("J_21")[0], text);
					value = 2;
				}
			}
			string pathTongCong = "Pluc_01_TT146/PLPhieuTHDoanhThu/BKDoanhThu/tongCong";
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathTongCong);
			HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			double num5 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num6 = HTKKFile.ColumnLettersToIndex("V");
			int num7 = HTKKFile.ColumnLettersToIndex("F");
			double value2 = 0.0;
			for (int num8 = 0; num8 < hTKKCMCTableInfo.RowsValues.Count; num8++)
			{
				string text4 = winFpSpread.Sheets[0].Cells[dataLocationTemplateRowStart + num8 - 1, num7 - 1].Text;
				double num9 = Convert.ToDouble(winFpSpread.Sheets[0].Cells[dataLocationTemplateRowStart + num8 - 1, num6 - 1].Value);
				if (text4 == maCaSiNo)
				{
					value2 = num9;
					break;
				}
			}
			if (Math.Round(value2, 0) != Math.Round(num5 / 1.1, 0))
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, text3);
				if (value != 2)
				{
					value = 1;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ClearAndLockCellTK(SheetView sViewTK, HTKKCMCSheetInfo sInfoTK, bool isBindData, bool isCheckN34)
	{
		try
		{
			List<int> list = new List<int>();
			list.Add(idxCol_S);
			list.Add(idxCol_V);
			list.Add(idxCol_AA);
			list.Add(idxCol_AD);
			List<int> list2 = list;
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfoTK.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "01_TTDB_I");
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
			{
				for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsInfo.Cells.Count; num3++)
				{
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells[num3];
					if (list2.Contains(hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1))
					{
						if (!isBindData && isCheckN34)
						{
							SetDefaultValueToTableCell(sViewTK, hTKKCMCTableCellInfo, num + num2);
						}
						if (idxCol_AD != hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1)
						{
							sViewTK.Cells[num + num2, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Locked = isCheckN34;
						}
					}
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfoTK.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "01_TTDB_II");
			int num4 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int num2 = 0; num2 < hTKKCMCTableInfo2.RowsValues.Count; num2++)
			{
				for (int num3 = 0; num3 < hTKKCMCTableInfo2.RowsInfo.Cells.Count; num3++)
				{
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo2.RowsInfo.Cells[num3];
					if (list2.Contains(hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1))
					{
						if (!isBindData && isCheckN34)
						{
							SetDefaultValueToTableCell(sViewTK, hTKKCMCTableCellInfo, num4 + num2);
						}
						if (idxCol_AD != hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1)
						{
							sViewTK.Cells[num4 + num2, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Locked = isCheckN34;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ClearCellNumberPL1(SheetView sViewPL1, HTKKCMCSheetInfo sInfoPL1)
	{
		try
		{
			for (int i = 0; i < sInfoPL1.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfoPL1.Tables[i];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				if (hTKKCMCTableInfo.RowsValues.Count <= 0)
				{
					continue;
				}
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					for (int k = 0; k < hTKKCMCTableInfo.RowsInfo.Cells.Count; k++)
					{
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells[k];
						ICellType cellType = sViewPL1.Cells[num + j, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].CellType;
						if (hTKKCMCTableCellInfo.Controltype == "16" || (hTKKCMCTableCellInfo.Controltype == "0" && cellType.ToString() == "NumberCellType"))
						{
							string value = ((hTKKCMCTableCellInfo.DefaultValue != null) ? hTKKCMCTableCellInfo.DefaultValue : "0");
							sViewPL1.Cells[num + j, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value = value;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void SetDefaultValueToTableCell(SheetView sView, HTKKCMCTableCellInfo cellInfo, int idxRow)
	{
		try
		{
			string text = "";
			if (cellInfo.DefaultValue != null)
			{
				text = cellInfo.DefaultValue;
			}
			else if (cellInfo.Controltype != null)
			{
				text = ((!(cellInfo.Controltype == "16")) ? "" : "0");
			}
			else
			{
				ICellType cellType = sView.Cells[idxRow, cellInfo.CellTemplate.ColIndex - 1].CellType;
				if (cellType != null)
				{
					text = ((!(cellType.ToString() == "NumberCellType")) ? "" : "0");
				}
			}
			sView.Cells[idxRow, cellInfo.CellTemplate.ColIndex - 1].Text = text;
			sView.Cells[idxRow, cellInfo.CellTemplate.ColIndex - 1].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void SetDefaultValueToRow(SheetView sViewPL1, HTKKCMCRowInfo rowInfo, int idxRow)
	{
		try
		{
			foreach (HTKKCMCTableCellInfo cell in rowInfo.Cells)
			{
				string text = "";
				int column = cell.CellTemplate.ColIndex - 1;
				if (cell.DefaultValue != null)
				{
					text = cell.DefaultValue;
				}
				else if (cell.Controltype != null)
				{
					text = ((!(cell.Controltype == "16")) ? "" : "0");
				}
				else
				{
					ICellType cellType = sViewPL1.Cells[idxRow, column].CellType;
					text = ((!(cellType.ToString() == "NumberCellType")) ? "" : "0");
				}
				sViewPL1.Cells[idxRow, column].Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
				{
					string key = array[1];
					string text = array[2].Trim();
					string text2 = array[4].Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						text = text + " (" + text2 + " %)";
					}
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDMThueDB = dictionary;
			dicDMThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_04d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_04da: Expected O, but got Unknown
		//IL_0508: Unknown result type (might be due to invalid IL or missing references)
		//IL_050f: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> temp = new Dictionary<string, string[]>();
			Dictionary<string, string> dic = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("1"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								if (!string.IsNullOrEmpty(item.thuesuat))
								{
									text = text + " (" + item.thuesuat + " %)";
								}
								GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
								if (!dic.ContainsKey(key))
								{
									dic.Add(key, text);
								}
								if (!temp.ContainsKey(text))
								{
									temp.Add(text, array);
								}
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("1") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							if (!string.IsNullOrEmpty(item.thuesuat))
							{
								text = text + " (" + item.thuesuat + " %)";
							}
							GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
							if (!dic.ContainsKey(key))
							{
								dic.Add(key, text);
							}
							if (!temp.ContainsKey(text))
							{
								temp.Add(text, array);
							}
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						string text2 = array[4].Trim();
						if (!string.IsNullOrEmpty(text2))
						{
							text = text + " (" + text2 + " %)";
						}
						if (!dic.ContainsKey(key))
						{
							dic.Add(key, text);
						}
						if (!temp.ContainsKey(text))
						{
							temp.Add(text, array);
						}
					}
				}
			}
			dicDMThueDB = temp;
			dicDMThueDBTmp = dic;
			return dic;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDichVuThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDichVuThueDBTmp.Count > 0)
			{
				return dicDichVuThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[1].StartsWith("2") && array[^1] == "1")
				{
					string key = array[1];
					string text = array[2].Trim();
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDichVuThueDB = dictionary;
			dicDichVuThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDichVuThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_0477: Unknown result type (might be due to invalid IL or missing references)
		//IL_047e: Expected O, but got Unknown
		//IL_04ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b3: Expected O, but got Unknown
		try
		{
			if (dicDichVuThueDBTmp.Count > 0)
			{
				return dicDichVuThueDBTmp;
			}
			Dictionary<string, string[]> temp = new Dictionary<string, string[]>();
			Dictionary<string, string> dic = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("2"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
								if (!dic.ContainsKey(key))
								{
									dic.Add(key, text);
								}
								if (!temp.ContainsKey(text))
								{
									temp.Add(text, array);
								}
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("2") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							GeneralFunctions.GetParentDanhMuc(ref dic, ref temp, item.macha, danhmuc.items);
							if (!dic.ContainsKey(key))
							{
								dic.Add(key, text);
							}
							if (!temp.ContainsKey(text))
							{
								temp.Add(text, array);
							}
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("2") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						if (!dic.ContainsKey(key))
						{
							dic.Add(key, text);
						}
						if (!temp.ContainsKey(text))
						{
							temp.Add(text, array);
						}
					}
				}
			}
			dicDichVuThueDB = temp;
			dicDichVuThueDBTmp = dic;
			return dic;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMTTDBDonVi(string pathFile)
	{
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		try
		{
			if (dicDMDonVi.Count > 0)
			{
				return dicDMDonVi;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			int num = 0;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["value"]).Value;
				num++;
				dictionary.Add(num.ToString(), value);
			}
			dicDMDonVi = dictionary;
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string[] array = new string[5] { "01_TTDB_I", "01_TTDB_II", "01_TTDB_III_A", "01_TTDB_III_B", "01_TTDB_III_C" };
			string[] array2 = null;
			string[] array3 = null;
			string[] array4 = array;
			foreach (string tableName in array4)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
				Predicate<HTKKCMCTableInfo> match = (HTKKCMCTableInfo table) => table.TableName == tableName;
				HTKKCMCTableInfo hTKKCMCTableInfo = tables.Find(match);
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType = (ComboBoxCellType)sheetView.Cells[dataLocationTemplateRowStart - 1, HTKKFile.ColumnLettersToIndex("G") - 1].Editor;
				array2 = comboBoxCellType.Items;
				array3 = comboBoxCellType.ItemData;
				string empty = string.Empty;
				string empty2 = string.Empty;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int num2 = dataLocationTemplateRowStart; num2 <= num; num2++)
				{
					empty2 = sheetView.Cells[num2 - 1, HTKKFile.ColumnLettersToIndex("F") - 1].Text;
					ICellType cellType = sheetView.Cells[num2 - 1, HTKKFile.ColumnLettersToIndex("G") - 1].CellType;
					if (cellType.ToString() == "TaxComboBoxCellType")
					{
						List<string> list = new List<string>(array3);
						if (list.Contains(empty2))
						{
							int num3 = list.IndexOf(empty2);
							empty = array2[num3];
							sheetView.Cells[num2 - 1, HTKKFile.ColumnLettersToIndex("G") - 1].Text = empty;
						}
					}
				}
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
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count - 1; i++)
			{
				for (int j = 0; j < hTKKCMCSheetInfo.Tables[i].RowsValues.Count; j++)
				{
					CellChange_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j, idxCol_J_PL1 + 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			if (activeRowIndex == 5 && activeColumnIndex == 1)
			{
				ShowBangKe("BangKe_THDT_TTDB.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
