using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ntnn01Nckt207
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
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AD") - 1].Locked = true;
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
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = PeriodInfo.kyKKhai;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AD") - 1].Locked = true;
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
				ShowBangKe("Bang_Ke_01NTNN.xls");
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
			int[] array = new int[2]
			{
				(int)iCol - 1,
				(int)iRow - 1
			};
			List<int[]> list = new List<int[]>();
			list.Add(GeneralFunctions.ArrayCellId("AI_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("AM_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("AU_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("AY_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("BC_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("BG_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("BM_" + iRow));
			List<int[]> list2 = list;
			string nameTable = "01_NTNN";
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (array[1] < num)
			{
				XuLyCellChange(winFpSpread.Sheets[0], array, list2[0], list2[1], list2[2], list2[3], list2[4], list2[5], list2[6]);
				return;
			}
			string value = GeneralFunctions.GetValue(winFpSpread.Sheets[0], "BF_33");
			int[] array2 = GeneralFunctions.ArrayCellId("S_33");
			string text = MoneyHelpers.MoneyToText(value);
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void XuLyCellChange(SheetView sView, int[] cellChangeID, int[] ct6, int[] ct7, int[] ct8, int[] ct9, int[] ct10, int[] ct11, int[] ct12)
	{
		try
		{
			string formula = string.Format("ROUND(AI{0}*AM{0}/100,0)", (cellChangeID[1] + 1).ToString());
			string formula2 = string.Format("IF(ROUND(AY{0}*BC{0}/100,0)-BG{0} < 0,0,ROUND(AY{0}*BC{0}/100,0)-BG{0})", (cellChangeID[1] + 1).ToString());
			if (ct6[0] == cellChangeID[0] && ct6[1] == cellChangeID[1])
			{
				sView.Cells[ct8[1], ct8[0]].Formula = formula;
			}
			else if (ct7[0] == cellChangeID[0] && ct7[1] == cellChangeID[1])
			{
				sView.Cells[ct8[1], ct8[0]].Formula = formula;
			}
			else if (ct8[0] == cellChangeID[0] && ct8[1] == cellChangeID[1])
			{
				sView.Cells[ct8[1], ct8[0]].Formula = "";
			}
			if (ct9[0] == cellChangeID[0] && ct9[1] == cellChangeID[1])
			{
				sView.Cells[ct12[1], ct12[0]].Formula = formula2;
				if (ConvertHelpers.ConvertToDouble(sView.Cells[ct9[1], ct9[0]].Value) < 0.0)
				{
					sView.Cells[ct11[1], ct11[0]].Value = 0;
					sView.Cells[ct11[1], ct11[0]].Locked = true;
				}
				else
				{
					sView.Cells[ct11[1], ct11[0]].Locked = false;
				}
			}
			else if (ct10[0] == cellChangeID[0] && ct10[1] == cellChangeID[1])
			{
				sView.Cells[ct12[1], ct12[0]].Formula = formula2;
			}
			else if (ct11[0] == cellChangeID[0] && ct11[1] == cellChangeID[1])
			{
				sView.Cells[ct12[1], ct12[0]].Formula = formula2;
			}
			else if (ct12[0] == cellChangeID[0] && ct12[1] == cellChangeID[1])
			{
				sView.Cells[ct12[1], ct12[0]].Formula = "";
			}
			string value = GeneralFunctions.GetValue(sView, "BF_33");
			int[] array = GeneralFunctions.ArrayCellId("S_33");
			string text = MoneyHelpers.MoneyToText(value);
			sView.Cells[array[1], array[0]].Text = text;
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
			int colunmIndexByColumName3 = HTKKFormater.GetColunmIndexByColumName("R");
			int colunmIndexByColumName4 = HTKKFormater.GetColunmIndexByColumName("X");
			int colunmIndexByColumName5 = HTKKFormater.GetColunmIndexByColumName("AD");
			int colunmIndexByColumName6 = HTKKFormater.GetColunmIndexByColumName("AI");
			int colunmIndexByColumName7 = HTKKFormater.GetColunmIndexByColumName("AM");
			int colunmIndexByColumName8 = HTKKFormater.GetColunmIndexByColumName("AU");
			int colunmIndexByColumName9 = HTKKFormater.GetColunmIndexByColumName("AY");
			int colunmIndexByColumName10 = HTKKFormater.GetColunmIndexByColumName("BC");
			int colunmIndexByColumName11 = HTKKFormater.GetColunmIndexByColumName("BG");
			int colunmIndexByColumName12 = HTKKFormater.GetColunmIndexByColumName("BM");
			int colunmIndexByColumName13 = HTKKFormater.GetColunmIndexByColumName("BQ");
			string[] array = new string[14];
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			bool flag = false;
			bool flag2 = false;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName);
				array[1] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName2);
				array[2] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName3);
				array[3] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName4);
				array[4] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName5);
				array[5] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName6);
				array[6] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName7);
				array[7] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName8);
				array[8] = "";
				array[9] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName9);
				array[10] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName10);
				array[11] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName11);
				array[12] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName12);
				array[13] = ConvertHelpers.ConvertCellValueToText(sView, num + i, colunmIndexByColumName13);
				bool flag3 = GeneralFunctions.HaveToValidateRow(rowsInfo, array);
				if (!flag3)
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName8);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName12);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName3);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName5);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName2);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName4);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName6);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName7);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName9);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName10);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName11);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName13);
					continue;
				}
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName8);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName12);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName3);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName5);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName2);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName4);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName6);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName7);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName9);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName10);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName11);
				HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName13);
				if (ConvertHelpers.ConvertToDouble(array[7]) != 0.0)
				{
					flag = true;
				}
				if (ConvertHelpers.ConvertToDouble(array[12]) != 0.0)
				{
					flag2 = true;
				}
				if (string.IsNullOrEmpty(array[0]) && flag3)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName);
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName3, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName3);
				}
				if (!string.IsNullOrEmpty(array[4]))
				{
					DateTime dateTime = DateTime.ParseExact(array[4], "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (dateTime > DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
						value = 2;
					}
					else if (dateTime < DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture))
					{
						HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName5);
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
					value = 2;
				}
				if (!ValidateBase.CheckMaSoThue(array[1]) && !string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_19")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName2);
				}
				double num2 = ConvertHelpers.ConvertToDouble(array[7]);
				double num3 = Math.Round(ConvertHelpers.ConvertToDouble(array[5]) * ConvertHelpers.ConvertToDouble(array[6]) / 100.0, 0);
				double num4 = ConvertHelpers.ConvertToDouble(array[12]);
				double num5 = Math.Round(ConvertHelpers.ConvertToDouble(array[9]) * ConvertHelpers.ConvertToDouble(array[10]) / 100.0, 0);
				num5 = ((!(num5 - ConvertHelpers.ConvertToDouble(array[11]) > 0.0)) ? 0.0 : (num5 - ConvertHelpers.ConvertToDouble(array[11])));
				if (num2 != num3)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName8, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_26")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName8);
				}
				if (num4 != num5)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName12, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_26")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName12);
				}
				double num6 = Convert.ToDouble(array[10]);
				double num7 = Convert.ToDouble(array[9]);
				double num8 = Convert.ToDouble(array[11]);
				if (num8 > Math.Round(num7 * num6 / 100.0, 0))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_28")));
					value = 2;
				}
				if (!string.IsNullOrEmpty(array[2].Replace("0", "")) && !string.IsNullOrEmpty(array[4].Replace("0", "")))
				{
					continue;
				}
				if (string.IsNullOrEmpty(array[3].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[5].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName6, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[6].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName7, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[7].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName8, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[9].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName9, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[10].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName10, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[11].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName11, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[12].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName12, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
				if (string.IsNullOrEmpty(array[13].Replace("0", "")))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, num + i, colunmIndexByColumName13, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (value == 0)
					{
						value = 1;
					}
				}
			}
			int num9 = num + hTKKCMCTableInfo.RowsValues.Count + 1;
			double num10 = ConvertHelpers.ConvertToDouble(sView.Cells[num9, colunmIndexByColumName8].Value);
			double num11 = ConvertHelpers.ConvertToDouble(sView.Cells[num9, colunmIndexByColumName12].Value);
			if (flag)
			{
				if (num10 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num9, colunmIndexByColumName8, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num9, colunmIndexByColumName8);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, num9, colunmIndexByColumName8);
			}
			if (flag2)
			{
				if (num11 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sView, num9, colunmIndexByColumName12, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num9, colunmIndexByColumName12);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, num9, colunmIndexByColumName12);
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
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < hTKKCMCSheetInfo.Tables[i].RowsValues.Count; j++)
				{
					CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("AM"));
					CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("BG"));
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
