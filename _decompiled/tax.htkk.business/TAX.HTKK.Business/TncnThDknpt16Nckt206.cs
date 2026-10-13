using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnThDknpt16Nckt206
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			ButtonCellType buttonCellType = new ButtonCellType();
			buttonCellType.TwoState = false;
			buttonCellType.Text = HTKKFile.GetMessage("0346");
			buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
			buttonCellType.TextColor = Color.Blue;
			sView.Cells[0, HTKKFile.ColumnLettersToIndex("D") - 1].CellType = buttonCellType;
			List<HTKKCMCRowValue> rowsValues = ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues;
			int column = GeneralFunctions.ArrayCellId("R_40")[0];
			int column2 = GeneralFunctions.ArrayCellId("T_40")[0];
			int column3 = GeneralFunctions.ArrayCellId("V_40")[0];
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int row = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + i - 1;
				if (ConvertHelpers.ConvertToString(sView.Cells[row, HTKKFile.ColumnLettersToIndex("P") - 1].Value) == "Khác")
				{
					sView.Cells[row, column].Locked = true;
					sView.Cells[row, column2].Locked = true;
					sView.Cells[row, column3].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void WinInsertCombox(SheetView sView, int iRow, int iCol, string[] Values)
	{
		try
		{
			if (iCol >= 1 && iRow >= 1)
			{
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = Values;
				sView.Cells[iRow - 1, iCol - 1].CellType = comboBoxCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("O");
			int num2 = HTKKFile.ColumnLettersToIndex("P");
			int num3 = HTKKFile.ColumnLettersToIndex("Q");
			int num4 = HTKKFile.ColumnLettersToIndex("R");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("T");
			int num7 = HTKKFile.ColumnLettersToIndex("U");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("H");
			int num10 = HTKKFile.ColumnLettersToIndex("I");
			if (iCol == num2)
			{
				WinInsertCombox(sheetView, (int)iRow, num4, HTKKWinDictionaryDiaDanh.TinhTP);
				WinInsertCombox(sheetView, (int)iRow, num6, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP["ALL"]);
				WinInsertCombox(sheetView, (int)iRow, num8, HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen["ALL"]);
				string value = "";
				HTKKWinDictionaryDiaDanh.QuocTich.TryGetValue(sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text, out value);
				if (value == "02")
				{
					HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "QuocTich", "True");
					sheetView.Cells[(int)iRow - 1, num4 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = "Khác";
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = "Khác";
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = "Khác";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "";
				}
				else
				{
					HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "QuocTich", "False");
					sheetView.Cells[(int)iRow - 1, num4 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol, out object Obj)
	{
		try
		{
			Obj = null;
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num2 = 0;
			if (idxCol == num && idxRow == num2)
			{
				Obj = getDuLieuQTT05PL3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			}
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_16TH_DangKyCapMa.xls");
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

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		//IL_0137: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0176: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = (DateTime.Now.Year - 1) * 12 + DateTime.Now.Month;
			int num2 = (DateTime.Now.Year - 1) * 4 + DateTime.Now.Month / 3;
			int num3 = int.Parse(PeriodInfo.Nam);
			if (PeriodInfo.KieuKy == "M")
			{
				int num4 = (num3 - 1) * 12 + int.Parse(PeriodInfo.Thang);
				if (num4 <= 0)
				{
					return false;
				}
				if (num4 > num)
				{
					frmMessage frmMessage = new frmMessage("0062", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				int num5 = (num3 - 1) * 4 + int.Parse(PeriodInfo.Quy);
				if (num5 > num2)
				{
					frmMessage frmMessage = new frmMessage("0045", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				int num6 = 8059;
				if (num5 < num6)
				{
					frmMessage frmMessage = new frmMessage("01XSBHDC0007", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		if (winFpSpread.TabStrip[0].BackColor != Color.White || winFpSpread.TabStrip[1].BackColor != Color.White)
		{
			return false;
		}
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		int num = 0;
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[78, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[85, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[73, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[84, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[76, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[82, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[83, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			List<int> list = new List<int>();
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("D");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("F");
			int num6 = HTKKFile.ColumnLettersToIndex("G");
			int num7 = HTKKFile.ColumnLettersToIndex("H");
			int num8 = HTKKFile.ColumnLettersToIndex("I");
			int num9 = HTKKFile.ColumnLettersToIndex("J");
			int num10 = HTKKFile.ColumnLettersToIndex("K");
			int num11 = HTKKFile.ColumnLettersToIndex("L");
			int num12 = HTKKFile.ColumnLettersToIndex("M");
			int num13 = HTKKFile.ColumnLettersToIndex("N");
			int num14 = HTKKFile.ColumnLettersToIndex("O");
			int num15 = HTKKFile.ColumnLettersToIndex("P");
			int num16 = HTKKFile.ColumnLettersToIndex("Q");
			int num17 = HTKKFile.ColumnLettersToIndex("R");
			int num18 = HTKKFile.ColumnLettersToIndex("S");
			int num19 = HTKKFile.ColumnLettersToIndex("T");
			int num20 = HTKKFile.ColumnLettersToIndex("U");
			int num21 = HTKKFile.ColumnLettersToIndex("V");
			int num22 = HTKKFile.ColumnLettersToIndex("W");
			int num23 = HTKKFile.ColumnLettersToIndex("X");
			double num24 = 0.0;
			bool flag = true;
			string mst = TTinChung.mst;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text10 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string text11 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string text12 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text13 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string valCotMstNPT = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text14 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num7 - 1].Text;
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Text;
					string text19 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Text;
					string text20 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num13 - 1].Text;
					string text21 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num14 - 1].Text;
					string text22 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num15 - 1].Text;
					string text23 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num16 - 1].Text;
					string text24 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num17 - 1].Text;
					string text25 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num18 - 1].Text;
					string text26 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num19 - 1].Text;
					string text27 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num20 - 1].Text;
					string text28 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num21 - 1].Text;
					string text29 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num22 - 1].Text;
					string text30 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num23 - 1].Text;
					string text31 = "";
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
					if (i == 0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num13 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num16 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num20 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num21 - 1);
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1);
					if (HTKKFormater.isNullString(text10) && HTKKFormater.isNullString(text11) && HTKKFormater.isNullString(text12) && HTKKFormater.isNullString(text13) && HTKKFormater.isNullString(valCotMstNPT) && HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(text16) && HTKKFormater.isNullString(text18) && HTKKFormater.isNullString(text19) && HTKKFormater.isNullString(text20) && HTKKFormater.isNullString(text22) && HTKKFormater.isNullString(text24) && HTKKFormater.isNullString(text26) && HTKKFormater.isNullString(text28) && HTKKFormater.isNullString(text29) && HTKKFormater.isNullString(text30))
					{
						continue;
					}
					flag = false;
					if (HTKKFormater.isNullString(text10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text11))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
						list.Add(1);
					}
					else if (!ValidateBase.CheckMaSoThue(text11))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
						list.Add(1);
					}
					else if (text11.Equals(mst))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_86")));
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text12))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
						list.Add(0);
					}
					if (i == 0)
					{
						if (HTKKFormater.isNullString(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text);
							list.Add(1);
						}
						else
						{
							bool success = false;
							DateTime t = HTKKFormater.ConvertToDateTime(text13, out success);
							if (success)
							{
								if (DateTime.Compare(t, DateTime.Now) > 0)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text5);
									list.Add(1);
								}
								else if (!HTKKFormater.isNullString(text29))
								{
									bool success2 = false;
									DateTime dateTime = HTKKFormater.ConvertToDateTime("01/" + text29, out success2);
									if (success2)
									{
										if ((t.Year == dateTime.Year && t.Month > dateTime.Month) || t.Year > dateTime.Year)
										{
											HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, text6);
											list.Add(1);
										}
										else
										{
											HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
											list.Add(0);
										}
									}
								}
								else
								{
									HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
									list.Add(0);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1, HTKKFile.GetMessage("0403"));
								list.Add(1);
							}
						}
					}
					if (i > 0)
					{
						HTKKFile.RemoveValidateBackGround(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
					}
					if (i == 0)
					{
						if (!HTKKFormater.isNullString(text16))
						{
							if (text16.Trim().IndexOf(" ") > -1)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
								list.Add(1);
							}
						}
						else if (!HTKKFormater.isNullString(text13) && DateTime.Now.Year - ConvertHelpers.ConvertToInt(text13.Split(new char[1] { '/' })[2]) > 15)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, text8);
							list.Add(1);
						}
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
					if (HTKKFormater.isNullString(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text);
						list.Add(1);
					}
					if (i == 0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num21 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num16 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num20 - 1);
						if (!HTKKFormater.isNullString(text13) && HTKKFormater.isNullString(text16))
						{
							if (HTKKFormater.isNullString(text19) && DateTime.Now.Year - ConvertHelpers.ConvertToInt(text13.Split(new char[1] { '/' })[2]) <= 15)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text);
								list.Add(1);
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
								list.Add(0);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
							list.Add(0);
						}
						if (HTKKFormater.isNullString(text16))
						{
							if (HTKKFormater.isNullString(text22))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, text);
								list.Add(1);
							}
							if (HTKKFormater.isNullString(text24))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1, text);
								list.Add(1);
							}
							if (HTKKFormater.isNullString(text26))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, text);
								list.Add(1);
							}
							if (HTKKFormater.isNullString(text28) && !text27.Equals("5050307"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num21 - 1, text);
								list.Add(1);
							}
						}
						if (!HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(text14))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text);
							num = 2;
						}
						if (!HTKKFormater.isNullString(text18) && HTKKFormater.isNullString(text17))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text);
							num = 2;
						}
						if (!HTKKFormater.isNullString(text22) && HTKKFormater.isNullString(text21))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num14 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num15 - 1, text);
							num = 2;
						}
						if (!HTKKFormater.isNullString(text24) && HTKKFormater.isNullString(text23) && text24 != "Khác")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num16 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1, text);
							num = 2;
						}
						if (!HTKKFormater.isNullString(text26) && HTKKFormater.isNullString(text25) && text26 != "Khác")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, text);
							num = 2;
						}
						if (!HTKKFormater.isNullString(text28) && HTKKFormater.isNullString(text27) && text28 != "Khác")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num20 - 1, text);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num21 - 1, text);
							num = 2;
						}
						if (!HTKKFormater.isNullString(text27) && !HTKKFormater.isNullString(text25))
						{
							if (text27.Length == 7 && text25.Length == 5)
							{
								if (!text27.Substring(0, 5).Equals(text25))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num20 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num21 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
									num = 2;
								}
							}
							else if (text27.Length != 7)
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num20 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_91")));
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num21 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_91")));
							}
						}
						if (!HTKKFormater.isNullString(text25) && !HTKKFormater.isNullString(text23))
						{
							if (text25.Length == 5 && text23.Length == 3)
							{
								if (!text25.Substring(0, 3).Equals(text23))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
									num = 2;
								}
							}
							else if (text25.Length != 5)
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num18 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_92")));
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num19 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_92")));
							}
						}
						if (!HTKKFormater.isNullString(text23) && !HTKKFormater.isNullString(text21) && text23.Length != 3)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num16 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_93")));
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num17 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_93")));
						}
					}
					if (i == 1)
					{
						if (HTKKFormater.isNullString(valCotMstNPT))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text);
							list.Add(1);
						}
						else if (!ValidateBase.CheckMaSoThue(valCotMstNPT))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text3);
							list.Add(1);
						}
						else if (!HTKKFormater.isNullString(text11) && !HTKKFormater.isNullString(text11) && text11.Equals(valCotMstNPT))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_87")));
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_87")));
							list.Add(1);
						}
					}
					DateTime t2 = DateTime.MinValue;
					DateTime t3 = DateTime.MinValue;
					bool success3 = false;
					bool success4 = false;
					if (HTKKFormater.isNullString(text29))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1, text);
						list.Add(1);
					}
					else
					{
						t2 = HTKKFormater.ConvertToDateTime("01/" + text29, out success3);
						if (success3)
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1);
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1, text9);
							list.Add(1);
						}
					}
					if (!HTKKFormater.isNullString(text30))
					{
						t3 = HTKKFormater.ConvertToDateTime("01/" + text30, out success4);
						if (success4)
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1);
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1, text9);
							list.Add(1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1);
					}
					if (!success3 || !success4)
					{
						continue;
					}
					if (DateTime.Compare(t2, t3) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1, text7);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1, text7);
						list.Add(1);
						continue;
					}
					if (HTKKFormater.isNullString(valCotMstNPT))
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1);
						continue;
					}
					List<HTKKCMCRowValue> listMstDuplicate = rowsValues.FindAll((HTKKCMCRowValue p) => p.RowValue[4].Equals(valCotMstNPT));
					int num25 = validMstNPT(rowsValues[j], listMstDuplicate);
					if (num25 > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_94")));
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_94")));
						list.Add(1);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num22 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num23 - 1);
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					num = 2;
					break;
				}
			}
			switch (num)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private List<HTKKCMCRowValue> getDuLieuQTT05PL3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0108: Unknown result type (might be due to invalid IL or missing references)
		//IL_010f: Expected O, but got Unknown
		//IL_017d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0184: Expected O, but got Unknown
		List<HTKKCMCRowValue> result = null;
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst;
			bool flag = false;
			if (Directory.Exists(text))
			{
				DirectoryInfo directoryInfo = new DirectoryInfo(text);
				FileInfo[] files = directoryInfo.GetFiles("*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (fileInfo.Name.IndexOf("05_3_QTT_TNCN") > -1 && fileInfo.Name.IndexOf(PeriodInfo.Nam) > 0)
					{
						text = fileInfo.FullName;
						flag = true;
						break;
					}
				}
			}
			if (!flag)
			{
				return new List<HTKKCMCRowValue>();
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
			HTKKCMCXML hTKKCMCXML = new HTKKCMCXML(text);
			XmlNodeList childNodes = ((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections").ChildNodes;
			foreach (XmlNode item in childNodes)
			{
				XmlNode val = item;
				if (!(val.Name == "Section"))
				{
					continue;
				}
				if (((XmlNode)val.Attributes["Dynamic"]).Value == "0")
				{
					XmlNodeList childNodes2 = val.SelectSingleNode("Cells").ChildNodes;
					foreach (XmlNode item2 in childNodes2)
					{
						XmlNode xmlCellNode = item2;
						hTKKCMCSheetInfo.StaticCells.Add(hTKKCMCSheetInfo.GetStaticCellInfo(xmlCellNode));
					}
				}
				else
				{
					hTKKCMCSheetInfo.Tables.Add(hTKKCMCSheetInfo.GetTable(val));
				}
			}
			if (hTKKCMCSheetInfo.Tables.Count > 0 && hTKKCMCSheetInfo.Tables[0].RowsValues.Count > 0)
			{
				for (int j = 0; j < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; j++)
				{
					hTKKCMCSheetInfo.Tables[0].RowsValues[j].RowValue.Add("");
				}
				result = hTKKCMCSheetInfo.Tables[0].RowsValues;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void InsertTableRow(SheetView sView, int frRow, int frCol, int toRow, int toCol, int colCount, HTKKCMCSheetInfo sInfo, int iCurTable)
	{
		sView.Rows.Add(frRow, 1);
		sView.CopyRange(frRow - 1, frCol, toRow, toCol, 1, colCount, dataOnly: false);
		sInfo.IncrementStaticCell(frRow, 1);
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table.Location.DataLocationTemplateRowStart > frRow)
			{
				table.Increment(1);
			}
		}
		sInfo.Tables[iCurTable].addNewDefaultRow(frRow - sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + 1);
	}

	public static void fpSpBindDataCell(SheetView sView, int iRow, int iCol, string style, string values)
	{
		try
		{
			if (iRow == 54)
			{
				int num = 0;
			}
			if (iRow <= 0 || iCol <= 0)
			{
				return;
			}
			switch (style)
			{
			case "2":
				if (values == "True")
				{
					sView.Cells[iRow - 1, iCol - 1].Value = true;
				}
				else
				{
					sView.Cells[iRow - 1, iCol - 1].Value = false;
				}
				return;
			case "7":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				return;
			case "6":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				return;
			case "5":
				sView.Cells[iRow - 1, iCol - 1].Text = values;
				return;
			}
			if (!sView.Cells[iRow - 1, iCol - 1].Locked && !string.IsNullOrEmpty(values) && !string.IsNullOrEmpty(sView.Cells[iRow - 1, iCol - 1].Formula))
			{
				sView.Cells[iRow - 1, iCol - 1].ResetFormula();
			}
			if (sView.Cells[iRow - 1, iCol - 1].CellType.GetType().Name == "NumberCellType")
			{
				sView.Cells[iRow - 1, iCol - 1].Value = ConvertHelpers.ConvertToDouble(values);
			}
			else
			{
				sView.Cells[iRow - 1, iCol - 1].Text = values;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("F");
			int num5 = HTKKFile.ColumnLettersToIndex("G");
			int num6 = HTKKFile.ColumnLettersToIndex("H");
			int num7 = HTKKFile.ColumnLettersToIndex("I");
			int num8 = HTKKFile.ColumnLettersToIndex("J");
			int num9 = HTKKFile.ColumnLettersToIndex("K");
			int num10 = HTKKFile.ColumnLettersToIndex("L");
			int num11 = HTKKFile.ColumnLettersToIndex("M");
			int num12 = HTKKFile.ColumnLettersToIndex("N");
			int num13 = HTKKFile.ColumnLettersToIndex("O");
			int num14 = HTKKFile.ColumnLettersToIndex("P");
			int num15 = HTKKFile.ColumnLettersToIndex("Q");
			int num16 = HTKKFile.ColumnLettersToIndex("R");
			int num17 = HTKKFile.ColumnLettersToIndex("S");
			int num18 = HTKKFile.ColumnLettersToIndex("T");
			int num19 = HTKKFile.ColumnLettersToIndex("U");
			int num20 = HTKKFile.ColumnLettersToIndex("V");
			int num21 = HTKKFile.ColumnLettersToIndex("W");
			int num22 = HTKKFile.ColumnLettersToIndex("X");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count - 1; i++)
			{
				for (int j = 0; j < hTKKCMCSheetInfo.Tables[i].RowsValues.Count; j++)
				{
					string text = sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Text;
					string value = "";
					HTKKWinDictionaryDiaDanh.dicQuanHuyen.TryGetValue(text, out value);
					sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Text = value;
					string text2 = sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num19 - 1].Text;
					string value2 = "";
					HTKKWinDictionaryDiaDanh.dicPhuongXa.TryGetValue(text2, out value2);
					sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num20 - 1].Text = value2;
					string text3 = sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num13 - 1].Text;
					if (text3 == "02")
					{
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Locked = true;
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Locked = true;
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num20 - 1].Locked = true;
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Value = "Khác";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Value = "Khác";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num20 - 1].Value = "Khác";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num15 - 1].Value = "";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num17 - 1].Value = "";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num19 - 1].Value = "";
					}
					else
					{
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num16 - 1].Locked = false;
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num18 - 1].Locked = false;
						sheetView.Cells[hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + j - 1, num20 - 1].Locked = false;
					}
				}
			}
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public int validMstNPT(HTKKCMCRowValue rowCurrent, List<HTKKCMCRowValue> listMstDuplicate)
	{
		int result = 0;
		try
		{
			List<HTKKCMCRowValue> list = listMstDuplicate.FindAll((HTKKCMCRowValue p) => p.RowValue[20].Equals(rowCurrent.RowValue[20]) && p.RowValue[21].Equals(rowCurrent.RowValue[21]));
			if (list.Count > 1)
			{
				return 2;
			}
			string[] array = rowCurrent.RowValue.ToArray();
			string text = array[20];
			string text2 = array[21];
			if (HTKKFormater.isNullString(text) || HTKKFormater.isNullString(text2))
			{
				return 0;
			}
			int num = ConvertHelpers.ConvertToInt(text.Split(new char[1] { '/' })[0]);
			int num2 = ConvertHelpers.ConvertToInt(text2.Split(new char[1] { '/' })[0]);
			int num3 = ConvertHelpers.ConvertToInt(text.Split(new char[1] { '/' })[1]);
			int num4 = ConvertHelpers.ConvertToInt(text2.Split(new char[1] { '/' })[1]);
			DateTime dateTime = DateTime.ParseExact("01/" + text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime dateTime2 = DateTime.ParseExact("01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture).AddMonths(1).AddDays(-1.0);
			for (int num5 = 0; num5 < listMstDuplicate.Count; num5++)
			{
				string[] array2 = listMstDuplicate[num5].RowValue.ToArray();
				string text3 = array2[20];
				string text4 = array2[21];
				if (!HTKKFormater.isNullString(text3) && !HTKKFormater.isNullString(text4) && (!(text == text3) || !(text2 == text4)))
				{
					if (text == text3 && text2 == text4)
					{
						result = 1;
						break;
					}
					int num6 = ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[0]);
					int num7 = ConvertHelpers.ConvertToInt(text4.Split(new char[1] { '/' })[0]);
					int num8 = ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[1]);
					int num9 = ConvertHelpers.ConvertToInt(text4.Split(new char[1] { '/' })[1]);
					if (num3 == num8 && num4 == num9)
					{
						result = 2;
						break;
					}
					DateTime dateTime3 = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime4 = DateTime.ParseExact("01/" + text4, "dd/MM/yyyy", CultureInfo.CurrentCulture).AddMonths(1).AddDays(-1.0);
					if ((dateTime >= dateTime3 && dateTime <= dateTime4) || (dateTime3 >= dateTime && dateTime3 <= dateTime2))
					{
						result = 3;
						break;
					}
					if (dateTime4.Year == dateTime.Year)
					{
						result = 3;
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}
}
