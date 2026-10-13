using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BlpBc01Nckt200
{
	private Dictionary<string, string[]> dicDMBienLaiPhi = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMBienLaiPhiTmp = new Dictionary<string, string>();

	private string cellId_J32 = "J_32";

	private string cellId_J30 = "J_30";

	private int[] CellId_J32 => GeneralFunctions.ArrayCellId(cellId_J32);

	private int[] CellId_J30 => GeneralFunctions.ArrayCellId(cellId_J30);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			sView.Cells[GeneralFunctions.ArrayCellId("H_17")[1], GeneralFunctions.ArrayCellId("H_17")[0]].Value = ((PeriodInfo.KieuKy == "K") ? 1 : 0);
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
			if (idxRow == 4 && idxCol == 1)
			{
				ShowBangKe("Bangke_01BCAC_BLP.xls");
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

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				int num3 = HTKKFile.ColumnLettersToIndex("L") - 1;
				int num4 = HTKKFile.ColumnLettersToIndex("M") - 1;
				if (idxRow >= num && idxRow <= num2 && (idxCol == num3 || idxCol == num4))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("M");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("K");
			int num5 = HTKKFile.ColumnLettersToIndex("Q");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num6 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num7 = num6 + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "";
			if (iRow - 1 < num6 || iRow - 1 > num7)
			{
				return;
			}
			if (iCol == num || iCol == num2)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				string text2 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, num2 - 1].Text;
				int num8 = ConvertHelpers.ConvertToInt(text2);
				int num9 = ConvertHelpers.ConvertToInt(text3);
				if (!HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
				{
					if (num8 == 0 && num9 == 0)
					{
						sheetView.Cells[(int)iRow - 1, num5 - 1].Value = "0";
					}
					else if (num8 == 0 && num9 != 0)
					{
						sheetView.Cells[(int)iRow - 1, num5 - 1].Value = (num9 - num8).ToString();
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num5 - 1].Value = (num9 - num8 + 1).ToString();
					}
				}
				else if (HTKKFormater.isNullString(text2) && HTKKFormater.isNullString(text3))
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = "";
				}
			}
			if (iCol == num4)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieuBLP(text);
			}
			if (iCol == num3)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = text.ToUpper();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LeaveCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, long iNewRow, long iNewCol)
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

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("H");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			if (iCol == num)
			{
				string[] array = dicDMBienLaiPhi[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length == 3)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[2];
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[1];
				}
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
			}
			return true;
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
			int num = 0;
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[11, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[13, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			List<int> list = new List<int>();
			int num2 = HTKKFile.ColumnLettersToIndex("H");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("D");
			int num5 = HTKKFile.ColumnLettersToIndex("E");
			int num6 = HTKKFile.ColumnLettersToIndex("F");
			int num7 = HTKKFile.ColumnLettersToIndex("G");
			int num8 = HTKKFile.ColumnLettersToIndex("I");
			int num9 = HTKKFile.ColumnLettersToIndex("J");
			int num10 = HTKKFile.ColumnLettersToIndex("K");
			int num11 = HTKKFile.ColumnLettersToIndex("L");
			int num12 = HTKKFile.ColumnLettersToIndex("M");
			int num13 = HTKKFile.ColumnLettersToIndex("Q");
			bool flag = true;
			int num14 = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text13 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string text14 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num7 - 1].Text;
					string text19 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Text;
					string text20 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Text;
					string text21 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Text;
					string text22 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Text;
					string text23 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Text;
					string text24 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num13 - 1].Text;
					if (HTKKFormater.isNullString(text13) && HTKKFormater.isNullString(text14) && HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(text16) && HTKKFormater.isNullString(text17) && HTKKFormater.isNullString(text18) && HTKKFormater.isNullString(text19) && HTKKFormater.isNullString(text20) && HTKKFormater.isNullString(text21) && HTKKFormater.isNullString(text22) && HTKKFormater.isNullString(text23) && ConvertHelpers.ConvertToDouble1(text24) == 0.0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
						HTKKFile.RemoveValidate(sheetView, CellId_J32[1], CellId_J32[0]);
						HTKKFile.RemoveValidate(sheetView, CellId_J30[1], CellId_J30[0]);
						continue;
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
					string[] array = GeneralFunctions.checkTrung(winFpSpread, num14, "L", "M", "B", "J", "I", "K", "C", isMST: true, 1);
					string text25 = "";
					flag = false;
					if (HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text20) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text22) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text23) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text17))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
						list.Add(0);
					}
					if (!HTKKFormater.isNullString(text18))
					{
						if (!ValidateBase.ValidateFullDate(text18, "dd/MM/yyyy"))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text11);
							list.Add(1);
						}
						else
						{
							DateTime t = ConvertHelpers.ConvertToDateTime(text18, "dd/MM/yyyy");
							DateTime t2 = ConvertHelpers.ConvertToDateTime(PeriodInfo.DenNgay, "dd/MM/yyyy");
							DateTime t3 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
							if (DateTime.Compare(t, t2) > 0 || DateTime.Compare(t, t3) < 0)
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text7);
								list.Add(0);
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
								list.Add(0);
							}
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text5);
						list.Add(1);
					}
					if (HTKKFormater.isNullString(text19) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text5);
						list.Add(1);
					}
					else
					{
						text25 = GeneralFunctions.CheckMauSoBLP(text19, "0", text13);
						if (text25 == "1" && !HTKKFormater.isNullString(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text2);
							list.Add(1);
						}
						else if (text25 == "2" && !HTKKFormater.isNullString(text13))
						{
							if (text21.Length > 0)
							{
								if (text21.Substring(text21.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienBLPDT(text19, text13, text21).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text12);
										list.Add(1);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text);
									list.Add(1);
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text);
								list.Add(1);
							}
						}
						else if (text21.Length > 0 && text21.Substring(text21.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienBLPDT(text19, text13, text21).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text12);
							list.Add(1);
						}
					}
					if (HTKKFormater.isNullString(text22) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
					}
					if (HTKKFormater.isNullString(text23) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text5);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
					}
					if (!HTKKFormater.isNullString(text22) && !HTKKFormater.isNullString(text23))
					{
						if (ConvertHelpers.ConvertToInt(text22) > ConvertHelpers.ConvertToInt(text23))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text4);
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text4);
							list.Add(1);
						}
						else if (array != null && array.Length > 1)
						{
							for (int k = 0; k < array.Length; k++)
							{
								if (dataLocationTemplateRowStart + j - 1 == ConvertHelpers.ConvertToInt(array[k]))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text8);
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text8);
									list.Add(1);
									break;
								}
							}
						}
					}
					if (text22.ToIntOrDefault(-1) == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
						num = 2;
					}
					if (text23.ToIntOrDefault(-1) == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
						num = 2;
					}
					if (HTKKFormater.isNullString(text21) && !HTKKFormater.isNullString(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text5);
						list.Add(1);
					}
					else
					{
						text21 = text21.ToUpper();
						sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Text = text21.ToUpper();
						text25 = GeneralFunctions.CheckSoBLP(text21, text13, DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture) >= new DateTime(2017, 1, 1));
						if ((text25 == "1" || text25 == "2") && !HTKKFormater.isNullString(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text3);
							list.Add(1);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
							list.Add(0);
						}
					}
					if (HTKKFormater.isNullString(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text5);
						list.Add(1);
					}
					else if (!ValidateBase.CheckMaSoThue(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text6);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						list.Add(0);
					}
					if (array.Length <= 1)
					{
						continue;
					}
					for (int l = 0; l < array.Length; l++)
					{
						if (num14 + j == Convert.ToInt32(array[l]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num14 + j, num11 - 1, message);
							HTKKFile.SetNodeBackGroundPink(sheetView, num14 + j, num12 - 1, message);
							list.Add(1);
							break;
						}
					}
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[0];
			string text26 = sheetView.Cells[CellId_J30[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J30[0]].Text;
			string text27 = sheetView.Cells[CellId_J32[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J32[0]].Text;
			if (HTKKFormater.isNullString(text26))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_J30[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J30[0], text5);
				list.Add(1);
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_J30[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J30[0]);
				list.Add(0);
			}
			if (HTKKFormater.isNullString(text27))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_J32[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J32[0], text5);
				list.Add(1);
			}
			else if (!ValidateBase.ValidateFullDate(text27, "dd/MM/yyyy"))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_J32[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J32[0], text11);
				list.Add(1);
			}
			else
			{
				DateTime t4 = ConvertHelpers.ConvertToDateTime(text27, "dd/MM/yyyy");
				DateTime t5 = ConvertHelpers.ConvertToDateTime(DateTime.Now.ToString("dd/MM/yyyy"), "dd/MM/yyyy");
				DateTime t3 = ConvertHelpers.ConvertToDateTime(PeriodInfo.TuNgay, "dd/MM/yyyy");
				DateTime t2 = ConvertHelpers.ConvertToDateTime(PeriodInfo.DenNgay, "dd/MM/yyyy");
				if (DateTime.Compare(t4, t5) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_J32[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J32[0], text10);
					list.Add(1);
				}
				else if (DateTime.Compare(t4, t3) < 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_J32[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J32[0], text9);
					list.Add(1);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_J32[1] + hTKKCMCTableInfo2.RowsValues.Count - 1, CellId_J32[0]);
					list.Add(0);
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("H");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num3 - 1].Text;
				if (text.Length > 5)
				{
					sheetView.Cells[num + i, num2 - 1].Text = text.Substring(0, 5);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMBienLaiPhi(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMBienLaiPhiTmp.Count > 0)
			{
				return dicDMBienLaiPhiTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[1];
				if (!dictionary2.ContainsKey(key))
				{
					dictionary2.Add(key, value2);
				}
				if (!dictionary.ContainsKey(key))
				{
					dictionary.Add(key, array);
				}
			}
			dicDMBienLaiPhi = dictionary;
			dicDMBienLaiPhiTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}
}
