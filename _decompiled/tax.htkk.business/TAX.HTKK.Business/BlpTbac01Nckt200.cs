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

public class BlpTbac01Nckt200
{
	private Dictionary<string, string[]> dicDMBienLaiPhi = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMBienLaiPhiTmp = new Dictionary<string, string>();

	private string pathCQT = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/tenCQTTNhanTBao";

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("R") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string formula = sView.Cells[25, HTKKFile.ColumnLettersToIndex("AA") - 1].Formula;
			int num = HTKKFile.ColumnLettersToIndex("AA");
			int num2 = ConvertHelpers.ConvertToInt(formula.Replace("AA", ""));
			if (HTKKFormater.isNullString(sView.Cells[num2 - 1, num - 1].Text))
			{
				sView.Cells[num2 - 1, num - 1].Text = PeriodInfo.kyKKhai;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCQT));
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = TTinChung.tenCQTNoiNop;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 2 && idxCol == 1)
			{
				ShowBangKe("Bangke_01TBAC_BLP.xls");
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
				if (idxRow >= num && idxRow <= num2 && (idxCol == idxCol_TuSo || idxCol == idxCol_DenSo))
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
			int num = HTKKFile.ColumnLettersToIndex("P");
			int num2 = HTKKFile.ColumnLettersToIndex("R");
			int num3 = HTKKFile.ColumnLettersToIndex("L");
			int num4 = HTKKFile.ColumnLettersToIndex("J");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num5 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num6 = num5 + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text = "";
			if (iRow - 1 >= num5 && iRow - 1 <= num6)
			{
				if (iCol == num || iCol == num2)
				{
					text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				}
				if (iCol == num3)
				{
					text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieuBLP(text);
				}
				if (iCol == num4)
				{
					text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = text.ToUpper();
				}
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
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			if (iCol == num)
			{
				string[] array = dicDMBienLaiPhi[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length == 3)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[2];
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = array[1];
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
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[34, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[0];
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			List<int> list = new List<int>();
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("L");
			int num5 = HTKKFile.ColumnLettersToIndex("N");
			int num6 = HTKKFile.ColumnLettersToIndex("P");
			int num7 = HTKKFile.ColumnLettersToIndex("R");
			int num8 = HTKKFile.ColumnLettersToIndex("T");
			int num9 = HTKKFile.ColumnLettersToIndex("W");
			int num10 = HTKKFile.ColumnLettersToIndex("AA");
			int num11 = HTKKFile.ColumnLettersToIndex("AD");
			int num12 = HTKKFile.ColumnLettersToIndex("AE");
			bool flag = true;
			int num13 = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart - 1;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num13, "P", "R", "B", "D", "J", "L", "", isMST: false, 1);
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
					string text13 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text;
					string text14 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string text16 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text17 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string text18 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num6 - 1].Text;
					string text19 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num7 - 1].Text;
					string text20 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Text;
					string text21 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num9 - 1].Text;
					string text22 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Text;
					string text23 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Text;
					string text24 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Text;
					if (HTKKFormater.isNullString(text13) && HTKKFormater.isNullString(text14) && HTKKFormater.isNullString(text15) && HTKKFormater.isNullString(text16) && HTKKFormater.isNullString(text18) && HTKKFormater.isNullString(text19) && HTKKFormater.isNullString(text20) && HTKKFormater.isNullString(text21) && HTKKFormater.isNullString(text22) && HTKKFormater.isNullString(text23) && HTKKFormater.isNullString(text24))
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
						continue;
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
					string text25 = "";
					flag = false;
					if (string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num - 1, text8);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num - 1);
						list.Add(0);
					}
					if (string.IsNullOrEmpty(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text8);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						list.Add(0);
					}
					if (string.IsNullOrEmpty(text15))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text8);
						list.Add(1);
					}
					else
					{
						text25 = GeneralFunctions.CheckMauSoBLP(text15, "0", text13);
						if (text25 == "1")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text2);
							list.Add(1);
						}
						else if (text25 == "2")
						{
							if (text16.Length > 0)
							{
								if (text16.Substring(text16.Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienBLPDT(text15, text13, text16).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text12);
										list.Add(1);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
									list.Add(1);
								}
							}
						}
						else if (text16.Length > 0 && text16.Substring(text16.Length - 1).Equals("E") && GeneralFunctions.CheckSoLienBLPDT(text15, text13, text16).Equals("3"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text12);
							list.Add(1);
						}
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text8);
						list.Add(1);
					}
					else
					{
						text25 = GeneralFunctions.CheckSoBLP(text16, text13, DateTime.ParseExact(PeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.InvariantCulture) >= new DateTime(2017, 1, 1));
						if (text25 == "1" || text25 == "2")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1, text3);
							list.Add(1);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4 - 1);
							list.Add(0);
						}
					}
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
					if (HTKKFormater.isNullString(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text8);
						list.Add(1);
					}
					if (HTKKFormater.isNullString(text19))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text8);
						list.Add(1);
					}
					if (ConvertHelpers.ConvertToInt(text18) > ConvertHelpers.ConvertToInt(text19))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text4);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text4);
						list.Add(1);
					}
					if (text18.ToIntOrDefault(-1) == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
						value = 2;
					}
					if (text19.ToIntOrDefault(-1) == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
						value = 2;
					}
					if (HTKKFormater.isNullString(text20))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text8);
					}
					else
					{
						DateTime dateTime = ConvertHelpers.ConvertToDateTime(text20, "dd/MM/yyyy");
						DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(PeriodInfo.kyKKhai, "dd/MM/yyyy");
						if (DateTime.Compare(dateTime, dateTime2) < 0 && text16.Length == 6)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text5);
							list.Add(1);
						}
						else if (DateTime.Compare(dateTime, dateTime2) < 0 && text16.Length == 8)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text6);
							list.Add(1);
						}
						else if ((dateTime - dateTime2).TotalDays < 5.0 && text16.Length == 6)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text5);
							list.Add(1);
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
							list.Add(0);
						}
					}
					if (!string.IsNullOrEmpty(text16))
					{
						if ("P".Equals(text16.Trim().Substring(text16.Trim().Length - 1, 1)))
						{
							if (HTKKFormater.isNullString(text21))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, text10);
								list.Add(1);
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
								list.Add(0);
							}
							if (HTKKFormater.isNullString(text22))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text10);
								list.Add(1);
							}
							else
							{
								DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(text22, "dd/MM/yyyy");
								DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(PeriodInfo.kyKKhai, "dd/MM/yyyy");
								if ((dateTime2 - dateTime3).TotalDays < 0.0)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1, text7);
									list.Add(1);
								}
								else
								{
									HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num10 - 1);
									list.Add(0);
								}
							}
							if (string.IsNullOrEmpty(text23))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1, text10);
								list.Add(1);
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num11 - 1);
								list.Add(0);
							}
							if (string.IsNullOrEmpty(text24))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text10);
								list.Add(1);
							}
							else if (!ValidateBase.CheckMaSoThue(text24))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text9);
								list.Add(1);
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
								list.Add(0);
							}
						}
						else if (!string.IsNullOrEmpty(text24))
						{
							if (!ValidateBase.CheckMaSoThue(text24))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1, text9);
								list.Add(1);
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num12 - 1);
								list.Add(0);
							}
						}
					}
					if (array.Length <= 1)
					{
						continue;
					}
					for (int k = 0; k < array.Length; k++)
					{
						if (num13 + j == Convert.ToInt32(array[k]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num13 + j, num6 - 1, message);
							HTKKFile.SetNodeBackGroundPink(sheetView, num13 + j, num7 - 1, message);
							list.Add(1);
							break;
						}
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					value = 2;
					break;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
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
}
