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

public class TncnDs05Nckt216
{
	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 2 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_05DS.xls");
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

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			HTKKFile.ErrorMessage = "";
			int iwarningColor = ((levelWarning.Count > 0) ? levelWarning[0] : 0);
			checkTKChinh(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, iwarningColor);
			if (!string.IsNullOrEmpty(HTKKFile.ErrorMessage) && ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count > 500)
			{
				winFpSpread.ResumeLayout();
				Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
				foreach (Process process in processesByName)
				{
					process.Kill();
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

	private void checkTKChinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int iwarningColor)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int colIndex = HTKKFile.ColumnLettersToIndex("C") - 1;
			int colIndex2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int colIndex3 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int colIndex4 = HTKKFile.ColumnLettersToIndex("F") - 1;
			int colIndex5 = HTKKFile.ColumnLettersToIndex("G") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("H") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int colIndex6 = HTKKFile.ColumnLettersToIndex("J") - 1;
			int colIndex7 = HTKKFile.ColumnLettersToIndex("K") - 1;
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			string text = PeriodInfo.TuThang.Split(new char[1] { '/' })[0];
			string text2 = PeriodInfo.DenThang.Split(new char[1] { '/' })[0];
			int num4 = (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				List<string> arrARowValue = hTKKCMCTableInfo.RowsValues[i].RowValue;
				if (i == 0 && !GeneralFunctions.HaveToValidateRow(rowsInfo, arrARowValue.ToArray()))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
					continue;
				}
				string value = arrARowValue[0];
				if (string.IsNullOrEmpty(value))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					iwarningColor = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
				}
				HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
				if (!string.IsNullOrEmpty(arrARowValue[1].Trim()))
				{
					if (arrARowValue[1].Trim().Length == 10)
					{
						if (!ValidateBase.CheckMaSoThue(arrARowValue[1].Trim()))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
							iwarningColor = 2;
						}
						else
						{
							int num5 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[1].Trim()) && !string.IsNullOrEmpty(p.RowValue[1]));
							if (num5 > 1)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
								iwarningColor = 2;
							}
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
						iwarningColor = 2;
					}
				}
				HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
				if (!string.IsNullOrEmpty(arrARowValue[2]))
				{
					if (arrARowValue[2].Contains(" "))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						iwarningColor = 2;
					}
					else
					{
						int num5 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(arrARowValue[2]) && !string.IsNullOrEmpty(p.RowValue[2]));
						if (num5 > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
							iwarningColor = 2;
						}
					}
				}
				if (string.IsNullOrEmpty(arrARowValue[1].Trim()) && string.IsNullOrEmpty(arrARowValue[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
					iwarningColor = 2;
				}
				double num6 = ConvertHelpers.ConvertToDouble(arrARowValue[3]);
				if (num6 > 0.0)
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					iwarningColor = 2;
				}
				double num7 = ConvertHelpers.ConvertToDouble(arrARowValue[4]);
				if (num7 % 100000.0 == 0.0)
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex5);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex5, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28")));
					iwarningColor = 2;
				}
				string[] array = PeriodInfo.TuThang.Split(new char[1] { '/' });
				string[] array2 = PeriodInfo.DenThang.Split(new char[1] { '/' });
				double num8 = num4 * 1000000;
				double num9 = ConvertHelpers.ConvertToDouble(arrARowValue[7]);
				if (num9 > num8)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex6, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
					iwarningColor = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
				}
				double num10 = ConvertHelpers.ConvertToDouble(arrARowValue[8]);
				if (num10 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex7, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					iwarningColor = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
				}
			}
			switch (iwarningColor)
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
}
