using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tsmh01TT41Nckt300
{
	private string fomulaColW = "R###+AD###";

	private string MA_COLOR_LOCKED = "#e8f5ff";

	private string MA_COLOR_UNLOCK = "#FFFFFF";

	private string MA_CCCD = "2080";

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 11 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_01_1_TSMH.xls");
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

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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
		try
		{
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (!string.IsNullOrEmpty(HTKKFile.ErrorMessage) && (ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || ToKhaiInfo.SheetInfos[1].Tables[1].RowsValues.Count > 500 || ToKhaiInfo.SheetInfos[1].Tables[2].RowsValues.Count > 500))
			{
				winFpSpread.ResumeLayout();
				Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
				foreach (Process process in processesByName)
				{
					process.Kill();
				}
			}
			return result;
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
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("G") - 1;
			int column = HTKKFile.ColumnLettersToIndex("U") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("T") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("P") - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			string message = HTKKFile.GetMessage("0047");
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < tables[i].RowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5);
				}
			}
			int num6 = -1;
			int[] array = GeneralFunctions.ArrayCellId("AD_4");
			string text = sheetView.Cells[array[1], array[0]].Text.Trim();
			for (int i = 0; i < tables.Count; i++)
			{
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				int num7 = tables[i].Location.DataLocationTemplateRowStart - 1;
				for (int k = 0; k < tables[i].RowsValues.Count; k++)
				{
					List<string> rowValue = tables[i].RowsValues[k].RowValue;
					string text2 = sheetView.Cells[num7 + k, array[0]].Text.Trim();
					if (text2 == "1")
					{
						num6 = i;
						break;
					}
				}
				if (num6 >= 0)
				{
					break;
				}
			}
			if (num6 == -1)
			{
				for (int i = 0; i < tables.Count; i++)
				{
					int num8 = tables[i].Location.DataLocationTemplateRowStart - 1;
					for (int k = 0; k < tables[i].RowsValues.Count; k++)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num8 + k, num, message2, k + 1);
						value = 2;
					}
				}
			}
			else
			{
				int num7 = 0;
				StringBuilder stringBuilder = new StringBuilder();
				for (int i = 0; i < tables.Count; i++)
				{
					num7 = tables[i].Location.DataLocationTemplateRowStart - 1;
					HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
					List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
					List<string> list = new List<string>();
					List<string> list2 = new List<string>();
					for (int k = 0; k < rowsValues.Count; k++)
					{
						List<string> rowValue = tables[i].RowsValues[k].RowValue;
						string text2 = sheetView.Cells[num7 + k, array[0]].Text.Trim();
						if (!(text2 == "1"))
						{
							continue;
						}
						string text3 = sheetView.Cells[num7 + k, num].Text;
						string text4 = sheetView.Cells[num7 + k, num2].Text.Trim();
						string text5 = sheetView.Cells[num7 + k, column].Text.Trim();
						string text6 = sheetView.Cells[num7 + k, num3].Text.Trim();
						string value2 = sheetView.Cells[num7 + k, num4].Text.Trim();
						string text7 = sheetView.Cells[num7 + k, num5].Text.Trim();
						if (string.IsNullOrEmpty(text3))
						{
							stringBuilder.Append("Lỗi Họ tên tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), k + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text4) || !string.IsNullOrEmpty(text6))
						{
							if (!string.IsNullOrEmpty(text4) && !string.IsNullOrEmpty(text6))
							{
								stringBuilder.Append("Lỗi nhập cả [08] và [09] tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")), k + 1);
								stringBuilder.Append("Lỗi nhập cả [08] và [09] tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")), k + 1);
								value = 2;
							}
						}
						else
						{
							stringBuilder.Append("Lỗi thiếu [08] hoặc [09] tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), k + 1);
							stringBuilder.Append("Lỗi thiếu [08] hoặc [09] tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), k + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text4) && !ValidateBase.CheckMaSoThue(text4))
						{
							stringBuilder.Append("Lỗi sai cấu trúc MST tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num2, message);
							value = 2;
						}
						if (text5.Equals(MA_CCCD) && !string.IsNullOrEmpty(text6) && !GeneralFunctions.checkCCCD(text6))
						{
							stringBuilder.Append("Lỗi sai cấu trúc CCCD tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")), k + 1);
							value = 2;
						}
						if (i == 2 && string.IsNullOrEmpty(value2))
						{
							stringBuilder.Append("Lỗi thiếu thông tin cá nhân kiểm soát tại dòng số " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")), k + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text7) && !GeneralFunctions.IsValidDateTime(text7))
						{
							stringBuilder.Append("Lỗi nhập thời điểm không hợp lệ " + (k + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num7 + k, num5, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")), k + 1);
							value = 2;
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
