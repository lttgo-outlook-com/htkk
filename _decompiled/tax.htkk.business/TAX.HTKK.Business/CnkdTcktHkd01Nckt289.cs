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

public class CnkdTcktHkd01Nckt289
{
	private int[] cell_hoptackd = GeneralFunctions.ArrayCellId("E_10");

	private int[] cell_chothuebds = GeneralFunctions.ArrayCellId("E_12");

	private int[] cell_xaphuongPl2 = GeneralFunctions.ArrayCellId("F_30");

	private int[] cell_maXaphuongPl2 = GeneralFunctions.ArrayCellId("W_30");

	private int[] cell_tinhPl2 = GeneralFunctions.ArrayCellId("E_32");

	private int[] cell_maTinhPl2 = GeneralFunctions.ArrayCellId("W_32");

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 8 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_01_1_CNKD_TCKT_HKD.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 8 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_01_2_CNKD_TCKT_HKD.xls");
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

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int index = 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			SheetView sheetView = winFpSpread.Sheets[index];
			int num = HTKKFile.ColumnLettersToIndex("L") - 1;
			int column = HTKKFile.ColumnLettersToIndex("N") - 1;
			if (iCol - 1 == num)
			{
				sheetView.Cells[(int)iRow - 1, column].Locked = sheetView.Cells[(int)iRow - 1, num].Value.ToIntOrDefault() > 500000000;
				sheetView.Cells[(int)iRow - 1, column].Value = sheetView.Cells[(int)iRow - 1, num].Value.ToIntOrDefault() <= 500000000;
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
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			bool flag = sheetView.Cells[cell_hoptackd[1], cell_hoptackd[0]].Value.ToBoolOrDefault();
			bool flag2 = sheetView.Cells[cell_chothuebds[1], cell_chothuebds[0]].Value.ToBoolOrDefault();
			HTKKFile.RemoveValidateBackGround1(sheetView, cell_hoptackd[1], cell_hoptackd[0]);
			HTKKFile.RemoveValidateBackGround1(sheetView, cell_chothuebds[1], cell_chothuebds[0]);
			if ((flag && flag2) || (!flag && !flag2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell_chothuebds[1], cell_chothuebds[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
				HTKKFile.SetNodeBackGroundPink(sheetView, cell_hoptackd[1], cell_hoptackd[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
				value = 2;
			}
			else if (flag && ToKhaiInfo.SheetInfos[1] == null)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell_hoptackd[1], cell_hoptackd[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
				value = 2;
			}
			else if (flag2 && ToKhaiInfo.SheetInfos[2] == null)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell_chothuebds[1], cell_chothuebds[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
				value = 2;
			}
			if (PeriodInfo.ToKhaiPhatSinh > 0 && ToKhaiInfo.SheetInfos[2] == null)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell_chothuebds[1], cell_chothuebds[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
				value = 2;
			}
			if (ToKhaiInfo.SheetInfos[1] != null && ToKhaiInfo.SheetInfos[2] != null)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell_chothuebds[1], cell_chothuebds[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
				value = 2;
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
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
			int num = 1;
			if (ToKhaiInfo.SheetInfos[num] == null)
			{
				return;
			}
			SheetView sView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			string message = HTKKFile.GetMessage("0047");
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30"));
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int colIndex = HTKKFile.ColumnLettersToIndex("E") - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[num].Tables.First((HTKKCMCTableInfo t) => t.TablePath == "PLuc_HKD_CNKD/CTietBKeHKD_CNKD");
			List<HTKKCMCRowValue> rows = hTKKCMCTableInfo.RowsValues;
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + num3 - 1, colIndex);
			}
			int i;
			for (i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				if (string.IsNullOrEmpty(rows[i].RowValue[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colIndex, message3);
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThueCCCD(rows[i].RowValue[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colIndex, message);
					value = 2;
				}
				if (hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue r) => r.RowValue[1] == rows[i].RowValue[1]) > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colIndex, message2);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
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
			int num = 2;
			if (ToKhaiInfo.SheetInfos[num] == null)
			{
				return;
			}
			SheetView sView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			string message = HTKKFile.GetMessage("0047");
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_23"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30"));
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int colIndex = HTKKFile.ColumnLettersToIndex("G") - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[num].Tables.First((HTKKCMCTableInfo t) => t.TablePath == "PLuc_CANHAN_TTS/CTietBKeCNTTS");
			List<HTKKCMCRowValue> rows = hTKKCMCTableInfo.RowsValues;
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
			{
				HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + num3 - 1, colIndex);
			}
			int i;
			for (i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				if (string.IsNullOrEmpty(rows[i].RowValue[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colIndex, message3);
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThueCCCD(rows[i].RowValue[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colIndex, message);
					value = 2;
				}
				if (hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue r) => r.RowValue[1] == rows[i].RowValue[1]) > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, colIndex, message2);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			sView.Cells[cell_hoptackd[1], cell_hoptackd[0]].Value = ToKhaiInfo.SheetInfos[1] != null;
			sView.Cells[cell_chothuebds[1], cell_chothuebds[0]].Value = ToKhaiInfo.SheetInfos[2] != null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			sView.Cells[cell_xaphuongPl2[1], cell_xaphuongPl2[0]].Value = PeriodInfo.TenXa;
			sView.Cells[cell_maXaphuongPl2[1], cell_maXaphuongPl2[0]].Value = PeriodInfo.MaXa;
			sView.Cells[cell_tinhPl2[1], cell_tinhPl2[0]].Value = PeriodInfo.TenTinh;
			sView.Cells[cell_maTinhPl2[1], cell_maTinhPl2[0]].Value = PeriodInfo.MaTinh;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
