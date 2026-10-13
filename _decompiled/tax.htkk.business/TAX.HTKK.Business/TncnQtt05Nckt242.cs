using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnQtt05Nckt242
{
	private int odrColC = HTKKFile.ColumnLettersToIndex("C");

	private int odrColD = HTKKFile.ColumnLettersToIndex("D");

	private int odrColQ = HTKKFile.ColumnLettersToIndex("Q");

	private int odrColF = HTKKFile.ColumnLettersToIndex("F");

	private int odrColG = HTKKFile.ColumnLettersToIndex("G");

	private int odrColH = HTKKFile.ColumnLettersToIndex("H");

	private int odrColI = HTKKFile.ColumnLettersToIndex("I");

	private int odrColJ = HTKKFile.ColumnLettersToIndex("J");

	private int odrColK = HTKKFile.ColumnLettersToIndex("K");

	private int odrColL = HTKKFile.ColumnLettersToIndex("L");

	private int odrColM = HTKKFile.ColumnLettersToIndex("M");

	private int odrColN = HTKKFile.ColumnLettersToIndex("N");

	private int odrColO = HTKKFile.ColumnLettersToIndex("O");

	private int odrColR = HTKKFile.ColumnLettersToIndex("R");

	private string CT16formula = "ROUND(H###*0.1,0)";

	private string CT17formula = "IF(OR(F###=\"1\",F###=\"x\",F###=\"X\",F###=TRUE),ROUND(I###*0.2*0.5,0),ROUND(I###*0.1*0.5,0))";

	private string CT22PL051Formula = "ROUND(IF(OR(G###=\"1\",G###=\"x\",G###=\"X\",G###=TRUE),IF($G$5=\"2009\",IF(BD###-R###>0,BD###-R###,0),IF(AND($G$5=\"2012\",AP###<=5000000),IF((AQ###-R###)/2>0,(AQ###-R###)/2,0),IF(AQ###-R###>0,AQ###-R###,0))),0),0)";

	private string CT23PL051Formula = "IF(OR(G###=\"1\",G###=\"x\",G###=\"X\",G###=TRUE),IF(T###-Q###<0,ABS(T###-Q###),0),0)";

	private string CT24PL051Formula = "IF(OR(G###=\"1\",G###=\"x\",G###=\"X\",G###=TRUE),IF(T###-Q###>=0,T###-Q###,0),0)";

	private bool IsQTKTN = false;

	private bool isGiamDan = true;

	private string cellId_chkQTKTN = "D_4";

	private string cellId_lyDo = "D_6";

	private int[] CellId_chkQTKTN => GeneralFunctions.ArrayCellId(cellId_chkQTKTN);

	private int[] CellId_lyDo => GeneralFunctions.ArrayCellId(cellId_lyDo);

	private int[] CellID_Ct21 => GeneralFunctions.ArrayCellId("I_36");

	private int[] CellID_Ct22 => GeneralFunctions.ArrayCellId("I_37");

	private int[] CellID_Ct21_PL01 => GeneralFunctions.ArrayCellId("Q_36");

	private int[] CellID_Ct22_PL01 => GeneralFunctions.ArrayCellId("Q_37");

	private int[] CellID_Ct24 => GeneralFunctions.ArrayCellId("I_39");

	private int[] CellID_Ct24_PL => GeneralFunctions.ArrayCellId("Q_39");

	private int[] CellID_Ct25 => GeneralFunctions.ArrayCellId("I_40");

	private int[] CellID_Ct26 => GeneralFunctions.ArrayCellId("I_41");

	private int[] CellID_Ct26_PL => GeneralFunctions.ArrayCellId("Q_41");

	private int[] CellID_Ct27 => GeneralFunctions.ArrayCellId("I_42");

	private int[] CellID_Ct28 => GeneralFunctions.ArrayCellId("I_43");

	private int[] CellID_Ct29 => GeneralFunctions.ArrayCellId("I_44");

	private int[] CellID_Ct29_PL => GeneralFunctions.ArrayCellId("Q_44");

	private int[] CellID_Ct30 => GeneralFunctions.ArrayCellId("I_45");

	private int[] CellID_Ct30_PL => GeneralFunctions.ArrayCellId("Q_45");

	private int[] CellID_Ct31 => GeneralFunctions.ArrayCellId("I_46");

	private int[] CellID_Ct32 => GeneralFunctions.ArrayCellId("I_47");

	private int[] CellID_Ct32_PL => GeneralFunctions.ArrayCellId("Q_47");

	private int[] CellID_Ct33 => GeneralFunctions.ArrayCellId("I_48");

	private int[] CellID_Ct33_PL => GeneralFunctions.ArrayCellId("Q_48");

	private int[] CellID_Ct35 => GeneralFunctions.ArrayCellId("I_50");

	private int[] CellID_Ct35_PL => GeneralFunctions.ArrayCellId("Q_50");

	private int[] CellID_Ct36 => GeneralFunctions.ArrayCellId("I_51");

	private int[] CellID_Ct36_PL => GeneralFunctions.ArrayCellId("Q_51");

	private int[] CellID_Ct37 => GeneralFunctions.ArrayCellId("I_52");

	private int[] CellID_Ct38 => GeneralFunctions.ArrayCellId("I_53");

	private int[] CellID_Ct38_PL => GeneralFunctions.ArrayCellId("Q_53");

	private int[] CellID_Ct39 => GeneralFunctions.ArrayCellId("I_54");

	private int[] CellID_Ct40 => GeneralFunctions.ArrayCellId("I_55");

	private int[] CellID_Ct40_PL => GeneralFunctions.ArrayCellId("Q_55");

	private int[] CellID_Ct41 => GeneralFunctions.ArrayCellId("I_56");

	private int[] CellID_Ct41_PL => GeneralFunctions.ArrayCellId("Q_56");

	private int[] CellID_Ct42 => GeneralFunctions.ArrayCellId("I_60");

	private int[] CellID_Ct43 => GeneralFunctions.ArrayCellId("I_61");

	private int[] CellID_Ct44 => GeneralFunctions.ArrayCellId("I_62");

	private int[] CellId_E21 => GeneralFunctions.ArrayCellId("E_21");

	private int[] CellId_E85 => GeneralFunctions.ArrayCellId("E_85");

	public int idxCellTpl_ct25 => 8;

	public int idxCellTpl_ct27 => 10;

	public int idxCellTpl_ct28 => 11;

	public int idxCellTpl_ct29 => 12;

	public int idxCellTpl_ct30 => 13;

	public int idxCellTpl_ct31 => 14;

	public int idxCellTpl_ct32 => 15;

	public int idxCellTpl_ct33 => 16;

	public int idxCellTpl_ct34 => 17;

	public int idxCellTpl_ct35 => 18;

	public int idxCellTpl_ct36 => 19;

	public int idxCellTpl_ct37 => 20;

	public int idxCellTpl_ct38 => 21;

	public int idxCellTpl_ct39 => 22;

	public int idxCellTpl_ct40 => 23;

	public int idxCellTpl_ct41 => 24;

	public int idxCellTpl_ct42 => 25;

	public int idxCellTpl_ct43 => 26;

	public int idxCellTpl_ct44 => 27;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string text = PeriodInfo.TuThang.Split(new char[1] { '/' })[0];
			string text2 = PeriodInfo.DenThang.Split(new char[1] { '/' })[0];
			int num = (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
			sView.Cells[3, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = num;
			IsQTKTN = num < 12;
			int row = 43;
			int num2 = 44;
			int row2 = 46;
			int num3 = 47;
			int row3 = 49;
			int num4 = 50;
			GeneralFunctions.ChangeLockedCell(sView, CellId_chkQTKTN[1], CellId_chkQTKTN[0], !IsQTKTN);
			if (sView.Cells[row, odrColI - 1].Locked)
			{
				sView.Cells[row, odrColI - 1].Value = double.Parse(hTKKCMCSheetInfo.StaticCells[idxCellTpl_ct29].Value);
			}
			GeneralFunctions.ChangeLockedCell(sView, num2, odrColI - 1, !IsQTKTN);
			sView.Cells[num2, odrColI - 1].Value = double.Parse(hTKKCMCSheetInfo.StaticCells[idxCellTpl_ct30].Value);
			if (sView.Cells[row2, odrColI - 1].Locked)
			{
				sView.Cells[row2, odrColI - 1].Value = double.Parse(hTKKCMCSheetInfo.StaticCells[idxCellTpl_ct32].Value);
			}
			GeneralFunctions.ChangeLockedCell(sView, num3, odrColI - 1, !IsQTKTN);
			sView.Cells[num3, odrColI - 1].Value = double.Parse(hTKKCMCSheetInfo.StaticCells[idxCellTpl_ct33].Value);
			if (sView.Cells[row3, odrColI - 1].Locked)
			{
				sView.Cells[row3, odrColI - 1].Value = double.Parse(hTKKCMCSheetInfo.StaticCells[idxCellTpl_ct35].Value);
			}
			GeneralFunctions.ChangeLockedCell(sView, num4, odrColI - 1, !IsQTKTN);
			sView.Cells[num4, odrColI - 1].Value = double.Parse(hTKKCMCSheetInfo.StaticCells[idxCellTpl_ct36].Value);
			if ((bool)sView.Cells[CellId_chkQTKTN[1], CellId_chkQTKTN[0]].Value)
			{
				sView.Cells[CellId_lyDo[1], CellId_lyDo[0]].Locked = false;
			}
			else
			{
				sView.Cells[CellId_lyDo[1], CellId_lyDo[0]].Locked = true;
			}
			sView.Cells[CellID_Ct35[1], CellID_Ct35[0]].Locked = false;
			if (!(PeriodInfo.LoaiToKhai == "B"))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[CellId_chkQTKTN[1], CellId_chkQTKTN[0]].Value)
			{
				winFpSpread.Sheets[0].Cells[CellId_lyDo[1], CellId_lyDo[0]].Locked = false;
				return;
			}
			winFpSpread.Sheets[0].Cells[CellId_lyDo[1], CellId_lyDo[0]].ResetValue();
			winFpSpread.Sheets[0].Cells[CellId_lyDo[1], CellId_lyDo[0]].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 2 && idxCol == 3)
			{
				ShowBangKe("Bang_Ke_05_3_QTT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		//IL_03cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_08f5: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (idxRow == 2 && idxCol == 3)
			{
				ShowBangKe("Bang_Ke_05_1_QTT.xls");
			}
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("S");
			int num6 = HTKKFile.ColumnLettersToIndex("B");
			int num7 = 17;
			int num8 = 18;
			int num9 = 20;
			string text = "ROUND(IF(OR(G###=\"1\",G###=\"x\",G###=\"X\",G###=TRUE),IF($G$5=\"2009\",IF(BD###-R###>0,BD###-R###,0),IF(AND($G$5=\"2012\",AP###<=5000000),IF((AQ###-R###)/2>0,(AQ###-R###)/2,0),IF(AQ###-R###>0,AQ###-R###,0))),0),0)";
			string text2 = "IF(OR(G###=\"1\",G###=\"x\",G###=\"X\",G###=TRUE),IF(T###-Q###<0,ABS(T###-Q###),0),0)";
			string text3 = "IF(OR(G###=\"1\",G###=\"x\",G###=\"X\",G###=TRUE),IF(T###-Q###>=0,T###-Q###,0),0)";
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart;
			if (idxRow >= dataLocationTemplateRowStart - 1 && idxCol == num2 - 1)
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[idxRow, idxCol].Value))
				{
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("T") - 1].Formula = text.Replace("###", (idxRow + 1).ToString());
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = text2.Replace("###", (idxRow + 1).ToString());
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = text3.Replace("###", (idxRow + 1).ToString());
				}
				else
				{
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("T") - 1].ResetFormula();
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("V") - 1].ResetFormula();
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("T") - 1].Value = 0;
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("U") - 1].Value = 0;
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
				}
			}
			if (idxRow >= dataLocationTemplateRowStart - 1 && idxCol == num5 - 1)
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[idxRow, idxCol].Value))
				{
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = false;
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = false;
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = true;
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = true;
					sheetView.Cells[idxRow, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
				}
				sheetView.SetActiveCell(idxRow, idxCol);
			}
			if (idxRow == num9 - 1 && idxCol == num2 - 1)
			{
				frmThongTinIn frmThongTinIn = new frmThongTinIn();
				((Form)frmThongTinIn).ShowDialog();
				int options = frmThongTinIn.GetOptions();
				string start = frmThongTinIn.GetStart();
				string end = frmThongTinIn.GetEnd();
				winFpSpread.SuspendLayout();
				sheetView.AutoCalculation = false;
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					switch (options)
					{
					case 1:
					{
						object[,] array = new object[rowsValues.Count, 1];
						for (int j = 0; j < rowsValues.Count; j++)
						{
							array[j, 0] = true;
						}
						sheetView.SetArray(dataLocationTemplateRowStart2 - 1, num2 - 1, array);
						break;
					}
					case 2:
					{
						object[,] array = new object[rowsValues.Count, 1];
						for (int j = 0; j < rowsValues.Count; j++)
						{
							array[j, 0] = false;
						}
						sheetView.SetArray(dataLocationTemplateRowStart2 - 1, num2 - 1, array);
						break;
					}
					case 3:
					{
						int num10 = ConvertHelpers.ConvertToInt(start);
						int num11 = ConvertHelpers.ConvertToInt(end);
						if (num10 < 1)
						{
							num10 = 1;
						}
						if (ConvertHelpers.ConvertToInt(num11) > rowsValues.Count)
						{
							num11 = rowsValues.Count;
						}
						object[,] array = new object[rowsValues.Count, 1];
						for (int j = 0; j < rowsValues.Count; j++)
						{
							if (j >= num10 - 1 && j <= num11 - 1)
							{
								array[j, 0] = true;
							}
							else
							{
								array[j, 0] = false;
							}
						}
						sheetView.SetArray(dataLocationTemplateRowStart2 - 1, num2 - 1, array);
						break;
					}
					}
				}
				winFpSpread.ResumeLayout(performLayout: true);
			}
			int num12 = 0;
			if (idxRow == num8 - 1 && idxCol == num3 - 1)
			{
				sheetView.SortRange(hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[0].CellTemplate.ColIndex - 1, hTKKCMCSheetInfo.Tables[0].RowsValues.Count, HTKKFile.ColumnLettersToIndex("BK") - 1, byRows: true, new SortInfo[1]
				{
					new SortInfo(hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[4].CellTemplate.ColIndex - 1, isGiamDan)
				});
				isGiamDan = !isGiamDan;
			}
			else if (idxRow == num7 - 1 && idxCol == num - 1)
			{
				sheetView.SortRange(hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[0].CellTemplate.ColIndex - 1, hTKKCMCSheetInfo.Tables[0].RowsValues.Count, HTKKFile.ColumnLettersToIndex("BK") - 1, byRows: true, new SortInfo[1]
				{
					new SortInfo(hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[0].CellTemplate.ColIndex - 1, isGiamDan)
				});
				isGiamDan = !isGiamDan;
			}
			else if (idxRow == num7 - 1 && idxCol == num4 - 1)
			{
				sheetView.SortRange(hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[0].CellTemplate.ColIndex - 1, hTKKCMCSheetInfo.Tables[0].RowsValues.Count, HTKKFile.ColumnLettersToIndex("BK") - 1, byRows: true, new SortInfo[1]
				{
					new SortInfo(hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[13].CellTemplate.ColIndex - 1, isGiamDan)
				});
				isGiamDan = !isGiamDan;
			}
			if (idxRow == 14 && idxCol == num6 - 1 && ConvertHelpers.ConvertToBool(sheetView.Cells[idxRow, idxCol].Value))
			{
				frmMessage frmMessage = new frmMessage("0523", "2", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				if (!frmMessage.Yes)
				{
					sheetView.Cells[idxRow, idxCol].Value = "FALSE";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			winFpSpread.Sheets[1].AutoCalculation = true;
			winFpSpread.ResumeLayout(performLayout: true);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 1 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_05_2_QTT.xls");
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("K");
			int num4 = 17;
			int num5 = 18;
			int num6 = 20;
			int num7 = HTKKFile.ColumnLettersToIndex("F");
			string text = "IF(OR(F###=\"1\",F###=\"x\",F###=\"X\",F###=TRUE),ROUND((G###-J###)*0.2- M###,0),ROUND((G###-J###)*0.1 - M###,0))";
			if (idxCol == num7 - 1)
			{
				sheetView2.Cells[idxRow, odrColM - 1].Formula = CT17formula.Replace("###", (idxRow + 1).ToString());
				sheetView2.Cells[idxRow, odrColK - 1].Formula = text.Replace("###", (idxRow + 1).ToString());
				sheetView2.SetActiveCell(idxRow - 1, idxCol);
			}
			sheetView.Cells[CellID_Ct24[1], CellID_Ct24[0]].Value = sheetView.Cells[CellID_Ct24_PL[1], CellID_Ct24_PL[0]].Value;
			if (idxRow == num4 - 1 && idxCol == num - 1)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					for (int j = 0; j < rowsValues.Count; j++)
					{
						for (int k = j + 1; k < rowsValues.Count; k++)
						{
							if (isGiamDan)
							{
								if (string.Compare(sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text, sheetView2.Cells[dataLocationTemplateRowStart + k - 1, num - 1].Text) >= 0)
								{
									continue;
								}
								for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
								{
									HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
									if (HTKKFormater.isNullString(sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Locked)
									{
										sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula = "";
										sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Formula = "";
									}
									string text2 = sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Text;
									sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Text = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Text;
									sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Text = text2;
								}
							}
							else
							{
								if (string.Compare(sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text, sheetView2.Cells[dataLocationTemplateRowStart + k - 1, num - 1].Text) <= 0)
								{
									continue;
								}
								for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
								{
									HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
									if (HTKKFormater.isNullString(sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Locked)
									{
										sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula = "";
										sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Formula = "";
									}
									string text2 = sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Text;
									sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Text = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Text;
									sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Text = text2;
								}
							}
						}
					}
				}
				isGiamDan = !isGiamDan;
				if (idxRow >= 21)
				{
					sheetView2.Cells[idxRow, odrColM - 1].Formula = CT17formula.Replace("###", idxRow.ToString());
					sheetView2.Cells[idxRow, odrColK - 1].Formula = text.Replace("###", idxRow.ToString());
				}
			}
			if (idxRow != num5 - 1 || (idxCol != num2 - 1 && idxCol != num3 - 1))
			{
				return;
			}
			int num8 = idxCol + 1;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					sheetView2.Cells[dataLocationTemplateRowStart + j - 1, odrColK - 1].Formula = "";
					for (int k = j + 1; k < rowsValues.Count; k++)
					{
						if (isGiamDan)
						{
							if (!(ConvertHelpers.ConvertToDouble1(sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Text) > ConvertHelpers.ConvertToDouble1(sheetView2.Cells[dataLocationTemplateRowStart + k - 1, num8 - 1].Text)))
							{
								continue;
							}
							for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
							{
								HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
								if (HTKKFormater.isNullString(sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Locked)
								{
									sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula = "";
									sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Formula = "";
								}
								object value = sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Value;
								sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Value = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Value;
								sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Value = value;
							}
						}
						else
						{
							if (!(ConvertHelpers.ConvertToDouble1(sheetView2.Cells[dataLocationTemplateRowStart + j - 1, num8 - 1].Text) < ConvertHelpers.ConvertToDouble1(sheetView2.Cells[dataLocationTemplateRowStart + k - 1, num8 - 1].Text)))
							{
								continue;
							}
							for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
							{
								HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
								if (HTKKFormater.isNullString(sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Locked)
								{
									sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Formula = "";
									sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Formula = "";
								}
								object value = sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Value;
								sheetView2.Cells[dataLocationTemplateRowStart + k - 1, cellTemplate.ColIndex - 1].Value = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Value;
								sheetView2.Cells[dataLocationTemplateRowStart + j - 1, cellTemplate.ColIndex - 1].Value = value;
							}
						}
					}
				}
			}
			isGiamDan = !isGiamDan;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ClearData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			resetCT22_TK(winFpSpread, ToKhaiInfo);
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
			resetCT22_TK(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			resetCT22_TK(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart;
			SheetView sheetView2 = winFpSpread.Sheets[0];
			winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("I_36")[1], GeneralFunctions.ArrayCellId("I_36")[0]].Value = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1] != "") + ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[0] != "");
			sheetView2.Cells[CellID_Ct22[1], CellID_Ct22[0]].Value = sheetView2.Cells[CellID_Ct22_PL01[1], CellID_Ct22_PL01[0]].Value;
			sheetView2.Cells[CellID_Ct24[1], CellID_Ct24[0]].Value = sheetView2.Cells[CellID_Ct24_PL[1], CellID_Ct24_PL[0]].Value;
			sheetView2.Cells[CellID_Ct26[1], CellID_Ct26[0]].Value = sheetView2.Cells[CellID_Ct26_PL[1], CellID_Ct26_PL[0]].Value;
			sheetView2.Cells[CellID_Ct29[1], CellID_Ct29[0]].Value = sheetView2.Cells[CellID_Ct29_PL[1], CellID_Ct29_PL[0]].Value;
			sheetView2.Cells[CellID_Ct30[1], CellID_Ct30[0]].Value = sheetView2.Cells[CellID_Ct30_PL[1], CellID_Ct30_PL[0]].Value;
			sheetView2.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView2.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
			sheetView2.Cells[CellID_Ct33[1], CellID_Ct33[0]].Value = sheetView2.Cells[CellID_Ct33_PL[1], CellID_Ct33_PL[0]].Value;
			sheetView2.Cells[CellID_Ct35[1], CellID_Ct35[0]].ResetFormula();
			sheetView2.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView2.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
			SheetView sheetView3 = winFpSpread.Sheets[3];
			int dataLocationTemplateRowStart2 = ToKhaiInfo.SheetInfos[3].Tables[0].Location.DataLocationTemplateRowStart;
			int num = ToKhaiInfo.SheetInfos[3].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count;
			for (int num2 = dataLocationTemplateRowStart2; num2 < num; num2++)
			{
				int num3 = 0;
				int num4 = 0;
				int num5 = 0;
				int num6 = 0;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart;
			if (iRow >= dataLocationTemplateRowStart)
			{
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("I_36")[1], GeneralFunctions.ArrayCellId("I_36")[0]].Value = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1] != "") + ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[0] != "");
				SheetView sheetView2 = winFpSpread.Sheets[0];
				if (iCol == odrColD)
				{
					sheetView2.Cells[CellID_Ct22[1], CellID_Ct22[0]].Value = sheetView2.Cells[CellID_Ct22_PL01[1], CellID_Ct22_PL01[0]].Value;
				}
				else if (iCol == odrColQ)
				{
					sheetView2.Cells[CellID_Ct24[1], CellID_Ct24[0]].Value = sheetView2.Cells[CellID_Ct24_PL[1], CellID_Ct24_PL[0]].Value;
					sheetView2.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView2.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
					sheetView2.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView2.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
				}
				else if (iCol == odrColJ)
				{
					sheetView2.Cells[CellID_Ct26[1], CellID_Ct26[0]].Value = sheetView2.Cells[CellID_Ct26_PL[1], CellID_Ct26_PL[0]].Value;
				}
				else if (iCol == odrColH)
				{
					sheetView2.Cells[CellID_Ct29[1], CellID_Ct29[0]].Value = sheetView2.Cells[CellID_Ct29_PL[1], CellID_Ct29_PL[0]].Value;
					sheetView2.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView2.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
				}
				else if (iCol == odrColR)
				{
					sheetView2.Cells[CellID_Ct38[1], CellID_Ct38[0]].ResetFormula();
					sheetView2.Cells[CellID_Ct38[1], CellID_Ct38[0]].Value = sheetView2.Cells[CellID_Ct38_PL[1], CellID_Ct38_PL[0]].Value;
				}
				int[] array = new int[7] { odrColH, odrColJ, odrColL, odrColM, odrColN, odrColO, odrColI };
				if (Array.IndexOf(array, (int)iCol) != -1)
				{
					sheetView.Cells[(int)iRow - 1, odrColR - 1].Formula = "AR" + iRow;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "IF(OR(F###=\"1\",F###=\"x\",F###=\"X\",F###=TRUE),IF(ROUND((G###-J###)*0.2- M###,0) < 0, 0, ROUND((G###-J###)*0.2- M###,0)),IF(ROUND((G###-J###)*0.1 - M###,0) < 0, 0, ROUND((G###-J###)*0.1 - M###,0)))";
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[2].Tables[0].Location.DataLocationTemplateRowStart;
			if (iRow >= dataLocationTemplateRowStart)
			{
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("I_36")[1], GeneralFunctions.ArrayCellId("I_36")[0]].Value = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1] != "") + ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[0] != "");
				SheetView sheetView = winFpSpread.Sheets[0];
				SheetView sheetView2 = winFpSpread.Sheets[2];
				if (iCol == odrColF)
				{
					sheetView.Cells[CellID_Ct24[1], CellID_Ct24[0]].Value = sheetView.Cells[CellID_Ct24_PL[1], CellID_Ct24_PL[0]].Value;
					sheetView.Cells[CellID_Ct29[1], CellID_Ct29[0]].Value = sheetView.Cells[CellID_Ct29_PL[1], CellID_Ct29_PL[0]].Value;
					sheetView.Cells[CellID_Ct30[1], CellID_Ct30[0]].Value = sheetView.Cells[CellID_Ct30_PL[1], CellID_Ct30_PL[0]].Value;
					sheetView.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
					sheetView.Cells[CellID_Ct33[1], CellID_Ct33[0]].Value = sheetView.Cells[CellID_Ct33_PL[1], CellID_Ct33_PL[0]].Value;
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].ResetFormula();
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
					sheetView.Cells[CellID_Ct38[1], CellID_Ct38[0]].ResetFormula();
					sheetView.Cells[CellID_Ct38[1], CellID_Ct38[0]].Value = sheetView.Cells[CellID_Ct38_PL[1], CellID_Ct38_PL[0]].Value;
				}
				else if (iCol == odrColK)
				{
					sheetView.Cells[CellID_Ct24[1], CellID_Ct24[0]].Value = sheetView.Cells[CellID_Ct24_PL[1], CellID_Ct24_PL[0]].Value;
					sheetView.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
					sheetView.Cells[CellID_Ct33[1], CellID_Ct33[0]].Value = sheetView.Cells[CellID_Ct33_PL[1], CellID_Ct33_PL[0]].Value;
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].ResetFormula();
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
				}
				else if (iCol == odrColJ)
				{
					sheetView.Cells[CellID_Ct26[1], CellID_Ct26[0]].Value = sheetView.Cells[CellID_Ct26_PL[1], CellID_Ct26_PL[0]].Value;
				}
				else if (iCol == odrColG || iCol == odrColC)
				{
					sheetView.Cells[CellID_Ct29[1], CellID_Ct29[0]].Value = sheetView.Cells[CellID_Ct29_PL[1], CellID_Ct29_PL[0]].Value;
					sheetView.Cells[CellID_Ct30[1], CellID_Ct30[0]].Value = sheetView.Cells[CellID_Ct30_PL[1], CellID_Ct30_PL[0]].Value;
					sheetView.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
					sheetView.Cells[CellID_Ct33[1], CellID_Ct33[0]].Value = sheetView.Cells[CellID_Ct33_PL[1], CellID_Ct33_PL[0]].Value;
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].ResetFormula();
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
				}
				else if (iCol == odrColH)
				{
					sheetView.Cells[CellID_Ct29[1], CellID_Ct29[0]].Value = sheetView.Cells[CellID_Ct29_PL[1], CellID_Ct29_PL[0]].Value;
					sheetView.Cells[CellID_Ct30[1], CellID_Ct30[0]].Value = sheetView.Cells[CellID_Ct30_PL[1], CellID_Ct30_PL[0]].Value;
					sheetView.Cells[CellID_Ct32[1], CellID_Ct32[0]].Value = sheetView.Cells[CellID_Ct32_PL[1], CellID_Ct32_PL[0]].Value;
					sheetView.Cells[CellID_Ct33[1], CellID_Ct33[0]].Value = sheetView.Cells[CellID_Ct33_PL[1], CellID_Ct33_PL[0]].Value;
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].ResetFormula();
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
					sheetView.Cells[CellID_Ct40[1], CellID_Ct40[0]].ResetFormula();
					sheetView.Cells[CellID_Ct40[1], CellID_Ct40[0]].Value = sheetView.Cells[CellID_Ct40_PL[1], CellID_Ct40_PL[0]].Value;
					sheetView.Cells[CellID_Ct41[1], CellID_Ct41[0]].ResetFormula();
					sheetView.Cells[CellID_Ct41[1], CellID_Ct41[0]].Value = sheetView.Cells[CellID_Ct41_PL[1], CellID_Ct41_PL[0]].Value;
				}
				else if (iCol == odrColL)
				{
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].ResetFormula();
					sheetView.Cells[CellID_Ct35[1], CellID_Ct35[0]].Value = sheetView.Cells[CellID_Ct35_PL[1], CellID_Ct35_PL[0]].Value;
				}
				else if (iCol == odrColM)
				{
					sheetView.Cells[CellID_Ct38[1], CellID_Ct38[0]].ResetFormula();
					sheetView.Cells[CellID_Ct38[1], CellID_Ct38[0]].Value = sheetView.Cells[CellID_Ct38_PL[1], CellID_Ct38_PL[0]].Value;
				}
				if (iCol == odrColH)
				{
					sheetView2.Cells[(int)iRow - 1, odrColL - 1].Formula = CT16formula.Replace("###", iRow.ToString());
				}
				if ((iCol == odrColF) | (iCol == odrColI))
				{
					sheetView2.Cells[(int)iRow - 1, odrColM - 1].Formula = CT17formula.Replace("###", iRow.ToString());
				}
				if ((iCol == odrColF) | (iCol == odrColG) | (iCol == odrColJ) | (iCol == odrColM))
				{
					sheetView2.Cells[(int)iRow - 1, odrColK - 1].Formula = text.Replace("###", iRow.ToString());
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			FinishImportExcel(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void UpdateFormula_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			FinishImportExcel(winFpSpread, ToKhaiInfo);
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (!string.IsNullOrEmpty(HTKKFile.ErrorMessage) && (ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count > 500 || ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count > 500))
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
			StringBuilder stringBuilder = new StringBuilder();
			SheetView sheetView = winFpSpread.Sheets[0];
			string message = HTKKFile.GetMessage(sheetView, CellId_E85);
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			bool flag = (bool)sheetView.Cells[CellId_chkQTKTN[1], CellId_chkQTKTN[0]].Value;
			string text = sheetView.Cells[CellId_lyDo[1], CellId_lyDo[0]].Text;
			if (flag)
			{
				if (string.IsNullOrEmpty(text))
				{
					stringBuilder.Append("Lỗi quyết toán không tròn năm: " + (CellId_lyDo[1] + 1) + "\n" + Environment.NewLine);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_lyDo[1], CellId_lyDo[0], message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_lyDo[1], CellId_lyDo[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_lyDo[1], CellId_lyDo[0]);
			}
			switch (value)
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			StringBuilder stringBuilder = new StringBuilder();
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[42, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[50, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[56, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[66, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[67, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[69, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[79, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[81, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[83, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text18 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[84, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text19 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[91, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text20 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[92, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text21 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[93, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text22 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text23 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[96, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text24 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[97, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text25 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[128, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text26 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[124, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text27 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[125, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[1];
			string message = HTKKFile.GetMessage(sheetView, CellId_E21);
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[2].Tables[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int colIndex = hTKKCMCTableInfo2.RowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int num3 = hTKKCMCTableInfo2.RowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int colIndex2 = hTKKCMCTableInfo2.RowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num4 = hTKKCMCTableInfo2.RowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int num5 = hTKKCMCTableInfo2.RowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int colIndex3 = hTKKCMCTableInfo2.RowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int colIndex4 = hTKKCMCTableInfo2.RowsInfo.Cells[7].CellTemplate.ColIndex - 1;
			int colIndex5 = hTKKCMCTableInfo2.RowsInfo.Cells[8].CellTemplate.ColIndex - 1;
			int colIndex6 = hTKKCMCTableInfo2.RowsInfo.Cells[9].CellTemplate.ColIndex - 1;
			int num6 = hTKKCMCTableInfo2.RowsInfo.Cells[10].CellTemplate.ColIndex - 1;
			int num7 = hTKKCMCTableInfo2.RowsInfo.Cells[11].CellTemplate.ColIndex - 1;
			int colIndex7 = hTKKCMCTableInfo2.RowsInfo.Cells[12].CellTemplate.ColIndex - 1;
			int colIndex8 = hTKKCMCTableInfo2.RowsInfo.Cells[14].CellTemplate.ColIndex - 1;
			int colIndex9 = hTKKCMCTableInfo2.RowsInfo.Cells[15].CellTemplate.ColIndex - 1;
			int num8 = hTKKCMCTableInfo2.RowsInfo.Cells[17].CellTemplate.ColIndex - 1;
			int colIndex10 = hTKKCMCTableInfo2.RowsInfo.Cells[18].CellTemplate.ColIndex - 1;
			int colIndex11 = hTKKCMCTableInfo2.RowsInfo.Cells[19].CellTemplate.ColIndex - 1;
			int column = hTKKCMCTableInfo2.RowsInfo.Cells[16].CellTemplate.ColIndex - 1;
			int num9 = int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]);
			int num10 = int.Parse(PeriodInfo.TuThang.Split(new char[1] { '/' })[0]);
			int num11 = (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
			int num12 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			string mst = TTinChung.mst;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, 13, HTKKFile.ColumnLettersToIndex("E"), "0");
			string text28 = "";
			string defaultValue = hTKKCMCTableInfo2.RowsInfo.Cells[13].DefaultValue;
			string defaultValue2 = hTKKCMCTableInfo2.RowsInfo.Cells[17].DefaultValue;
			string defaultValue3 = hTKKCMCTableInfo2.RowsInfo.Cells[18].DefaultValue;
			string defaultValue4 = hTKKCMCTableInfo2.RowsInfo.Cells[19].DefaultValue;
			int count = hTKKCMCTableInfo2.RowsValues.Count;
			int num13 = 0;
			int num14 = 0;
			int num15 = 1;
			int num16 = 100000;
			int num17 = 0;
			int num18 = 0;
			int num19 = 0;
			int num20 = 0;
			num17 = ((count < num16) ? count : num16);
			num18 = ((count < num16 * 2) ? (count - num17) : (num16 * 2));
			num19 = ((count < num16 * 3) ? (count - num18) : (num16 * 3));
			if (count >= num16 * 4)
			{
				num20 = num16 * 4;
			}
			else
			{
				num20 = count - num19;
			}
			int num21 = num17;
			if (num21 > 0)
			{
				for (int i = 0; i < num21; i++)
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
					HTKKFile.RemoveValidate(sheetView, num + i, num3);
					HTKKFile.RemoveValidate(sheetView, num + i, num4);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex5);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex8);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex9);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, colIndex10);
					HTKKFile.RemoveValidateBackGround(sheetView, num + i, colIndex11);
					string[] arrARowValue = hTKKCMCTableInfo2.RowsValues[i].RowValue.ToArray();
					if (i == 0 && !GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo2.RowsInfo, arrARowValue))
					{
						continue;
					}
					if (string.IsNullOrEmpty(arrARowValue[1]))
					{
						stringBuilder.Append("Lỗi Họ tên tại dòng số " + (i + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex, message, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
					}
					bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[num + i, num4].Value);
					bool flag2 = ConvertHelpers.ConvertToBool(sheetView.Cells[num + i, column].Value);
					if (flag)
					{
						if (string.IsNullOrEmpty(arrARowValue[2].Trim()))
						{
							stringBuilder.Append("Lỗi MST không được để trống: " + (i + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text5, i + 1);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, text10, i + 1);
							value = 2;
						}
					}
					else if (string.IsNullOrEmpty(arrARowValue[2].Trim()) && string.IsNullOrEmpty(arrARowValue[3]))
					{
						stringBuilder.Append("Lỗi MST không được để trống: " + (i + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text5, i + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text5);
						value = 2;
					}
					else if (string.IsNullOrEmpty(arrARowValue[2].Trim()) && flag2)
					{
						stringBuilder.Append("Lỗi MST không được để trống: " + (i + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text4, i + 1);
						value = 2;
					}
					int num22;
					if (arrARowValue[2].Trim().Length > 0)
					{
						if (arrARowValue[2].Trim().Length == 10 || arrARowValue[2].Trim().Length == 12)
						{
							if (!ValidateBase.CheckMaSoThueCCCD(arrARowValue[2].Trim()))
							{
								stringBuilder.Append("Lỗi sai cấu trúc MST: " + (i + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text, i + 1);
								value = 2;
							}
							else if (mst.Equals(arrARowValue[2].Trim()))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text22);
								value = 2;
							}
							else if (!HTKKFormater.isNullString(valueCustom) && valueCustom.Equals(arrARowValue[2].Trim()))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text23, i + 1);
								value = 2;
							}
							else
							{
								num22 = hTKKCMCTableInfo2.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(arrARowValue[2].Trim()));
								if (arrARowValue[2].Trim().Equals(TTinChung.mst))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")), i + 1);
									value = 2;
								}
								if (num22 > 1)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")), i + 1);
									value = 2;
								}
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text, i + 1);
							value = 2;
						}
					}
					if (!string.IsNullOrEmpty(arrARowValue[3]))
					{
						if (arrARowValue[3].Contains(" "))
						{
							stringBuilder.Append("Lỗi CMT/Hộ chiếu : " + (i + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text3, i + 1);
							value = 2;
						}
						else
						{
							num22 = hTKKCMCTableInfo2.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[3].Equals(arrARowValue[3]));
							if (num22 > 1)
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
								if (value != 2)
								{
									value = 1;
								}
							}
						}
					}
					if (!string.IsNullOrEmpty(arrARowValue[8]) && ConvertHelpers.ConvertToInt(arrARowValue[8]) > 100)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex5, text25, i + 1);
						value = 2;
					}
					int num23 = 0;
					text28 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AJ") - 1].Text;
					if (text28.Trim().Equals("1"))
					{
						stringBuilder.Append("Lỗi ct12 < ct11 : " + (i + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, text6, i + 1);
						value = 2;
					}
					else if (num12 > 2018 || (num12 == 2018 && num10 >= 8))
					{
						string obj = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("I") - 1].Value.ToString();
						if (ConvertHelpers.ConvertToDouble(obj) > 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, text27, i + 1);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
						num23++;
					}
					text28 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BH") - 1].Text;
					if (text28.Trim().Equals("1"))
					{
						stringBuilder.Append("Lỗi ct13 < ct11 : " + (i + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex4, text12, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
						num23++;
					}
					if (num23 > 1)
					{
						text28 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BK") - 1].Text;
						if (text28.Trim().Equals("1"))
						{
							stringBuilder.Append("Lỗi ct12 + ct13 <= ct11 : " + (i + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, text15, i + 1);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex4, text15);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
							HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
						}
					}
					double num24 = 1000000.0 * ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("Q") - 1].Value);
					if (ConvertHelpers.ConvertToDouble(arrARowValue[12]) > num24)
					{
						stringBuilder.Append("Lỗi ct18 : " + (i + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex7, text13, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
					}
					text28 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AL") - 1].Text;
					if (arrARowValue[14].Trim().Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex8, text17, i + 1);
						value = 2;
					}
					else if (text28.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex8, text7);
						if (value != 2)
						{
							value = 1;
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex8);
					}
					text28 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BC") - 1].Text;
					if (arrARowValue[15].Trim().Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex9, text17, i + 1);
						value = 2;
					}
					else if (num12 > 2018 || (num12 == 2018 && num10 >= 8))
					{
						string obj2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("R") - 1].Value.ToString();
						if (ConvertHelpers.ConvertToDouble(obj2) > 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex9, text26);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, colIndex9);
						}
					}
					else if (text28.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex9, text20);
						if (value != 2)
						{
							value = 1;
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex9);
					}
					if (!arrARowValue[18].Trim(new char[1] { '0' }).Equals("") && !arrARowValue[19].Trim(new char[1] { '0' }).Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex10, text24, i + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex11, text24, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidateBackGround(sheetView, num + i, colIndex10);
						HTKKFile.RemoveValidateBackGround(sheetView, num + i, colIndex11);
					}
					text28 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BB") - 1].Text;
					if (text28.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex6, text11, i + 1);
						value = 2;
					}
					else if (ConvertHelpers.ConvertToDouble(arrARowValue[8]) > 0.0 && ConvertHelpers.ConvertToDouble(arrARowValue[9]) <= 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex6, text21, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
					}
					if (!flag || string.IsNullOrEmpty(arrARowValue[2].Trim()))
					{
						continue;
					}
					num22 = ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[2].Trim()));
					if (num22 >= 1 && sheetView.Cells[num + i, num3].BackColor != Color.Pink)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
						if (value != 2)
						{
							value = 1;
						}
					}
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[42, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[69, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[81, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[84, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[91, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[96, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[126, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[127, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[2];
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			int colIndex = hTKKCMCTableInfo2.RowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			int num2 = hTKKCMCTableInfo2.RowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int colIndex2 = hTKKCMCTableInfo2.RowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int num3 = hTKKCMCTableInfo2.RowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num4 = hTKKCMCTableInfo2.RowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int colIndex3 = hTKKCMCTableInfo2.RowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int colIndex4 = hTKKCMCTableInfo2.RowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int colIndex5 = hTKKCMCTableInfo2.RowsInfo.Cells[7].CellTemplate.ColIndex - 1;
			int colIndex6 = hTKKCMCTableInfo2.RowsInfo.Cells[8].CellTemplate.ColIndex - 1;
			int colIndex7 = hTKKCMCTableInfo2.RowsInfo.Cells[9].CellTemplate.ColIndex - 1;
			int colIndex8 = hTKKCMCTableInfo2.RowsInfo.Cells[10].CellTemplate.ColIndex - 1;
			string mst = TTinChung.mst;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, 12, HTKKFile.ColumnLettersToIndex("D"), "0");
			string text18 = "";
			int num5 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			string tuThang = PeriodInfo.TuThang;
			int num6 = ConvertHelpers.ConvertToInt(tuThang.Split(new char[1] { '/' })[0]);
			string defaultValue = hTKKCMCTableInfo2.RowsInfo.Cells[11].DefaultValue;
			int count = hTKKCMCTableInfo2.RowsValues.Count;
			int num7 = 0;
			int num8 = 0;
			int num9 = 1;
			int num10 = 100000;
			int num11 = 0;
			int num12 = 0;
			int num13 = 0;
			int num14 = 0;
			num11 = ((count < num10) ? count : num10);
			num12 = ((count < num10 * 2) ? (count - num11) : (num10 * 2));
			num13 = ((count < num10 * 3) ? (count - num12) : (num10 * 3));
			if (count >= num10 * 4)
			{
				num14 = num10 * 4;
			}
			else
			{
				num14 = count - num13;
			}
			int num15 = num11;
			if (num15 > 0)
			{
				for (int i = 0; i < num15; i++)
				{
					string[] arrARowValue = hTKKCMCTableInfo2.RowsValues[i].RowValue.ToArray();
					if (i == 0 && !GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo2.RowsInfo, arrARowValue))
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
						HTKKFile.RemoveValidate(sheetView, num + i, num2);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex5);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex8);
						continue;
					}
					if (string.IsNullOrEmpty(arrARowValue[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex, text3, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
					}
					HTKKFile.RemoveValidate(sheetView, num + i, num2);
					text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AC") - 1].Text;
					if (text18.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
						value = 1;
					}
					if (arrARowValue[1].Trim().Length > 0)
					{
						if (arrARowValue[1].Trim().Length == 10 || arrARowValue[1].Trim().Length == 12)
						{
							if (!ValidateBase.CheckMaSoThueCCCD(arrARowValue[1].Trim()))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text14, i + 1);
								value = 2;
							}
							else if (mst.Equals(arrARowValue[1].Trim()))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text5);
								value = 2;
							}
							else if (!HTKKFormater.isNullString(valueCustom) && valueCustom.Equals(arrARowValue[1].Trim()))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text15, i + 1);
								value = 2;
							}
							else
							{
								int num16 = hTKKCMCTableInfo2.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[1].Trim()));
								if (arrARowValue[1].Trim().Equals(TTinChung.mst))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")), i + 1);
									value = 2;
								}
								if (num16 > 1)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24")), i + 1);
									value = 2;
								}
								num16 = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[1].Trim()) && ConvertHelpers.ConvertToBool(p.RowValue[3]));
								if (num16 >= 1 && sheetView.Cells[num + i, num2].BackColor != Color.Pink)
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
									if (value != 2)
									{
										value = 1;
									}
								}
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text4, i + 1);
							value = 2;
						}
					}
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
					if (string.IsNullOrEmpty(arrARowValue[1].Trim()) && string.IsNullOrEmpty(arrARowValue[2].Trim()))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text7, i + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text7);
						value = 2;
					}
					if (arrARowValue[2].Contains(" "))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text6, i + 1);
						value = 2;
					}
					else
					{
						int num16 = hTKKCMCTableInfo2.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(arrARowValue[2].Trim()) & !HTKKFormater.isNullString(p.RowValue[2]));
						if (num16 > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41")), i + 1);
							value = 2;
						}
					}
					text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text;
					if (text18.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, text11, i + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex4, text11);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex5, text11);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex3);
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex5);
						if (num5 > 2018 || (num5 == 2018 && num6 >= 8))
						{
							string obj = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("I") - 1].Value.ToString();
							if (ConvertHelpers.ConvertToDouble(obj) > 0.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex4, text16, i + 1);
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, colIndex4);
						}
					}
					text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Z") - 1].Text;
					if (text18.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex6, text8, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
					}
					text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AT") - 1].Text;
					if (text18.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex7, text13, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
					}
					text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AR") - 1].Text;
					if (text18.Trim().Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex8, text10, i + 1);
						value = 2;
					}
					else if (num5 > 2018 || (num5 == 2018 && num6 >= 8))
					{
						string obj2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToString();
						if (ConvertHelpers.ConvertToDouble(obj2) > 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex8, text17, i + 1);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, colIndex8);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex8);
					}
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
			throw ex;
		}
	}

	public void Validate_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[70, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[71, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[72, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[73, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[74, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[75, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[76, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[78, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[83, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[85, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[86, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[87, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[89, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[96, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[115, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			SheetView sheetView = winFpSpread.Sheets[3];
			int value = 0;
			levelWarning.TryGetValue(3, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[2].Tables[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[0];
			string mst = TTinChung.mst;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, 13, HTKKFile.ColumnLettersToIndex("E"), "0");
			int num = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
			int colIndex = hTKKCMCTableInfo3.RowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			int num2 = hTKKCMCTableInfo3.RowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int num3 = hTKKCMCTableInfo3.RowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int colIndex2 = hTKKCMCTableInfo3.RowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num4 = hTKKCMCTableInfo3.RowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int colIndex3 = hTKKCMCTableInfo3.RowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int colIndex4 = hTKKCMCTableInfo3.RowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int num5 = hTKKCMCTableInfo3.RowsInfo.Cells[7].CellTemplate.ColIndex - 1;
			int colIndex5 = hTKKCMCTableInfo3.RowsInfo.Cells[8].CellTemplate.ColIndex - 1;
			int colIndex6 = hTKKCMCTableInfo3.RowsInfo.Cells[9].CellTemplate.ColIndex - 1;
			int num6 = hTKKCMCTableInfo3.RowsInfo.Cells[10].CellTemplate.ColIndex - 1;
			int num7 = hTKKCMCTableInfo3.RowsInfo.Cells[11].CellTemplate.ColIndex - 1;
			int colIndex7 = hTKKCMCTableInfo3.RowsInfo.Cells[12].CellTemplate.ColIndex - 1;
			int colIndex8 = hTKKCMCTableInfo3.RowsInfo.Cells[13].CellTemplate.ColIndex - 1;
			int colIndex9 = hTKKCMCTableInfo3.RowsInfo.Cells[14].CellTemplate.ColIndex - 1;
			int colIndex10 = hTKKCMCTableInfo3.RowsInfo.Cells[15].CellTemplate.ColIndex - 1;
			int colIndex11 = hTKKCMCTableInfo3.RowsInfo.Cells[16].CellTemplate.ColIndex - 1;
			int colIndex12 = hTKKCMCTableInfo3.RowsInfo.Cells[17].CellTemplate.ColIndex - 1;
			int colIndex13 = hTKKCMCTableInfo3.RowsInfo.Cells[18].CellTemplate.ColIndex - 1;
			int colIndex14 = hTKKCMCTableInfo3.RowsInfo.Cells[19].CellTemplate.ColIndex - 1;
			int num8 = hTKKCMCTableInfo3.RowsInfo.Cells[20].CellTemplate.ColIndex - 1;
			int num9 = hTKKCMCTableInfo3.RowsInfo.Cells[21].CellTemplate.ColIndex - 1;
			string denThang = PeriodInfo.DenThang;
			DateTime dateTime = ConvertHelpers.ConvertToDateTime("01/" + denThang, "dd/MM/yyyy");
			string nam = PeriodInfo.Nam;
			DateTime t = ConvertHelpers.ConvertToDateTime("01/01/" + nam, "dd/MM/yyyy");
			int count = hTKKCMCTableInfo3.RowsValues.Count;
			int num10 = 0;
			int num11 = 0;
			int num12 = 1;
			int num13 = 100000;
			int num14 = 0;
			int num15 = 0;
			int num16 = 0;
			int num17 = 0;
			num14 = ((count < num13) ? count : num13);
			num15 = ((count < num13 * 2) ? (count - num14) : (num13 * 2));
			num16 = ((count < num13 * 3) ? (count - num15) : (num13 * 3));
			if (count >= num13 * 4)
			{
				num17 = num13 * 4;
			}
			else
			{
				num17 = count - num16;
			}
			int num18 = num14;
			if (num18 > 0)
			{
				for (int i = 0; i < num18; i++)
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
					HTKKFile.RemoveValidate(sheetView, num + i, num2);
					HTKKFile.RemoveValidate(sheetView, num + i, num3);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
					HTKKFile.RemoveValidate(sheetView, num + i, num4);
					HTKKFile.RemoveValidate(sheetView, num + i, num5);
					HTKKFile.RemoveValidate(sheetView, num + i, num8);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex5);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex6);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex7);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex8);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex9);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex10);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex11);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex12);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex13);
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex14);
					string[] arrARowValue = hTKKCMCTableInfo3.RowsValues[i].RowValue.ToArray();
					if (i == 0 && !GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo3.RowsInfo, arrARowValue))
					{
						continue;
					}
					if (string.IsNullOrEmpty(arrARowValue[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex, text, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, colIndex);
					}
					if (string.IsNullOrEmpty(arrARowValue[1].Trim()))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text9, i + 1);
						value = 2;
					}
					else if (arrARowValue[1].Trim().Length == 10 || arrARowValue[1].Trim().Length == 12)
					{
						if (!ValidateBase.CheckMaSoThueCCCD(arrARowValue[1].Trim()))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text2, i + 1);
							value = 2;
						}
						else if (mst.Equals(arrARowValue[1].Trim()))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text3, i + 1);
							value = 2;
						}
						else if (!HTKKFormater.isNullString(valueCustom) && valueCustom.Equals(arrARowValue[1].Trim()))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text15, i + 1);
							value = 2;
						}
						else
						{
							bool flag = false;
							int num19 = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(arrARowValue[1].Trim()));
							if (num19 > 0)
							{
								flag = true;
							}
							else
							{
								num19 = ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[1].Trim()));
								if (num19 > 0)
								{
									flag = true;
								}
								else if (!flag && sheetView.Cells[num + i, num2].BackColor != Color.Pink)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_90")), i + 1);
									value = 2;
								}
							}
							num19 = hTKKCMCTableInfo3.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[4].Equals(arrARowValue[1].Trim()));
							if (num19 >= 1 && sheetView.Cells[num + i, num2].BackColor != Color.Pink)
							{
								HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
								if (value != 2)
								{
									value = 1;
								}
							}
							else if (sheetView.Cells[num + i, num2].BackColor != Color.Pink)
							{
								HTKKFile.RemoveValidate(sheetView, num + i, num2);
							}
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text2, i + 1);
						value = 2;
					}
					string text17 = sheetView.Cells[num + i, num3].Text;
					if (string.IsNullOrEmpty(text17))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, text4, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, num3);
					}
					HTKKFile.RemoveValidate(sheetView, num + i, colIndex2);
					if (!string.IsNullOrEmpty(arrARowValue[3]))
					{
						if (HTKKFormater.IsDateTime(arrARowValue[3], "dd/MM/yyyy", out var dtDate) && dtDate >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
						{
							if (dtDate > DateTime.Now.Date)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text5, i + 1);
								value = 2;
							}
							else if (!arrARowValue[20].Trim().Equals(""))
							{
								int num20 = 0;
								int num21 = 0;
								num20 = ConvertHelpers.ConvertToInt(arrARowValue[20].Split(new char[1] { '/' })[0]);
								num21 = ConvertHelpers.ConvertToInt(arrARowValue[20].Split(new char[1] { '/' })[1]);
								if (dtDate.Year > num21 || (dtDate.Year == num21 && dtDate.Month > num20))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text11, i + 1);
									value = 2;
								}
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex2, text10, i + 1);
							value = 2;
						}
					}
					HTKKFile.RemoveValidate(sheetView, num + i, num4);
					if (arrARowValue[4].Length > 0)
					{
						if (arrARowValue[4].Length == 10 || arrARowValue[4].Length == 12)
						{
							if (!ValidateBase.CheckMaSoThueCCCD(arrARowValue[4]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, text6, i + 1);
								value = 2;
							}
							else
							{
								if (arrARowValue[4] == mst)
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
									if (value != 2)
									{
										value = 1;
									}
								}
								int num19 = hTKKCMCTableInfo3.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[4]));
								if (num19 >= 1 && sheetView.Cells[num + i, num4].BackColor != Color.Pink)
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
									if (value != 2)
									{
										value = 1;
									}
								}
								else
								{
									int num22 = hTKKCMCTableInfo3.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[4].Equals(arrARowValue[4]));
									if (num22 > 1 && sheetView.Cells[num + i, num4].BackColor != Color.Pink)
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_124")), i + 1);
										value = 2;
									}
								}
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num4, text6, i + 1);
							value = 2;
						}
					}
					string text18 = sheetView.Cells[num + i, num5].Text;
					if (!string.IsNullOrEmpty(text18) && text18.Contains(" "))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num5, text7, i + 1);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, num5);
					}
					if (!HTKKFormater.isNullString(arrARowValue[6]) && HTKKFormater.isNullString(arrARowValue[5]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex3, text9);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_110")), i + 1);
						value = 2;
					}
					if (!HTKKFormater.isNullString(arrARowValue[9]) && HTKKFormater.isNullString(arrARowValue[8]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex5, text9);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex6, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_111")), i + 1);
						value = 2;
					}
					if (!HTKKFormater.isNullString(arrARowValue[13]) && HTKKFormater.isNullString(arrARowValue[12]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_112")), i + 1);
						value = 2;
					}
					if (!HTKKFormater.isNullString(arrARowValue[15]) && HTKKFormater.isNullString(arrARowValue[14]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex10, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_113")), i + 1);
						value = 2;
					}
					if (!HTKKFormater.isNullString(arrARowValue[17]) && HTKKFormater.isNullString(arrARowValue[16]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex12, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_114")), i + 1);
						value = 2;
					}
					if (!HTKKFormater.isNullString(arrARowValue[19]) && HTKKFormater.isNullString(arrARowValue[18]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_115")), i + 1);
						value = 2;
					}
					if (!HTKKFormater.isNullString(arrARowValue[18]) && !HTKKFormater.isNullString(arrARowValue[16]))
					{
						if (arrARowValue[18].Length == 7 && arrARowValue[16].Length == 5)
						{
							if (!arrARowValue[18].Substring(0, 5).Equals(arrARowValue[16]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex13, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_101")), i + 1);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_101")));
								value = 2;
							}
						}
						else if (arrARowValue[18].Length != 7)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex13, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_104")));
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_104")));
						}
					}
					if (!HTKKFormater.isNullString(arrARowValue[16]) && !HTKKFormater.isNullString(arrARowValue[14]))
					{
						if (arrARowValue[16].Length == 5 && arrARowValue[14].Length == 3)
						{
							if (!arrARowValue[16].Substring(0, 3).Equals(arrARowValue[14]))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex11, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_102")), i + 1);
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex12, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_102")));
								value = 2;
							}
						}
						else if (arrARowValue[16].Length != 5)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex11, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_105")));
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex12, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_105")));
						}
					}
					if (!HTKKFormater.isNullString(arrARowValue[14]) && !HTKKFormater.isNullString(arrARowValue[12]))
					{
						if (arrARowValue[14].Length == 3 && !arrARowValue[14].Equals("999") && !arrARowValue[12].Equals("01"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_103")), i + 1);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex10, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_103")));
							value = 2;
						}
						else if (arrARowValue[14].Length != 3)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_106")));
							HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, colIndex10, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_106")));
						}
					}
					HTKKFile.RemoveValidate(sheetView, num + i, num8);
					HTKKFile.RemoveValidate(sheetView, num + i, num9);
					string text19 = sheetView.Cells[num + i, num8].Text;
					string text20 = sheetView.Cells[num + i, num9].Text;
					DateTime dtDate2;
					if (string.IsNullOrEmpty(text19))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, text9);
						value = 2;
					}
					else if (HTKKFormater.IsDateTime("01/" + text19, "dd/MM/yyyy", out dtDate2))
					{
						if (DateTime.Compare(dtDate2, t) < 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_123")), i + 1);
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, text10, i + 1);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text20))
					{
						if (HTKKFormater.IsDateTime("01/" + text20, "dd/MM/yyyy", out dtDate2))
						{
							if (dtDate2 > dateTime)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, text16, i + 1);
								value = 2;
							}
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, text10, i + 1);
							value = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, text9, i + 1);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text19) && !string.IsNullOrEmpty(text20) && HTKKFormater.IsDateTime("01/" + text19, "dd/MM/yyyy", out var dtDate3) && HTKKFormater.IsDateTime("01/" + text20, "dd/MM/yyyy", out var dtDate4) && (dtDate3.Year > dtDate4.Year || (dtDate3.Year == dtDate4.Year && dtDate3.Month > dtDate4.Month)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num8, text8, i + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num9, text8);
						value = 2;
					}
					if (string.IsNullOrEmpty(arrARowValue[9]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex6, text9, i + 1);
						value = 2;
					}
					if (arrARowValue[12].Trim().Equals("01") && !arrARowValue[13].Trim().Equals(""))
					{
						if (arrARowValue[14].Equals("999"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex10, text10, i + 1);
							value = 2;
						}
						if (arrARowValue[16].Equals("99999"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex12, text10);
							value = 2;
						}
						if (arrARowValue[18].Equals("9999999"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex14, text10);
							value = 2;
						}
					}
					if (!arrARowValue[15].Equals("") || !arrARowValue[17].Equals("") || !arrARowValue[19].Equals(""))
					{
						if (arrARowValue[13].Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex8, text9, i + 1);
							value = 2;
						}
						if (arrARowValue[15].Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex10, text9);
							value = 2;
						}
						if (arrARowValue[17].Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex12, text9);
							value = 2;
						}
						if (arrARowValue[19].Equals(""))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colIndex14, text9);
							value = 2;
						}
					}
					int num23 = 0;
					int num24 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
					int num25 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
					num23 = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(arrARowValue[1].Trim()));
					if (num23 == 0)
					{
						num23 = ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(arrARowValue[1].Trim()));
					}
					if (num23 == 0 && sheetView.Cells[num + i, num2].BackColor != Color.Pink)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, text14, i + 1);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[3].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[3].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[3].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static List<HTKKCMCRowValue> GetDataFromExcelFile(string filePath, HTKKCMCRowInfo rowInfo, int DataLocationTaiBangKeRowStart, int iColCheck)
	{
		try
		{
			string strEnd = "";
			return GeneralFunctions.GetDataFromExcelFile(filePath, rowInfo, DataLocationTaiBangKeRowStart, strEnd, iColCheck);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool IsBetweenMonth(DateTime checkDate, DateTime beginDate, DateTime endDate)
	{
		try
		{
			return (checkDate >= beginDate) & (checkDate <= endDate);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public DateTime GetMonth(string thang)
	{
		try
		{
			thang = thang.Remove(0, 2).Insert(0, "01");
			return DateTime.ParseExact(thang, "ddMMyyyy", null);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ButtonCellType buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[16, HTKKFile.ColumnLettersToIndex("D") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[16, HTKKFile.ColumnLettersToIndex("D") - 1].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[17, HTKKFile.ColumnLettersToIndex("H") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[17, HTKKFile.ColumnLettersToIndex("H") - 1].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[16, HTKKFile.ColumnLettersToIndex("Q") - 1].Text;
		buttonCellType.WordWrap = true;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[16, HTKKFile.ColumnLettersToIndex("Q") - 1].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[19, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\Check.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[19, HTKKFile.ColumnLettersToIndex("G") - 1].CellType = buttonCellType;
		int count = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count;
		for (int i = 0; i < count; i++)
		{
			int row = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart + i - 1;
			if (ConvertHelpers.ConvertToBool(sView.Cells[row, HTKKFile.ColumnLettersToIndex("S") - 1].Text))
			{
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = false;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = false;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("T") - 1].ResetFormula();
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("V") - 1].ResetFormula();
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("T") - 1].Value = ConvertHelpers.ConvertToDouble(ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues[i].RowValue[17]);
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("U") - 1].Value = ConvertHelpers.ConvertToDouble(ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues[i].RowValue[18]);
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("V") - 1].Value = ConvertHelpers.ConvertToDouble(ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues[i].RowValue[19]);
			}
			int num = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			string[] array = PeriodInfo.TuThang.Split(new char[1] { '/' });
			int num2 = ConvertHelpers.ConvertToInt(array[0]);
			if (num > 2018 || (num == 2018 && num2 >= 8))
			{
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = true;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("I") - 1].Value = "0";
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("R") - 1].ResetFormula();
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("R") - 1].Value = "0";
			}
		}
		GeneralFunctions.fpSpReindexingSheetTemplate2(sView, ToKhaiInfo.SheetInfos[1]);
	}

	public void AutoBindData_2(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ButtonCellType buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		buttonCellType.Text = sView.Cells[16, HTKKFile.ColumnLettersToIndex("C") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		sView.Cells[16, HTKKFile.ColumnLettersToIndex("C") - 1].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[17, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[17, HTKKFile.ColumnLettersToIndex("G") - 1].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[17, HTKKFile.ColumnLettersToIndex("K") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[17, HTKKFile.ColumnLettersToIndex("K") - 1].CellType = buttonCellType;
		if (sView.Cells[0, HTKKFile.ColumnLettersToIndex("C") - 1].CellType.GetType().Name == "ButtonCellType")
		{
			buttonCellType = (ButtonCellType)sView.Cells[0, HTKKFile.ColumnLettersToIndex("C") - 1].CellType;
			string text = buttonCellType.Text;
			if (!text.Equals(""))
			{
				buttonCellType = new ButtonCellType();
				buttonCellType.TwoState = false;
				buttonCellType.Text = text;
				buttonCellType.ShadowSize = 0;
				buttonCellType.TextColor = Color.Blue;
				sView.Cells[1, HTKKFile.ColumnLettersToIndex("C") - 1].CellType = buttonCellType;
				buttonCellType = new ButtonCellType();
				buttonCellType.TwoState = false;
				buttonCellType.ShadowSize = 0;
				sView.Cells[0, HTKKFile.ColumnLettersToIndex("C") - 1].CellType = buttonCellType;
				sView.Cells[0, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
			}
		}
		int count = ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count;
		for (int i = 0; i < count; i++)
		{
			int row = ToKhaiInfo.SheetInfos[2].Tables[0].Location.DataLocationTemplateRowStart + i - 1;
			int num = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			string[] array = PeriodInfo.TuThang.Split(new char[1] { '/' });
			int num2 = ConvertHelpers.ConvertToInt(array[0]);
			if (num > 2018 || (num == 2018 && num2 >= 8))
			{
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = true;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("I") - 1].Value = "0";
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].ResetFormula();
				sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Value = "0";
			}
		}
		GeneralFunctions.fpSpReindexingSheetTemplate2(sView, ToKhaiInfo.SheetInfos[2]);
	}

	public void AutoBindData_3(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			List<HTKKCMCRowValue> rowsValues = ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues;
			int column = GeneralFunctions.ArrayCellId("S_22")[0];
			int column2 = GeneralFunctions.ArrayCellId("U_22")[0];
			int column3 = GeneralFunctions.ArrayCellId("W_22")[0];
			for (int i = 0; i < rowsValues.Count; i++)
			{
				int row = ToKhaiInfo.SheetInfos[3].Tables[0].Location.DataLocationTemplateRowStart + i - 1;
				if (ConvertHelpers.ConvertToString(sView.Cells[row, HTKKFile.ColumnLettersToIndex("Q") - 1].Value) == "Khác")
				{
					sView.Cells[row, column].Locked = true;
					sView.Cells[row, column2].Locked = true;
					sView.Cells[row, column3].Locked = true;
				}
			}
			GeneralFunctions.fpSpReindexingSheetTemplate2(sView, ToKhaiInfo.SheetInfos[3]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			int num = HTKKFile.ColumnLettersToIndex("Q");
			int num2 = HTKKFile.ColumnLettersToIndex("S");
			int num3 = HTKKFile.ColumnLettersToIndex("U");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			if (iCol == num)
			{
				WinInsertCombox(sheetView, (int)iRow, num2, HTKKWinDictionaryDiaDanh.TinhTP);
				WinInsertCombox(sheetView, (int)iRow, num3, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP["ALL"]);
				WinInsertCombox(sheetView, (int)iRow, num4, HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen["ALL"]);
				string value = "";
				HTKKWinDictionaryDiaDanh.QuocTich.TryGetValue(sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text, out value);
				if (value == "02")
				{
					HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "QuocTich", "True");
					sheetView.Cells[(int)iRow - 1, num2 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num4 - 1].Locked = true;
					sheetView.Cells[(int)iRow - 1, num2 - 1].Value = "Khác";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "Khác";
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = "Khác";
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = "999";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "99999";
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = "9999999";
				}
				else
				{
					HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "QuocTich", "False");
					sheetView.Cells[(int)iRow - 1, num2 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num3 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num4 - 1].Locked = false;
					sheetView.Cells[(int)iRow - 1, num2 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num4 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = "";
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

	private void resetCT22_TK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[CellID_Ct22[1], CellID_Ct22[0]].Value = 0;
			sheetView.Cells[CellID_Ct38[1], CellID_Ct38[0]].ResetFormula();
			sheetView.Cells[CellID_Ct38[1], CellID_Ct38[0]].Value = sheetView.Cells[CellID_Ct38_PL[1], CellID_Ct38_PL[0]].Value;
			int num = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => !HTKKFormater.isNullString(p.RowValue[1]));
			int num2 = ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count((HTKKCMCRowValue p) => !HTKKFormater.isNullString(p.RowValue[0]));
			winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("I_36")[1], GeneralFunctions.ArrayCellId("I_36")[0]].Value = num + num2;
			winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("I_37")[1], GeneralFunctions.ArrayCellId("I_37")[0]].Value = num;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			for (int i = 0; i < ToKhaiInfo.SheetInfos.Count - 1; i++)
			{
				GeneralFunctions.fpSpReindexingSheetTemplate2(winFpSpread.Sheets[i], ToKhaiInfo.SheetInfos[i]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
