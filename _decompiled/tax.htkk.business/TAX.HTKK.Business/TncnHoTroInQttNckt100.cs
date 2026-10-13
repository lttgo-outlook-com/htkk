using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnHoTroInQttNckt100
{
	public static Dictionary<string, bool> lstValue = new Dictionary<string, bool>();

	private bool isGiamDan = true;

	private int[] CellID_CheckChon => GeneralFunctions.ArrayCellId("C_18");

	private int[] CellID_SapXepHoTen => GeneralFunctions.ArrayCellId("D_18");

	private int[] CellID_SapXepQuyetToan => GeneralFunctions.ArrayCellId("AH_18");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		ButtonCellType buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = true;
		buttonCellType.Text = sView.Cells[CellID_SapXepHoTen[1], CellID_SapXepHoTen[0]].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.WordWrap = true;
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[CellID_SapXepHoTen[1], CellID_SapXepHoTen[0]].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[CellID_SapXepQuyetToan[1], CellID_SapXepQuyetToan[0]].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\sort2.bmp");
		buttonCellType.WordWrap = true;
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[CellID_SapXepQuyetToan[1], CellID_SapXepQuyetToan[0]].CellType = buttonCellType;
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sView.Cells[CellID_CheckChon[1], CellID_CheckChon[0]].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\Check.bmp");
		buttonCellType.WordWrap = true;
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sView.Cells[CellID_CheckChon[1], CellID_CheckChon[0]].CellType = buttonCellType;
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		//IL_0114: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = 17;
			int num5 = 18;
			int num6 = 20;
			bool value = ConvertHelpers.ConvertToBool(winFpSpread.Sheets[0].Cells[idxRow, idxCol].Value);
			string key = $"{idxRow}_{idxCol}";
			if (lstValue.ContainsKey(key))
			{
				lstValue[key] = value;
			}
			else
			{
				lstValue.Add(key, value);
			}
			int dataLocationTemplateRowStart = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart;
			if (idxRow == CellID_CheckChon[1] && idxCol == CellID_CheckChon[0])
			{
				frmThongTinIn frmThongTinIn = new frmThongTinIn();
				((Form)frmThongTinIn).ShowDialog();
				int options = frmThongTinIn.GetOptions();
				string start = frmThongTinIn.GetStart();
				string end = frmThongTinIn.GetEnd();
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					for (int j = 0; j < rowsValues.Count; j++)
					{
						switch (options)
						{
						case 1:
							sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, idxCol].Value = "True";
							break;
						case 2:
							sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, idxCol].Value = "False";
							break;
						case 3:
							if (j + 1 >= ConvertHelpers.ConvertToInt(start) && j + 1 <= ConvertHelpers.ConvertToInt(end) && ConvertHelpers.ConvertToInt(start) > 0 && ConvertHelpers.ConvertToInt(end) <= rowsValues.Count)
							{
								sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, idxCol].Value = "True";
							}
							else
							{
								sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, idxCol].Value = "False";
							}
							break;
						}
					}
				}
			}
			if (idxRow == CellID_SapXepHoTen[1] && idxCol == CellID_SapXepHoTen[0])
			{
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					for (int j = 0; j < rowsValues.Count; j++)
					{
						for (int k = j + 1; k < rowsValues.Count; k++)
						{
							if (isGiamDan)
							{
								if (string.Compare(sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, CellID_SapXepHoTen[0]].Text, sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, CellID_SapXepHoTen[0]].Text) >= 0)
								{
									continue;
								}
								for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
								{
									HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
									if (HTKKFormater.isNullString(sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Locked)
									{
										sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula = "";
										sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Formula = "";
									}
									string text = sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Text;
									sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Text = sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Text;
									sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Text = text;
								}
							}
							else
							{
								if (string.Compare(sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, CellID_SapXepHoTen[0]].Text, sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, CellID_SapXepHoTen[0]].Text) <= 0)
								{
									continue;
								}
								for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
								{
									HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
									if (HTKKFormater.isNullString(sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Locked)
									{
										sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula = "";
										sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Formula = "";
									}
									string text = sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Text;
									sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Text = sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Text;
									sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Text = text;
								}
							}
						}
					}
				}
				isGiamDan = !isGiamDan;
			}
			if (idxRow != CellID_SapXepQuyetToan[1] || idxCol != CellID_SapXepQuyetToan[0])
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart2 = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					for (int k = j + 1; k < rowsValues.Count; k++)
					{
						if (isGiamDan)
						{
							if (string.Compare(sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, CellID_SapXepQuyetToan[0]].Text, sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, CellID_SapXepQuyetToan[0]].Text) >= 0)
							{
								continue;
							}
							for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
							{
								HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
								if (HTKKFormater.isNullString(sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Locked)
								{
									sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula = "";
									sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Formula = "";
								}
								object value2 = sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Value;
								sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Value;
								sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Value = value2;
							}
						}
						else
						{
							if (string.Compare(sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, CellID_SapXepQuyetToan[0]].Text, sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, CellID_SapXepQuyetToan[0]].Text) <= 0)
							{
								continue;
							}
							for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
							{
								HTKKCMCCellBasicInfo cellTemplate = hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate;
								if (HTKKFormater.isNullString(sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula) && !sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Locked)
								{
									sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Formula = "";
									sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Formula = "";
								}
								object value2 = sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Value;
								sheetView.Cells[dataLocationTemplateRowStart2 + k - 1, cellTemplate.ColIndex - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Value;
								sheetView.Cells[dataLocationTemplateRowStart2 + j - 1, cellTemplate.ColIndex - 1].Value = value2;
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

	public void SetFPSpreadExp(FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKHeaderInfo HTKKHeaderInfo)
	{
		try
		{
			for (int i = 0; i < winFpExp.Sheets.Count - 2; i++)
			{
				winFpExp.Sheets[i].Cells[12, HTKKFile.ColumnLettersToIndex("T") - 1].Formula = "";
				winFpExp.Sheets[i].Cells[12, HTKKFile.ColumnLettersToIndex("T") - 1].Text = PeriodInfo.Nam;
			}
			winFpExp.ActiveSheetIndex = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
			double num2 = 9000000.0;
			if (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year > 2019)
			{
				num2 = 11000000.0;
			}
			double num3 = 0.0;
			if (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year < 2020 || ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year >= 2020)
			{
				num3 = (double)num * num2;
			}
			else
			{
				int num4 = 0;
				if (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year == 2019)
				{
					num4 = 12 - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
				}
				int num5 = num - num4;
				num3 = num4 * 9000000 + num5 * 11000000;
			}
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				return;
			}
			List<Dictionary<string, string>> pl05_1Bk = getPl05_1Bk(PeriodInfo, ToKhaiInfo, HeaderInfo);
			if (pl05_1Bk == null || pl05_1Bk.Count == 0)
			{
				return;
			}
			ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Clear();
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			for (int i = 0; i < pl05_1Bk.Count; i++)
			{
				HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
				hTKKCMCRowValue.RowValue.Add("False");
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["D_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["E_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["F_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["H_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["I_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["J_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["K_22"]);
				hTKKCMCRowValue.RowValue.Add(ConvertHelpers.ConvertToString(num3));
				double num6 = ConvertHelpers.ConvertToDouble(pl05_1Bk[i]["L_22"]) - num3;
				if (num6 < 0.0)
				{
					num6 = 0.0;
				}
				hTKKCMCRowValue.RowValue.Add(ConvertHelpers.ConvertToString(num6));
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["M_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["N_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["O_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["P_22"]);
				double num7 = ConvertHelpers.ConvertToDouble(pl05_1Bk[i]["R_22"]) + ConvertHelpers.ConvertToDouble(pl05_1Bk[i]["T_22"]);
				hTKKCMCRowValue.RowValue.Add(ConvertHelpers.ConvertToString(num7));
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["Q_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["R_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["V_22"]);
				hTKKCMCRowValue.RowValue.Add(pl05_1Bk[i]["U_22"]);
				hTKKCMCRowValue.RowValue.Add((pl05_1Bk[i]["V_22"] == "0") ? "Không" : "Có");
				hTKKCMCRowValue.RowValue.Add(ConvertHelpers.ConvertToString(pl05_1Bk[i]["L_22"]));
				hTKKCMCTableInfo.RowsValues.Add(hTKKCMCRowValue);
			}
			ToKhaiInfo.SheetInfos[0].IncrementStaticCell(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.TableLength + hTKKCMCTableInfo.RowsValues.Count - 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private List<Dictionary<string, string>> getPl05_1Bk(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo)
	{
		try
		{
			string xMLFileName = GeneralFunctions.GetXMLFileName(HTKKConfig.DATAFILES + "\\" + headerInfo.mst, string.Format("05_1_QTT_TNCN_{0}_{1}_{2}", periodInfo.Nam, periodInfo.TuThang.Replace("/", ""), periodInfo.DenThang.Replace("/", "")), "242", "xml");
			int num = toKhaiInfo.DataFileNameSaveExtend.Split(new char[1] { '_' }).Length - 1;
			bool flag = toKhaiInfo.DataFileNameSaveExtend.Split(new char[1] { '_' })[num] == "00";
			List<string> list = new List<string>();
			list.Add("C_22");
			list.Add("D_22");
			list.Add("E_22");
			list.Add("F_22");
			list.Add("G_22");
			list.Add("H_22");
			list.Add("I_22");
			list.Add("J_22");
			list.Add("K_22");
			list.Add("L_22");
			list.Add("M_22");
			list.Add("N_22");
			list.Add("O_22");
			list.Add("P_22");
			list.Add("Q_22");
			list.Add("R_22");
			list.Add("S_22");
			list.Add("T_22");
			list.Add("U_22");
			list.Add("V_22");
			List<string> lstCellId = list;
			return GeneralFunctions.GetRowDataValueInXml(xMLFileName, lstCellId);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int iwarningcolor, out object outPut)
	{
		outPut = false;
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				if (ConvertHelpers.ConvertToBool(array[0]))
				{
					outPut = true;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			outPut = false;
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
