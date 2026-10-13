using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Mbai01Nckt225
{
	private string cellId_CT22_ct4 = "R_45";

	private string cellId_CT22_ct5 = "Y_45";

	private int[] CellId_CT22_ct4 => GeneralFunctions.ArrayCellId(cellId_CT22_ct4);

	private int[] CellId_CT22_ct5 => GeneralFunctions.ArrayCellId(cellId_CT22_ct5);

	private int idxCT23_ColC => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCT23_ColR => HTKKFile.ColumnLettersToIndex("R") - 1;

	private int idxCT23_ColY => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public Dictionary<string, string> GetDMMucLePhiMonBaiCT22(string pathFile)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		try
		{
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 0)
				{
					string key = array[0];
					string value2 = array[1];
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, value2);
					}
				}
			}
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMMucLePhiMonBaiCT23(string pathFile)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		try
		{
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 0)
				{
					string key = array[0];
					string value2 = array[1];
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, value2);
					}
				}
			}
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("R");
			int num2 = 20;
			if (PeriodInfo.dicVar.Count <= 0)
			{
				return;
			}
			bool flag = Convert.ToBoolean(PeriodInfo.dicVar["CoSoMoiThanhLap"]);
			string text = (string)PeriodInfo.dicVar["NgayThanhLap"];
			sheetView.Cells[num2 - 1, num - 1].Value = flag;
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayLapTKChinhThuc"));
			if (flag)
			{
				if (ConvertHelpers.ConvertToDouble1(PeriodInfo.SoLan) > 0.0)
				{
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("C") - 1].Formula = "";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "0";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("D") - 1].Formula = "";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "0";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text = "0";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("F") - 1].Text = "1";
				}
				else
				{
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("C") - 1].Formula = "";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "0";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("D") - 1].Formula = "";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "0";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text = "1";
					sheetView.Cells[31, HTKKFile.ColumnLettersToIndex("F") - 1].Text = "0";
				}
				sheetView.Cells[34, HTKKFile.ColumnLettersToIndex("K") - 1].Text = "x";
				sheetView.Cells[34, HTKKFile.ColumnLettersToIndex("M") - 1].Text = text;
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Locked = true;
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].BackColor = Color.FromArgb(233, 245, 254);
			}
			else
			{
				sheetView.Cells[34, HTKKFile.ColumnLettersToIndex("K") - 1].Text = "";
				sheetView.Cells[34, HTKKFile.ColumnLettersToIndex("M") - 1].Text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			if (activeSheetIndex == 0)
			{
				SheetView sViewTK = winFpSpread.Sheets[0];
				if (iCol - 1 == CellId_CT22_ct4[0] && iRow < dataLocationTemplateRowStart + rowsValues.Count)
				{
					XuLyCT22_ct5(sViewTK, (int)iRow - 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			int num = HTKKFile.ColumnLettersToIndex("G");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			double num2 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[CellId_CT22_ct4[1], CellId_CT22_ct4[0]].Value);
			double num3 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[CellId_CT22_ct5[1], CellId_CT22_ct5[0]].Value);
			bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[GeneralFunctions.ArrayCellId("R_20")[1], GeneralFunctions.ArrayCellId("R_20")[0]].Value);
			double num4 = 0.0;
			if (num2 > 1000000000.0)
			{
				num4 = 3000000.0;
			}
			else
			{
				num4 = 2000000.0;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num5 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num5 + i, idxCT23_ColC].Text;
				array[1] = sheetView.Cells[num5 + i, num - 1].Text;
				array[2] = sheetView.Cells[num5 + i, idxCT23_ColR].Text;
				array[3] = sheetView.Cells[num5 + i, idxCT23_ColY].Text;
				array[4] = sheetView.Cells[num5 + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToString();
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array) && !flag)
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColC);
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColY);
					HTKKFile.RemoveValidate(sheetView, num5 + i, num - 1);
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColR);
					continue;
				}
				if (!string.IsNullOrEmpty(array[0]))
				{
					if (!ValidateBase.CheckMaSoThue(array[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, idxCT23_ColC, message3);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColC);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColC);
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, num - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num5 + i, num - 1);
				}
				if (flag)
				{
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, num - 1, message);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num5 + i, num - 1);
					}
					if (array[4].Replace("0", "").IsNullOrEmpty())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, idxCT23_ColY, message);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColY);
					}
				}
				else if (!string.IsNullOrEmpty(array[0]) || !string.IsNullOrEmpty(array[1]) || !string.IsNullOrEmpty(array[2].Replace("0", "")))
				{
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, num - 1, message);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num5 + i, num - 1);
					}
					if (string.IsNullOrEmpty(array[4].Replace("0", "")))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + i, idxCT23_ColY, message);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num5 + i, idxCT23_ColY);
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayLapTKChinhThuc"));
			if (string.IsNullOrEmpty(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message);
				value = 2;
			}
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
		}
	}

	private void XuLyCT22_ct5(SheetView sViewTK, int iRow)
	{
		try
		{
			double num = Convert.ToDouble(sViewTK.Cells[iRow, CellId_CT22_ct4[0]].Value);
			double num2 = 10000000000.0;
			if (iRow > 44)
			{
				if (ConvertHelpers.ConvertToDouble1(sViewTK.Cells[iRow, CellId_CT22_ct5[0]].Value) == 0.0)
				{
					sViewTK.Cells[iRow, CellId_CT22_ct5[0]].Value = 1000000;
				}
			}
			else if (num > num2)
			{
				sViewTK.Cells[iRow, CellId_CT22_ct5[0]].Value = 3000000;
			}
			else
			{
				sViewTK.Cells[iRow, CellId_CT22_ct5[0]].Value = 2000000;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
