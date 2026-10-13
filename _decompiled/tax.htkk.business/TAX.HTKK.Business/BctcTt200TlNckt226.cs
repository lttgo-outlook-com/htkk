using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Xml;
using FarPoint.Win;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt200TlNckt226
{
	private int rowBeginLT = 53;

	private int rowEndLT = 171;

	private int rowBeginKLT = 172;

	private int rowEndKLT = 249;

	private int rowBeginLTExp = 54;

	private int rowEndLTExp = 171;

	private int rowBeginKLTExp = 172;

	private int rowEndKLTExp = 249;

	private HTKKCMCToKhaiInfo tKhaiInfoOld = new HTKKCMCToKhaiInfo();

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

	public void ExportXML_1(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("PL_KQHDSXKD/PhanKy/ngayLap");
			if (val != null)
			{
				string innerText = val.InnerText;
				if (innerText.Length >= 8)
				{
					val.InnerText = innerText.Substring(4, 4) + "-" + innerText.Substring(2, 1).PadLeft(2, '0') + "-" + innerText.Substring(0, 1).PadLeft(2, '0');
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_2(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
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

	public void ExportXML_3(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
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

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			int column = HTKKFile.ColumnLettersToIndex("R") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("W") - 1;
			int[] array = new int[2] { 152, 153 };
			double num = double.Parse(sheetView.Cells[array[0] - 1, column].Value.ToString());
			double num2 = double.Parse(sheetView.Cells[array[1] - 1, column].Value.ToString());
			if (num + num2 == 0.0)
			{
				sheetView.Cells[array[0] - 2, column].Locked = false;
			}
			double num3 = double.Parse(sheetView.Cells[array[0] - 1, column2].Value.ToString());
			double num4 = double.Parse(sheetView.Cells[array[1] - 1, column2].Value.ToString());
			if (num3 + num4 == 0.0)
			{
				sheetView.Cells[array[0] - 2, column2].Locked = false;
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("B_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("C_38");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
				else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else if (value == 0)
				{
					winFpSpread.TabStrip[key].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 0;
			SheetView sView = winFpSpread.Sheets[num];
			int num2 = 0;
			int num3 = HTKKFile.ColumnLettersToIndex("AC");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			int num5 = HTKKFile.ColumnLettersToIndex("O");
			int num6 = 38;
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int[] array = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, Convert.ToBoolean(valueCustom));
			if (num2 == 2 && winFpSpread.TabStrip[num].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
			}
			else if (num2 == 1 && winFpSpread.TabStrip[num].BackColor == Color.White)
			{
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
			}
			else if (num2 == 0)
			{
				winFpSpread.TabStrip[num].BackColor = Color.White;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 1;
			SheetView sView = winFpSpread.Sheets[num];
			int[] array = GeneralFunctions.ArrayCellId("B_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 2;
			SheetView sView = winFpSpread.Sheets[num];
			int[] array = GeneralFunctions.ArrayCellId("C_38");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("B_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("C_38");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("R") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("W") - 1;
			int[] array = new int[2];
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			string text2 = "R_62;W_62;R_71;W_71;R_75;W_75;R_90;W_90;R_94;W_94;R_97;W_97;R_100;W_100;R_103;W_103;R_111;W_111;R_157;W_157";
			string text3 = "I_40;J_40;I_41;J_41";
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text2.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "I###+O###";
			if ((int)iCol == HTKKFile.ColumnLettersToIndex("I"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Formula = text.Replace("###", iRow.ToString());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "R###+AN###";
			if ((int)iCol == HTKKFile.ColumnLettersToIndex("R"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AB") - 1].Formula = text.Replace("###", iRow.ToString());
			}
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isReadDataFile, bool isCheckLayDl, bool callIsCheck)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			string text2 = pathDataFileOld[1];
			string text3 = pathDataFileOld[2];
			string text4 = text3;
			string text5 = "";
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			string text6 = "";
			if (sheetIndex == 0)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AI");
			}
			if (sheetIndex == 1)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("I");
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("N");
				num4 = HTKKFile.ColumnLettersToIndex("K");
				num5 = HTKKFile.ColumnLettersToIndex("O");
				num6 = HTKKFile.ColumnLettersToIndex("P");
				num7 = HTKKFile.ColumnLettersToIndex("L");
				num8 = HTKKFile.ColumnLettersToIndex("Q");
				text6 = "I###+O###";
			}
			if (sheetIndex == 2)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AM");
				num4 = HTKKFile.ColumnLettersToIndex("AB");
				num5 = HTKKFile.ColumnLettersToIndex("AN");
				num6 = HTKKFile.ColumnLettersToIndex("AO");
				num7 = HTKKFile.ColumnLettersToIndex("AG");
				num8 = HTKKFile.ColumnLettersToIndex("AP");
				text6 = "R###+AN###";
			}
			if (File.Exists(text4) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text4);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex != 118)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
			}
			else
			{
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
					}
				}
			}
			if (File.Exists(text))
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if ((sheetIndex == 1 || sheetIndex == 2) && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num4)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (sheetIndex == 1 && ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num4 && !isReadDataFile)
				{
					sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = 0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetFPSpreadExp(FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKHeaderInfo HTKKHeaderInfo)
	{
		try
		{
			int num = 0;
			for (num = 0; num < winFpExp.Sheets.Count - 1; num++)
			{
				SheetView sheetView = winFpExp.Sheets[num];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = sheetInfos[num];
				if (num == 0 && PeriodInfo.DonViTinhMa == "VND")
				{
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("L") - 1].Text = "";
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].Text = "";
					Color color = default(Color);
					color = Color.FromArgb(240, 240, 240);
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].BackColor = color;
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].ForeColor = color;
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].Border = new BevelBorder(BevelBorderType.Raised, color, color);
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("P") - 1].Border = new BevelBorder(BevelBorderType.Raised, color, color);
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("Q") - 1].Border = new BevelBorder(BevelBorderType.Raised, color, color);
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("R") - 1].Border = new BevelBorder(BevelBorderType.Raised, color, color);
					sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				}
			}
			winFpExp.ActiveSheetIndex = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private List<string> getPathDataFileOld(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex)
	{
		List<string> list = new List<string>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			string item = text;
			string item2 = text;
			string text2 = "";
			string text3 = "";
			string text4 = "";
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text3 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text4 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			string text5 = HTKKConfig.DATAFILES + "\\" + HeaderInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text5);
			if (sheetIndex == 0)
			{
				string value = ToKhaiInfo.SheetInfos[sheetIndex].DataFile.Replace("_DD", "").Replace("_TL", "") + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_";
				if (directoryInfo.Exists)
				{
					FileInfo[] files = directoryInfo.GetFiles();
					if (files.Length > 0)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(value))
							{
								item2 = text5 + "\\" + fileInfo.Name;
								break;
							}
						}
					}
				}
			}
			if (sheetIndex == 1 || sheetIndex == 2)
			{
				text = text + "_Q0" + PeriodInfo.Quy + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text3 + "_" + text4 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				int num = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
				int month = GeneralFunctions.GetNgayDauQuy(num, num2, 1, 1).Month;
				Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(month, num2, HeaderInfo.NgayBatDauNamTaiChinh);
				if (num == 1)
				{
					num = 4;
					num2--;
				}
				else
				{
					num--;
				}
				string value2 = "";
				quyNamByThangNam.TryGetValue(num.ToString(), out value2);
				if (num2 == ConvertHelpers.ConvertToInt(value2))
				{
					string value3 = ToKhaiInfo.SheetInfos[sheetIndex].DataFile + "_Q0" + num + value2;
					string value4 = ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
					if (directoryInfo.Exists)
					{
						FileInfo[] files = directoryInfo.GetFiles();
						if (files.Length > 0)
						{
							FileInfo[] array = files;
							foreach (FileInfo fileInfo in array)
							{
								if (fileInfo.Name.Contains(value3) && fileInfo.Name.Contains(value4))
								{
									item = text5 + "\\" + fileInfo.Name;
									break;
								}
							}
						}
					}
				}
			}
			list.Add(text);
			list.Add(item);
			list.Add(item2);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	private void checkSoNamTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheck)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string path = pathDataFileOld[2];
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			if (sheetIndex == 0)
			{
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AI");
			}
			if (sheetIndex == 1)
			{
				path = pathDataFileOld[0];
				num2 = HTKKFile.ColumnLettersToIndex("I");
				num3 = HTKKFile.ColumnLettersToIndex("J");
				num4 = HTKKFile.ColumnLettersToIndex("N");
				num5 = HTKKFile.ColumnLettersToIndex("K");
				num6 = HTKKFile.ColumnLettersToIndex("O");
				num7 = HTKKFile.ColumnLettersToIndex("P");
				num8 = HTKKFile.ColumnLettersToIndex("L");
				num9 = HTKKFile.ColumnLettersToIndex("Q");
			}
			if (sheetIndex == 2)
			{
				path = pathDataFileOld[0];
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AM");
				num5 = HTKKFile.ColumnLettersToIndex("AB");
				num6 = HTKKFile.ColumnLettersToIndex("AN");
				num7 = HTKKFile.ColumnLettersToIndex("AO");
				num8 = HTKKFile.ColumnLettersToIndex("AG");
				num9 = HTKKFile.ColumnLettersToIndex("AP");
			}
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				string valueCustom;
				object value;
				if (File.Exists(path) && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num3 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
				{
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
					}
				}
				if (File.Exists(pathDataFileOld[0]) && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num8 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Locked)
				{
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num8, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1, text);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1);
					}
				}
				if (!File.Exists(pathDataFileOld[1]) || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num5 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
				{
					continue;
				}
				valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num5, "NumberCellType");
				value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value;
				if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
				{
					if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].BackColor != Color.Pink)
					{
						num = 1;
						HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1, text2);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1);
				}
			}
			if (num == 2 && winFpSpread.TabStrip[sheetIndex].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[sheetIndex].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[sheetIndex].BackColor == Color.White)
			{
				winFpSpread.TabStrip[sheetIndex].BackColor = Color.Yellow;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
