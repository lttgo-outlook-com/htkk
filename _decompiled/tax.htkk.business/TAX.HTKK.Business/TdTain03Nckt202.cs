using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TdTain03Nckt202
{
	private string cellId_Tong = "V_46";

	private int[] CellId_Tong => GeneralFunctions.ArrayCellId(cellId_Tong);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("F_25");
			int[] array2 = GeneralFunctions.ArrayCellId("H_24");
			int[] array3 = GeneralFunctions.ArrayCellId("Z_24");
			string text = ConvertHelpers.ConvertToString(sView.Cells[array3[1], array3[0]].Value);
			string text2 = ConvertHelpers.ConvertToString(sView.Cells[array2[1], array2[0]].Value);
			if (string.IsNullOrEmpty(text))
			{
				string text3 = "";
				int month = 0;
				if (PeriodInfo.KieuKy == "M")
				{
					month = int.Parse(PeriodInfo.Thang);
				}
				else if (PeriodInfo.KieuKy == "Q")
				{
					month = int.Parse(PeriodInfo.Quy) * 3 - 2;
				}
				else if (PeriodInfo.KieuKy == "Y")
				{
					month = 12;
				}
				int year = int.Parse(PeriodInfo.Nam);
				DateTime dtPeriod = new DateTime(year, month, 1);
				if (isGetDmMoi(dtPeriod, PeriodInfo))
				{
					sView.Cells[array2[1], array2[0]].Value = "Nước thiên nhiên dùng cho sản xuất thủy điện";
					sView.Cells[array3[1], array3[0]].Value = "V5";
					sView.Cells[array[1], array[0]].Value = GetThueSuatByLoaiTN(PeriodInfo, "Nước thiên nhiên dùng cho sản xuất thủy điện");
				}
				else
				{
					sView.Cells[array2[1], array2[0]].Value = "Nước thiên nhiên dùng cho sản xuất thủy điện";
					sView.Cells[array3[1], array3[0]].Value = "0702";
					sView.Cells[array[1], array[0]].Value = GetThueSuatByLoaiTN(PeriodInfo, "Nước thiên nhiên dùng cho sản xuất thủy điện");
				}
			}
			else
			{
				sView.Cells[array2[1], array2[0]].Value = GetTenTaiNguyenByMaLoaiTN(PeriodInfo, text);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					LayThemThongTinCQTChiCuc(sView, dataLocationTemplateRowStart + j);
					LayThemThongTinTinhHuyen_PL(sView, dataLocationTemplateRowStart + j);
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
				}
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
			int[] array = GeneralFunctions.ArrayCellId("F_25");
			int[] array2 = GeneralFunctions.ArrayCellId("H_24");
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if ((activeColumnIndex == array2[0]) & (activeRowIndex == array2[1]))
			{
				string loaiTN = winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Value.ToString();
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Value = GetThueSuatByLoaiTN(PeriodInfo, loaiTN);
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
			string text = string.Format("ROUND(I{0}*K{0}*$E$25/100,0)", iRow);
			int[] cellChangeID = new int[2]
			{
				(int)iCol - 1,
				(int)iRow - 1
			};
			List<int[]> list = new List<int[]>();
			list.Add(GeneralFunctions.ArrayCellId("H_24"));
			list.Add(GeneralFunctions.ArrayCellId("I_" + iRow));
			list.Add(GeneralFunctions.ArrayCellId("K_" + iRow));
			List<int[]> list2 = list;
			XuLyCellChange(winFpSpread.Sheets[0], cellChangeID, list2[0], list2[1], list2[2]);
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
			if (iCol == HTKKFile.ColumnLettersToIndex("U"))
			{
				double num = 0.0;
				bool flag = false;
				int num2 = -1;
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					if (iRow == hTKKCMCTableInfo.Header.StaticCells[5].CellTemplate.RowIndex)
					{
						num2 = i;
						flag = true;
					}
					num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
				}
				if (flag)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_03_1_TD_TAIN/tongCong"));
					winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2];
					int num3 = ((ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count - 1 > num2) ? (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2 + 1].Location.DataLocationTemplateRowStart - 1) : (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].RowsValues.Count));
					for (int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num4 < num3; num4++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num6 * (num5 / 100.0);
					}
				}
			}
			else if (iCol == HTKKFile.ColumnLettersToIndex("R"))
			{
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					double num7 = 0.0;
					int num2 = -1;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					int num3 = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].RowsValues.Count;
					if (iRow < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].Location.DataLocationTemplateRowStart || iRow > num3)
					{
						continue;
					}
					for (int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num4 < num3; num4++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						num7 += num5;
						if (iRow == num4)
						{
							num2 = i;
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num6 * (num5 / 100.0);
						}
					}
					if (num2 > -1)
					{
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].Header.StaticCells[4].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Value = num7;
					}
				}
			}
			int num8 = HTKKFile.ColumnLettersToIndex("P");
			int num9 = HTKKFile.ColumnLettersToIndex("AJ");
			if (iCol == num8 || iCol == num9)
			{
				LayThemThongTinCQTChiCuc(winFpSpread.Sheets[1], (int)iRow);
				LayThemThongTinTinhHuyen_PL(winFpSpread.Sheets[1], (int)iRow);
			}
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sView = winFpSpread.Sheets[1];
			if (iCol == HTKKFile.ColumnLettersToIndex("C") || iCol == HTKKFile.ColumnLettersToIndex("H") || iCol == HTKKFile.ColumnLettersToIndex("P"))
			{
				int tableIndex = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
				GeneralFunctions.SetLaiDongCon(sView, sInfo, tableIndex, -1, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
				GeneralFunctions.SetLaiDongCon(sView, sInfo, tableIndex, -1, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			double num = 0.0;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
				num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_03_1_TD_TAIN/tongCong"));
			winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void XuLyCellChange(SheetView sView, int[] cellChangeID, int[] ct22, int[] ct25, int[] ct26)
	{
		try
		{
			string formula = string.Format("ROUND(I{0}*K{0}*$F$25/100,0)", (cellChangeID[1] + 1).ToString());
			int column = HTKKFile.ColumnLettersToIndex("O") - 1;
			if (ct22[0] == cellChangeID[0] && ct22[1] == cellChangeID[1])
			{
				sView.Cells[cellChangeID[1], column].Formula = formula;
			}
			else if (ct25[0] == cellChangeID[0] && ct25[1] == cellChangeID[1])
			{
				sView.Cells[cellChangeID[1], column].Formula = formula;
			}
			else if (ct26[0] == cellChangeID[0] && ct26[1] == cellChangeID[1])
			{
				sView.Cells[cellChangeID[1], column].Formula = formula;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public string GetThueSuatByLoaiTN(HTKKPeriodInfo PeriodInfo, string loaiTN)
	{
		//IL_02bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c2: Expected O, but got Unknown
		//IL_02ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f3: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = "";
			int month = 0;
			if (PeriodInfo.KieuKy == "M")
			{
				month = int.Parse(PeriodInfo.Thang);
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				month = int.Parse(PeriodInfo.Quy) * 3 - 2;
			}
			else if (PeriodInfo.KieuKy == "Y")
			{
				month = 12;
			}
			int year = int.Parse(PeriodInfo.Nam);
			DateTime dtPeriod = new DateTime(year, month, 1);
			if (PeriodInfo.KieuKy == "Y")
			{
				string nam = PeriodInfo.Nam;
				text = (nam.Equals("2008") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_052008_01", "CatalogueQTTN_2008_01") : (nam.Equals("2010") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072010_01", "CatalogueQTTN_2010_01") : (nam.Equals("2014") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_022014_01", "CatalogueQTTN_2014_01") : ((!nam.Equals("2016")) ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0] : GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072016_01", "CatalogueQTTN_2016_01")))));
			}
			else
			{
				text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
			}
			if (isGetDmMoi(dtPeriod, PeriodInfo))
			{
				Serializer serializer = new Serializer();
				string input = File.ReadAllText(text);
				danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
				if (danhmuc != null && danhmuc.items.Count > 0)
				{
					List<item> items = danhmuc.items;
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
						string text2 = array[0];
						string text3 = array[1];
						string result = array[4];
						if (text3.Equals("V5"))
						{
							return result;
						}
					}
				}
			}
			else
			{
				XmlDocument val = new XmlDocument();
				val.Load(text);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item2 in elementsByTagName)
				{
					XmlNode val2 = item2;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					string text3 = array[4];
					string result = array[2];
					if (result.Equals(loaiTN))
					{
						return text3;
					}
				}
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public string GetTenTaiNguyenByMaLoaiTN(HTKKPeriodInfo PeriodInfo, string maLoaiTN)
	{
		try
		{
			Dictionary<string, string> dMTaiNguyen = GetDMTaiNguyen(PeriodInfo);
			foreach (string key in dMTaiNguyen.Keys)
			{
				if (key.Equals(maLoaiTN))
				{
					return dMTaiNguyen[key];
				}
			}
			return string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			checkTKChinh(winFpSpread.Sheets[0], winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			if (winFpSpread.Sheets[1] != null && winFpSpread.Sheets[1].Visible)
			{
				checkPL3_1(winFpSpread.Sheets[1], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void checkTKChinh(SheetView sView, SheetView sViewPL, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int colunmIndexByColumName = HTKKFormater.GetColunmIndexByColumName("C");
			int colunmIndexByColumName2 = HTKKFormater.GetColunmIndexByColumName("F");
			int colunmIndexByColumName3 = HTKKFormater.GetColunmIndexByColumName("I");
			int colunmIndexByColumName4 = HTKKFormater.GetColunmIndexByColumName("K");
			int colunmIndexByColumName5 = HTKKFormater.GetColunmIndexByColumName("V");
			string[] array = new string[5];
			List<string> list = new List<string>();
			List<string> list2 = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertToString(sView.Cells[num + i, colunmIndexByColumName].Value);
				array[1] = ConvertHelpers.ConvertToString(sView.Cells[num + i, colunmIndexByColumName2].Value);
				array[2] = ConvertHelpers.ConvertToString(sView.Cells[num + i, colunmIndexByColumName3].Value);
				array[3] = ConvertHelpers.ConvertToString(sView.Cells[num + i, colunmIndexByColumName4].Value);
				array[4] = ConvertHelpers.ConvertToString(sView.Cells[num + i, colunmIndexByColumName5].Value);
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName2);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName3);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName4);
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName5);
					continue;
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName);
				}
				if (!ValidateBase.CheckMaSoThue(array[1]) && !string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sView, num + i, colunmIndexByColumName2, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, num + i, colunmIndexByColumName2);
				}
			}
			if (!sViewPL.Visible)
			{
				HTKKFile.RemoveValidate(sView, CellId_Tong[1], CellId_Tong[0]);
				if (value == 1)
				{
					value = 0;
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void checkPL3_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			if (tables.Count > 0)
			{
				HTKKCMCStaticOfTable header = tables[0].Header;
				string[] array = new string[8];
				string defaultValue = tables[0].Header.StaticCells[0].DefaultValue;
				string defaultValue2 = tables[0].Header.StaticCells[3].DefaultValue;
				string defaultValue3 = tables[0].Header.StaticCells[6].DefaultValue;
				string defaultValue4 = tables[0].Header.StaticCells[7].DefaultValue;
				array[0] = defaultValue;
				array[3] = defaultValue2;
				array[6] = defaultValue3;
				array[7] = defaultValue4;
				string[] array2 = new string[8];
				string defaultValue5 = tables[0].RowsInfo.Cells[0].DefaultValue;
				string defaultValue6 = tables[0].RowsInfo.Cells[6].DefaultValue;
				string defaultValue7 = tables[0].RowsInfo.Cells[7].DefaultValue;
				array2[0] = defaultValue5;
				array2[6] = defaultValue6;
				array2[7] = defaultValue7;
				HTKKCMCStaticOfTable hTKKCMCStaticOfTable = new HTKKCMCStaticOfTable();
				HTKKCMCRowInfo hTKKCMCRowInfo = new HTKKCMCRowInfo();
				int num = 0;
				for (int i = 0; i < tables.Count; i++)
				{
					hTKKCMCStaticOfTable = tables[i].Header;
					hTKKCMCRowInfo = tables[i].RowsInfo;
					num = tables[i].Location.DataLocationTemplateRowStart - 1;
					array[1] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1].Value);
					array[2] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1].Value);
					array[4] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.ColIndex - 1].Value);
					array[5] = ConvertHelpers.ConvertToString(sView.Cells[hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1].Value);
					bool flag = GeneralFunctions.HaveToValidateHeaderRow(header, array);
					bool flag2 = false;
					double num2 = 0.0;
					for (int j = 0; j < tables[i].RowsValues.Count; j++)
					{
						array2[1] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Value);
						array2[2] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1].Value);
						array2[3] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1].Value);
						array2[4] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1].Value);
						array2[5] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1].Value);
						array2[6] = ConvertHelpers.ConvertToString(sView.Cells[num + j, hTKKCMCRowInfo.Cells[6].CellTemplate.ColIndex - 1].Value);
						bool visible = sView.Cells[num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Row.Visible;
						if (GeneralFunctions.HaveToValidateRow(hTKKCMCRowInfo, array2) && visible)
						{
							flag2 = true;
						}
						else
						{
							if (!flag)
							{
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidateBackGround(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
								continue;
							}
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidateBackGround(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
						if (!string.IsNullOrEmpty(array2[2]))
						{
							if (!ValidateBase.CheckMaSoThue(array2[2]))
							{
								HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
						}
						if (string.IsNullOrEmpty(array2[3]))
						{
							HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
						}
						num2 += ConvertHelpers.ConvertToDouble(array2[4]);
						if (ConvertHelpers.ConvertToDouble(array2[5]) == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidateBackGround(sView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
					}
					if (!flag2 && !flag)
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
						continue;
					}
					if (num2 != 100.0)
					{
						for (int j = 0; j < tables[i].RowsValues.Count; j++)
						{
							HTKKFile.SetNodeBackGroundPink(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
					}
					else
					{
						for (int j = 0; j < tables[i].RowsValues.Count; j++)
						{
							HTKKFile.RemoveValidate(sView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
						}
					}
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
					}
					if (!string.IsNullOrEmpty(array[2]))
					{
						if (!ValidateBase.CheckMaSoThue(array[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
					}
					if (ConvertHelpers.ConvertToDouble(array[5]) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
					}
				}
			}
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void SupportImportXML_1(HTKKCMCToKhaiInfo ToKhaiInfo, XmlNode xmlSheet)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[4];
			if (hTKKCMCTableInfo == null)
			{
				return;
			}
			hTKKCMCTableInfo.RowsInfo.RowDefaultDisplay = "0";
			hTKKCMCTableInfo.RowsValues[0].RowDisplay = "0";
			XmlNodeList val = xmlSheet.SelectNodes(hTKKCMCTableInfo.TablePath.Remove(0, hTKKCMCTableInfo.TablePath.IndexOf('/') + 1));
			for (int i = 0; i < val.Count; i++)
			{
				for (int j = 0; j < hTKKCMCTableInfo.RowsInfo.Cells.Count; j++)
				{
					XmlNode val2 = val[i].SelectSingleNode(hTKKCMCTableInfo.RowsInfo.Cells[j].Path);
					if (val2 != null)
					{
						val2.InnerText = hTKKCMCTableInfo.RowsInfo.Cells[j].DefaultValue;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void SupportImportXML_2(HTKKCMCToKhaiInfo ToKhaiInfo, XmlNode xmlSheet)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			if (hTKKCMCTableInfo != null)
			{
				hTKKCMCTableInfo.RowsInfo.RowDefaultDisplay = "0";
				hTKKCMCTableInfo.RowsValues[0].RowDisplay = "0";
				XmlNodeList val = xmlSheet.SelectNodes(hTKKCMCTableInfo.TablePath.Remove(0, hTKKCMCTableInfo.TablePath.IndexOf('/') + 1));
				for (int i = 0; i < val.Count; i++)
				{
					for (int j = 0; j < hTKKCMCTableInfo.RowsInfo.Cells.Count; j++)
					{
						XmlNode val2 = val[i].SelectSingleNode(hTKKCMCTableInfo.RowsInfo.Cells[j].Path);
						if (val2 != null)
						{
							val2.InnerText = hTKKCMCTableInfo.RowsInfo.Cells[j].DefaultValue;
						}
					}
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[4];
			if (hTKKCMCTableInfo2 == null)
			{
				return;
			}
			hTKKCMCTableInfo2.RowsInfo.RowDefaultDisplay = "0";
			hTKKCMCTableInfo2.RowsValues[0].RowDisplay = "0";
			XmlNodeList val3 = xmlSheet.SelectNodes(hTKKCMCTableInfo2.TablePath.Remove(0, hTKKCMCTableInfo2.TablePath.IndexOf('/') + 1));
			for (int i = 0; i < val3.Count; i++)
			{
				for (int j = 0; j < hTKKCMCTableInfo2.RowsInfo.Cells.Count; j++)
				{
					XmlNode val2 = val3[i].SelectSingleNode(hTKKCMCTableInfo2.RowsInfo.Cells[j].Path);
					if (val2 != null)
					{
						val2.InnerText = hTKKCMCTableInfo2.RowsInfo.Cells[j].DefaultValue;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private Dictionary<string, string> GetDMTaiNguyen(HTKKPeriodInfo PeriodInfo)
	{
		//IL_02be: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c5: Expected O, but got Unknown
		//IL_02f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f9: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = "";
			int month = 0;
			if (PeriodInfo.KieuKy == "M")
			{
				month = int.Parse(PeriodInfo.Thang);
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				month = int.Parse(PeriodInfo.Quy) * 3 - 2;
			}
			else if (PeriodInfo.KieuKy == "Y")
			{
				month = 12;
			}
			int year = int.Parse(PeriodInfo.Nam);
			DateTime dtPeriod = new DateTime(year, month, 1);
			if (PeriodInfo.KieuKy == "Y")
			{
				string nam = PeriodInfo.Nam;
				text = (nam.Equals("2008") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_052008_01", "CatalogueQTTN_2008_01") : (nam.Equals("2010") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072010_01", "CatalogueQTTN_2010_01") : (nam.Equals("2014") ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_022014_01", "CatalogueQTTN_2014_01") : ((!nam.Equals("2016")) ? GeneralFunctions.GetCatalogueFileName(dtPeriod)[0] : GeneralFunctions.GetCatalogueFileName(dtPeriod)[0].Replace("Catalogue_072016_01", "CatalogueQTTN_2016_01")))));
			}
			else
			{
				text = GeneralFunctions.GetCatalogueFileName(dtPeriod)[0];
			}
			if (isGetDmMoi(dtPeriod, PeriodInfo))
			{
				Serializer serializer = new Serializer();
				string input = File.ReadAllText(text);
				danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
				if (danhmuc != null && danhmuc.items.Count > 0)
				{
					List<item> items = danhmuc.items;
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
						string text2 = array[0];
						string text3 = array[1];
						string value = array[2];
						if (text3.Equals("V5"))
						{
							dictionary.Add(text3, value);
						}
					}
				}
			}
			else
			{
				XmlDocument val = new XmlDocument();
				val.Load(text);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item2 in elementsByTagName)
				{
					XmlNode val2 = item2;
					string value2 = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value2, "###");
					string text2 = array[0];
					string text3 = array[1];
					string value = array[2];
					if (array[^1].Equals("1") && ((ConvertHelpers.ConvertToInt(text2) >= 404 && text3.Substring(0, 2).Equals("07")) || ConvertHelpers.ConvertToInt(text2) == 469 || (ConvertHelpers.ConvertToInt(text2) < 404 && text3.Substring(0, 2).Equals("09"))))
					{
						dictionary.Add(text3, value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("//tieuMucHachToan");
			if ((val != null) & !string.IsNullOrEmpty(PeriodInfo.TieuMucHachToanVal))
			{
				val.InnerText = PeriodInfo.TieuMucHachToanVal;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private bool isGetDmMoi(DateTime dtPeriod, HTKKPeriodInfo PeriodInfo)
	{
		bool result = false;
		DateTime dateTime = DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture);
		string denThang = PeriodInfo.DenThang;
		denThang = ((!HTKKFormater.isNullString(denThang)) ? ("01/" + denThang) : PeriodInfo.DenNgay);
		if (HTKKFormater.isNullString(denThang))
		{
			denThang = "01/01/" + PeriodInfo.Nam;
		}
		DateTime dateTime2 = DateTime.ParseExact(denThang, "dd/MM/yyyy", CultureInfo.CurrentCulture);
		if (PeriodInfo.KieuKy == "Y")
		{
			dateTime2 = new DateTime(dateTime2.Year, dateTime2.Month, 1).AddMonths(1).AddDays(-1.0);
			if (dtPeriod.Year >= dateTime.Year && dateTime2 >= dateTime)
			{
				return true;
			}
		}
		else if (PeriodInfo.KieuKy == "M")
		{
			if (dtPeriod.Year > dateTime.Year || (dtPeriod.Year >= dateTime.Year && dtPeriod.Month >= dateTime.Month && dateTime2 >= dateTime))
			{
				return true;
			}
		}
		else if (PeriodInfo.KieuKy == "Q")
		{
			DateTime dateTime3 = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if ((dateTime3 <= dateTime && dateTime2 >= dateTime) || dateTime3 > dateTime)
			{
				return true;
			}
		}
		else if (dtPeriod >= DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture))
		{
			return true;
		}
		return result;
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 1;
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					LayThemThongTinCQTChiCuc(sView, dataLocationTemplateRowStart + j);
					LayThemThongTinTinhHuyen_PL(sView, dataLocationTemplateRowStart + j);
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("C"), HTKKFile.ColumnLettersToIndex("AR"));
					GeneralFunctions.SetLaiDongCon(sView, hTKKCMCSheetInfo, i, dataLocationTemplateRowStart + j, HTKKFile.ColumnLettersToIndex("H"), HTKKFile.ColumnLettersToIndex("AS"));
					GeneralFunctions.UpdateFromSViewToSInfo(sView, dataLocationTemplateRowStart + j, -1, hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LayThemThongTinTinhHuyen_PL(SheetView sView, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AH");
			int num2 = HTKKFile.ColumnLettersToIndex("AK");
			int colTenTinh = HTKKFile.ColumnLettersToIndex("AN");
			int colMaTinh = HTKKFile.ColumnLettersToIndex("AO");
			int colTenHuyen = HTKKFile.ColumnLettersToIndex("AP");
			int colMaHuyen = HTKKFile.ColumnLettersToIndex("AQ");
			string text = sView.Cells[iRow - 1, num - 1].Text;
			string text2 = sView.Cells[iRow - 1, num2 - 1].Text;
			if (!string.IsNullOrEmpty(text) && text != "0")
			{
				text = text.Replace(".", "");
				GeneralFunctions.ThayDoiTinhTheoCQTDangBang2(text, text2, sView, colTenTinh, colMaTinh, colTenHuyen, colMaHuyen, iRow, awaySet: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LayThemThongTinCQTChiCuc(SheetView sView, int iRow)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AH");
			int colMaCQTChiCuc = HTKKFile.ColumnLettersToIndex("AK");
			int colTenCQTChiCuc = HTKKFile.ColumnLettersToIndex("AJ");
			string text = sView.Cells[iRow - 1, num - 1].Text;
			if (!string.IsNullOrEmpty(text) && text != "0")
			{
				text = text.Replace(".", "");
				GeneralFunctions.GetCQTChiCucTuCQTCuc(text, sView, colTenCQTChiCuc, colMaCQTChiCuc, iRow);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
