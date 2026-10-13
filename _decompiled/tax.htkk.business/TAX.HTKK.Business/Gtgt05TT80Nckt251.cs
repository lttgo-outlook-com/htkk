using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt05TT80Nckt251
{
	private string cellId_J54 = "J_54";

	private string cellId_M54 = "M_54";

	private string cellId_C54 = "C_54";

	private string cellId_I54 = "I_54";

	private string cellId_HDXayDungCt21 = "K_43";

	private string cellId_HDBdsCt21 = "K_44";

	private int[] cell09c = GeneralFunctions.ArrayCellId("F_18");

	private int[] cell09cMa = GeneralFunctions.ArrayCellId("AF_18");

	private int[] cell09b = GeneralFunctions.ArrayCellId("Q_18");

	private int[] cell09bMa = GeneralFunctions.ArrayCellId("AG_18");

	private int[] cell09a = GeneralFunctions.ArrayCellId("F_20");

	private int[] cell09aMa = GeneralFunctions.ArrayCellId("AF_20");

	private int[] CellId_Tinh => GeneralFunctions.ArrayCellId("F_18");

	private int[] CellId_J54 => GeneralFunctions.ArrayCellId(cellId_J54);

	private int[] CellId_M54 => GeneralFunctions.ArrayCellId(cellId_M54);

	private int[] CellId_C54 => GeneralFunctions.ArrayCellId(cellId_C54);

	private int[] CellId_I54 => GeneralFunctions.ArrayCellId(cellId_I54);

	private int[] CellId_HDXayDungCt21 => GeneralFunctions.ArrayCellId(cellId_HDXayDungCt21);

	private int[] CellId_HDBdsCt21 => GeneralFunctions.ArrayCellId(cellId_HDBdsCt21);

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		bool result = true;
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		foreach (KeyValuePair<int, int> item in levelWarning)
		{
			int key = item.Key;
			switch (item.Value)
			{
			case 2:
				winFpSpread.TabStrip[key].BackColor = Color.Pink;
				continue;
			case 1:
				if (winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					continue;
				}
				break;
			}
			winFpSpread.TabStrip[key].BackColor = Color.White;
		}
		return result;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			StringBuilder stringBuilder = new StringBuilder();
			SheetView sheetView = winFpSpread.Sheets[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string loaiToKhai = PeriodInfo.LoaiToKhai;
			double num = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_HDXayDungCt21[1], CellId_HDXayDungCt21[0]].Value);
			double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_HDBdsCt21[1], CellId_HDBdsCt21[0]].Value);
			HTKKFile.RemoveValidate(sheetView, CellId_HDXayDungCt21[1], CellId_HDXayDungCt21[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_HDBdsCt21[1], CellId_HDBdsCt21[0]);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQTDbhc(headerInfo, winFpSpread.Sheets[0], PeriodInfo, ToKhaiInfo.SheetInfos[0], cell09c[1] + 1, cell09a[1] + 1))
			{
				value = 2;
			}
			string text10 = sheetView.Cells[cell09cMa[1], cell09cMa[0]].Text;
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int colIndex = hTKKCMCTableInfo.RowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int colIndex2 = hTKKCMCTableInfo.RowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int colIndex3 = hTKKCMCTableInfo.RowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int num4 = hTKKCMCTableInfo.RowsInfo.Cells[6].CellTemplate.ColIndex - 1;
			int colIndex4 = hTKKCMCTableInfo.RowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			int colIndex5 = hTKKCMCTableInfo.RowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int colIndex6 = hTKKCMCTableInfo.RowsInfo.Cells[11].CellTemplate.ColIndex - 1;
			string text11 = "";
			string text12 = "";
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num3 + i, colIndex);
				HTKKFile.RemoveValidate(sheetView, num3 + i, colIndex2);
				HTKKFile.RemoveValidate(sheetView, num3 + i, colIndex3);
				HTKKFile.RemoveValidate(sheetView, num3 + i, colIndex5);
				HTKKFile.RemoveValidate(sheetView, num3 + i, colIndex6);
				HTKKFile.RemoveValidate(sheetView, num3 + i, colIndex4);
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				bool flag = false;
				for (int j = 1; j < array.Length; j++)
				{
					if (!string.IsNullOrEmpty(array[j]) && j != array.Length - 1)
					{
						flag = true;
						break;
					}
					if (j == array.Length - 1 && ConvertHelpers.ConvertToDouble(array[^1]) != 0.0)
					{
						flag = true;
					}
				}
				if (string.IsNullOrEmpty(array[1]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex, text8);
					value = 2;
				}
				text11 = array[2];
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(text11, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text11))
				{
					if (flag)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex2, text8);
						value = 2;
					}
				}
				else if (dateTime == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex2, text4);
					value = 2;
				}
				else if (dateTime > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex2, text);
					value = 2;
				}
				if (string.IsNullOrEmpty(array[3]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex5, text8);
					value = 2;
				}
				text12 = array[4];
				if (!string.IsNullOrEmpty(text12) && !ValidateBase.CheckMaSoThue(text12))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex3, text2, i + 1);
					value = 2;
				}
				string text13 = array[6];
				if (!string.IsNullOrEmpty(text13) && !string.IsNullOrEmpty(text10) && text13 != text10)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex4, text9, i + 1);
					value = 2;
				}
				else if (string.IsNullOrEmpty(text13) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex4, text8);
					value = 2;
				}
				if (string.IsNullOrEmpty(array[11]) && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 + i, colIndex6, text8);
					value = 2;
				}
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

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "100");
			if (hTKKCMCCellInfo != null)
			{
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = PeriodInfo.TenTinh;
				sheetView.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Value = PeriodInfo.MaTinh;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "101");
			if (hTKKCMCCellInfo2 != null)
			{
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = PeriodInfo.TenXa;
				sheetView.Cells[hTKKCMCCellInfo2.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo2.CellSelectedValue.ColIndex - 1].Value = PeriodInfo.MaXa;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}
}
