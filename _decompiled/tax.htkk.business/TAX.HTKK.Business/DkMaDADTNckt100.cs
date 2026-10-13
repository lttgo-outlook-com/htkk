using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkMaDADTNckt100
{
	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			string msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[34, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string msg5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("T");
			int num3 = HTKKFile.ColumnLettersToIndex("AD");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("Y");
			int num7 = HTKKFile.ColumnLettersToIndex("P");
			int num8 = HTKKFile.ColumnLettersToIndex("V");
			int num9 = HTKKFile.ColumnLettersToIndex("O");
			int num10 = HTKKFile.ColumnLettersToIndex("R");
			int num11 = HTKKFile.ColumnLettersToIndex("X");
			int num12 = HTKKFile.ColumnLettersToIndex("Z");
			int num13 = HTKKFile.ColumnLettersToIndex("E");
			for (int i = 0; i < rowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1);
				string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text3))
				{
					DateTime t = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					if (DateTime.Compare(t, DateTime.Now) > 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num2 - 1, msg2);
						value = 2;
					}
				}
				string text4 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text4))
				{
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo, num, text4);
					if (rowIndexInSViewByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, msg);
						value = 2;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Value))
				{
					string text5 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num9 - 1].Text.Trim();
					string text6 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Text;
					string text7 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num8 - 1].Text;
					if (!text6.Equals(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, msg3);
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, msg3);
						value = 2;
					}
					string text8 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Text;
				}
				else
				{
					string text8 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num11 - 1].Text;
					string maTinhMoiByMaTinhCu = GeneralFunctions.GetMaTinhMoiByMaTinhCu(text8.Substring(0, 3));
					string text6 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Text;
					if (text8 != "82500" && text8 != "82700" && (!maTinhMoiByMaTinhCu.Equals(text6) || text8.Substring(3, 2) != "00"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num10 - 1, msg4);
						value = 2;
					}
				}
				if (CheckRowError(sheetView, hTKKCMCTableInfo.RowsInfo, dataLocationTemplateRowStart + i))
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num13 - 1].Value = false;
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num13 - 1, msg5);
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
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		int num = HTKKFile.ColumnLettersToIndex("E");
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		if (idxRow != 18 || idxCol != num - 1)
		{
			return;
		}
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
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
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
				sheetView.SetArray(dataLocationTemplateRowStart - 1, num - 1, array);
				break;
			}
			case 2:
			{
				object[,] array = new object[rowsValues.Count, 1];
				for (int j = 0; j < rowsValues.Count; j++)
				{
					array[j, 0] = false;
				}
				sheetView.SetArray(dataLocationTemplateRowStart - 1, num - 1, array);
				break;
			}
			case 3:
			{
				int num2 = ConvertHelpers.ConvertToInt(start);
				int num3 = ConvertHelpers.ConvertToInt(end);
				if (num2 < 1)
				{
					num2 = 1;
				}
				if (ConvertHelpers.ConvertToInt(num3) > rowsValues.Count)
				{
					num3 = rowsValues.Count;
				}
				object[,] array = new object[rowsValues.Count, 1];
				for (int j = 0; j < rowsValues.Count; j++)
				{
					if (j >= num2 - 1 && j <= num3 - 1)
					{
						array[j, 0] = true;
					}
					else
					{
						array[j, 0] = false;
					}
				}
				sheetView.SetArray(dataLocationTemplateRowStart - 1, num - 1, array);
				break;
			}
			}
		}
		winFpSpread.ResumeLayout(performLayout: true);
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		setMaDADTXoa(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		SheetView sheetView = winFpSpread.Sheets[0];
		ButtonCellType buttonCellType = new ButtonCellType();
		buttonCellType = new ButtonCellType();
		buttonCellType.TwoState = false;
		buttonCellType.Text = sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		buttonCellType.Picture = Image.FromFile("..\\Pictures\\Check.bmp");
		buttonCellType.ButtonColor = Color.FromArgb(240, 240, 240);
		sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].CellType = buttonCellType;
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			if (iRow - 1 >= num && iRow - 1 <= num2)
			{
				if (iCol == num3)
				{
				}
				if (iCol == num3)
				{
					string text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					setMaDADTXoa(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
	}

	private void setMaDADTXoa(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0090: Unknown result type (might be due to invalid IL or missing references)
		//IL_0097: Expected O, but got Unknown
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AC");
			string text = "";
			string empty = string.Empty;
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + TTinChung.mst);
			string text2 = "02_GTGT~01_DNHT";
			string[] array = text2.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string text3 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text3 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					XmlDocument val = new XmlDocument();
					val.Load(fileInfo.FullName);
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/maHoSoKhaiDuAnDauTu']");
					if (val2.Count > 0)
					{
						text = text + "~" + ((XmlNode)val2[0].Attributes["Value"]).Value;
					}
				}
			}
			string[] array3 = text.Split(new char[1] { '~' });
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int num2 = location.DataLocationTemplateRowStart - 1;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			if (array3.Length <= 0)
			{
				return;
			}
			for (int k = 0; k < rowsValues.Count; k++)
			{
				bool flag = false;
				string[] array4 = rowsValues[k].RowValue.ToArray();
				array2 = array3;
				foreach (string value in array2)
				{
					if (array4[1].Equals(value))
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					sheetView.Cells[num2 + k, num - 1].Text = "0";
				}
				else
				{
					sheetView.Cells[num2 + k, num - 1].Text = "1";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckRowError(SheetView sView, HTKKCMCRowInfo rowInfo, int iRow)
	{
		bool result = false;
		foreach (HTKKCMCTableCellInfo cell in rowInfo.Cells)
		{
			if (sView.Cells[iRow - 1, cell.CellTemplate.ColIndex - 1].BackColor == Color.Pink)
			{
				result = true;
				break;
			}
		}
		return result;
	}
}
