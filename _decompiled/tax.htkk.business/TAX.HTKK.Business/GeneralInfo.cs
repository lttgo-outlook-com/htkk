using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class GeneralInfo
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int row = 37;
			int column = HTKKFile.ColumnLettersToIndex("P") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			if ((bool)sView.Cells[row, column].Value)
			{
				sView.Cells[39, column2].Locked = false;
				sView.Cells[37, column2].Locked = false;
			}
			else
			{
				sView.Cells[39, column2].Locked = true;
				sView.Cells[37, column2].Locked = true;
			}
			row = 42;
			column = HTKKFile.ColumnLettersToIndex("C") - 1;
			column2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			int[] array = new int[12]
			{
				44, 46, 48, 52, 58, 60, 62, 64, 66, 68,
				70, 72
			};
			if ((bool)sView.Cells[row, column].Value)
			{
				for (int i = 0; i < array.Length; i++)
				{
					sView.Cells[array[i], column2].Locked = false;
				}
			}
			else
			{
				for (int i = 0; i < array.Length; i++)
				{
					sView.Cells[array[i], column2].Locked = true;
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("C_30");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array2[1] + 1, array2[0] + 1, "0");
			if (HTKKFormater.isNullString(valueCustom))
			{
				sView.Cells[array2[1], array2[0]].Text = "01/01";
			}
		}
		catch (Exception)
		{
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1;
			int num2 = winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1;
			int row;
			int column;
			int column2;
			if (num == 38 && num2 == HTKKFile.ColumnLettersToIndex("P"))
			{
				row = 37;
				column = HTKKFile.ColumnLettersToIndex("P") - 1;
				column2 = HTKKFile.ColumnLettersToIndex("C") - 1;
				if ((bool)winFpSpread.Sheets[0].Cells[row, column].Value)
				{
					winFpSpread.Sheets[0].Cells[39, column2].Locked = false;
					winFpSpread.Sheets[0].Cells[37, column2].Locked = false;
				}
				else
				{
					winFpSpread.Sheets[0].Cells[39, column2].Locked = true;
					winFpSpread.Sheets[0].Cells[37, column2].Locked = true;
					winFpSpread.Sheets[0].Cells[39, column2].ResetValue();
					winFpSpread.Sheets[0].Cells[37, column2].ResetValue();
				}
			}
			if (num != 43 || num2 != HTKKFile.ColumnLettersToIndex("C"))
			{
				return;
			}
			row = 42;
			column = HTKKFile.ColumnLettersToIndex("C") - 1;
			column2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			int[] array = new int[12]
			{
				44, 46, 48, 52, 58, 60, 62, 64, 66, 68,
				70, 72
			};
			if ((bool)winFpSpread.Sheets[0].Cells[row, column].Value)
			{
				for (int i = 0; i < array.Length; i++)
				{
					winFpSpread.Sheets[0].Cells[array[i], column2].Locked = false;
				}
				return;
			}
			for (int i = 0; i < array.Length; i++)
			{
				winFpSpread.Sheets[0].Cells[array[i], column2].Locked = true;
				winFpSpread.Sheets[0].Cells[array[i], column2].ResetValue();
			}
		}
		catch (Exception)
		{
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int iwarningcolor)
	{
		Validate_0(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
		return true;
	}

	public void Validate_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 0;
			List<int> list = new List<int>();
			list.Add(4);
			list.Add(6);
			list.Add(10);
			list.Add(16);
			list.Add(30);
			list.Add(34);
			list.Add(36);
			list.Add(38);
			list.Add(40);
			list.Add(45);
			list.Add(47);
			list.Add(49);
			list.Add(53);
			list.Add(59);
			list.Add(67);
			list.Add(69);
			list.Add(71);
			list.Add(73);
			List<int> list2 = list;
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			List<string> list3 = new List<string>();
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_15")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_16")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_17")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_18")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_19")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_22")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_36")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_26")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_27")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_28")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_34")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_31")));
			list3.Add(HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30")));
			List<string> list4 = list3;
			if (bool.Parse(sView.Cells[42, HTKKFile.ColumnLettersToIndex("C") - 1].Value.ToString()))
			{
				for (int i = 9; i < list2.Count; i++)
				{
					if (sView.Cells[list2[i] - 1, num2 - 1].Text == "")
					{
						HTKKFile.SetNodeBackGroundPink(sView, list2[i] - 1, num2 - 1, list4[i]);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, list2[i] - 1, num2 - 1);
					}
				}
			}
			else
			{
				for (int i = 9; i < list2.Count; i++)
				{
					HTKKFile.RemoveValidate(sView, list2[i] - 1, num2 - 1);
				}
			}
			if (bool.Parse(sView.Cells[37, HTKKFile.ColumnLettersToIndex("P") - 1].Value.ToString()))
			{
				for (int i = 7; i < 9; i++)
				{
					if (sView.Cells[list2[i] - 1, num2 - 1].Text == "")
					{
						HTKKFile.SetNodeBackGroundPink(sView, list2[i] - 1, num2 - 1, list4[i]);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sView, list2[i] - 1, num2 - 1);
					}
				}
			}
			else
			{
				for (int i = 7; i < 9; i++)
				{
					HTKKFile.RemoveValidate(sView, list2[i] - 1, num2 - 1);
				}
			}
			for (int i = 0; i < 7; i++)
			{
				if (sView.Cells[list2[i] - 1, num2 - 1].Text == "")
				{
					HTKKFile.SetNodeBackGroundPink(sView, list2[i] - 1, num2 - 1, list4[i]);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, list2[i] - 1, num2 - 1);
				}
			}
			int[] array = GeneralFunctions.ArrayCellId("T_2");
			int[] array2 = GeneralFunctions.ArrayCellId("C_45");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "0");
			string[] array3 = new string[13]
			{
				"C", "D", "E", "F", "G", "H", "I", "J", "K", "L",
				"N", "O", "P"
			};
			for (int j = 0; j < array3.Length; j++)
			{
				HTKKFile.RemoveValidateBackGround(sView, array[1], HTKKFile.ColumnLettersToIndex(array3[j]) - 1);
			}
			if (!HTKKFormater.isNullString(valueCustom) && !ValidateBase.CheckMaSoThueCCCD(valueCustom))
			{
				for (int k = 0; k < array3.Length; k++)
				{
					HTKKFile.SetNodeBackGroundPink(sView, array[1], HTKKFile.ColumnLettersToIndex(array3[k]) - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
				}
				num = 2;
			}
			HTKKFile.RemoveValidate(sView, array2[1], array2[0]);
			if (bool.Parse(sView.Cells[42, HTKKFile.ColumnLettersToIndex("C") - 1].Value.ToString()))
			{
				string valueCustom2 = GeneralFunctions.GetValueCustom(sView, array2[1] + 1, array2[0] + 1, "0");
				if (HTKKFormater.isNullString(valueCustom2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, array2[1], array2[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_24")));
					num = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(valueCustom2))
				{
					HTKKFile.SetNodeBackGroundPink(sView, array2[1], array2[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				else if (valueCustom2.Equals(valueCustom))
				{
					HTKKFile.SetNodeBackGroundYellow(sView, array2[1], array2[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_29")));
					if (num == 0)
					{
						num = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, array2[1], array2[0]);
				}
			}
			if (bool.Parse(sView.Cells[42, HTKKFile.ColumnLettersToIndex("C") - 1].Value.ToString()))
			{
				int[] array4 = GeneralFunctions.ArrayCellId("C_69");
				HTKKFile.RemoveValidate(sView, array4[1], array4[0]);
				string valueCustom3 = GeneralFunctions.GetValueCustom(sView, array4[1] + 1, array4[0] + 1, "0");
				if (!HTKKFormater.isNullString(valueCustom3))
				{
					DateTime dtDate = DateTime.Now;
					if (!HTKKFormater.IsDateTime(valueCustom3, "dd/MM/yyyy", out dtDate))
					{
						HTKKFile.SetNodeBackGroundPink(sView, array4[1], array4[0], HTKKFile.GetMessage("0403"));
						num = 2;
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sView, array4[1], array4[0], HTKKFile.GetMessage("0403"));
					num = 2;
				}
			}
			int[] array5 = GeneralFunctions.ArrayCellId("C_40");
			if (bool.Parse(sView.Cells[37, HTKKFile.ColumnLettersToIndex("P") - 1].Value.ToString()))
			{
				string valueCustom4 = GeneralFunctions.GetValueCustom(sView, array5[1] + 1, array5[0] + 1, "0");
				if (HTKKFormater.isNullString(valueCustom4))
				{
					HTKKFile.SetNodeBackGroundPink(sView, array5[1], array5[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
					num = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(valueCustom4))
				{
					HTKKFile.SetNodeBackGroundPink(sView, array5[1], array5[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, array5[1], array5[0]);
				}
			}
			switch (num)
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
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_0123: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("T_2");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "0");
			int[] array2 = GeneralFunctions.ArrayCellId("C_30");
			if ((int)iRow != array2[1] + 1 || (int)iCol != array2[0] + 1)
			{
				return;
			}
			string valueCustom2 = GeneralFunctions.GetValueCustom(sView, array2[1] + 1, array2[0] + 1, "0");
			if (HTKKFormater.isNullString(valueCustom2))
			{
				return;
			}
			string text = HTKKConfig.DATAFILES + "\\" + valueCustom + "\\" + ToKhaiInfo.SheetInfos[0].DataFile + ".xml";
			if (File.Exists(text))
			{
				string valueByAttributes = HTKKFile.GetValueByAttributes(text, "/Sections/Section/Cells/Cell[@Path='p26']");
				if (!valueCustom2.Equals(valueByAttributes))
				{
					frmMessage frmMessage = new frmMessage("0073", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
