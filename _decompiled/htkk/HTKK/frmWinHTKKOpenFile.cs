using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Excel;
using FarPoint.Win.Spread;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmWinHTKKOpenFile : Form
{
	private FpSpread winFpSpread;

	private IContainer components = null;

	private RadioButton rdThemDuLieu;

	private RadioButton rdThemXoaDuLieuDaCo;

	private Panel panel1;

	private TextBox txtFileNameLoading;

	private Button btnDong;

	private Button btnTaiDuLieu;

	private Button btnBrow;

	private Panel panel2;

	private Label label3;

	public frmWinHTKKOpenFile(FpSpread winFpSpread)
	{
		InitializeComponent();
		this.winFpSpread = winFpSpread;
		if (WinSettings.MenuID == "U8")
		{
			((Control)rdThemDuLieu).Enabled = false;
		}
	}

	private void btnBrow_Click(object sender, EventArgs e)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Invalid comparison between Unknown and I4
		OpenFileDialog val = new OpenFileDialog();
		((FileDialog)val).DefaultExt = "xlsx";
		((FileDialog)val).Filter = "Excel file (*.xls, *.xlsx)|*.xls; *.xlsx";
		((FileDialog)val).AddExtension = true;
		((FileDialog)val).RestoreDirectory = true;
		((FileDialog)val).Title = "Chọn bảng kê";
		DialogResult val2 = ((CommonDialog)val).ShowDialog();
		if ((int)val2 == 1)
		{
			((Control)txtFileNameLoading).Text = ((FileDialog)val).FileName;
		}
	}

	private void btnDong_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void form_DoWorkTaiBangKe(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].TaiDuLieuLon == "1")
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			TaiDuLieuLon(winFpSpread, ((Control)txtFileNameLoading).Text, rdThemDuLieu.Checked);
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
		else if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].TaiDuLieuRatLon == "1")
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			TaiDuLieuRatLon(winFpSpread, ((Control)txtFileNameLoading).Text, rdThemDuLieu.Checked);
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
		else
		{
			TaiBangKe(winFpSpread, ((Control)txtFileNameLoading).Text, rdThemDuLieu.Checked);
		}
	}

	private void TaiDuLieuLon(FpSpread winFpSpread, string fileName, bool themDuLieu)
	{
		//IL_017a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_012f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0228: Unknown result type (might be due to invalid IL or missing references)
		//IL_022f: Expected O, but got Unknown
		//IL_0477: Unknown result type (might be due to invalid IL or missing references)
		//IL_05fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0605: Expected O, but got Unknown
		MethodInvoker val = null;
		if (HTKKFormater.isNullString(fileName))
		{
			winFpSpread.ResumeLayout();
			frmMessage frmMessage = new frmMessage("0023", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		FileInfo file = new FileInfo(fileName);
		if (HTKKFile.IsFileLocked(file))
		{
			winFpSpread.ResumeLayout();
			frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		FpSpread winFpSpreadBK = new FpSpread();
		((ISupportInitialize)winFpSpreadBK).BeginInit();
		winFpSpread.Name = "winFpSpreadBK";
		((ISupportInitialize)winFpSpreadBK).EndInit();
		try
		{
			Stream stream = new FileStream(fileName, FileMode.Open);
			bool flag = winFpSpreadBK.OpenExcel(stream, ExcelOpenFlags.DataOnly);
			stream.Close();
			if (!flag)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
		}
		catch (Exception ex)
		{
			winFpSpread.ResumeLayout();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		winFpSpreadBK.ActiveSheetIndex = 0;
		SheetView sheetView = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex];
		string status = "";
		HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
		List<HTKKCMCTableInfo> tables = sInfo.Tables;
		winFpSpread.ActiveSheet.AutoCalculation = false;
		if (WinSettings.MenuID == "M4")
		{
			if (val == null)
			{
				val = (MethodInvoker)delegate
				{
					//IL_02db: Unknown result type (might be due to invalid IL or missing references)
					//IL_040a: Unknown result type (might be due to invalid IL or missing references)
					//IL_0411: Expected O, but got Unknown
					for (int num2 = sInfo.Tables.Count - 1; num2 >= 0; num2--)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables[num2];
						int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
						HTKKCMCRowInfo rowsInfo2 = hTKKCMCTableInfo2.RowsInfo;
						GeneralFunctions.RemoveRowOnSView(winFpSpread.ActiveSheet, hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo2.RowsValues.Count - 1);
						foreach (HTKKCMCTableInfo table in sInfo.Tables)
						{
							if (table.Location.DataLocationTemplateRowStart > dataLocationTemplateRowStart2)
							{
								table.Increment((hTKKCMCTableInfo2.RowsValues.Count - 1) * -1);
							}
						}
						if (hTKKCMCTableInfo2.Footer != null && hTKKCMCTableInfo2.Footer.StaticCells.Count > 0)
						{
							hTKKCMCTableInfo2.IncrementFooter(dataLocationTemplateRowStart2, (hTKKCMCTableInfo2.RowsValues.Count - 1) * -1);
						}
						sInfo.IncrementStaticCell(dataLocationTemplateRowStart2, (hTKKCMCTableInfo2.RowsValues.Count - 1) * -1);
						if (!themDuLieu)
						{
							sInfo.Tables[num2].RowsValues.Clear();
						}
					}
					int index = 0;
					HTKKCMCTableInfo hTKKCMCTableInfo3 = sInfo.Tables[index];
					int dataLocationTaiBangKeRowStart = ((hTKKCMCTableInfo3.Location.DataLocationTaiBangKeRowStart >= 0) ? hTKKCMCTableInfo3.Location.DataLocationTaiBangKeRowStart : 0);
					XmlNode xmlRootNode = ((XmlNode)new HTKKCMCXML(GeneralFunctions.GetXMLFileName(WinSettings.ToKhaiInfo.InterfaceTemplateXml, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].DataFile, WinSettings.ToKhaiInfo.VersionXML, "xml")).XDocument).SelectSingleNode("/Sections");
					Dictionary<int, List<HTKKCMCRowValue>> dataFromExcelFileGroupF7LargeData = WinSpreadProcess.GetDataFromExcelFileGroupF7LargeData(winFpSpread, winFpSpreadBK, ref sInfo, xmlRootNode, dataLocationTaiBangKeRowStart, 0, themDuLieu, ref status, winFpSpread.ActiveSheetIndex);
					if (!themDuLieu)
					{
						foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
						{
							table2.RowsValues.Clear();
						}
					}
					if (!HTKKFormater.isNullString(status))
					{
						winFpSpread.ResumeLayout();
						string outMessage2 = "Dữ liệu bảng kê không đúng định dạng tại cell " + status;
						frmMessage frmMessage2 = new frmMessage("", outMessage2, "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
					}
					else
					{
						for (int num2 = 0; num2 < sInfo.Tables.Count; num2++)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables[num2];
							int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
							List<HTKKCMCRowValue> list2 = (dataFromExcelFileGroupF7LargeData.ContainsKey(num2) ? dataFromExcelFileGroupF7LargeData[num2] : new List<HTKKCMCRowValue>());
							if (list2.Count > 0)
							{
								hTKKCMCTableInfo2.RowsValues.AddRange(list2);
							}
							else
							{
								HTKKCMCRowValue rowValueDefault2 = GeneralFunctions.GetRowValueDefault(hTKKCMCTableInfo2.RowsInfo);
								hTKKCMCTableInfo2.RowsValues.Add(rowValueDefault2);
							}
							winFpSpread.TabStrip[winFpSpread.ActiveSheetIndex].BackColor = Color.White;
							WinSpreadProcess.InsertTableRowsLargeDataNoSView(dataLocationTemplateRowStart2, -1, hTKKCMCTableInfo2.RowsValues.Count + dataLocationTemplateRowStart2 - 1, -1, hTKKCMCTableInfo2.RowsValues.Count - 1, -1, sInfo);
						}
						MethodInvoker val3 = null;
						int i;
						for (i = 0; i < sInfo.Tables.Count; i++)
						{
							FpSpread fpSpread2 = winFpSpread;
							if (val3 == null)
							{
								val3 = (MethodInvoker)delegate
								{
									WinSpreadProcess.ProcessingLargeDataOnSpread(winFpSpread.ActiveSheet, winFpSpread.ActiveSheetIndex, sInfo.Tables[i]);
								};
							}
							((Control)fpSpread2).Invoke((Delegate)(object)val3);
						}
						winFpSpread.ActiveSheet.AutoCalculation = true;
						WinDoWorkEvent.FinishImportExcel(winFpSpread, WinSettings.ToKhaiInfo);
					}
				};
			}
			((Control)this).Invoke((Delegate)(object)val);
			return;
		}
		for (int num = sInfo.Tables.Count - 1; num >= 0; num--)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			GeneralFunctions.RemoveRowOnSView(winFpSpread.ActiveSheet, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count - 1);
			foreach (HTKKCMCTableInfo table3 in sInfo.Tables)
			{
				if (table3.Location.DataLocationTemplateRowStart > dataLocationTemplateRowStart)
				{
					table3.Increment((hTKKCMCTableInfo.RowsValues.Count - 1) * -1);
				}
			}
			if (hTKKCMCTableInfo.Footer != null && hTKKCMCTableInfo.Footer.StaticCells.Count > 0)
			{
				hTKKCMCTableInfo.IncrementFooter(dataLocationTemplateRowStart, (hTKKCMCTableInfo.RowsValues.Count - 1) * -1);
			}
			sInfo.IncrementStaticCell(dataLocationTemplateRowStart, (hTKKCMCTableInfo.RowsValues.Count - 1) * -1);
			if (!themDuLieu)
			{
				sInfo.Tables[num].RowsValues.Clear();
			}
		}
		Dictionary<int, List<HTKKCMCRowValue>> dataFromExcelLarge = WinSpreadProcess.GetDataFromExcelLarge(winFpSpreadBK, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex], WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables, themDuLieu, ref status);
		if (!HTKKFormater.isNullString(status))
		{
			winFpSpread.ResumeLayout();
			string outMessage = "Dữ liệu bảng kê không đúng định dạng tại cell " + status;
			frmMessage frmMessage = new frmMessage("", outMessage, "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		for (int num = 0; num < sInfo.Tables.Count; num++)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			List<HTKKCMCRowValue> list = (dataFromExcelLarge.ContainsKey(num) ? dataFromExcelLarge[num] : new List<HTKKCMCRowValue>());
			if (list.Count > 0)
			{
				hTKKCMCTableInfo.RowsValues.AddRange(list);
			}
			else
			{
				HTKKCMCRowValue rowValueDefault = GeneralFunctions.GetRowValueDefault(hTKKCMCTableInfo.RowsInfo);
				hTKKCMCTableInfo.RowsValues.Add(rowValueDefault);
			}
			if (!themDuLieu)
			{
				winFpSpread.TabStrip[winFpSpread.ActiveSheetIndex].BackColor = Color.White;
				WinSpreadProcess.InsertTableRowsLargeDataNoSView(dataLocationTemplateRowStart, -1, hTKKCMCTableInfo.RowsValues.Count + dataLocationTemplateRowStart - 1, -1, hTKKCMCTableInfo.RowsValues.Count - 1, -1, sInfo);
			}
			else
			{
				WinSpreadProcess.InsertTableRowsLargeDataNoSView(dataLocationTemplateRowStart, -1, hTKKCMCTableInfo.RowsValues.Count + dataLocationTemplateRowStart - 1, -1, hTKKCMCTableInfo.RowsValues.Count, -1, sInfo);
			}
		}
		MethodInvoker val2 = null;
		int t;
		for (t = 0; t < sInfo.Tables.Count; t++)
		{
			FpSpread fpSpread = winFpSpread;
			if (val2 == null)
			{
				val2 = (MethodInvoker)delegate
				{
					WinSpreadProcess.ProcessingLargeDataOnSpread(winFpSpread.ActiveSheet, winFpSpread.ActiveSheetIndex, sInfo.Tables[t]);
				};
			}
			((Control)fpSpread).Invoke((Delegate)(object)val2);
		}
		winFpSpread.ActiveSheet.AutoCalculation = true;
		WinDoWorkEvent.FinishImportExcel(winFpSpread, WinSettings.ToKhaiInfo);
	}

	private void TaiBangKe(FpSpread winFpSpread, string fileName, bool themDuLieu)
	{
		//IL_017a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_012f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0408: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b54: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b5e: Expected O, but got Unknown
		//IL_0bb9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cdd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d62: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e38: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e3f: Expected O, but got Unknown
		MethodInvoker val = null;
		if (HTKKFormater.isNullString(fileName))
		{
			winFpSpread.ResumeLayout();
			frmMessage frmMessage = new frmMessage("0023", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		FileInfo file = new FileInfo(fileName);
		if (HTKKFile.IsFileLocked(file))
		{
			winFpSpread.ResumeLayout();
			frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		FpSpread winFpSpreadBK = new FpSpread();
		((ISupportInitialize)winFpSpreadBK).BeginInit();
		winFpSpread.Name = "winFpSpreadBK";
		((ISupportInitialize)winFpSpreadBK).EndInit();
		try
		{
			Stream stream = new FileStream(fileName, FileMode.Open);
			bool flag = winFpSpreadBK.OpenExcel(stream, ExcelOpenFlags.DataOnly);
			stream.Close();
			if (!flag)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
		}
		catch (Exception ex)
		{
			winFpSpread.ResumeLayout();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		winFpSpreadBK.ActiveSheetIndex = 0;
		if (WinSettings.MenuID == "33")
		{
			bool flag2 = false;
			if (winFpSpread.ActiveSheetIndex == 1)
			{
				string text = winFpSpreadBK.Sheets[winFpSpreadBK.ActiveSheetIndex].Cells[4, HTKKFile.ColumnLettersToIndex("L") - 1].Text.Trim();
				string text2 = winFpSpreadBK.Sheets[winFpSpreadBK.ActiveSheetIndex].Cells[4, HTKKFile.ColumnLettersToIndex("M") - 1].Text.Trim();
				if (!HTKKFormater.isNullString(text) && text == "[17]" && !HTKKFormater.isNullString(text2) && text2 == "[18]")
				{
					flag2 = true;
				}
			}
			if (winFpSpread.ActiveSheetIndex == 2)
			{
				string text = winFpSpreadBK.Sheets[winFpSpreadBK.ActiveSheetIndex].Cells[4, HTKKFile.ColumnLettersToIndex("L") - 1].Text.Trim();
				string text2 = winFpSpreadBK.Sheets[winFpSpreadBK.ActiveSheetIndex].Cells[4, HTKKFile.ColumnLettersToIndex("M") - 1].Text.Trim();
				if (!HTKKFormater.isNullString(text) && text == "[17]" && HTKKFormater.isNullString(text2))
				{
					flag2 = true;
				}
			}
			if (winFpSpread.ActiveSheetIndex == 3)
			{
				string text = winFpSpreadBK.Sheets[winFpSpreadBK.ActiveSheetIndex].Cells[38, HTKKFile.ColumnLettersToIndex("W") - 1].Text.Trim();
				if (!HTKKFormater.isNullString(text) && text == "[21]")
				{
					flag2 = true;
				}
			}
			if (!flag2)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0452", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
		}
		HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, " start import - " + DateTime.Now.ToString("hh:mm:ss"));
		Stopwatch stopwatch = Stopwatch.StartNew();
		stopwatch.Reset();
		stopwatch.Start();
		winFpSpread.ActiveSheet.AutoCalculation = false;
		int num = 0;
		if (WinSettings.MenuID == "G7" || WinSettings.MenuID == "D4" || WinSettings.MenuID == "G5" || WinSettings.MenuID.Equals("Z3"))
		{
			if (WinSettings.MenuID == "G5" && winFpSpread.ActiveSheetIndex == 5)
			{
				num = 1;
			}
			if (WinSettings.MenuID == "G7" && winFpSpread.ActiveSheetIndex == 5)
			{
				num = 1;
			}
			if (WinSettings.MenuID == "D4" && winFpSpread.ActiveSheetIndex == 3)
			{
				num = 1;
			}
			if (WinSettings.MenuID.Equals("Z3") && winFpSpread.ActiveSheetIndex == 1)
			{
				num = 1;
			}
		}
		if (WinSettings.MenuID == "Q8" && winFpSpread.ActiveSheetIndex == 3)
		{
			num = 1;
		}
		if (WinSettings.MenuID == "U5" && winFpSpread.ActiveSheetIndex == 2)
		{
			num = 1;
		}
		string flagStop = "aa";
		if (WinSettings.MenuID == "95")
		{
			flagStop = "II. Đăng ký thay đổi về người phụ thuộc";
		}
		if (WinSettings.MenuID == "68")
		{
			flagStop = "II";
		}
		if (WinSettings.MenuID == "79")
		{
			flagStop = "Tổng";
		}
		if (WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5")
		{
			flagStop = "Tổng";
		}
		if (WinSettings.MenuID == "31")
		{
			flagStop = "Tổng cộng";
		}
		if (WinSettings.MenuID == "G8")
		{
			flagStop = "Tổng cộng";
		}
		if (WinSettings.MenuID == "R4")
		{
			flagStop = "Tổng cộng";
		}
		if (WinSettings.MenuID == "X8")
		{
			flagStop = "Tổng cộng";
		}
		if (WinSettings.MenuID == "U7")
		{
			flagStop = "Tổng cộng";
		}
		if (WinSettings.MenuID == "Q8")
		{
			flagStop = "II. MÁY MÓC, THIẾT BỊ";
		}
		if (WinSettings.MenuID == "03")
		{
			flagStop = "aa";
		}
		if (WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8")
		{
			flagStop = "II. Người phụ thuộc đăng ký thuế bằng giấy khai sinh:";
		}
		if (WinSettings.MenuID == "L4")
		{
			flagStop = "II,III";
		}
		if (WinSettings.MenuID == "P6")
		{
			flagStop = "II";
		}
		if (WinSettings.MenuID == "X7")
		{
			flagStop = "II,Tiểu mục 1099 - Thuế thu nhập doanh nghiệp khác";
		}
		if (WinSettings.MenuID == "P2")
		{
			flagStop = "II";
		}
		if (WinSettings.MenuID == "X6")
		{
			flagStop = "II";
		}
		if (WinSettings.MenuID == "N7" && winFpSpread.ActiveSheetIndex == 1)
		{
			flagStop = "II. Thông tin chung nhà cung cấp nước ngoài";
		}
		if (WinSettings.MenuID == "N7" && winFpSpread.ActiveSheetIndex == 3)
		{
			flagStop = "II. Thông tin tài khoản nhà cung cấp nước ngoài";
		}
		if (WinSettings.MenuID == "L8" && num < 6)
		{
			flagStop = "Tổng cộng,I,II,III,IV,V";
		}
		if (WinSettings.MenuID == "99" && num < 3 && winFpSpread.ActiveSheetIndex == 3)
		{
			flagStop = "aa,bb,cc";
		}
		if (WinSettings.MenuID == "D1")
		{
			flagStop = "II. Hàng hóa, dịch vụ bán ra trong kỳ";
		}
		if (WinSettings.MenuID == "Z5")
		{
			flagStop = "II. Hàng hóa, dịch vụ bán ra trong kỳ";
		}
		if (WinSettings.MenuID == "P7")
		{
			flagStop = "Thông tin về các công ty mẹ thuộc đối tượng áp dụng quy định IIR,Thông tin về các đơn vị hợp thành thuộc đối tượng áp dụng quy định IIR";
			num = 1;
		}
		if (WinSettings.MenuID == "Q5")
		{
			flagStop = "1.3.1,1.3.2,1.3.3, 1.3.4,3.1.2.2,3.1.2.3,3.2.3.b,3.5.1,4.1.1,4.1.2,4.1.3,4.1.4,4.2.1,4.2.2,4.2.3,4.3.1,4.3.2,4.3.3,4.4.1";
			num = 0;
		}
		if (WinSettings.MenuID == "Y8")
		{
			int[] array = GeneralFunctions.ArrayCellId("M_6");
			int[] array2 = GeneralFunctions.ArrayCellId("M_8");
			bool flag3 = winFpSpread.ActiveSheet.Cells[array[1], array[0]].Value.ToBoolOrDefault();
			bool flag4 = winFpSpread.ActiveSheet.Cells[array2[1], array2[0]].Value.ToBoolOrDefault();
			flagStop = "";
			if (flag4)
			{
				num = 1;
			}
			if (flag3)
			{
				num = 0;
			}
			if (!flag3 && !flag4)
			{
				return;
			}
		}
		int DataLocationTaiBangKeRowStart = 0;
		int rowStartNext = 0;
		int num2 = 0;
		string status = "";
		HTKKCMCTableInfo tableInfo;
		List<HTKKCMCRowValue> rowsValues;
		bool checkThemBangF7;
		SheetView activeSheet;
		while (true)
		{
			tableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num];
			if (rowStartNext == 0)
			{
				DataLocationTaiBangKeRowStart = tableInfo.Location.DataLocationTaiBangKeRowStart;
			}
			else
			{
				DataLocationTaiBangKeRowStart = rowStartNext;
			}
			rowsValues = null;
			stopwatch.Stop();
			checkThemBangF7 = false;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
			((Control)this).Invoke((Delegate)(MethodInvoker)delegate
			{
				winFpSpread.SuspendLayout();
				winFpSpread.ActiveSheet.AutoCalculation = false;
				XmlNode xmlRootNode = ((XmlNode)new HTKKCMCXML(GeneralFunctions.GetXMLFileName(WinSettings.ToKhaiInfo.InterfaceTemplateXml, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].DataFile, WinSettings.ToKhaiInfo.VersionXML, "xml")).XDocument).SelectSingleNode("/Sections");
				checkThemBangF7 = WinSpreadProcess.GetDataFromExcelFileGroupF7(winFpSpread, winFpSpreadBK, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex], xmlRootNode, DataLocationTaiBangKeRowStart, rowStartNext, themDuLieu, ref status);
			});
			if (checkThemBangF7)
			{
				if (!string.IsNullOrEmpty(status))
				{
					winFpSpread.ResumeLayout();
					string outMessage = status;
					frmMessage frmMessage = new frmMessage("", outMessage, "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
				else
				{
					ResetValueCombobox(winFpSpread);
					WinDoWorkEvent.FinishImportExcel(winFpSpread, WinSettings.ToKhaiInfo);
					winFpSpread.ResumeLayout();
				}
				return;
			}
			if (!checkThemBangF7)
			{
				rowsValues = WinSpreadProcess.GetDataFromExcelFile(winFpSpreadBK, tableInfo.RowsInfo, DataLocationTaiBangKeRowStart, flagStop, out rowStartNext, out status);
			}
			if (rowsValues != null && rowsValues.Count == 0)
			{
				HTKKCMCRowValue rowValueDefault = GeneralFunctions.GetRowValueDefault(tableInfo.RowsInfo);
				if (rowValueDefault != null)
				{
					rowsValues.Add(rowValueDefault);
				}
			}
			if (!HTKKFormater.isNullString(status))
			{
				winFpSpread.ResumeLayout();
				string outMessage = "Dữ liệu bảng kê không đúng định dạng tại cell " + status;
				frmMessage frmMessage = new frmMessage("", outMessage, "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (rowsValues == null || rowsValues.Count == 0)
			{
				return;
			}
			if (WinSettings.ToKhaiInfo.MenuID == "B1" && rowsValues.Count >= 90000)
			{
				frmMessage frmMessage = new frmMessage("0626", "6", "2", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				if (frmMessage.Cancel)
				{
					winFpSpread.ResumeLayout();
					return;
				}
			}
			long elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "rowsValues - " + elapsedMilliseconds);
			stopwatch.Reset();
			stopwatch.Start();
			if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].SheetName + WinSettings.ToKhaiInfo.MenuID))
			{
				_ = rowsValues.Count;
			}
			bool flag5 = 0 == 0;
			if (val == null)
			{
				val = (MethodInvoker)delegate
				{
					if (themDuLieu)
					{
						int num5 = 0;
						if (tableInfo.RowsValues.Count == 1 && !GeneralFunctions.HaveToValidateRow(tableInfo.RowsInfo, tableInfo.RowsValues.FirstOrDefault().RowValue.ToArray()))
						{
							num5 = 1;
						}
						WinSpreadProcess.InsertTableRows(winFpSpread.ActiveSheet, tableInfo.Location.DataLocationTemplateRowStart, -1, tableInfo.RowsValues.Count + tableInfo.Location.DataLocationTemplateRowStart - 1, -1, rowsValues.Count - num5, -1, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex]);
						if (tableInfo.RowsValues.Count == 1 && !GeneralFunctions.HaveToValidateRow(tableInfo.RowsInfo, tableInfo.RowsValues.FirstOrDefault().RowValue.ToArray()))
						{
							tableInfo.RowsValues.Clear();
						}
					}
					else
					{
						WinSpreadProcess.InsertTableRows(winFpSpread.ActiveSheet, tableInfo.Location.DataLocationTemplateRowStart, -1, tableInfo.RowsValues.Count + tableInfo.Location.DataLocationTemplateRowStart - 1, -1, rowsValues.Count, -1, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex]);
						WinSpreadProcess.RemoveDataRows(winFpSpread.ActiveSheet, tableInfo.Location.DataLocationTemplateRowStart, tableInfo.RowsValues.Count, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex]);
						tableInfo.RowsValues.Clear();
						winFpSpread.TabStrip[winFpSpread.ActiveSheetIndex].BackColor = Color.White;
					}
				};
			}
			((Control)this).Invoke((Delegate)(object)val);
			stopwatch.Stop();
			elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "InsertTableRows - " + elapsedMilliseconds);
			stopwatch.Reset();
			stopwatch.Start();
			activeSheet = winFpSpread.ActiveSheet;
			int i = num;
			if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(hTKKCMCSheetInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) && rowsValues.Count > 500)
			{
				activeSheet.FpSpread.SuspendLayout();
				activeSheet.AutoCalculation = false;
				WinSpreadProcess.BindDataLarge(activeSheet, hTKKCMCSheetInfo, i, rowsValues);
			}
			else
			{
				int count = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num].RowsInfo.Cells.Count;
				for (int num3 = 0; num3 < rowsValues.Count; num3++)
				{
					for (int num4 = 0; num4 < count; num4++)
					{
						try
						{
							if (tableInfo.RowsInfo.Cells[num4].Controltype == "16" && tableInfo.RowsInfo.Cells[num4].DefaultValue.IndexOf(",") < 0)
							{
								rowsValues[num3].RowValue[num4] = rowsValues[num3].RowValue[num4].Replace(",", "").Replace(".", "");
							}
							else if (tableInfo.RowsInfo.Cells[num4].Controltype == "16" && tableInfo.RowsInfo.Cells[num4].DefaultValue.IndexOf(",") >= 0 && rowsValues[num3].RowValue[num4].Contains("%"))
							{
								rowsValues[num3].RowValue[num4] = rowsValues[num3].RowValue[num4].Replace("%", "");
							}
							WinSpreadProcess.ResetColorCellDefault(winFpSpread.ActiveSheet, tableInfo.RowsInfo.Cells[num4].Validate.ColorDefault, tableInfo.Location.DataLocationTemplateRowStart + num3 + tableInfo.RowsValues.Count, tableInfo.RowsInfo.Cells[num4].CellTemplate.ColIndex);
							WinSpreadProcess.fpSpBindDataCell(winFpSpread.ActiveSheet, tableInfo.Location.DataLocationTemplateRowStart + num3 + tableInfo.RowsValues.Count, tableInfo.RowsInfo.Cells[num4].CellTemplate.ColIndex, tableInfo.RowsInfo.Cells[num4].Controltype, rowsValues[num3].RowValue[num4]);
							if (WinSettings.MenuID == "68")
							{
								WinDoWorkEvent.CellFormatImport(winFpSpread, winFpSpread.ActiveSheetIndex, tableInfo.Location.DataLocationTemplateRowStart + num3 + tableInfo.RowsValues.Count, tableInfo.RowsInfo.Cells[num4].CellTemplate.ColIndex);
							}
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
					}
				}
			}
			winFpSpread.ActiveSheet.AutoCalculation = true;
			stopwatch.Stop();
			elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "fpSpBindDataCell - " + elapsedMilliseconds);
			stopwatch.Reset();
			stopwatch.Start();
			if (tableInfo.RowsValues.Count > 1)
			{
				tableInfo.RowsValues.InsertRange(tableInfo.RowsValues.Count - 1, rowsValues);
			}
			else
			{
				tableInfo.RowsValues.InsertRange(tableInfo.RowsValues.Count, rowsValues);
			}
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "InsertRange - " + elapsedMilliseconds);
			stopwatch.Stop();
			elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			stopwatch.Reset();
			stopwatch.Start();
			if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(hTKKCMCSheetInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) && rowsValues.Count > 500)
			{
				WinSpreadProcess.fpSpQuickReindexingTableTemplate(winFpSpread.ActiveSheet, tableInfo, tableInfo.RowsValues.Count);
				HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "fpSpReindexingTableTemplate - " + elapsedMilliseconds);
			}
			else
			{
				WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.ActiveSheet, tableInfo);
				HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "fpSpReindexingTableTemplate - " + elapsedMilliseconds);
			}
			stopwatch.Stop();
			elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Stop   import - " + DateTime.Now.ToString("hh:mm:ss"));
			if (WinSettings.MenuID == "95" && num == 0 && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				continue;
			}
			if ((WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8") && num == 0 && num < hTKKCMCSheetInfo.Tables.Count - 2)
			{
				num++;
				continue;
			}
			if (WinSettings.MenuID == "68" && num == 0 && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				continue;
			}
			if (WinSettings.MenuID == "79" && hTKKCMCSheetInfo.SheetIndex == 1)
			{
				switch (num)
				{
				case 0:
					num++;
					continue;
				case 1:
					num++;
					continue;
				case 2:
					num++;
					continue;
				}
			}
			if ((WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5") && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				continue;
			}
			if (WinSettings.MenuID == "P6" && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				flagStop = "II";
				continue;
			}
			if (WinSettings.MenuID == "X7" && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				flagStop = "II,Tiểu mục 1099 - Thuế thu nhập doanh nghiệp khác";
				continue;
			}
			if (WinSettings.MenuID == "P2" && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				continue;
			}
			if (WinSettings.MenuID == "X6" && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				continue;
			}
			if (WinSettings.MenuID == "31" && num < 2)
			{
				num++;
				flagStop = "Tổng cộng";
				continue;
			}
			if (WinSettings.MenuID == "G8" && num < 3)
			{
				num++;
				flagStop = "Tổng cộng";
				continue;
			}
			if (WinSettings.MenuID == "R4" && num < 3)
			{
				num++;
				flagStop = "Tổng cộng";
				continue;
			}
			if (WinSettings.MenuID == "X8" && num < 4)
			{
				num++;
				flagStop = "Tổng cộng";
				continue;
			}
			if (WinSettings.MenuID == "U7" && num < 2)
			{
				num++;
				flagStop = "Tổng cộng";
				continue;
			}
			if (WinSettings.MenuID == "N7" && winFpSpread.ActiveSheetIndex == 1 && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				flagStop = "II. Thông tin chung nhà cung cấp nước ngoài";
				continue;
			}
			if (WinSettings.MenuID == "Q8" && winFpSpread.ActiveSheetIndex == 2 && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				flagStop = "II. MÁY MÓC, THIẾT BỊ";
				if (num == 1)
				{
					rowStartNext--;
				}
				continue;
			}
			if (WinSettings.MenuID == "N7" && winFpSpread.ActiveSheetIndex == 3 && num < hTKKCMCSheetInfo.Tables.Count - 1)
			{
				num++;
				flagStop = "II. Thông tin tài khoản nhà cung cấp nước ngoài";
				continue;
			}
			if (WinSettings.MenuID == "L4" && num < 2)
			{
				num++;
				flagStop = "II,III";
				continue;
			}
			if (WinSettings.MenuID == "03")
			{
				switch (num)
				{
				case 0:
					num++;
					flagStop = "aa";
					continue;
				case 1:
					num++;
					flagStop = "bb";
					continue;
				case 2:
					num++;
					flagStop = "cc";
					continue;
				case 3:
					num++;
					flagStop = "dd";
					continue;
				case 4:
					num++;
					flagStop = "ee";
					continue;
				case 5:
					num++;
					flagStop = "ff";
					continue;
				case 6:
					num++;
					flagStop = "gg";
					continue;
				case 7:
					num++;
					flagStop = "hh";
					continue;
				}
			}
			if (WinSettings.MenuID == "L8" && num < 5)
			{
				num++;
				flagStop = "Tổng cộng,I,II,III,IV,V";
				if (num == 1)
				{
					rowStartNext += 5;
				}
			}
			else if (WinSettings.MenuID == "99" && num < 3 && winFpSpread.ActiveSheetIndex == 3)
			{
				num++;
				flagStop = "aa,bb,cc";
			}
			else if (WinSettings.MenuID == "D1" && num == 0 && winFpSpread.ActiveSheetIndex == 7)
			{
				num++;
				flagStop = "II. Hàng hóa, dịch vụ bán ra trong kỳ";
			}
			else if (WinSettings.MenuID == "Z5" && num == 0 && winFpSpread.ActiveSheetIndex == 6)
			{
				num++;
				flagStop = "II. Hàng hóa, dịch vụ bán ra trong kỳ";
			}
			else if (WinSettings.MenuID == "P7" && (num == 1 || num == 2))
			{
				num++;
			}
			else if (WinSettings.MenuID == "Q5" && num < 20)
			{
				if (num == 3)
				{
					num++;
				}
				num++;
			}
			else
			{
				if (!(WinSettings.MenuID == "Y8") || num != 1)
				{
					break;
				}
				num++;
				rowStartNext = 0;
			}
		}
		activeSheet.FpSpread.Sheets[0].AutoCalculation = true;
		ResetValueCombobox(winFpSpread);
		WinDoWorkEvent.FinishImportExcel(winFpSpread, WinSettings.ToKhaiInfo);
		HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "All Done - " + DateTime.Now.ToString("hh:mm:ss"));
	}

	private void ResetValueCombobox(FpSpread winFpSpread)
	{
		//IL_0066: Unknown result type (might be due to invalid IL or missing references)
		//IL_006d: Expected O, but got Unknown
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c5: Expected O, but got Unknown
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			string text = "";
			string text2 = "";
			string text3 = "";
			int num = 0;
			string text4 = "";
			string text5 = "";
			string pathCtToKhaiCheck = "";
			string pathCtMaToKhaiCheck = "";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex];
			XmlDocument val = new XmlDocument();
			val.Load("ValidCombobox.xml");
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Items/Item");
			if (val2 == null || val2.Count == 0)
			{
				return;
			}
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				try
				{
					text4 = ((val3.Attributes["Dynamic"] != null) ? ((XmlNode)val3.Attributes["Dynamic"]).Value.Trim() : "");
					num = ((val3.Attributes["SheetIndex"] != null) ? Convert.ToInt32(((XmlNode)val3.Attributes["SheetIndex"]).Value.Trim()) : (-1));
					text3 = ((val3.Attributes["ID"] != null) ? ((XmlNode)val3.Attributes["ID"]).Value.Trim() : "");
					pathCtToKhaiCheck = ((val3.Attributes["CtPath"] != null) ? ((XmlNode)val3.Attributes["CtPath"]).Value.Trim() : "");
					pathCtMaToKhaiCheck = ((val3.Attributes["CtPathMa"] != null) ? ((XmlNode)val3.Attributes["CtPathMa"]).Value.Trim() : "");
					text5 = ((val3.Attributes["TablePath"] != null) ? ((XmlNode)val3.Attributes["TablePath"]).Value.Trim() : "");
					if (!text3.Equals(WinSettings.MenuID) || num != activeSheetIndex)
					{
						continue;
					}
					if (text4.Equals("0"))
					{
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCtToKhaiCheck));
						if (hTKKCMCCellInfo != null)
						{
							text = winFpSpread.ActiveSheet.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Text;
						}
						continue;
					}
					for (int num2 = 0; num2 < hTKKCMCSheetInfo.Tables.Count; num2++)
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[num2];
						if (!text5.Equals(hTKKCMCTableInfo.TablePath))
						{
							continue;
						}
						HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
						HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
						int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
						List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
						for (int num3 = 0; num3 < rowsValues.Count; num3++)
						{
							HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowsInfo.Cells.Find((HTKKCMCTableCellInfo cell) => cell.Path.Equals(pathCtToKhaiCheck));
							HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = rowsInfo.Cells.Find((HTKKCMCTableCellInfo cell) => cell.Path.Equals(pathCtMaToKhaiCheck));
							if (hTKKCMCTableCellInfo == null)
							{
								continue;
							}
							if (pathCtMaToKhaiCheck.Equals(""))
							{
								text = winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart + num3 - 1, HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo.indexColumnSelectedValue) - 1].Text;
								continue;
							}
							text = winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart + num3 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Text;
							text2 = winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart + num3 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text;
							if (!isValidMa(text, text2, hTKKCMCTableCellInfo.Controltype))
							{
								winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart + num3 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Text = "";
								winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart + num3 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text = "";
							}
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool isValidMa(string ma, string ten, string controlType)
	{
		try
		{
			bool result = false;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string value;
			if (controlType.Equals("68"))
			{
				dictionary = WinSpreadProcess.GetDanhMuc("GetDMHopDong1");
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && ten.Equals(value))
				{
					return true;
				}
			}
			else if (controlType.Equals("44"))
			{
				dictionary = WinSpreadProcess.GetDanhMuc("GetDMMucDich");
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && ten.Equals(value))
				{
					return true;
				}
			}
			else if (controlType.Equals("45"))
			{
				dictionary = WinSpreadProcess.GetDanhMuc("GetDMTrangThai");
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && ten.Equals(value))
				{
					return true;
				}
			}
			else if (controlType.Equals("3"))
			{
				dictionary = HTKKWinDictionaryDiaDanh.QuanHe;
				if (dictionary != null && dictionary.TryGetValue(ten, out value) && ma.Equals(value))
				{
					return true;
				}
			}
			else if (controlType.IndexOf("DC_") > -1)
			{
				dictionary = WinSpreadProcess.GetDMDungChung(controlType);
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && ten.Equals(value))
				{
					return true;
				}
			}
			else if (controlType.Equals("87"))
			{
				dictionary = WinSpreadProcess.DMNnkdNtnn();
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && ten.Equals(value))
				{
					return true;
				}
			}
			else if (controlType.Equals("76"))
			{
				dictionary = WinSpreadProcess.GetDanhMuc("GetDMHoaDonDT");
				if (dictionary != null && dictionary.TryGetValue(ten, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ma.Equals(value2))
					{
						return true;
					}
				}
			}
			else if (controlType.Equals("73"))
			{
				dictionary = WinSpreadProcess.DMLoaiTienTe();
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			else if (controlType.Equals("56"))
			{
				dictionary = WinSpreadProcess.GetDMQuocGia();
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			else if (controlType.Equals("93"))
			{
				dictionary = WinSpreadProcess.GetDMNganHang();
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			else if (controlType.Equals("98"))
			{
				dictionary = WinSpreadProcess.GetDMNganHangKBNN();
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			else if (controlType.Equals("106"))
			{
				dictionary = WinSpreadProcess.GetDMDynamic("DM_NganHang_Moi.xml");
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			else if (controlType.Equals("108"))
			{
				dictionary = WinSpreadProcess.GetDMDynamic("DM_CQHQ.xml");
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			else
			{
				if (!controlType.Equals("67"))
				{
					return true;
				}
				dictionary = WinSpreadProcess.GetDMQuocTichFull();
				if (dictionary != null && dictionary.TryGetValue(ma, out value) && value != "")
				{
					string value2 = value.Split(new char[1] { '-' })[0].Trim();
					if (ten.Equals(value2))
					{
						return true;
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	private void btnTaiDuLieu_Click(object sender, EventArgs e)
	{
		try
		{
			HTKKFile.LogRun("start tai bang ke", DateTime.Now.ToString("hh:mm:ss"));
			winFpSpread.ActiveSheet.AutoCalculation = false;
			winFpSpread.SuspendLayout();
			ProgressForm progressForm = new ProgressForm("Đang tải dữ liệu...");
			progressForm.Argument = false;
			progressForm.DoWork += form_DoWorkTaiBangKe;
			WinSpreadProcess.StartForm(progressForm);
			if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].TaiDuLieuLon != "1" && WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].TaiDuLieuRatLon != "1")
			{
				WinDoWorkEvent.UpdateObjectValue(winFpSpread, new Dictionary<int, int>());
			}
			if (WinSettings.MenuID == "M4")
			{
				foreach (HTKKCMCSheetInfo sheetInfo in WinSettings.ToKhaiInfo.SheetInfos)
				{
					GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex], sheetInfo);
				}
			}
			winFpSpread.ResumeLayout();
			((Component)(object)progressForm).Dispose();
			((Form)this).Close();
		}
		catch (Exception ex)
		{
			winFpSpread.ResumeLayout();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			((Form)this).Close();
		}
	}

	private void frmWinHTKKOpenFile_Load(object sender, EventArgs e)
	{
	}

	private void frmWinHTKKOpenFile_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
			FormProcess.ChildFormClosed((Form)(object)WinSettings.myWinToKhai, (Control)(object)WinSettings.myWinToKhai.panelChild);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void TaiDuLieuRatLon(FpSpread winFpSpread, string fileName, bool themDuLieu)
	{
		//IL_016a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0120: Unknown result type (might be due to invalid IL or missing references)
		//IL_018a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0194: Expected O, but got Unknown
		int num = 0;
		int activeSheetIndex = winFpSpread.ActiveSheetIndex;
		string status = "";
		HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex];
		if (!themDuLieu)
		{
			WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex] = WinFnProcess.GetDataFromDatafileLarge(activeSheetIndex, getOnlyTemplate: true);
			sInfo = WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex];
		}
		List<HTKKCMCTableInfo> tables = sInfo.Tables;
		FpSpread winFpSpreadBK = new FpSpread();
		((ISupportInitialize)winFpSpreadBK).BeginInit();
		winFpSpread.Name = "winFpSpreadBK";
		((ISupportInitialize)winFpSpreadBK).EndInit();
		try
		{
			Stream stream = new FileStream(fileName, FileMode.Open);
			bool flag = winFpSpreadBK.OpenExcel(stream, ExcelOpenFlags.DataOnly);
			stream.Close();
			if (!flag)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
		}
		catch (Exception ex)
		{
			winFpSpread.ResumeLayout();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		winFpSpreadBK.ActiveSheetIndex = 0;
		((Control)this).Invoke((Delegate)(MethodInvoker)delegate
		{
			//IL_014d: Unknown result type (might be due to invalid IL or missing references)
			if (!themDuLieu)
			{
				foreach (HTKKCMCTableInfo table in sInfo.Tables)
				{
					table.RowsValues.Clear();
				}
			}
			foreach (HTKKCMCTableInfo table2 in sInfo.Tables)
			{
				if (table2.RowsValues.Count == 1 && !GeneralFunctions.checkHaveNewDataOfTable(table2.RowsInfo, table2.RowsValues[0].RowValue))
				{
					table2.RowsValues.Clear();
				}
			}
			sInfo = WinSpreadProcess.GetDataFromExcelLarge1(winFpSpreadBK, sInfo, ref status);
			if (!HTKKFormater.isNullString(status))
			{
				winFpSpread.ResumeLayout();
				string outMessage = "Dữ liệu bảng kê không đúng định dạng tại cell " + status + "\r\nĐề nghị NNT vào sheet HDSD trên bảng kê excel để xem lại định dạng đúng";
				frmMessage frmMessage2 = new frmMessage("", outMessage, "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
			}
			else
			{
				foreach (HTKKCMCTableInfo table3 in sInfo.Tables)
				{
					if (table3.RowsValues.Count == 0)
					{
						HTKKCMCRowValue rowValueDefault = GeneralFunctions.GetRowValueDefault(table3.RowsInfo);
						table3.RowsValues.Add(rowValueDefault);
					}
				}
				foreach (HTKKCMCTableInfo table4 in sInfo.Tables)
				{
					int num2 = 1;
					for (int i = 0; i < table4.RowsValues.Count; i++)
					{
						table4.RowsValues[i].RowValue[0] = num2.ToString();
						num2++;
					}
				}
				WinDoWorkEvent.FinishImportExcel(winFpSpread, WinSettings.ToKhaiInfo);
			}
		});
	}

	private void xoaDuLieuSview(HTKKCMCSheetInfo sInfo)
	{
		for (int num = sInfo.Tables.Count - 1; num >= 0; num--)
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			GeneralFunctions.RemoveRowOnSView(winFpSpread.ActiveSheet, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart, hTKKCMCTableInfo.RowsValues.Count - 1);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > dataLocationTemplateRowStart)
				{
					table.Increment((hTKKCMCTableInfo.RowsValues.Count - 1) * -1);
				}
			}
			if (hTKKCMCTableInfo.Footer != null && hTKKCMCTableInfo.Footer.StaticCells.Count > 0)
			{
				hTKKCMCTableInfo.IncrementFooter(dataLocationTemplateRowStart, (hTKKCMCTableInfo.RowsValues.Count - 1) * -1);
			}
			sInfo.IncrementStaticCell(dataLocationTemplateRowStart, (hTKKCMCTableInfo.RowsValues.Count - 1) * -1);
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0427: Unknown result type (might be due to invalid IL or missing references)
		//IL_0431: Expected O, but got Unknown
		//IL_04de: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e8: Expected O, but got Unknown
		//IL_060d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0617: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmWinHTKKOpenFile));
		rdThemDuLieu = new RadioButton();
		rdThemXoaDuLieuDaCo = new RadioButton();
		panel1 = new Panel();
		btnBrow = new Button();
		txtFileNameLoading = new TextBox();
		btnDong = new Button();
		btnTaiDuLieu = new Button();
		panel2 = new Panel();
		label3 = new Label();
		((Control)panel1).SuspendLayout();
		((Control)panel2).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)rdThemDuLieu).AutoSize = true;
		((Control)rdThemDuLieu).Location = new Point(7, 7);
		((Control)rdThemDuLieu).Name = "rdThemDuLieu";
		((Control)rdThemDuLieu).Size = new Size(172, 17);
		((Control)rdThemDuLieu).TabIndex = 0;
		((Control)rdThemDuLieu).Text = "Thêm dữ liệu vào dữ liệu đã có";
		((ButtonBase)rdThemDuLieu).UseVisualStyleBackColor = true;
		((Control)rdThemXoaDuLieuDaCo).AutoSize = true;
		rdThemXoaDuLieuDaCo.Checked = true;
		((Control)rdThemXoaDuLieuDaCo).Location = new Point(7, 30);
		((Control)rdThemXoaDuLieuDaCo).Name = "rdThemXoaDuLieuDaCo";
		((Control)rdThemXoaDuLieuDaCo).Size = new Size(193, 17);
		((Control)rdThemXoaDuLieuDaCo).TabIndex = 1;
		rdThemXoaDuLieuDaCo.TabStop = true;
		((Control)rdThemXoaDuLieuDaCo).Text = "Thêm dữ liệu mới, xóa dữ liệu đã có";
		((ButtonBase)rdThemXoaDuLieuDaCo).UseVisualStyleBackColor = true;
		panel1.BorderStyle = (BorderStyle)1;
		((Control)panel1).Controls.Add((Control)(object)btnBrow);
		((Control)panel1).Controls.Add((Control)(object)txtFileNameLoading);
		((Control)panel1).Controls.Add((Control)(object)rdThemDuLieu);
		((Control)panel1).Controls.Add((Control)(object)rdThemXoaDuLieuDaCo);
		((Control)panel1).Location = new Point(5, 27);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(354, 84);
		((Control)panel1).TabIndex = 2;
		((Control)btnBrow).Location = new Point(302, 52);
		((Control)btnBrow).Name = "btnBrow";
		((Control)btnBrow).Size = new Size(44, 25);
		((Control)btnBrow).TabIndex = 3;
		((Control)btnBrow).Text = "...";
		((ButtonBase)btnBrow).UseVisualStyleBackColor = true;
		((Control)btnBrow).Click += btnBrow_Click;
		((Control)txtFileNameLoading).Location = new Point(7, 53);
		((Control)txtFileNameLoading).Name = "txtFileNameLoading";
		((Control)txtFileNameLoading).Size = new Size(288, 20);
		((Control)txtFileNameLoading).TabIndex = 2;
		((Control)btnDong).Location = new Point(272, 118);
		((Control)btnDong).Name = "btnDong";
		((Control)btnDong).Size = new Size(80, 25);
		((Control)btnDong).TabIndex = 5;
		((Control)btnDong).Text = "Đó&ng";
		((ButtonBase)btnDong).UseVisualStyleBackColor = true;
		((Control)btnDong).Click += btnDong_Click;
		((Control)btnTaiDuLieu).Location = new Point(186, 118);
		((Control)btnTaiDuLieu).Name = "btnTaiDuLieu";
		((Control)btnTaiDuLieu).Size = new Size(80, 25);
		((Control)btnTaiDuLieu).TabIndex = 4;
		((Control)btnTaiDuLieu).Text = "&Tải dữ liệu";
		((ButtonBase)btnTaiDuLieu).UseVisualStyleBackColor = true;
		((Control)btnTaiDuLieu).Click += btnTaiDuLieu_Click;
		((Control)panel2).BackgroundImage = (Image)componentResourceManager.GetObject("panel2.BackgroundImage");
		((Control)panel2).BackgroundImageLayout = (ImageLayout)3;
		((Control)panel2).Controls.Add((Control)(object)label3);
		((Control)panel2).Dock = (DockStyle)1;
		((Control)panel2).Location = new Point(0, 0);
		((Control)panel2).Name = "panel2";
		((Control)panel2).Size = new Size(364, 21);
		((Control)panel2).TabIndex = 27;
		((Control)label3).AutoSize = true;
		((Control)label3).BackColor = Color.Transparent;
		((Control)label3).Font = new Font("Tahoma", 8.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label3).ForeColor = Color.White;
		((Control)label3).Location = new Point(1, 3);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(80, 13);
		((Control)label3).TabIndex = 0;
		((Control)label3).Text = "Thêm dữ liệu";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(364, 151);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel2);
		((Control)this).Controls.Add((Control)(object)btnDong);
		((Control)this).Controls.Add((Control)(object)btnTaiDuLieu);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmWinHTKKOpenFile";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thêm dữ liệu";
		((Form)this).FormClosed += new FormClosedEventHandler(frmWinHTKKOpenFile_FormClosed);
		((Form)this).Load += frmWinHTKKOpenFile_Load;
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)panel2).ResumeLayout(false);
		((Control)panel2).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
