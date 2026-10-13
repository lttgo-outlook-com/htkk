using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using FarPoint.Excel;
using FarPoint.Win.Spread;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmWinSelectExport : Form
{
	private FpSpread winFpSpread = null;

	private static string PathSave = string.Empty;

	private static XmlDocument xmlExportContent;

	private IContainer components = null;

	private Panel panel1;

	private Button btnClose;

	private Button btnExport;

	private RadioButton radKetXuatExcel;

	private RadioButton radKetXuatXML;

	private FpSpread fpSpreadExp;

	private SheetView fpSpreadExp_Sheet1;

	private Panel panel2;

	private Label label3;

	public frmWinSelectExport(FpSpread _winFpSpread)
	{
		InitializeComponent();
		winFpSpread = _winFpSpread;
		winFpSpread.ResumeLayout();
	}

	private void btnClose_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void form_DoWorkExportExcel(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		try
		{
			if (!Directory.Exists(HTKKConfig.PATH_TEMP))
			{
				Directory.CreateDirectory(HTKKConfig.PATH_TEMP);
			}
			winFpSpread.SaveExcel(PathSave, ExcelSaveFlags.DataOnly);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void form_DoWorkExportXML(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		try
		{
			if (!Directory.Exists(HTKKConfig.PATH_TEMP))
			{
				Directory.CreateDirectory(HTKKConfig.PATH_TEMP);
			}
			string text = WinDoWorkEvent.WinDoWorkExportXML(winFpSpread, PathSave);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnExport_Click(object sender, EventArgs e)
	{
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Expected O, but got Unknown
		//IL_0312: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Invalid comparison between Unknown and I4
		//IL_0379: Unknown result type (might be due to invalid IL or missing references)
		//IL_037f: Invalid comparison between Unknown and I4
		//IL_05ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_0464: Unknown result type (might be due to invalid IL or missing references)
		//IL_0531: Unknown result type (might be due to invalid IL or missing references)
		//IL_050c: Unknown result type (might be due to invalid IL or missing references)
		//IL_070e: Unknown result type (might be due to invalid IL or missing references)
		//IL_06e9: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			winFpSpread.SuspendLayout();
			((Form)this).Close();
			SaveFileDialog val = new SaveFileDialog();
			((FileDialog)val).AddExtension = true;
			((FileDialog)val).RestoreDirectory = true;
			if (!Directory.Exists(HTKKConfig.PATH_TEMP))
			{
				try
				{
					Directory.CreateDirectory(HTKKConfig.PATH_TEMP);
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			((FileDialog)val).InitialDirectory = HTKKConfig.PATH_TEMP;
			if (radKetXuatExcel.Checked)
			{
				((FileDialog)val).Title = "Tệp Excel kết xuất tới thư mục " + HTKKConfig.PATH_TEMP;
				((FileDialog)val).FileName = WinDoWorkEvent.GetFileNameExcel(WinSettings.ToKhaiInfo, WinSettings.TTPeriodInfo, WinSettings.HTKKHeaderInfo);
				((FileDialog)val).DefaultExt = "xlsx";
				((FileDialog)val).Filter = "(*.xlsx)|*.xlsx|(*.xls)|*.xls";
				if ((int)((CommonDialog)val).ShowDialog() == 1)
				{
					try
					{
						for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
						{
							if (winFpSpread.Sheets[i].SheetName.Contains("/"))
							{
								winFpSpread.Sheets[i].SheetName = winFpSpread.Sheets[i].SheetName.Replace("/", "__");
							}
						}
						winFpSpread.SuspendLayout();
						HTKKFile.LogRun("Export Excel", "Start Process data");
						WinDoWorkEvent.SetFPSpreadExpExcel(winFpSpread, winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
						HTKKFile.LogRun("Export Excel", "Start save data");
						if (((FileDialog)val).FileName.Contains(".xlsx"))
						{
							winFpSpread.SaveExcel(((FileDialog)val).FileName, ExcelSaveFlags.UseOOXMLFormat);
						}
						else
						{
							winFpSpread.SaveExcel(((FileDialog)val).FileName, ExcelSaveFlags.NoFormulas);
						}
						HTKKFile.LogRun("Export Excel", "Start Restore data");
						WinDoWorkEvent.SetFPSpreadExpExcelAfterExport(winFpSpread, winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
						HTKKFile.LogRun("Export Excel", "End");
						for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
						{
							if (winFpSpread.Sheets[i].SheetName.Contains("__"))
							{
								winFpSpread.Sheets[i].SheetName = winFpSpread.Sheets[i].SheetName.Replace("__", "/");
							}
						}
						winFpSpread.ResumeLayout();
					}
					catch (Exception ex)
					{
						winFpSpread.ResumeLayout();
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
					frmMessage frmMessage = new frmMessage("0153", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
				((Component)(object)val).Dispose();
				val = null;
			}
			if (radKetXuatXML.Checked)
			{
				((FileDialog)val).Title = "Tệp XML kết xuất tới thư mục " + HTKKConfig.PATH_TEMP;
				((FileDialog)val).FileName = WinDoWorkEvent.getFileName(WinSettings.ToKhaiInfo, WinSettings.TTPeriodInfo, WinSettings.HTKKHeaderInfo);
				((FileDialog)val).Filter = "Xml file (*.xml)|*.xml";
				if ((int)((CommonDialog)val).ShowDialog() == 1)
				{
					PathSave = ((FileDialog)val).FileName;
					winFpSpread.SuspendLayout();
					GC.Collect();
					GC.WaitForPendingFinalizers();
					if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
					{
						if (File.Exists(PathSave))
						{
							File.Delete(PathSave);
						}
						string text = "tmpExport.xml";
						string text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
						string text3 = text2 + "\\" + text;
						string value = WinDoWorkEvent.WinDoWorkExportXML_Large(winFpSpread, WinSettings.ToKhaiInfo.SheetInfos, text3);
						if (!string.IsNullOrEmpty(value))
						{
							frmMessage frmMessage = new frmMessage("0279", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						WinDoWorkEvent.removeNilAttrIfValueOk(text3, PathSave);
						if (string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) || (!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) & WinDoWorkEvent.ValidateXmlAgainstXsd(PathSave, WinSettings.ToKhaiInfo.XsdFileName)))
						{
							if (WinSettings.MenuID == "U8")
							{
								UpdateXmlValueByPath(PathSave, "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/soLan", "0");
							}
							frmMessage frmMessage = new frmMessage("0280", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
						else
						{
							frmMessage frmMessage = new frmMessage("0279", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							if (File.Exists(PathSave))
							{
								File.Delete(PathSave);
							}
						}
						return;
					}
					string text4 = "";
					text4 = WinDoWorkEvent.WinDoWorkExportXML(winFpSpread, PathSave);
					winFpSpread.ResumeLayout();
					if (string.IsNullOrEmpty(text4))
					{
						frmMessage frmMessage = new frmMessage("0280", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						editNode(PathSave);
					}
					else
					{
						bool flag = false;
						HTKKCMCSheetInfo hTKKCMCSheetInfo = ((WinSettings.ToKhaiInfo.SheetInfos.Count > 0) ? WinSettings.ToKhaiInfo.SheetInfos[0] : null);
						if (hTKKCMCSheetInfo != null && WinSettings.MenuID != "Q5" && WinSettings.MenuID != "Q6")
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo.Tables.Count > 0) ? hTKKCMCSheetInfo.Tables[0] : null);
							if (hTKKCMCTableInfo != null)
							{
								HTKKCMCRowValue hTKKCMCRowValue = ((hTKKCMCTableInfo.RowsValues.Count > 0) ? hTKKCMCTableInfo.RowsValues[0] : null);
								if (hTKKCMCRowValue != null && !GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, hTKKCMCRowValue.RowValue.ToArray()))
								{
									flag = true;
								}
							}
						}
						if (flag && WinSettings.TTinChung.loaiTKhai != "B")
						{
							frmMessage frmMessage = new frmMessage("0460", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
						else
						{
							frmMessage frmMessage = new frmMessage("0279", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
					}
				}
			}
			if (winFpSpread.IsLayoutSuspended)
			{
				winFpSpread.ResumeLayout();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			((Form)this).Close();
		}
	}

	public void editNode(string file)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ((WinSettings.ToKhaiInfo.SheetInfos.Count > 0) ? WinSettings.ToKhaiInfo.SheetInfos[0] : null);
			if (hTKKCMCSheetInfo == null || !(WinSettings.MenuID == "I7"))
			{
				return;
			}
			string text = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDung01LP_TT63/ChiTietLePhiTT63/ma_LoaiLePhi_TT63";
			XDocument val = XDocument.Load(file);
			XNamespace ns = XNamespace.op_Implicit("http://kekhaithue.gdt.gov.vn/TKhaiThue");
			string[] array = text.Split(new char[1] { '/' });
			List<XElement> list = new List<XElement>();
			list.Add(val.Root);
			List<XElement> list2 = list;
			for (int i = 1; i < array.Length; i++)
			{
				string nodeName = array[i];
				list2 = list2.SelectMany((XElement x) => ((XContainer)x).Elements(ns + nodeName)).ToList();
				if (list2.Count == 0)
				{
					HTKKFile.Log("Không tìm thấy path: " + nodeName);
					return;
				}
			}
			foreach (XElement item in list2)
			{
				item.Value = item.Value.Replace("-", "");
			}
			val.Save(file);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmWinSelectExport_Load(object sender, EventArgs e)
	{
	}

	private void frmWinSelectExport_FormClosed(object sender, FormClosedEventArgs e)
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

	public static bool UpdateXmlValueByPath(string filePath, string xmlPath, string newValue)
	{
		try
		{
			if (string.IsNullOrEmpty(filePath) || filePath.Trim().Length == 0 || string.IsNullOrEmpty(xmlPath) || xmlPath.Trim().Length == 0 || !File.Exists(filePath))
			{
				return false;
			}
			XDocument val = XDocument.Load(filePath, (LoadOptions)1);
			if (val.Root == null)
			{
				return false;
			}
			string[] pathParts = xmlPath.Split(new char[1] { '/' }, StringSplitOptions.RemoveEmptyEntries);
			if (pathParts.Length == 0)
			{
				return false;
			}
			XElement val2 = val.Root;
			int num = (string.Equals(pathParts[0], val2.Name.LocalName, StringComparison.Ordinal) ? 1 : 0);
			int i;
			for (i = num; i < pathParts.Length; i++)
			{
				val2 = ((XContainer)val2).Elements().FirstOrDefault((Func<XElement, bool>)((XElement element) => string.Equals(element.Name.LocalName, pathParts[i], StringComparison.Ordinal)));
				if (val2 == null)
				{
					return false;
				}
			}
			val2.Value = newValue ?? string.Empty;
			val.Save(filePath, (SaveOptions)1);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
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
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_03f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03fe: Expected O, but got Unknown
		//IL_04ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b5: Expected O, but got Unknown
		//IL_05fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0606: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmWinSelectExport));
		panel1 = new Panel();
		radKetXuatExcel = new RadioButton();
		radKetXuatXML = new RadioButton();
		btnClose = new Button();
		btnExport = new Button();
		fpSpreadExp = new FpSpread();
		fpSpreadExp_Sheet1 = new SheetView();
		panel2 = new Panel();
		label3 = new Label();
		((Control)panel1).SuspendLayout();
		((ISupportInitialize)fpSpreadExp).BeginInit();
		((ISupportInitialize)fpSpreadExp_Sheet1).BeginInit();
		((Control)panel2).SuspendLayout();
		((Control)this).SuspendLayout();
		panel1.BorderStyle = (BorderStyle)1;
		((Control)panel1).Controls.Add((Control)(object)radKetXuatExcel);
		((Control)panel1).Controls.Add((Control)(object)radKetXuatXML);
		((Control)panel1).Location = new Point(5, 27);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(286, 66);
		((Control)panel1).TabIndex = 3;
		((Control)radKetXuatExcel).AutoSize = true;
		((Control)radKetXuatExcel).Location = new Point(18, 35);
		((Control)radKetXuatExcel).Name = "radKetXuatExcel";
		((Control)radKetXuatExcel).Size = new Size(93, 17);
		((Control)radKetXuatExcel).TabIndex = 1;
		((Control)radKetXuatExcel).Text = "Kết xuất Excel";
		((ButtonBase)radKetXuatExcel).UseVisualStyleBackColor = true;
		((Control)radKetXuatXML).AutoSize = true;
		radKetXuatXML.Checked = true;
		((Control)radKetXuatXML).Location = new Point(18, 9);
		((Control)radKetXuatXML).Name = "radKetXuatXML";
		((Control)radKetXuatXML).Size = new Size(89, 17);
		((Control)radKetXuatXML).TabIndex = 0;
		radKetXuatXML.TabStop = true;
		((Control)radKetXuatXML).Text = "Kết xuất XML";
		((ButtonBase)radKetXuatXML).UseVisualStyleBackColor = true;
		((Control)btnClose).Location = new Point(149, 97);
		((Control)btnClose).Name = "btnClose";
		((Control)btnClose).Size = new Size(80, 25);
		((Control)btnClose).TabIndex = 5;
		((Control)btnClose).Text = "Đó&ng";
		((ButtonBase)btnClose).UseVisualStyleBackColor = true;
		((Control)btnClose).Click += btnClose_Click;
		((Control)btnExport).Location = new Point(62, 97);
		((Control)btnExport).Name = "btnExport";
		((Control)btnExport).Size = new Size(80, 25);
		((Control)btnExport).TabIndex = 4;
		((Control)btnExport).Text = "&Kết xuất";
		((ButtonBase)btnExport).UseVisualStyleBackColor = true;
		((Control)btnExport).Click += btnExport_Click;
		((Control)fpSpreadExp).AccessibleDescription = "";
		fpSpreadExp.BorderStyle = (BorderStyle)0;
		((Control)fpSpreadExp).Location = new Point(0, 156);
		fpSpreadExp.Name = "fpSpreadExp";
		fpSpreadExp.Sheets.AddRange(new SheetView[1] { fpSpreadExp_Sheet1 });
		((Control)fpSpreadExp).Size = new Size(10, 10);
		((Control)fpSpreadExp).TabIndex = 4;
		((Control)fpSpreadExp).Visible = false;
		fpSpreadExp_Sheet1.Reset();
		fpSpreadExp_Sheet1.SheetName = "Sheet1";
		((Control)panel2).BackgroundImage = (Image)componentResourceManager.GetObject("panel2.BackgroundImage");
		((Control)panel2).BackgroundImageLayout = (ImageLayout)3;
		((Control)panel2).Controls.Add((Control)(object)label3);
		((Control)panel2).Dock = (DockStyle)1;
		((Control)panel2).Location = new Point(0, 0);
		((Control)panel2).Name = "panel2";
		((Control)panel2).Size = new Size(295, 21);
		((Control)panel2).TabIndex = 27;
		((Control)label3).AutoSize = true;
		((Control)label3).BackColor = Color.Transparent;
		((Control)label3).Font = new Font("Tahoma", 8.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label3).ForeColor = Color.White;
		((Control)label3).Location = new Point(1, 3);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(75, 13);
		((Control)label3).TabIndex = 0;
		((Control)label3).Text = "Kết xuất file";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(295, 126);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel2);
		((Control)this).Controls.Add((Control)(object)btnClose);
		((Control)this).Controls.Add((Control)(object)btnExport);
		((Control)this).Controls.Add((Control)(object)fpSpreadExp);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).MaximizeBox = false;
		((Control)this).Name = "frmWinSelectExport";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Kết xuất file";
		((Form)this).Load += frmWinSelectExport_Load;
		((Form)this).FormClosed += new FormClosedEventHandler(frmWinSelectExport_FormClosed);
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((ISupportInitialize)fpSpreadExp).EndInit();
		((ISupportInitialize)fpSpreadExp_Sheet1).EndInit();
		((Control)panel2).ResumeLayout(false);
		((Control)panel2).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
