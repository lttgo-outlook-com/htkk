using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

public class PrintCertainPage
{
	public static PrintDocument printDoc;

	private FpSpread winFpSpread;

	private PrintInfo portraitPI;

	private PrintInfo landscapePI;

	private int number = 1;

	private int currentNumber = 1;

	private int totalPrint = 0;

	private int totalPage = 0;

	private int count = 0;

	private int pagePrint = 0;

	private int pageCheck = 1;

	private int pageCount = 0;

	private int pageQueryCount = 0;

	private int numberOfSheet = 0;

	private int currentSheet = 0;

	private List<int> pagesRequired;

	public PrintCertainPage()
	{
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		printDoc = new PrintDocument();
		portraitPI = new PrintInfo();
		landscapePI = new PrintInfo();
		setPropertiesPI(portraitPI, PrintOrientation.Portrait, 150);
		setPropertiesPI(landscapePI, PrintOrientation.Landscape, 100);
	}

	private void setPropertiesPI(PrintInfo _printInfo, PrintOrientation _orientation, int footerHeight)
	{
		try
		{
			_printInfo.PdfWriteMode = PdfWriteMode.New;
			_printInfo.PdfWriteTo = PdfWriteTo.File;
			_printInfo.ShowBorder = false;
			_printInfo.ShowColumnHeader = PrintHeader.Hide;
			_printInfo.ShowRowHeader = PrintHeader.Hide;
			_printInfo.Orientation = _orientation;
			_printInfo.SmartPrintPagesTall = 1;
			_printInfo.SmartPrintPagesWide = 1;
			_printInfo.Preview = true;
			_printInfo.FooterHeight = footerHeight;
			_printInfo.Margin = new PrintMargin(_printInfo.Margin.Left, _printInfo.Margin.Top, _printInfo.Margin.Right, _printInfo.Margin.Bottom, _printInfo.Margin.Header, _printInfo.Margin.Footer);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public PrintCertainPage(FpSpread _winFpSpread, string printerName, List<int> _pagesRequired, int _totalPage)
	{
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		try
		{
			printDoc = new PrintDocument();
			portraitPI = new PrintInfo();
			landscapePI = new PrintInfo();
			winFpSpread = _winFpSpread;
			totalPage = _totalPage;
			printDoc.PrinterSettings.PrinterName = printerName;
			pagesRequired = _pagesRequired;
			setPropertiesPI(portraitPI, PrintOrientation.Portrait, 150);
			setPropertiesPI(landscapePI, PrintOrientation.Landscape, 100);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Printing()
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Expected O, but got Unknown
		try
		{
			printDoc.PrintPage += new PrintPageEventHandler(printDoc_PrintPage);
			printDoc.QueryPageSettings += new QueryPageSettingsEventHandler(printDoc_QueryPageSettings);
			for (int i = 0; i < winFpSpread.Sheets.Count; i++)
			{
				if (winFpSpread.Sheets[i].PrintInfo.Orientation == PrintOrientation.Portrait)
				{
					winFpSpread.Sheets[i].PrintInfo = portraitPI;
				}
				else if (winFpSpread.Sheets[i].PrintInfo.Orientation == PrintOrientation.Landscape)
				{
					winFpSpread.Sheets[i].PrintInfo = landscapePI;
				}
			}
			printDoc.Print();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void PrintPreview()
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		//IL_00fc: Unknown result type (might be due to invalid IL or missing references)
		PrintPreviewDialog val = new PrintPreviewDialog();
		printDoc.PrintPage += new PrintPageEventHandler(printDoc_PrintPage);
		printDoc.QueryPageSettings += new QueryPageSettingsEventHandler(printDoc_QueryPageSettings);
		try
		{
			for (int i = 0; i < winFpSpread.Sheets.Count; i++)
			{
				if (winFpSpread.Sheets[i].PrintInfo.Orientation == PrintOrientation.Portrait)
				{
					winFpSpread.Sheets[i].PrintInfo = portraitPI;
				}
				else if (winFpSpread.Sheets[i].PrintInfo.Orientation == PrintOrientation.Landscape)
				{
					winFpSpread.Sheets[i].PrintInfo = landscapePI;
				}
			}
			val.Document = printDoc;
			((Form)val).WindowState = (FormWindowState)2;
			((Form)val).ShowDialog();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void printDoc_PrintPage(object sender, PrintPageEventArgs e)
	{
		Rectangle rect = new Rectangle(80, 100, e.PageBounds.Width - 50, e.PageBounds.Height - 50);
		for (int i = numberOfSheet; i < winFpSpread.Sheets.Count; i++)
		{
			bool flag = false;
			for (int j = number; j <= count; j++)
			{
				count = winFpSpread.GetOwnerPrintPageCount(e.Graphics, rect, numberOfSheet);
				pagePrint = number;
				if (pagePrint > count)
				{
					pagePrint = 1;
				}
				if (pagesRequired.Contains(pageCount))
				{
					winFpSpread.Sheets[numberOfSheet].PrintInfo.Footer = "Page " + pageCount.ToString() + "/" + totalPage;
					winFpSpread.OwnerPrintDraw(e.Graphics, rect, numberOfSheet, pagePrint);
					flag = true;
					if (pagePrint < count)
					{
						pageCheck++;
						currentNumber = number;
						currentNumber++;
					}
					else
					{
						pageCheck = 1;
					}
					totalPrint++;
					break;
				}
			}
			if (flag)
			{
				break;
			}
		}
		if (totalPrint < pagesRequired.Count)
		{
			e.HasMorePages = true;
			if (pagePrint >= count)
			{
				currentSheet = numberOfSheet;
				currentSheet++;
				currentNumber = 1;
			}
		}
		else
		{
			e.HasMorePages = false;
		}
	}

	private void printDoc_QueryPageSettings(object sender, QueryPageSettingsEventArgs e)
	{
		try
		{
			for (numberOfSheet = currentSheet; numberOfSheet < winFpSpread.Sheets.Count; numberOfSheet++)
			{
				bool flag = false;
				count = winFpSpread.GetPrintPageCount(numberOfSheet);
				for (number = currentNumber; number <= count; number++)
				{
					pageCount++;
					if (pagesRequired.Contains(pageCount))
					{
						if (winFpSpread.Sheets[numberOfSheet].PrintInfo.Orientation == PrintOrientation.Portrait)
						{
							e.PageSettings.Landscape = false;
							flag = true;
						}
						else
						{
							e.PageSettings.Landscape = true;
							flag = true;
						}
						break;
					}
					if (pageCheck >= count && count == 1)
					{
						break;
					}
				}
				if (flag)
				{
					break;
				}
				currentNumber = 1;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
