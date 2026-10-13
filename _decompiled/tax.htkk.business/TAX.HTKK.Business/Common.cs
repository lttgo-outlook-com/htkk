using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Business;

public class Common
{
	public List<CacheFormula> formulas = new List<CacheFormula>();

	public void setCacheFormula(int indexSheet, int row, int col, string formula)
	{
		CacheFormula cacheFormula = new CacheFormula();
		cacheFormula.indexSheet = indexSheet;
		cacheFormula.row = row;
		cacheFormula.col = col;
		cacheFormula.formula = formula;
		formulas.Add(cacheFormula);
	}

	public void setCacheFormula(SheetView sView, int indexSheet, int row, int col)
	{
		string formula = sView.Cells[row - 1, col - 1].Formula;
		CacheFormula cacheFormula = new CacheFormula();
		cacheFormula.indexSheet = indexSheet;
		cacheFormula.row = row;
		cacheFormula.col = col;
		cacheFormula.formula = formula;
		formulas.Add(cacheFormula);
	}

	public string getCacheFormula(int indexSheet, int oldRow, int oldCol, int replaceToNewRow)
	{
		try
		{
			if (formulas.Count == 0)
			{
				return "";
			}
			CacheFormula cacheFormula = formulas.FirstOrDefault((CacheFormula x) => x.indexSheet == indexSheet && x.row == oldRow && x.col == oldCol);
			if (replaceToNewRow >= 0)
			{
				return cacheFormula.formula.Replace(oldRow.ToString(), replaceToNewRow.ToString());
			}
			return cacheFormula.formula;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}
}
