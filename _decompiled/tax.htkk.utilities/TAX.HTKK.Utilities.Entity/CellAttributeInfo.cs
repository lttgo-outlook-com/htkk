namespace TAX.HTKK.Utilities.Entity;

public class CellAttributeInfo
{
	public string C1 { get; set; }

	public string C2 { get; set; }

	public string Controltype { get; set; }

	public string ParentValue { get; set; }

	public CellAttributeInfo()
	{
		C1 = string.Empty;
		C2 = string.Empty;
		Controltype = string.Empty;
		ParentValue = string.Empty;
	}
}
