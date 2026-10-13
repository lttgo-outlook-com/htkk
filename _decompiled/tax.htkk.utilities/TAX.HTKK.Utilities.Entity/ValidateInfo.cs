namespace TAX.HTKK.Utilities.Entity;

public class ValidateInfo
{
	public string CellIndex { get; set; }

	public string TypeValidate { get; set; }

	public string ListCellRange { get; set; }

	public string Range { get; set; }

	public string MessageError { get; set; }

	public ValidateInfo()
	{
		CellIndex = string.Empty;
		TypeValidate = string.Empty;
		ListCellRange = string.Empty;
		Range = string.Empty;
		MessageError = string.Empty;
	}
}
