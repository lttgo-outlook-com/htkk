using System.Drawing;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCValidate
{
	public string LonHon { get; set; }

	public string NhoHon { get; set; }

	public string MinValue { get; set; }

	public string MaxValue { get; set; }

	public string MaxLen { get; set; }

	public string Bang { get; set; }

	public string MessageID { get; set; }

	public Color ColorDefault { get; set; }

	public bool IsReset { get; set; }

	public HTKKCMCValidate()
	{
		LonHon = string.Empty;
		NhoHon = string.Empty;
		MinValue = string.Empty;
		MaxValue = string.Empty;
		MaxLen = string.Empty;
		Bang = string.Empty;
		MessageID = string.Empty;
		ColorDefault = Color.Black;
		IsReset = true;
	}

	public HTKKCMCValidate(HTKKCMCValidate validate)
	{
		if (validate.LonHon != null)
		{
			LonHon = validate.LonHon;
		}
		else
		{
			LonHon = string.Empty;
		}
		if (validate.NhoHon != null)
		{
			NhoHon = validate.NhoHon;
		}
		else
		{
			NhoHon = string.Empty;
		}
		if (validate.MinValue != null)
		{
			MinValue = validate.MinValue;
		}
		else
		{
			MinValue = string.Empty;
		}
		if (validate.MaxValue != null)
		{
			MaxValue = validate.MaxValue;
		}
		else
		{
			MaxValue = string.Empty;
		}
		if (validate.MaxLen != null)
		{
			MaxLen = validate.MaxLen;
		}
		else
		{
			MaxLen = string.Empty;
		}
		if (validate.Bang != null)
		{
			Bang = validate.Bang;
		}
		else
		{
			Bang = string.Empty;
		}
		ColorDefault = validate.ColorDefault;
		IsReset = validate.IsReset;
	}
}
