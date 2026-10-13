using System;

namespace TAX.HTKK.Business;

public class DateRangeText
{
	public DateTime Start;

	public DateTime End;

	public string Text;

	public DateRangeText(DateTime start, DateTime end, string text)
	{
		Start = start;
		End = end;
		Text = text;
	}
}
