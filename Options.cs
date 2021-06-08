using CommandLine;

namespace KnockOutlook
{
	class Options
	{
		[Option("operation", Required = true)]
		public string Operation
		{
			get;
			set;
		}

		[Option("keyword")]
		public string Keyword
		{
			get;
			set;
		}

		[Option("id")]
		public string EntryID
		{
			get;
			set;
		}

		[Option("bypass")]
		public bool Bypass
		{
			get;
			set;
		}
	}
}
