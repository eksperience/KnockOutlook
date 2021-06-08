using System.Collections.Generic;

namespace KnockOutlook.Classes
{
	class Mail
	{
		public string ID;
		public string Timestamp;
		public string Subject;
		public string From;
		public List<string> To;
		public List<string> Attachments;

		public Mail(string id, string timestamp, string subject, string from)
		{
			ID = id;
			Timestamp = timestamp;
			Subject = subject;
			From = from;
			To = new List<string>();
		}
	}
}
