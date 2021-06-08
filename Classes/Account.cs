using System.Collections.Generic;

namespace KnockOutlook.Classes
{
	class Account
	{
		public string Address;
		public List<Contact> Contacts;
		public List<Mail> Mails;

		public Account(string address)
		{
			Address = address;
		}
	}
}
