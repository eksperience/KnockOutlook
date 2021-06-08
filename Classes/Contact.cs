namespace KnockOutlook.Classes
{
	class Contact
	{
		public string Name;
		public string Email;

		public Contact(string name, string email)
		{
			if (name != email)
			{
				Name = name;
			}

			Email = email;
		}
	}
}
