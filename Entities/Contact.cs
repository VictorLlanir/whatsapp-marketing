namespace Whatsapp.Marketing.Entities
{
    public class Contact
    {
        public Contact(string name, string number)
        {
            Name = name;
            Number = number;
        }
        public string Name { get; private set; }
        public string Number { get; private set; }
    }
}