namespace Whatsapp.Marketing.Entities
{
    public class Contact
    {
        public Contact(string name, string number, string secondNumber)
        {
            Name = name;
            Number = number;
            SecondNumber = secondNumber;
        }
        public string Name { get; private set; }
        public string Number { get; private set; }
        public string SecondNumber { get; private set; }
    }
}