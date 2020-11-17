namespace Whatsapp.Marketing.Entities
{
    public class Contact
    {
        public Contact(string name, string number, string secondNumber, int row)
        {
            Name = name;
            Number = number;
            SecondNumber = secondNumber;
            Row = row;
        }
        public string Name { get; private set; }
        public string Number { get; private set; }
        public string SecondNumber { get; private set; }
        public int Row { get; private set; }

        public override string ToString()
        {
            return $"{Name} - {Number} / {SecondNumber} - Linha: {Row}";
        }
    }
}