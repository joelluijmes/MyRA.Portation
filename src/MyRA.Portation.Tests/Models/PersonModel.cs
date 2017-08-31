using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.Models
{
    [ExcelSheet(SheetName = SHEET_NAME)]
    internal sealed class PersonModel
    {
        public const string SHEET_NAME = "Persons";

        [ExcelProperty]
        public int Id { get; set; }

        [ExcelProperty]
        public string Firstname { get; set; }

        [ExcelProperty]
        public string Lastname { get; set; }

        public string Phone { get; set; }

        private bool Equals(PersonModel other)
        {
            return string.Equals(Firstname, other.Firstname) && Id == other.Id && string.Equals(Lastname, other.Lastname);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;

            return obj is PersonModel && Equals((PersonModel)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Firstname != null ? Firstname.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ Id;
                hashCode = (hashCode * 397) ^ (Lastname != null ? Lastname.GetHashCode() : 0);
                return hashCode;
            }
        }

        public static bool operator ==(PersonModel left, PersonModel right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(PersonModel left, PersonModel right)
        {
            return !Equals(left, right);
        }
    }
}
