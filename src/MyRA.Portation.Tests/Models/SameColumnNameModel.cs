using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.Models
{
    internal sealed class SameColumnNameModel
    {
        public const string COLUMN_NAME = "Name";

        [ExcelProperty(1, ColumnName = COLUMN_NAME)]
        public string Firstname { get; set; }

        [ExcelProperty(2, ColumnName = COLUMN_NAME)]
        public string Lastname { get; set; }

        private bool Equals(SameColumnNameModel other)
        {
            return string.Equals(Firstname, other.Firstname) && string.Equals(Lastname, other.Lastname);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;

            return obj is SameColumnNameModel && Equals((SameColumnNameModel) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Firstname != null ? Firstname.GetHashCode() : 0) * 397) ^ (Lastname != null ? Lastname.GetHashCode() : 0);
            }
        }

        public static bool operator ==(SameColumnNameModel left, SameColumnNameModel right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(SameColumnNameModel left, SameColumnNameModel right)
        {
            return !Equals(left, right);
        }
    }
}
