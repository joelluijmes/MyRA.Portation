using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.Models
{
    internal sealed class ItemModel
    {
        [ExcelProperty]
        public string Name { get; set; }

        private bool Equals(ItemModel other)
        {
            return string.Equals(Name, other.Name);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;

            return obj is ItemModel && Equals((ItemModel) obj);
        }

        public override int GetHashCode()
        {
            return Name != null ? Name.GetHashCode() : 0;
        }

        public static bool operator ==(ItemModel left, ItemModel right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(ItemModel left, ItemModel right)
        {
            return !Equals(left, right);
        }
    }
}