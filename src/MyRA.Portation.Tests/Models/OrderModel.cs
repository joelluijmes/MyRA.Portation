using System.Collections.Generic;
using System.Linq;
using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.Models
{
    [ExcelSheet(SheetName = SHEET_NAME)]
    internal sealed class OrderModel
    {
        public const string SHEET_NAME = "Order";
        public const string PERSON_SHEET_NAME = "Customer";

        [ExcelSheet(SheetName = PERSON_SHEET_NAME)]
        public PersonModel Person { get; set; }

        [ExcelSheet]
        public IList<ItemModel> Items { get; set; }

        private bool Equals(OrderModel other)
        {
            return Items.SequenceEqual(other.Items) && Equals(Person, other.Person);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;

            return obj is OrderModel && Equals((OrderModel) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Items != null ? Items.GetHashCode() : 0) * 397) ^ (Person != null ? Person.GetHashCode() : 0);
            }
        }

        public static bool operator ==(OrderModel left, OrderModel right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OrderModel left, OrderModel right)
        {
            return !Equals(left, right);
        }
    }
}
