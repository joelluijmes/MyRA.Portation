using System.Collections.Generic;
using System.Linq;
using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.Models
{
    [ExcelSheet(SheetName = SHEET_NAME)]
    internal sealed class StockModel
    {
        public const string SHEET_NAME = "Stock info";

        [ExcelSheet]
        public IList<ItemModel> Items { get; set; }

        [ExcelProperty]
        public string Warehouse { get; set; }

        private bool Equals(StockModel other)
        {
            return Items.SequenceEqual(other.Items) && string.Equals(Warehouse, other.Warehouse);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;

            return obj is StockModel && Equals((StockModel) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Items != null ? Items.GetHashCode() : 0) * 397) ^ (Warehouse != null ? Warehouse.GetHashCode() : 0);
            }
        }

        public static bool operator ==(StockModel left, StockModel right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(StockModel left, StockModel right)
        {
            return !Equals(left, right);
        }
    }
}
