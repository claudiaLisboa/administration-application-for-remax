using System.Collections.Generic;
using System.Linq;

namespace AdmAppRemax.Business
{
    public class SaleList
    {
        private Dictionary<int, Sale> _list = new Dictionary<int, Sale>();

        public Dictionary<int, Sale>.ValueCollection Elements
        {
            get
            {
                return _list.Values;
            }
        }

        public int Quantity
        {
            get
            {
                return _list.Values.Count();
            }
        }

        public bool Add(Sale sale)
        {
            if (_list.ContainsKey(sale.SaleId))
            {
                return false;
            }
            else
            {
                _list.Add(sale.SaleId, sale);
                return true;
            }
        }

        public bool Exists(int saleId)
        {
            return _list.ContainsKey(saleId);
        }

        public Sale Find(int saleId)
        {
            if (Exists(saleId) == true)
            {
                return _list[saleId];
            }
            else
            {
                return null;
            }
        }

        public bool Remove(int saleId)
        {
            return _list.Remove(saleId);
        }
    }
}
