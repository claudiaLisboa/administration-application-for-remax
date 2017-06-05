using System.Collections.Generic;
using System.Linq;

namespace AdmAppRemax.Business
{
    public class HouseHistoryList
    {
        private Dictionary<int, HouseHistory> _list = new Dictionary<int, HouseHistory>();

        public Dictionary<int, HouseHistory>.ValueCollection Elements
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

        public bool Add(HouseHistory houseHistory)
        {
            throw new System.NotImplementedException();
        }

        public HouseHistory Find()
        {
            throw new System.NotImplementedException();
        }

        public bool Remove()
        {
            throw new System.NotImplementedException();
        }
    }
}
