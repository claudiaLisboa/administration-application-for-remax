namespace AdmAppRemax.Business
{
    public class Sale
    {
        private int _saleId;
        private House _house;
        private Employee _agent;
        private Client _buyer;
        private Client _seller;
        private double _price;

        // Default constructor
        public Sale()
        {
            _saleId = default(int);
            _house = default(House);
            _agent = default(Employee);
            _buyer = default(Client);
            _seller = default(Client);
            _price = default(double);
        }

        public int SaleId
        {
            get { return _saleId; }
            set { _saleId = value; }
        }

        public double Price
        {
            get { return _price; }
            set { _price = value; }
        }

        public Employee Agent
        {
            get { return _agent; }
            set { _agent = value; }
        }

        public Client Buyer
        {
            get { return _buyer; }
            set { _buyer = value; }
        }

        public House House
        {
            get { return _house; }
            set { _house = value; }
        }

        public Client Seller
        {
            get { return _seller; }
            set { _seller = value; }
        }
    }
}
