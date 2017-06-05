namespace AdmAppRemax.Business
{
    public class Address
    {
        private int _addressId;
        private string _street;
        private string _apartmentNo;
        private string _country;
        private string _postalCode;

        // Default constructor
        public Address()
        {
            _addressId = default(int);
            _street = string.Empty;
            _apartmentNo = string.Empty;
            _country = string.Empty;
            _postalCode = string.Empty;
        }

        public int AddressId
        {
            get { return _addressId; }
            set { _addressId = value; }
        }

        public string ApartmentNo
        {
            get { return _apartmentNo; }
            set { _apartmentNo = value; }
        }

        public string Country
        {
            get { return _country; }
            set { _country = value; }
        }

        public string PostalCode
        {
            get { return _postalCode; }
            set { _postalCode = value; }
        }

        public string Street
        {
            get { return _street; }
            set { _street = value; }
        }
    }
}
