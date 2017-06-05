namespace AdmAppRemax.Business
{
    public class House
    {
        private int _houseId;
        private BuildingType _buildingType;
        private string _street;
        private string _apartmentNo;
        private string _country;
        private string _postalCode;
        private string _description;
        private string _makingYear;
        private double _area;
        private int _bathrooms;
        private int _bedrooms;
        private double _price;
        private double _tax;
        private Client _seller;

        // Default constructor
        public House()
        {
            _houseId = default(int);
            _buildingType = BuildingType.House;
            _street = string.Empty;
            _apartmentNo = string.Empty;
            _country = string.Empty;
            _postalCode = string.Empty;
            _description = string.Empty;
            _makingYear = string.Empty;
            _area = default(double);
            _bathrooms = default(int);
            _bedrooms = default(int);
            _price = default(double);
            _tax = default(double);
            _seller = default(Client);
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

        public double Area
        {
            get { return _area; }
            set { _area = value; }
        }

        public int Bathrooms
        {
            get { return _bathrooms; }
            set { _bathrooms = value; }
        }

        public int Bedrooms
        {
            get { return _bedrooms; }
            set { _bedrooms = value; }
        }

        public BuildingType BuildingType
        {
            get { return _buildingType; }
            set { _buildingType = value; }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        public int HouseId
        {
            get { return _houseId; }
            set { _houseId = value; }
        }

        public string MakingYear
        {
            get { return _makingYear; }
            set { _makingYear = value; }
        }

        public double Price
        {
            get { return _price; }
            set { _price = value; }
        }

        public double Tax
        {
            get { return _tax; }
            set { _tax = value; }
        }

        public Client Seller
        {
            get { return _seller; }
            set { _seller = value; }
        }
    }
}
