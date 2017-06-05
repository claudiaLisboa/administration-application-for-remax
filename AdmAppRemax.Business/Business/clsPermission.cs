namespace AdmAppRemax.Business
{
    public class Permission
    {
        private bool _manageEmployees;
        private bool _manageAllClients;
        private bool _manageOwnClients;
        private bool _manageAllHouses;
        private bool _manageOwnHouses;
        private bool _manageAllSales;
        private bool _manageOwnSales;

        // Default constructor
        public Permission()
        {
            _manageEmployees = false;
            _manageAllClients = false;
            _manageOwnClients = false;
            _manageAllHouses = false;
            _manageOwnHouses = false;
            _manageAllSales = false;
            _manageOwnSales = false;
        }

        public bool ManageEmployees
        {
            get { return _manageEmployees; }
            set { _manageEmployees = value; }
        }

        public bool ManageAllClients
        {
            get { return _manageAllClients; }
            set { _manageAllClients = value; }
        }

        public bool ManageOwnClients
        {
            get { return _manageOwnClients; }
            set { _manageOwnClients = value; }
        }

        public bool ManageAllHouses
        {
            get { return _manageAllHouses; }
            set { _manageAllHouses = value; }
        }

        public bool ManageOwnHouses
        {
            get { return _manageOwnHouses; }
            set { _manageOwnHouses = value; }
        }

        public bool ManageAllSales
        {
            get { return _manageAllSales; }
            set { _manageAllSales = value; }
        }

        public bool ManageOwnSales
        {
            get { return _manageOwnSales; }
            set { _manageOwnSales = value; }
        }
    }
}