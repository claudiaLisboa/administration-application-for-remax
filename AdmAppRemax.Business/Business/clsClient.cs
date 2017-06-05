namespace AdmAppRemax.Business
{
    public class Client : Person
    {
        private int _clientId;
        private ClientType _clientType;
        private Employee _agent;

        // Default constructor
        public Client(): base()
        {
            _clientId = default(int);
            _clientType = ClientType.Buyer;
            _agent = default(Employee);
        }

        public int ClientId
        {
            get { return _clientId; }
            set { _clientId = value; }
        }

        public ClientType ClientType
        {
            get { return _clientType; }
            set { _clientType = value; }
        }

        public Employee Agent
        {
            get { return _agent; }
            set { _agent = value; }
        }
    }
}
