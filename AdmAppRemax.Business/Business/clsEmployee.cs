namespace AdmAppRemax.Business
{
    public class Employee : Person
    {
        private int _employeeId;
        private Position _position;
        private double _salary;
        private string _username;
        private string _password;
        private Permission _permissions;

        // Default constructor
        public Employee(): base()
        {
            _employeeId = default(int);
            _position = Position.Agent;
            _salary = default(double);
            _username = string.Empty;
            _password = string.Empty;
            _permissions = default(Permission);
        }

        public int EmployeeId
        {
            get { return _employeeId; }
            set { _employeeId = value; }
        }

        public Position Position
        {
            get { return _position; }
            set { _position = value; }
        }

        public double Salary
        {
            get { return _salary; }
            set { _salary = value; }
        }

        public string Username
        {
            get { return _username; }
            set { _username = value; }
        }

        public string Password
        {
            set { _password = value; }
        }

        public Permission Permissions
        {
            get { return _permissions; }
            set { _permissions = value; }
        }

        public bool CheckPassword(string password)
        {
            return (password == _password);
        }
    }
}
