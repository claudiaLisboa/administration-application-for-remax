using AdmAppRemax.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace AdmAppRemax.Business
{
    public class EmployeeList
    {
        private Dictionary<int, Employee> _list;

        // Default constructor
        public EmployeeList()
        {
            LoadEmployees();
        }

        private void LoadEmployees()
        {
            _list = new Dictionary<int, Employee>();
            DataSet ds = DataSource.GetEmployees();

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no employees.
                return;
            }

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                Permission permissions = new Permission();
                permissions.ManageEmployees = (bool)row["ManageEmployees"];
                permissions.ManageAllClients = (bool)row["ManageAllClients"];
                permissions.ManageOwnClients = (bool)row["ManageOwnClients"];
                permissions.ManageAllHouses = (bool)row["ManageAllHouses"];
                permissions.ManageOwnHouses = (bool)row["ManageOwnHouses"];
                permissions.ManageAllSales = (bool)row["ManageAllSales"];
                permissions.ManageOwnSales = (bool)row["ManageOwnSales"];
                Employee employee = new Employee();
                employee.EmployeeId = (int)row["EmployeeId"];
                employee.FirstName = row["FirstName"].ToString();
                employee.LastName = row["LastName"].ToString();
                employee.BirthDate = (DateTime)row["BirthDate"];
                employee.Email = row["Email"].ToString();
                employee.Phone = row["Phone"].ToString();
                employee.Position = (Position)row["PositionId"];
                employee.Salary = (double)row["Salary"];
                employee.Username = row["Username"].ToString();
                employee.Password = row["UserPassword"].ToString();
                employee.Permissions = permissions;

                _list.Add(employee.EmployeeId, employee);
            }
        }

        public Dictionary<int, Employee>.ValueCollection Elements
        {
            get
            {
                return _list.Values;
            }
        }

        public Dictionary<int, Employee>.KeyCollection Keys
        {
            get
            {
                return _list.Keys;
            }
        }

        public int Quantity
        {
            get
            {
                return _list.Values.Count();
            }
        }

        public ErrorAppRemax Add(Employee employee)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            if (_list.ContainsKey(employee.EmployeeId))
            {
                err.ErrorCode = -1;
                err.ErrorMessage = "There is another Employee with this same Employee Id.";
                return err;
            }

            // Checking whether there is another Employee with the same Username.
            if (FindByUsername(employee.Username) != null)
            {
                err.ErrorCode = -1;
                err.ErrorMessage = "There is another Employee with this same Username.";
                return err;
            }

            // result is the number of affected rows.
            // The initial password is equal to the Username.
            int result = DataSource.AddEmployee(employee.FirstName, employee.LastName, employee.BirthDate,
                employee.Email, employee.Phone, (int)employee.Position, employee.Salary, employee.Username, employee.Username);

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                err.ErrorCode = DataSource.ErrorCode();
                err.ErrorMessage = DataSource.ErrorMessage();
                return err;
            }

            if (result == 1)
            {
                // If the adding works fine only one row should be affected.

                // Reloading _list.
                LoadEmployees();

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The add operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }

        public bool Exists(int employeeId)
        {
            return _list.ContainsKey(employeeId);
        }

        public Employee FindByEmployeeId(int employeeId)
        {
            if (Exists(employeeId) == true)
            {
                return _list[employeeId];
            }
            else
            {
                return null;
            }
        }

        public Employee FindByUsername(string username)
        {
            foreach (Employee employee in this.Elements)
            {
                if (employee.Username == username)
                {
                    return employee;
                }
            }

            return null;
        }

        public List<Employee> GetEmployeesByPosition(Position position)
        {
            // Using LINQ to filter the internal dictionary _list based on the provided Position.
            var result = from dictItem in _list.AsEnumerable()
                         where dictItem.Value.Position == position
                         orderby dictItem.Value.FullName
                         select dictItem.Value;
            // Converting the result above to List<Employee> and returning it.
            return result.ToList();
        }

        public DataSet GetPositions()
        {
            DataSet ds = DataSource.GetPositions();

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return null;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no Positions.
                return null;
            }

            return ds;
        }

        public ErrorAppRemax Remove(int employeeId)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            // **** Checking whether the Employee being removed is assigned to any Client as an Agent.
            // Retrieving the Employee.
            Employee employee = FindByEmployeeId(employeeId);
            // Retrieving the list of Clients for which the Employee is assigned as an Agent.
            List<Client> lstClients = ( new ClientList() ).GetClientsByAgent(employee);
            // If there is any Client then the Employee cannot be removed.
            if (lstClients.Count != 0)
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"This Employee cannot be removed. She/he is the Agent of one or more Clients.";
                return err;
            }

            // result is the number of affected rows.
            int result = DataSource.DeleteEmployee(employeeId);

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                err.ErrorCode = DataSource.ErrorCode();
                err.ErrorMessage = DataSource.ErrorMessage();
                return err;
            }

            if (result == 1)
            {
                // If the deleting works fine only one row should be affected.

                _list.Remove(employeeId);

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The delete operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }

        public ErrorAppRemax Save(Employee employee)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            // **** Checking whether the Employee:
            // 1. Is being saved with a Position different than Agent.
            // 2. Is already assigned to any Client as an Agent.
            // If this is the case then the Employee cannot be changed to another Position.

            // Retrieving the Position of the Employee being saved.
            Position position = employee.Position;
            // If the Employee is not being saved as an Agent...
            if (position != Position.Agent)
            {
                // ...need to check whether she/he has any client assigned.

                // Retrieving the list of Clients for which the Employee is assigned as an Agent.
                List<Client> lstClients = (new ClientList()).GetClientsByAgent(employee);

                // If there is any Client then the Employee cannot be changed to a Position other than Agent.
                if (lstClients.Count != 0)
                {
                    err.ErrorCode = -1;
                    err.ErrorMessage = $"This Employee is the Agent of {lstClients.Count} Clients. The Position cannot be changed.";
                    return err;
                }
            }

            // Updating the Employee.
            // result is the number of affected rows.
            int result = DataSource.UpdateEmployee(employee.EmployeeId, employee.FirstName, employee.LastName, employee.BirthDate,
                employee.Email, employee.Phone, (int)employee.Position, employee.Salary);

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                err.ErrorCode = DataSource.ErrorCode();
                err.ErrorMessage = DataSource.ErrorMessage();
                return err;
            }

            if (result == 1)
            {
                // If the updating works fine only one row should be affected.

                // Locating the corresponding Employee object inside _list.
                Employee existingEmployee = FindByEmployeeId(employee.EmployeeId);
                // Updating the Employee object inside _list.
                existingEmployee.FirstName = employee.FirstName;
                existingEmployee.LastName = employee.LastName;
                existingEmployee.BirthDate = employee.BirthDate;
                existingEmployee.Email = employee.Email;
                existingEmployee.Phone = employee.Phone;
                existingEmployee.Position = employee.Position;
                existingEmployee.Salary = employee.Salary;

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The save operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }
    }
}
