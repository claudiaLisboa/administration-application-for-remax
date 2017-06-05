using AdmAppRemax.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace AdmAppRemax.Business
{
    public class ClientList
    {
        private Dictionary<int, Client> _list;

        // Default constructor
        public ClientList()
        {
            LoadClients();
        }

        public ClientList(Employee agent)
        {
            LoadClients(agent);
        }

        private void LoadClients()
        {
            _list = new Dictionary<int, Client>();

            DataSet ds = DataSource.GetClients();

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no Clients.
                return;
            }

            LoadDictionary(ds);
        }

        private void LoadClients(Employee agent)
        {
            _list = new Dictionary<int, Client>();

            DataSet ds = DataSource.GetClientsByAgent(agent.EmployeeId);

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no Clients.
                return;
            }

            LoadDictionary(ds);
        }

        private void LoadDictionary(DataSet ds)
        {
            // Retrieving the list of Employees to find the Agent assigned to each Client.
            EmployeeList employeeList = new EmployeeList();

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                Client client = new Client();
                client.ClientId = (int)row["ClientId"];
                client.FirstName = row["FirstName"].ToString();
                client.LastName = row["LastName"].ToString();
                client.BirthDate = (DateTime)row["BirthDate"];
                client.Email = row["Email"].ToString();
                client.Phone = row["Phone"].ToString();
                client.ClientType = (ClientType)row["ClientTypeId"];
                client.Agent = employeeList.FindByEmployeeId( (int)row["AgentId"] );

                _list.Add(client.ClientId, client);
            }
        }

        public Dictionary<int, Client>.ValueCollection Elements
        {
            get
            {
                return _list.Values;
            }
        }

        public Dictionary<int, Client>.KeyCollection Keys
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

        public ErrorAppRemax Add(Client client)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            if (_list.ContainsKey(client.ClientId))
            {
                err.ErrorCode = -1;
                err.ErrorMessage = "There is another Client with this same Client Id.";
                return err;
            }

            // result is the number of affected rows.
            int result = DataSource.AddClient(client.FirstName, client.LastName, client.BirthDate,
                client.Email, client.Phone, (int)client.ClientType, client.Agent.EmployeeId);

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
                LoadClients();

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The add operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }

        public bool Exists(int clientId)
        {
            return _list.ContainsKey(clientId);
        }

        public Client FindByClientId(int clientId)
        {
            if (Exists(clientId) == true)
            {
                return _list[clientId];
            }
            else
            {
                return null;
            }
        }

        public List<Client> GetClientsByAgent(Employee agent)
        {
            // Using LINQ to filter the internal dictionary _list based on the provided Agent.
            var result = from dictItem in _list.AsEnumerable()
                         where dictItem.Value.Agent.EmployeeId == agent.EmployeeId
                         select dictItem.Value;
            // Converting the result above to List<Client> and returning it.
            return result.ToList();
        }

        public List<Client> GetClientsByClientType(ClientType clientType)
        {
            // Using LINQ to filter the internal dictionary _list based on the provided Client Type.
            var result = from dictItem in _list.AsEnumerable()
                         where dictItem.Value.ClientType == clientType
                         orderby dictItem.Value.FullName
                         select dictItem.Value;
            // Converting the result above to List<Client> and returning it.
            return result.ToList();
        }

        public DataSet GetClientTypes()
        {
            DataSet ds = DataSource.GetClientTypes();

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return null;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no positions.
                return null;
            }

            return ds;
        }

        public ErrorAppRemax Remove(int clientId)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            // **** Checking whether the Client being removed is the Seller of any House.
            // Retrieving the Client.
            Client client = FindByClientId(clientId);
            // Retrieving the list of Houses for which the Client is the Seller.
            List<House> lstHouses = ( new HouseList() ).GetHousesBySeller(client);
            // If there is any House then the Client cannot be removed.
            if (lstHouses.Count != 0)
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"This Client cannot be removed. She/he is the Seller of one or more Houses.";
                return err;
            }

            // result is the number of affected rows.
            int result = DataSource.DeleteClient(clientId);

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

                _list.Remove(clientId);

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The delete operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }

        public ErrorAppRemax Save(Client client)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            // result is the number of affected rows.
            int result = DataSource.UpdateClient(client.ClientId, client.FirstName, client.LastName, client.BirthDate,
                client.Email, client.Phone, (int)client.ClientType, client.Agent.EmployeeId);

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

                // Locating the corresponding Client object inside _list.
                Client existingClient = FindByClientId(client.ClientId);
                // Updating the Client object inside _list.
                existingClient.FirstName = client.FirstName;
                existingClient.LastName = client.LastName;
                existingClient.BirthDate = client.BirthDate;
                existingClient.Email = client.Email;
                existingClient.Phone = client.Phone;
                existingClient.ClientType = client.ClientType;
                existingClient.Agent = client.Agent;

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
