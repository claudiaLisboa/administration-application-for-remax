using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace AdmAppRemax.Data
{
    public static class DataSource
    {
        private static OleDbConnection _dbConnection;

        private static int _errorCode = 0;
        private static string _errorMessage = string.Empty;

        public static bool Init()
        {
            ClearError();

            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\App_Data\dbRemax.mdb");

            if (!File.Exists(dbPath))
            {
                _errorCode = -1;
                _errorMessage = "The application database could not be found at the location below:\n\n" + dbPath;
                return false;
            }

            _dbConnection = null;
            _dbConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbPath);
            try
            {
                _dbConnection.Open();
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return false;
            }

            return true;
        }

        public static int ErrorCode()
        {
            return _errorCode;
        }

        public static string ErrorMessage()
        {
            return _errorMessage;
        }

        private static void ClearError()
        {
            _errorCode = 0;
            _errorMessage = string.Empty;
        }

        #region Client
        public static DataSet GetClients()
        {
            ClearError();

            string cmdText = "select ClientId, FirstName, LastName, BirthDate, Email, Phone, ClientTypeId, AgentId ";
            cmdText += "from Clients ";
            cmdText += "order by ClientId ";

            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter(cmdText, _dbConnection);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }

        public static DataSet GetClientsByAgent(int agentId)
        {
            ClearError();

            string cmdText = "select ClientId, FirstName, LastName, BirthDate, Email, Phone, ClientTypeId, AgentId ";
            cmdText += "from Clients ";
            cmdText += "where AgentId=@AgentId ";
            cmdText += "order by ClientId ";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdText, _dbConnection);
                cmd.Parameters.AddWithValue("@AgentId", Convert.ToInt32(agentId));

                OleDbDataAdapter adp = new OleDbDataAdapter(cmd);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }

        public static int AddClient(string firstName, string lastName, DateTime birthDate,
            string email, string phone, int clientTypeId, int agentId)
        {
            ClearError();

            string cmdString = "insert into Clients (FirstName, LastName, BirthDate, Email, Phone, ClientTypeId, AgentId) ";
            cmdString += "values (@FirstName, @LastName, @BirthDate, @Email, @Phone, @ClientTypeId, @AgentId) ";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@BirthDate", birthDate);
                cmd.Parameters.AddWithValue("@Email", email);
                cmd.Parameters.AddWithValue("@Phone", phone);
                cmd.Parameters.AddWithValue("@ClientTypeId", Convert.ToInt32(clientTypeId));
                cmd.Parameters.AddWithValue("@AgentId", Convert.ToInt32(agentId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }

        public static int DeleteClient(int clientId)
        {
            ClearError();

            string cmdString = "delete from Clients where ClientId = @ClientId";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@ClientId", Convert.ToInt32(clientId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }

        public static int UpdateClient(int clientId, string firstName, string lastName, DateTime birthDate,
            string email, string phone, int clientTypeId, int agentId)
        {
            ClearError();

            string cmdString = "update Clients set FirstName=@FirstName, LastName=@LastName, BirthDate=@BirthDate, ";
            cmdString += "Email=@Email, Phone=@Phone, ClientTypeId=@ClientTypeId, AgentId=@AgentId ";
            cmdString += "where ClientId=@ClientId ";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@BirthDate", birthDate);
                cmd.Parameters.AddWithValue("@Email", email);
                cmd.Parameters.AddWithValue("@Phone", phone);
                cmd.Parameters.AddWithValue("@ClientTypeId", Convert.ToInt32(clientTypeId));
                cmd.Parameters.AddWithValue("@AgentId", Convert.ToInt32(agentId));
                cmd.Parameters.AddWithValue("@ClientId", Convert.ToInt32(clientId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }
        #endregion

        #region ClientType
        public static DataSet GetClientTypes()
        {
            ClearError();

            string cmdText = "select ClientTypeId, Description ";
            cmdText += "from ClientTypes ";
            cmdText += "order by ClientTypeId ";

            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter(cmdText, _dbConnection);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }
        #endregion

        #region Employee
        public static DataSet GetEmployees()
        {
            ClearError();

            string cmdText = "select Employees.EmployeeId, Employees.FirstName, Employees.LastName, Employees.BirthDate, Employees.Email, ";
            cmdText += "Employees.Phone, Employees.PositionId, Employees.Salary, Employees.Username, Employees.UserPassword, ";
            cmdText += "PositionPermissions.ManageEmployees, ";
            cmdText += "PositionPermissions.ManageAllClients, PositionPermissions.ManageOwnClients, ";
            cmdText += "PositionPermissions.ManageAllHouses, PositionPermissions.ManageOwnHouses, ";
            cmdText += "PositionPermissions.ManageAllSales, PositionPermissions.ManageOwnSales ";
            cmdText += "from (Positions inner join Employees ";
            cmdText += "on Positions.PositionId = Employees.PositionId) ";
            cmdText += "inner join PositionPermissions ";
            cmdText += "on Positions.PositionId = PositionPermissions.PositionId ";
            cmdText += "order by Employees.EmployeeId ";

            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter(cmdText, _dbConnection);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }

        public static int AddEmployee(string firstName, string lastName, DateTime birthDate,
            string email, string phone, int positionId, double salary, string username, string userPassword)
        {
            ClearError();

            string cmdString = "insert into Employees (FirstName, LastName, BirthDate, Email, Phone, PositionId, Salary, Username, UserPassword) ";
            cmdString += "values (@FirstName, @LastName, @BirthDate, @Email, @Phone, @PositionId, @Salary, @Username, @UserPassword) ";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@BirthDate", birthDate);
                cmd.Parameters.AddWithValue("@Email", email);
                cmd.Parameters.AddWithValue("@Phone", phone);
                cmd.Parameters.AddWithValue("@PositionId", Convert.ToInt32(positionId));
                cmd.Parameters.AddWithValue("@Salary", Convert.ToDouble(salary));
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@UserPassword", userPassword);

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }

        public static int DeleteEmployee(int employeeId)
        {
            ClearError();

            string cmdString = "delete from Employees where EmployeeId = @EmployeeId";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@EmployeeId", Convert.ToInt32(employeeId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }

        public static int UpdateEmployee(int employeeId, string firstName, string lastName, DateTime birthDate,
            string email, string phone, int positionId, double salary)
        {
            ClearError();

            string cmdString = "update Employees set FirstName=@FirstName, LastName=@LastName, BirthDate=@BirthDate, ";
            cmdString += "Email=@Email, Phone=@Phone, PositionId=@PositionId, Salary=@Salary ";
            cmdString += "where EmployeeId=@EmployeeId";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@BirthDate", birthDate);
                cmd.Parameters.AddWithValue("@Email", email);
                cmd.Parameters.AddWithValue("@Phone", phone);
                cmd.Parameters.AddWithValue("@PositionId", Convert.ToInt32(positionId));
                cmd.Parameters.AddWithValue("@Salary", Convert.ToDouble(salary));
                cmd.Parameters.AddWithValue("@EmployeeId", Convert.ToInt32(employeeId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }
        #endregion

        #region Position
        public static DataSet GetPositions()
        {
            ClearError();

            string cmdText = "select PositionId, Description ";
            cmdText += "from Positions ";
            cmdText += "order by PositionId ";

            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter(cmdText, _dbConnection);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }
        #endregion

        #region House
        public static DataSet GetHouses()
        {
            ClearError();

            string cmdText = "select HouseId, BuildingTypeId, Street, ApartmentNo, PostalCode, Country, ";
            cmdText += "MakingYear, Bathrooms, Bedrooms, Area, Description, Price, Tax, SellerId ";
            cmdText += "from Houses ";
            cmdText += "order by HouseId ";

            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter(cmdText, _dbConnection);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }

        public static int AddHouse(int buildingTypeId, string street, string apartmentNo, string postalCode, string country,
            string makingYear, int bathrooms, int bedrooms, double area, string description, double price, double tax, int sellerId)
        {
            ClearError();

            string cmdString = "insert into Houses (BuildingTypeId, Street, ApartmentNo, PostalCode, ";
            cmdString += "Country, MakingYear, Bathrooms, Bedrooms, Area, Description, Price, Tax, SellerId) ";
            cmdString += "values (@BuildingTypeId, @Street, @ApartmentNo, @PostalCode, ";
            cmdString += "@Country, @MakingYear, @Bathrooms, @Bedrooms, @Area, @Description, @Price, @Tax, @SellerId) ";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@BuildingTypeId", Convert.ToInt32(buildingTypeId));
                cmd.Parameters.AddWithValue("@Street", street);
                cmd.Parameters.AddWithValue("@ApartmentNo", apartmentNo);
                cmd.Parameters.AddWithValue("@PostalCode", postalCode);
                cmd.Parameters.AddWithValue("@Country", country);
                cmd.Parameters.AddWithValue("@MakingYear", makingYear);
                cmd.Parameters.AddWithValue("@Bathrooms", Convert.ToInt32(bathrooms));
                cmd.Parameters.AddWithValue("@Bedrooms", Convert.ToInt32(bedrooms));
                cmd.Parameters.AddWithValue("@Area", area);
                cmd.Parameters.AddWithValue("@Description", description);
                cmd.Parameters.AddWithValue("@Price", price);
                cmd.Parameters.AddWithValue("@Tax", tax);
                cmd.Parameters.AddWithValue("@SellerId", Convert.ToInt32(sellerId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }

        public static int DeleteHouse(int houseId)
        {
            ClearError();

            string cmdString = "delete from Houses where HouseId = @HouseId";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@HouseId", Convert.ToInt32(houseId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }

        public static int UpdateHouse(int houseId, int buildingTypeId, string street, string apartmentNo, string postalCode, string country,
            string makingYear, int bathrooms, int bedrooms, double area, string description, double price, double tax, int sellerId)
        {
            ClearError();

            string cmdString = "update Houses set BuildingTypeId=@BuildingTypeId, Street=@Street, ApartmentNo=@ApartmentNo, ";
            cmdString += "PostalCode=@PostalCode, Country=@Country, MakingYear=@MakingYear, Bathrooms=@Bathrooms, Bedrooms=@Bedrooms, ";
            cmdString += "Area=@Area, Description=@Description, Price=@Price, Tax=@Tax, SellerId=@SellerId ";
            cmdString += "where HouseId=@HouseId";

            try
            {
                OleDbCommand cmd = new OleDbCommand(cmdString, _dbConnection);
                cmd.Parameters.AddWithValue("@BuildingTypeId", Convert.ToInt32(buildingTypeId));
                cmd.Parameters.AddWithValue("@Street", street);
                cmd.Parameters.AddWithValue("@ApartmentNo", apartmentNo);
                cmd.Parameters.AddWithValue("@PostalCode", postalCode);
                cmd.Parameters.AddWithValue("@Country", country);
                cmd.Parameters.AddWithValue("@MakingYear", makingYear);
                cmd.Parameters.AddWithValue("@Bathrooms", Convert.ToInt32(bathrooms));
                cmd.Parameters.AddWithValue("@Bedrooms", Convert.ToInt32(bedrooms));
                cmd.Parameters.AddWithValue("@Area", area);
                cmd.Parameters.AddWithValue("@Description", description);
                cmd.Parameters.AddWithValue("@Price", price);
                cmd.Parameters.AddWithValue("@Tax", tax);
                cmd.Parameters.AddWithValue("@SellerId", Convert.ToInt32(sellerId));
                cmd.Parameters.AddWithValue("@HouseId", Convert.ToInt32(houseId));

                // Returning the number of affected rows.
                int result = cmd.ExecuteNonQuery();
                return result;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return 0;
            }
        }
        #endregion

        #region BuildingType
        public static DataSet GetBuildingTypes()
        {
            ClearError();

            string cmdText = "select BuildingTypeId, Description ";
            cmdText += "from BuildingTypes ";
            cmdText += "order by BuildingTypeId ";

            try
            {
                OleDbDataAdapter adp = new OleDbDataAdapter(cmdText, _dbConnection);

                DataSet ds = new DataSet();

                adp.Fill(ds);

                return ds;
            }
            catch (Exception err)
            {
                _errorCode = err.HResult;
                _errorMessage = err.Message;
                return null;
            }
        }
        #endregion
    }
}
