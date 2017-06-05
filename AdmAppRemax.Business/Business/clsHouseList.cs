using AdmAppRemax.Data;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace AdmAppRemax.Business
{
    public class HouseList
    {
        private Dictionary<int, House> _list;

        // Default constructor
        public HouseList()
        {
            LoadHouses();
        }

        private void LoadHouses()
        {
            _list = new Dictionary<int, House>();
            DataSet ds = DataSource.GetHouses();

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no houses.
                return;
            }

            // Retrieving the list of Clients to find the Seller assigned to each House.
            ClientList clientList = new ClientList();

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                House house = new House();
                house.HouseId = (int)row["HouseId"];
                house.BuildingType = (BuildingType)row["BuildingTypeId"];
                house.Street = row["Street"].ToString();
                house.ApartmentNo = row["ApartmentNo"].ToString();
                house.PostalCode = row["PostalCode"].ToString();
                house.Country = row["Country"].ToString();
                house.MakingYear = row["MakingYear"].ToString();

                int bathrooms = default(int);
                int.TryParse(row["Bathrooms"].ToString(), out bathrooms);
                house.Bathrooms = bathrooms;

                int bedrooms = default(int);
                int.TryParse(row["Bedrooms"].ToString(), out bedrooms);
                house.Bedrooms = bedrooms;

                house.Area = (double)row["Area"];
                house.Description = row["Description"].ToString();
                house.Price = (double)row["Price"];
                house.Tax = (double)row["Tax"];
                house.Seller = clientList.FindByClientId( (int)row["SellerId"] );

                _list.Add(house.HouseId, house);
            }
        }

        public Dictionary<int, House>.ValueCollection Elements
        {
            get
            {
                return _list.Values;
            }
        }

        public Dictionary<int, House>.KeyCollection Keys
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

        public ErrorAppRemax Add(House house)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            if (_list.ContainsKey(house.HouseId))
            {
                err.ErrorCode = -1;
                err.ErrorMessage = "There is another House with this same House Id.";
                return err;
            }

            // result is the number of affected rows.
            int result = DataSource.AddHouse((int)house.BuildingType, house.Street, house.ApartmentNo, house.PostalCode,
                house.Country, house.MakingYear, house.Bathrooms, house.Bedrooms, house.Area, house.Description, house.Price,
                house.Tax, house.Seller.ClientId);

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
                LoadHouses();

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The add operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }

        public bool Exists(int houseId)
        {
            return _list.ContainsKey(houseId);
        }

        public House FindByHouseId(int houseId)
        {
            if (Exists(houseId) == true)
            {
                return _list[houseId];
            }
            else
            {
                return null;
            }
        }

        public List<House> GetHousesBySeller(Client seller)
        {
            // Using LINQ to filter the internal dictionary _list based on the provided Seller.
            var result = from dictItem in _list.AsEnumerable()
                         where dictItem.Value.Seller.ClientId == seller.ClientId
                         select dictItem.Value;
            // Converting the result above to List<House> and returning it.
            return result.ToList();
        }

        public DataSet GetBuildingTypes()
        {
            DataSet ds = DataSource.GetBuildingTypes();

            // Any error occurred?
            if (DataSource.ErrorCode() != 0)
            {
                //string msg = $"{DataSource.ErrorCode()} - {DataSource.ErrorMessage()}";
                //throw new System.ApplicationException(msg);

                return null;
            }

            if ((ds == null) || (ds.Tables[0].Rows.Count == 0))
            {
                // There are no building types.
                return null;
            }

            return ds;
        }

        public ErrorAppRemax Remove(int houseId)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            // TODO: check whether the House has ever taken part in any Sale transaction.

            // result is the number of affected rows.
            int result = DataSource.DeleteHouse(houseId);

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

                _list.Remove(houseId);

                return err;
            }
            else
            {
                err.ErrorCode = -1;
                err.ErrorMessage = $"The delete operation was not successful ({result} rows affected). Please check the database.";
                return err;
            }
        }

        public ErrorAppRemax Save(House house)
        {
            ErrorAppRemax err = new ErrorAppRemax();

            // result is the number of affected rows.
            int result = DataSource.UpdateHouse(house.HouseId, (int)house.BuildingType, house.Street, house.ApartmentNo,
                house.PostalCode, house.Country, house.MakingYear, house.Bathrooms, house.Bedrooms, house.Area, house.Description,
                house.Price, house.Tax, house.Seller.ClientId);

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

                // Locating the corresponding House object inside _list.
                House existingHouse = FindByHouseId(house.HouseId);
                // Updating the House object inside _list.
                existingHouse.HouseId = house.HouseId;
                existingHouse.BuildingType = house.BuildingType;
                existingHouse.Street = house.Street;
                existingHouse.ApartmentNo = house.ApartmentNo;
                existingHouse.PostalCode = house.PostalCode;
                existingHouse.Country = house.Country;
                existingHouse.MakingYear = house.MakingYear;
                existingHouse.Bathrooms = house.Bathrooms;
                existingHouse.Bedrooms = house.Bedrooms;
                existingHouse.Area = house.Area;
                existingHouse.Description = house.Description;
                existingHouse.Price = house.Price;
                existingHouse.Tax = house.Tax;
                existingHouse.Seller = house.Seller;

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
