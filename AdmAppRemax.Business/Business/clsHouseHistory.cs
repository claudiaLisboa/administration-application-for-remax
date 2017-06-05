using System;

namespace AdmAppRemax.Business
{
    public class HouseHistory
    {
        private House _house;
        private DateTime _startDate;
        private BuildingStatus _status;

        // Default constructor
        public HouseHistory()
        {
            _house = default(House);
            _startDate = default(DateTime);
            _status = default(BuildingStatus);
        }

        public House House
        {
            get { return _house; }
            set { _house = value; }
        }

        public DateTime StartDate
        {
            get { return _startDate; }
            set { _startDate = value; }
        }

        public BuildingStatus Status
        {
            get { return _status; }
            set { _status = value; }
        }
    }
}
