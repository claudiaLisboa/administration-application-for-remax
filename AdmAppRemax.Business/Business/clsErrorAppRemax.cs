namespace AdmAppRemax.Business
{
    public class ErrorAppRemax
    {
        private long _errorCode;
        private string _errorMessage;

        // Default constructor
        public ErrorAppRemax()
        {
            _errorCode = default(long);
            _errorMessage = string.Empty;
        }

        // Parameterized constructor
        public ErrorAppRemax(long errorCode, string errorMessage)
        {
            _errorCode = errorCode;
            _errorMessage = errorMessage;
        }

        public long ErrorCode
        {
            get { return _errorCode; }
            set { _errorCode = value; }
        }

        public string ErrorMessage
        {
            get { return _errorMessage; }
            set { _errorMessage = value; }
        }
    }
}
