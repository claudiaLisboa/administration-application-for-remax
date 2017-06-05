using AdmAppRemax.Data;
using System;
using System.Data;

namespace AdmAppRemax.Business
{
    public static class RemaxApplication
    {
        private static int _errorCode = 0;
        private static string _errorMessage = string.Empty;

        public static bool Init()
        {
            ClearError();

            if (!DataSource.Init())
            {
                _errorCode = DataSource.ErrorCode();
                _errorMessage = DataSource.ErrorMessage();
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
    }
}
