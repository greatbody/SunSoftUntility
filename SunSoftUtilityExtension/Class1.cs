using SunSoftUtility.DbHelper;

namespace SunSoftUtility.Extension
{
    public static class BaseExtension
    {
        public static QueryParameter AsQueryParameter(this object op)
        {
            return new QueryParameter(op);
        }
    }
}
