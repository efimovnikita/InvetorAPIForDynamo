using Inventor;

namespace DynamoEngine
{
    class Globals
    {
        private static Document _globalDoc = null;

        public static Document GlobalDoc
        {
            get { return _globalDoc; }

            set { _globalDoc = value; }
        }
    }
}
