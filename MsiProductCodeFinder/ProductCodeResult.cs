namespace MsiProductCodeFinder
{
    class ProductCodeResult
    {
        public bool Success { get; private set; }
        public string Result { get; private set; }

        public ProductCodeResult(bool success, string result)
        {
            Success = success;
            Result = result;
        }
    }
}
