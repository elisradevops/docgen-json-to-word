using JsonToWord.Services.Interfaces.ExcelServices;

namespace JsonToWord.Services.ExcelServices
{
    public class ExcelHelperService : IExcelHelperService
    {
        public string GetValueString(object value)
        {
            // Handle JsonElement type that comes from JSON deserialization
            if (value is System.Text.Json.JsonElement jsonElement)
            {
                if (jsonElement.ValueKind == System.Text.Json.JsonValueKind.String)
                {
                    return jsonElement.GetString();
                }
                return jsonElement.ToString();
            }

            // Handle any other object type
            return value?.ToString() ?? string.Empty;
        }

    }
}
