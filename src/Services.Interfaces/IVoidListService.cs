using System.Collections.Generic;

namespace JsonToWord.Services.Interfaces
{
    public interface IVoidListService
    {
        List<string> CreateVoidList(string docPath);
    }
}
