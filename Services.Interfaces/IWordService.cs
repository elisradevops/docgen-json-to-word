using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface IWordService
    {
        string Create(WordModel _wordModel);
    }
}
