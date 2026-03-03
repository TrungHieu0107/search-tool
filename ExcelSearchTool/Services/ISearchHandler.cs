using ExcelSearchTool.Models;

namespace ExcelSearchTool.Services;

public interface ISearchHandler
{
    Task SearchAsync(FileInfo file, SearchContext context);
}
