using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using JsonToWord.Services.Interfaces.ExcelServices;
using JsonToWord.Services.ExcelServices;

namespace JsonToWord
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            services.AddSwaggerGen();
            services.AddTransient<IAWSS3Service, AWSS3Service>();
            services.AddTransient<IWordService, WordService>();
            services.AddTransient<IExcelService, ExcelService>();
            services.AddTransient<IParagraphService, ParagraphService>();
            services.AddTransient<IDocumentService, DocumentService>();
            services.AddTransient<IPictureService, PictureService>();
            services.AddTransient<ITextService, TextService>();
            services.AddSingleton<ITableService, TableService>();
            services.AddSingleton<ITestReporterService, TestReporterService>();
            services.AddSingleton<IFileService, FileService>();
            services.AddSingleton<IUtilsService, UtilsService>();
            services.AddSingleton<IRunService, RunService>();
            services.AddSingleton<IHtmlService, HtmlService>();
            services.AddSingleton<IDocumentValidatorService, DocumentValidatorService>();
            services.AddSingleton<IContentControlService, ContentControlService>();
            services.AddSingleton<IExcelHelperService, ExcelHelperService>();
            services.AddTransient<IColumnService, ColumnService>();
            services.AddTransient<IReportDataService, ReportDataService>();
            services.AddTransient<ISpreadsheetService, SpreadsheetService>();
            services.AddTransient<IStylesheetService, StylesheetService>();
            services.AddTransient<IColumnService, ColumnService>();
            services.AddTransient<IVoidListService, VoidListService>();
        }

        // Remove ILoggingBuilder from the method signature
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            app.UseSwagger();

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("./swagger/v1/swagger.json", "Swagger API");
                c.RoutePrefix = string.Empty;
            });

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
