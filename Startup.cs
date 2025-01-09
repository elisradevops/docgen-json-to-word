using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;

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
            services.AddTransient<IParagraphService, ParagraphService>();
            services.AddTransient<IPictureService, PictureService>();
            services.AddTransient<ITextService, TextService>();
            services.AddSingleton<ITableService, TableService>();
            services.AddSingleton<IFileService, FileService>();
            services.AddSingleton<IUtilsService, UtilsService>();
            services.AddSingleton<IRunService, RunService>();
            services.AddSingleton<IListService, ListService>();
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
