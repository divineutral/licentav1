var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddMemoryCache();

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI(c => {
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "LicentaV1 API");
    c.RoutePrefix = "swagger";
});

app.UseStaticFiles();
app.UseDefaultFiles();
app.MapControllers();
app.Run();