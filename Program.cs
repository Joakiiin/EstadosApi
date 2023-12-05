var constructor = WebApplication.CreateBuilder(args);

constructor.Services.AddControllers();

constructor.Services.AddCors(opciones =>
{
    opciones.AddDefaultPolicy(builder =>
    {
        builder.AllowAnyOrigin()
               .AllowAnyHeader()
               .AllowAnyMethod();
    });
});
constructor.Services.AddEndpointsApiExplorer();
constructor.Services.AddSwaggerGen();
var aplicacion = constructor.Build();
if (aplicacion.Environment.IsDevelopment())
{
    aplicacion.UseSwagger();
    aplicacion.UseSwaggerUI();
}
aplicacion.UseHttpsRedirection();
aplicacion.UseCors();
aplicacion.UseAuthorization();
aplicacion.MapControllers();
aplicacion.Run();