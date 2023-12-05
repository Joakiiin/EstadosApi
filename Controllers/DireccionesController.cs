using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
namespace EstadosApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DireccionesController : ControllerBase
    {
        private readonly string RutaDeListaDeDirecciones = "C:/Users/joaquin.galindo/Documents/EstadosApi/CPdescarga.xls";
        [HttpGet("CodigoPostal/{CodigoPostalParametro}")]
        public ActionResult<IEnumerable<Dictionary<string, string>>> ObtenerDireccionPorCodigoPostal(string CodigoPostalParametro)
        {
            List<Dictionary<string, string>> DireccionesEncontradas = new List<Dictionary<string, string>>();
            using (FileStream ArchivoDeListaDirecciones = new FileStream(RutaDeListaDeDirecciones, FileMode.Open, FileAccess.Read))
            {
                var LibroDireccionesPorEstado = new HSSFWorkbook(ArchivoDeListaDirecciones);
                for (int NumeroDeSeccionDelLibro = 1; NumeroDeSeccionDelLibro < LibroDireccionesPorEstado.NumberOfSheets; NumeroDeSeccionDelLibro++)
                {
                    var Seccion = LibroDireccionesPorEstado.GetSheetAt(NumeroDeSeccionDelLibro);
                    var ColumnaDeCodigoPostal = ObtenerNombreDelCampoDeDireccion(Seccion, "d_codigo");
                    for (int NumeroDeFilaEnSeccion = 1; NumeroDeFilaEnSeccion <= Seccion.LastRowNum; NumeroDeFilaEnSeccion++)
                    {
                        var ValorEnCeldaCodigoPostal = Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ColumnaDeCodigoPostal)?.ToString();
                        if (ValorEnCeldaCodigoPostal != null && ValorEnCeldaCodigoPostal.Equals(CodigoPostalParametro, StringComparison.OrdinalIgnoreCase))
                        {
                           Dictionary<string, string> direccion = new Dictionary<string, string>
{
    { "d_codigo", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ColumnaDeCodigoPostal)?.ToString() ?? "" },
    { "d_estado", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ObtenerNombreDelCampoDeDireccion(Seccion, "d_estado"))?.ToString() ?? "" },
    { "D_mnpio", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ObtenerNombreDelCampoDeDireccion(Seccion, "D_mnpio"))?.ToString() ?? "" },
    { "d_tipo_asenta", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ObtenerNombreDelCampoDeDireccion(Seccion, "d_tipo_asenta"))?.ToString() ?? "" },
    { "d_asenta", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ObtenerNombreDelCampoDeDireccion(Seccion, "d_asenta"))?.ToString() ?? "" },
    { "d_ciudad", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ObtenerNombreDelCampoDeDireccion(Seccion, "d_ciudad"))?.ToString() ?? "" }
};
                            DireccionesEncontradas.Add(direccion);
                        }
                    }
                }
            }
            if (DireccionesEncontradas.Count > 0)
            {
                return Ok(DireccionesEncontradas);
            }
            else
            {
                return NotFound($"No se encontraron direcciones para el código {CodigoPostalParametro}");
            }
        }
        private int ObtenerNombreDelCampoDeDireccion(ISheet Secciones, string NombreColumna)
        {
            var FilaConLosNombresDeCampos = Secciones.GetRow(0);
            for (int NumeroDeColumnaConElCampoBuscado = 0; NumeroDeColumnaConElCampoBuscado < FilaConLosNombresDeCampos.LastCellNum; NumeroDeColumnaConElCampoBuscado++)
            {
                var NombreEnLaCelda = FilaConLosNombresDeCampos.GetCell(NumeroDeColumnaConElCampoBuscado)?.ToString();
                if (NombreEnLaCelda != null && NombreEnLaCelda.Equals(NombreColumna, StringComparison.OrdinalIgnoreCase))
                {
                    return NumeroDeColumnaConElCampoBuscado;
                }
            }
            throw new InvalidOperationException($"La columna {NombreColumna} no se encontró en el archivo Excel");
        }
    }
}
