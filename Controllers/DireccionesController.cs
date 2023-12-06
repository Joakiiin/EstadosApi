using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EstadosApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DireccionesController : ControllerBase
    {
        private readonly string RutaDeListaDeDirecciones = "C:/Users/joaquin.galindo/Documents/Repositorios/EstadosApi/EstadosApi/CPdescarga.xls";

        [HttpGet("CodigoPostal/{CodigoPostalParametro}")]
        public ActionResult ObtenerDireccionPorCodigoPostal(string CodigoPostalParametro, bool AgruparPorTipoDeAsentamiento = false)
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
                                { "d_ciudad", Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ObtenerNombreDelCampoDeDireccion(Seccion, "d_ciudad"))?.ToString() ?? "" },
                                { "pais", "México" }
                            };
                            DireccionesEncontradas.Add(direccion);
                        }
                    }
                }
            }
            if (DireccionesEncontradas.Count > 0)
            {
                if (AgruparPorTipoDeAsentamiento)
                {
                    var direccionesAgrupadas = AgruparPorTipoAsentamiento(DireccionesEncontradas);
                    return Ok(direccionesAgrupadas);
                }
                else
                {
                    return Ok(DireccionesEncontradas);
                }
            }
            else
            {
                return NotFound($"No se encontraron direcciones para el código {CodigoPostalParametro}");
            }
        }

        [HttpGet("BusquedaCP/{CriterioBusqueda}")]
        public ActionResult BuscarCodigoPostalPorCoincidencia(string CriterioBusqueda, int? limite = null)
        {
            HashSet<string> CodigosPostalesEncontrados = new HashSet<string>();

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
                        if (ValorEnCeldaCodigoPostal != null && ValorEnCeldaCodigoPostal.Contains(CriterioBusqueda, StringComparison.OrdinalIgnoreCase))
                        {
                            CodigosPostalesEncontrados.Add(ValorEnCeldaCodigoPostal);
                        }
                        if (limite.HasValue && CodigosPostalesEncontrados.Count >= limite)
                        {
                            break;
                        }
                    }
                    if (limite.HasValue && CodigosPostalesEncontrados.Count >= limite)
                    {
                        break;
                    }
                }
            }

            if (CodigosPostalesEncontrados.Count > 0)
            {
                return Ok(new
                {
                    response = new
                    {
                        cp = CodigosPostalesEncontrados.ToList()
                    }
                });
            }
            else
            {
                return NotFound($"No se encontraron códigos postales para el criterio de búsqueda {CriterioBusqueda}");
            }
        }

        [HttpGet("get_estados")]
        public ActionResult ObtenerEstados()
        {
            HashSet<string> EstadosEncontrados = new HashSet<string>();

            using (FileStream ArchivoDeListaDirecciones = new FileStream(RutaDeListaDeDirecciones, FileMode.Open, FileAccess.Read))
            {
                var LibroDireccionesPorEstado = new HSSFWorkbook(ArchivoDeListaDirecciones);
                for (int NumeroDeSeccionDelLibro = 1; NumeroDeSeccionDelLibro < LibroDireccionesPorEstado.NumberOfSheets; NumeroDeSeccionDelLibro++)
                {
                    var Seccion = LibroDireccionesPorEstado.GetSheetAt(NumeroDeSeccionDelLibro);
                    var ColumnaDeEstado = ObtenerNombreDelCampoDeDireccion(Seccion, "d_estado");

                    for (int NumeroDeFilaEnSeccion = 1; NumeroDeFilaEnSeccion <= Seccion.LastRowNum; NumeroDeFilaEnSeccion++)
                    {
                        var ValorEnCeldaEstado = Seccion.GetRow(NumeroDeFilaEnSeccion)?.GetCell(ColumnaDeEstado)?.ToString();
                        if (!string.IsNullOrEmpty(ValorEnCeldaEstado))
                        {
                            EstadosEncontrados.Add(ValorEnCeldaEstado);
                        }
                    }
                }
            }

            if (EstadosEncontrados.Count > 0)
            {
                return Ok(new
                {
                    response = new
                    {
                        estado = EstadosEncontrados.ToList()
                    }
                });
            }
            else
            {
                return NotFound("No se encontraron estados en el archivo.");
            }
        }

        private List<Dictionary<string, object>> AgruparPorTipoAsentamiento(List<Dictionary<string, string>> direcciones)
        {
            var agrupadas = new List<Dictionary<string, object>>();

            foreach (var direccion in direcciones)
            {
                var tipoAsentamiento = direccion["d_tipo_asenta"];
                var existente = agrupadas.Find(d => d["tipo_asentamiento"].ToString() == tipoAsentamiento);

                if (existente == null)
                {
                    var nuevoRegistro = new Dictionary<string, object>
                    {
                        { "cp", direccion["d_codigo"] },
                        { "asentamiento", new List<string> { direccion["d_asenta"] } },
                        { "tipo_asentamiento", tipoAsentamiento },
                        { "municipio", direccion["D_mnpio"] },
                        { "estado", direccion["d_estado"] },
                        { "ciudad", direccion["d_ciudad"] },
                        { "pais", "México" }
                    };
                    agrupadas.Add(nuevoRegistro);
                }
                else
                {
                    ((List<string>)existente["asentamiento"]).Add(direccion["d_asenta"]);
                }
            }

            return agrupadas;
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
