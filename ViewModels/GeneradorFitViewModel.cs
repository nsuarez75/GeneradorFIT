using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace GeneradorFIT.ViewModels
{
    public class FitData
    {
        public string Plant { get; set; }
        public string Line { get; set; }
        public string Guid { get; set; }
        public string PartReference { get; set; }
        public string Asset { get; set; }
        public string LayoutName { get; set; }
        public int NumAsset { get; set; }
        public string RefClient { get; set; }
        public string ProcessFeatureType { get; set; }
        public string Location { get; set; }
        public string Fid { get; set; }
        public string CriticalFeature { get; set; }
        public string Working { get; set; }
        public int Ref { get; set; }
        public string Tech { get; set; }
        public string Pointer { get; set; }
    }

    public partial class GeneradorFitViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _rutaOrigen = "";

        [ObservableProperty]
        private string _rutaDestino = "";

        [ObservableProperty]
        private int _referencias = 8;

        [ObservableProperty]
        private string _estado = "Listo";

        [ObservableProperty]
        private bool _isProcessing = false;

        public GeneradorFitViewModel()
        {
        }

        [RelayCommand]
        private void BuscarOrigen()
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*",
                Title = "Seleccionar archivo Excel origen"
            };

            if (ofd.ShowDialog() == true)
            {
                RutaOrigen = ofd.FileName;
                if (string.IsNullOrEmpty(RutaDestino))
                {
                    string dir = Path.GetDirectoryName(RutaOrigen);
                    string filename = Path.GetFileNameWithoutExtension(RutaOrigen);
                    RutaDestino = Path.Combine(dir, $"{filename}_FIT.xlsx");
                }
            }
        }

        [RelayCommand]
        private void BuscarDestino()
        {
            var sfd = new SaveFileDialog
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                Title = "Guardar nuevo archivo Excel como..."
            };

            if (sfd.ShowDialog() == true)
            {
                RutaDestino = sfd.FileName;
            }
        }

        [RelayCommand]
        private async Task GenerarAsync()
        {
            if (string.IsNullOrWhiteSpace(RutaOrigen) || !File.Exists(RutaOrigen))
            {
                MessageBox.Show("Por favor, selecciona un archivo Excel de origen válido.",
                    "Atención", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(RutaDestino))
            {
                MessageBox.Show("Por favor, selecciona una ruta de destino para guardar el archivo Excel.",
                    "Atención", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (Referencias <= 0)
            {
                MessageBox.Show("El número de referencias debe ser mayor que 0.",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SetProcessing(true);
            SetEstado("Iniciando...");

            await Task.Run(() =>
            {
                try
                {
                    SetEstado("Cargando archivo en memoria...");

                    // Usamos una aproximación más directa para evitar bloqueos por metadatos o eventos
                    using (var stream = new FileStream(RutaOrigen, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (var libroEntrada = new XLWorkbook(stream))
                        {
                            var datos = GenerarDatos(libroEntrada, Referencias);

                            if (datos.Count == 0)
                            {
                                throw new Exception("No se encontraron datos válidos en ninguna de las hojas PLC-RefX.");
                            }

                            // Una vez comprobadas todas las referencias, miramos si además existe
                            // la hoja opcional "PLC-Aux" para incorporar sus datos.
                            var datosAux = GenerarDatosAux(libroEntrada);

                            SetEstado("Generando libro de salida...");
                            using (var libroSalida = new XLWorkbook())
                            {
                                SetEstado("Hoja FIT...");
                                GenerarHojaFIT(libroSalida, datos, datosAux);

                                GenerarHojaConfigReference(libroSalida, Referencias, datos);

                                SetEstado("Guardando archivo...");
                                libroSalida.SaveAs(RutaDestino);
                            }

                            SetEstado("Generando DB User_TableInfo_DB...");
                            string rutaDB = Path.Combine(Path.GetDirectoryName(RutaDestino), "User_TableInfo_DB.db");
                            string contenidoDB = GenerarContenidoDBTableInfo(datos, datosAux);
                            File.WriteAllText(rutaDB, contenidoDB, new UTF8Encoding(true));

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                MessageBox.Show($"¡Éxito! Archivo generado en:\n{RutaDestino}\n\nDB generado en:\n{rutaDB}",
                                    "Generación Completada", MessageBoxButton.OK, MessageBoxImage.Information);
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"Error durante el proceso:\n{ex.Message}",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
                finally
                {
                    SetEstado("Listo");
                    SetProcessing(false);
                }
            });
        }

        // Las propiedades ObservableProperty están enlazadas a controles WPF (DispatcherObject).
        // GenerarAsync las actualiza desde el hilo de Task.Run, así que hay que reenviarlas
        // siempre al hilo de UI o WPF lanza InvalidOperationException.
        private void SetEstado(string texto)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
                Estado = texto;
            else
                dispatcher.Invoke(() => Estado = texto);
        }

        private void SetProcessing(bool valor)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
                IsProcessing = valor;
            else
                dispatcher.Invoke(() => IsProcessing = valor);
        }

        private List<FitData> GenerarDatos(XLWorkbook libroEntrada, int refs)
        {
            var datos = new List<FitData>();

            for (int refNum = 1; refNum <= refs; refNum++)
            {
                string nombreHoja = $"PLC-Ref{refNum}";
                if (!libroEntrada.Worksheets.TryGetWorksheet(nombreHoja, out var hoja))
                    continue;

                SetEstado($"Leyendo {nombreHoja}...");

                // En lugar de RowsUsed(), vamos a buscar el límite real de datos manualmente
                // para evitar que ClosedXML se pierda en miles de filas vacías con estilo.
                int ultimaFila = 0;
                var lastCell = hoja.LastCellUsed();
                if (lastCell != null)
                {
                    ultimaFila = lastCell.Address.RowNumber;
                }

                // Si el Excel tiene un formato infinito, limitamos a algo razonable o usamos una lógica de parada
                // El script original parece empezar en fila 2.
                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    // Actualizamos el estado cada 100 filas para no saturar la UI pero dar feedback
                    if (fila % 100 == 0)
                    {
                        SetEstado($"Leyendo {nombreHoja} (Fila {fila}/{ultimaFila})...");
                    }

                    // Columnas comunes a toda la fila (no dependen del grupo de asset).
                    // Se eliminaron del origen las columnas C (GUID) y V (ya iba vacía), por lo
                    // que todo lo que estaba después de cada una se desplaza una posición a la
                    // izquierda (dos posiciones para lo que iba después de V).
                    string plant = GetVal(hoja, fila, 1);           // A
                    string line = GetVal(hoja, fila, 2);            // B
                    string guid = "";                               // GUID ya no existe en el origen (era la columna C)
                    string partReference = GetVal(hoja, fila, 3);   // C (antes D)
                    string refClient = GetVal(hoja, fila, 16);          // P (antes Q)
                    string processFeatureType = GetVal(hoja, fila, 17); // Q (antes R)
                    string location = GetVal(hoja, fila, 18);           // R (antes S)
                    string fid = GetVal(hoja, fila, 19);                // S (antes T)
                    string criticalFeature = GetVal(hoja, fila, 20);    // T (antes U)
                    string tech = GetVal(hoja, fila, 23);               // W (antes Y)
                    string pointer = GetVal(hoja, fila, 24);            // X (antes Z)

                    // Columnas de LayoutName: E(5), H(8), K(11), N(14).
                    // El Asset de cada grupo está en la columna anterior (D, G, J, M) y el
                    // Working en la siguiente (F, I, L, O); ya no es una única columna M para todos.
                    int[] columnasLayoutName = { 5, 8, 11, 14 };
                    for (int assetIdx = 1; assetIdx <= 4; assetIdx++)
                    {
                        int col = columnasLayoutName[assetIdx - 1];
                        string layoutNameVal = GetVal(hoja, fila, col);

                        if (!string.IsNullOrWhiteSpace(layoutNameVal))
                        {
                            datos.Add(new FitData
                            {
                                Plant = plant,
                                Line = line,
                                Guid = guid,
                                PartReference = partReference,
                                Asset = GetVal(hoja, fila, col - 1),
                                LayoutName = layoutNameVal,
                                NumAsset = assetIdx,
                                RefClient = refClient,
                                ProcessFeatureType = processFeatureType,
                                Location = location,
                                Fid = fid,
                                CriticalFeature = criticalFeature,
                                Working = GetVal(hoja, fila, col + 1),
                                Ref = refNum,
                                Tech = tech,
                                Pointer = pointer
                            });
                        }
                    }

                    // Opcional: si encontramos 10 filas vacías seguidas, podríamos asumir fin de datos
                    // pero mantengamos la lógica de LastCellUsed por ahora.
                }
            }

            return datos;
        }

        // Hoja opcional "PLC-Aux": Asset en A, Working en B, Fid en C, desde la fila 2,
        // usando la misma lógica de LastCellUsed que las hojas PLC-RefX para saber cuántas
        // filas hay. Aquí no existe un LayoutName real distinto del Asset: se deja vacío
        // a propósito (la hoja FIT reflejará Asset relleno y LayoutName en blanco).
        private List<FitData> GenerarDatosAux(XLWorkbook libroEntrada)
        {
            var aux = new List<FitData>();

            if (!libroEntrada.Worksheets.TryGetWorksheet("PLC-Aux", out var hoja))
                return aux;

            SetEstado("Leyendo PLC-Aux...");

            int ultimaFila = 0;
            var lastCell = hoja.LastCellUsed();
            if (lastCell != null)
            {
                ultimaFila = lastCell.Address.RowNumber;
            }

            for (int fila = 2; fila <= ultimaFila; fila++)
            {
                string assetVal = GetVal(hoja, fila, 1);   // A
                string workingVal = GetVal(hoja, fila, 2); // B
                string fidVal = GetVal(hoja, fila, 3);     // C

                if (!string.IsNullOrWhiteSpace(assetVal))
                {
                    aux.Add(new FitData
                    {
                        Asset = assetVal,
                        Working = workingVal,
                        Fid = fidVal
                    });
                }
            }

            return aux;
        }

        private string GetVal(IXLWorksheet hoja, int row, int col)
        {
            try
            {
                var cell = hoja.Cell(row, col);

                // Usamos CachedValue (no Value) para replicar openpyxl con data_only=True:
                // leemos el último resultado que Excel calculó y guardó en el archivo, sin
                // pedirle a ClosedXML que recalcule la fórmula con su propio motor (que no
                // soporta todas las funciones de Excel y puede fallar o dar otro resultado).
                var valor = cell.CachedValue;

                if (valor.IsBlank || valor.IsError) return "";

                return valor.ToString(CultureInfo.InvariantCulture)?.Trim() ?? "";
            }
            catch
            {
                return "";
            }
        }

        private List<FitData> GenerarFidsLayoutName(List<FitData> datos, int numAsset, List<FitData> datosAux = null)
        {
            var fids = new List<FitData>();
            var fidsUnicos = new HashSet<string>();

            foreach (var dato in datos)
            {
                // El script original no descarta filas con FID vacío, solo evita
                // duplicar un mismo FID: replicamos eso exactamente (fidsUnicos admite "").
                if (dato.NumAsset == numAsset && fidsUnicos.Add(dato.Fid ?? ""))
                {
                    fids.Add(new FitData
                    {
                        Fid = dato.Fid,
                        LayoutName = dato.LayoutName,
                        Working = dato.Working,
                        RefClient = dato.RefClient,
                        Asset = dato.Asset
                    });
                }
            }

            // Datos de la hoja opcional "PLC-Aux": se añaden al final de la lista, pero
            // solo si este conjunto ya tiene datos reales de PLC-RefX; si un conjunto no
            // se usa en este archivo (fids vacío), no se le añade nada de PLC-Aux, para
            // no crear "columnas nuevas" donde antes no había ningún dato real. Al no
            // existir un LayoutName distinto del Asset en "PLC-Aux", usamos el propio
            // Asset como LayoutName para que el código AWL (que carga LayoutName en el
            // campo .Asset del DB) siga funcionando sin cambios.
            if (datosAux != null && fids.Count > 0)
            {
                foreach (var dato in datosAux)
                {
                    if (fidsUnicos.Add(dato.Fid ?? ""))
                    {
                        fids.Add(new FitData
                        {
                            Fid = dato.Fid,
                            LayoutName = dato.Asset,
                            Working = dato.Working,
                            Asset = dato.Asset
                        });
                    }
                }
            }

            return fids;
        }

        // Genera el texto fuente del DB "User_TableInfo_DB" (formato de exportación de
        // fuente externa de TIA Portal), con el mismo contenido que antes se volcaba en
        // la hoja "User Data Table Info": un array TableInfo_X por cada conjunto de
        // LayoutName (X = número de asset, sin la "L" que llevaba el código AWL anterior,
        // ya que las variables reales del DB son TableInfo_1, TableInfo_2, etc.).
        // Si un conjunto no tiene ningún FID configurado en el Excel de origen, no se
        // declara ni se rellena en el DB.
        private string GenerarContenidoDBTableInfo(List<FitData> datos, List<FitData> datosAux)
        {
            var grupos = new List<(int numAsset, List<FitData> fids)>();
            for (int numAsset = 1; numAsset <= 4; numAsset++)
            {
                var fids = GenerarFidsLayoutName(datos, numAsset, datosAux);
                if (fids.Count > 0)
                    grupos.Add((numAsset, fids));
            }

            var sb = new StringBuilder();
            sb.Append("DATA_BLOCK \"User_TableInfo_DB\"\r\n");
            sb.Append("{ S7_Optimized_Access := 'TRUE' }\r\n");
            sb.Append("VERSION : 0.1\r\n");
            sb.Append("NON_RETAIN\r\n");
            sb.Append("//Title_english \r\n");
            sb.Append("   VAR \r\n");

            foreach (var grupo in grupos)
            {
                sb.Append($"      TableInfo_{grupo.numAsset} {{ S7_SetPoint := 'False'}} : Array[1..{grupo.fids.Count}] of \"TableInfo_UDT\";\r\n");
            }

            sb.Append("   END_VAR\r\n");
            sb.Append("\r\n\r\n");
            sb.Append("BEGIN\r\n");

            foreach (var grupo in grupos)
            {
                for (int i = 0; i < grupo.fids.Count; i++)
                {
                    var fid = grupo.fids[i];
                    int indice = i + 1;
                    string featureId = string.IsNullOrWhiteSpace(fid.Fid) ? "0" : fid.Fid.Trim();

                    sb.Append($"   TableInfo_{grupo.numAsset}[{indice}].Asset := '{fid.LayoutName}';\r\n");
                    sb.Append($"   TableInfo_{grupo.numAsset}[{indice}].Working := '{fid.Working}';\r\n");
                    sb.Append($"   TableInfo_{grupo.numAsset}[{indice}].FeatureID := {featureId};\r\n");
                }
            }

            sb.Append("END_DATA_BLOCK\r\n");

            return sb.ToString();
        }

        private void GenerarHojaFIT(XLWorkbook libroSalida, List<FitData> datos, List<FitData> datosAux)
        {
            var hoja = libroSalida.Worksheets.Add("FIT");

            string[] cabeceras =
            {
                "Plant", "Line", "GUID", "PartReference", "Asset", "LayoutName",
                "Working", "FeatureReferenceClient", "Process | FeatureType",
                "Location", "FeatureId", "CriticalFeature"
            };

            for (int c = 0; c < cabeceras.Length; c++)
                hoja.Cell(1, c + 1).Value = cabeceras[c];

            // Una fila por cada combinación (fila origen × asset con LayoutName), tal cual está
            // en "datos": si un mismo FeatureId aparece en varios assets de la misma fila, o en
            // varias referencias (hojas PLC-RefX), cada combinación ya generó su propio FitData.
            int fila = 2;
            foreach (var dato in datos)
            {
                hoja.Cell(fila, 1).Value = dato.Plant;
                hoja.Cell(fila, 2).Value = dato.Line;
                hoja.Cell(fila, 3).Value = dato.Guid;
                hoja.Cell(fila, 4).Value = dato.PartReference;
                // Columna E (Asset) se deja sin rellenar a propósito.
                hoja.Cell(fila, 6).Value = dato.LayoutName;
                hoja.Cell(fila, 7).Value = dato.Working;
                hoja.Cell(fila, 8).Value = dato.RefClient;
                hoja.Cell(fila, 9).Value = dato.ProcessFeatureType;
                hoja.Cell(fila, 10).Value = dato.Location;
                hoja.Cell(fila, 11).Value = dato.Fid;
                hoja.Cell(fila, 12).Value = dato.CriticalFeature;
                fila++;
            }

            // Filas adicionales con los datos de la hoja opcional "PLC-Aux": la columna
            // Asset (E) se deja sin rellenar, igual que en las filas normales, y es la
            // columna LayoutName (F) la que se rellena con el Asset de "PLC-Aux". Plant y
            // Line se reutilizan de las filas normales (PLC-RefX), ya que "PLC-Aux" no
            // trae esas columnas y se asume una única planta/línea por archivo.
            // PartReference, Location y CriticalFeature son fijos para estas filas; el
            // resto de columnas no existen en "PLC-Aux" y quedan en blanco.
            string plantAux = datos.Count > 0 ? datos[0].Plant : "";
            string lineAux = datos.Count > 0 ? datos[0].Line : "";

            foreach (var dato in datosAux)
            {
                hoja.Cell(fila, 1).Value = plantAux;
                hoja.Cell(fila, 2).Value = lineAux;
                hoja.Cell(fila, 4).Value = "99";
                hoja.Cell(fila, 6).Value = dato.Asset;
                hoja.Cell(fila, 7).Value = dato.Working;
                hoja.Cell(fila, 10).Value = "1";
                hoja.Cell(fila, 11).Value = dato.Fid;
                hoja.Cell(fila, 12).Value = "False";
                fila++;
            }

            var rango = hoja.Range(1, 1, fila - 1, cabeceras.Length);
            rango.CreateTable("FIT");
        }

        private void GenerarHojaConfigReference(XLWorkbook libroSalida, int referencias, List<FitData> datos)
        {
            SetEstado("Configuración PLC...");
            var hojaConfig = libroSalida.Worksheets.Add("Config Reference");

            int filaTitulo = 4;

            for (int refNum = 0; refNum < referencias; refNum++)
            {
                int colActual = 1 + refNum;
                hojaConfig.Cell(filaTitulo, colActual).Value = $"Referencia {refNum + 1}";

                int filaCodigoActual = 5;

                foreach (var dato in datos)
                {
                    if (dato.NumAsset == 1 && dato.Ref == refNum + 1)
                    {
                        hojaConfig.Cell(filaCodigoActual, colActual).Value = $"//{dato.LayoutName} {dato.RefClient}";
                        filaCodigoActual++;

                        hojaConfig.Cell(filaCodigoActual, colActual).Value = $"L {dato.Fid}";
                        filaCodigoActual++;

                        hojaConfig.Cell(filaCodigoActual, colActual).Value = $"T #Ref_{dato.Ref}.Tech0{dato.Tech}[{dato.Pointer}].Name";
                        filaCodigoActual++;

                        hojaConfig.Cell(filaCodigoActual, colActual).Value = "";
                        filaCodigoActual++;
                    }
                }
                hojaConfig.Column(colActual).Width = 45;
            }
        }
    }
}
