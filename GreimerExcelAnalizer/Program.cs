using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class GreimerExcelAnalizer
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string filePath = @"C:\Users\WPOSS\Documents\LogsGreimerAnalizer\Consulta_Tabla_Auditoria_Init.xlsx"; // Reemplaza con la ruta de tu archivo Excel

        Dictionary<string, Dictionary<DateTime, List<(TimeSpan, string, string)>>> terminalInitTimes = new Dictionary<string, Dictionary<DateTime, List<(TimeSpan, string, string)>>>();

        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        {
            if (package.Workbook.Worksheets.Count > 0)
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Suponiendo que queremos la primera hoja

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Empezamos desde la segunda fila asumiendo que la primera fila contiene encabezados
                {
                    string terminalId = worksheet.Cells[row, 4].Value?.ToString(); // Columna "terminal_id"
                    DateTime fechaInicio = DateTime.Parse(worksheet.Cells[row, 5].Value.ToString()); // Columna "fecha_inicio"
                    TimeSpan horaInicio = DateTime.Parse(worksheet.Cells[row, 5].Value.ToString()).TimeOfDay; // Hora de inicio
                    string puertoPos = worksheet.Cells[row, 12].Value?.ToString(); // Columna "puerto_pos"
                    string mensajePos = worksheet.Cells[row, 9].Value?.ToString(); // Columna "msj_respuesta"

                    fechaInicio = fechaInicio.Date; // Tomamos solo la fecha sin la hora

                    if (!terminalInitTimes.ContainsKey(terminalId))
                    {
                        terminalInitTimes[terminalId] = new Dictionary<DateTime, List<(TimeSpan, string, string)>>();
                    }

                    if (!terminalInitTimes[terminalId].ContainsKey(fechaInicio))
                    {
                        terminalInitTimes[terminalId][fechaInicio] = new List<(TimeSpan, string, string)>();
                    }

                    terminalInitTimes[terminalId][fechaInicio].Add((horaInicio, puertoPos, mensajePos));
                }
            }
            else
            {
                Console.WriteLine("El archivo Excel no contiene ninguna hoja.");
                return;
            }
        }

        // Escribir resultados en un archivo de texto
        string outputFilePath = @"C:\Users\WPOSS\Documents\LogsGreimerAnalizer\resultado.txt";
        using (StreamWriter writer = new StreamWriter(outputFilePath))
        {
            foreach (var terminalEntry in terminalInitTimes)
            {
                string terminalId = terminalEntry.Key;
                Dictionary<DateTime, List<(TimeSpan, string, string)>> initTimes = terminalEntry.Value;

                writer.WriteLine($"Terminal ID: {terminalId}");
                foreach (var entry in initTimes)
                {
                    writer.WriteLine($"Fecha: {entry.Key.ToShortDateString()}, Inicializaciones: {entry.Value.Count}");
                    foreach (var timeAndPortAndMessage in entry.Value)
                    {
                        writer.WriteLine($"               -Hora: {timeAndPortAndMessage.Item1.ToString(@"hh\:mm\:ss")}, Port: {timeAndPortAndMessage.Item2}, Mensaje Respuesta: \"{timeAndPortAndMessage.Item3}\"");
                    }
                }
                writer.WriteLine();
            }
        }

        Console.WriteLine("Proceso completado. Se ha generado el archivo 'resultado.txt'");
        Console.ReadKey();
    }
}
