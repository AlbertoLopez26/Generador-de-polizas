using ClosedXML.Excel;
using NPOI.XSSF.UserModel; // Para trabajar con archivos XLSX
using NPOI.HSSF.UserModel; // Para trabajar con archivos XLS
using NPOI.SS.UserModel;
using DocumentFormat.OpenXml.Spreadsheet;

// Ruta del archivo Excel existente
string filePath = @"C:\Users\ACER A515-51-572H\Desktop\CONTRIBUYENTE\CANTON\GeneradorDePolizas\templatePoliza.xlsx";
string newfilepath = @"C:\Users\ACER A515-51-572H\Desktop\CONTRIBUYENTE\CANTON\GeneradorDePolizas\PolizasGeneradas\";
string filePathxls = @"C:\Users\ACER A515-51-572H\Desktop\CONTRIBUYENTE\CANTON\GeneradorDePolizas\PolizasGeneradas\PolizaXLS.xls";

// Cargar el archivo Excel
using (var workbook = new XLWorkbook(filePath))
{
    // Abrir la hoja de movimientos 
    var bbva = workbook.Worksheet(1);
    var santander = workbook.Worksheet(2);
    var iva = workbook.Worksheet(3);
    // Abrir la hoja de informacion de poliza
    var info = workbook.Worksheet(4);
    // Abrir la hoja donde se encuentra la poliza
    var poliza = workbook.Worksheet(5);

    // Obtener informacion relevante
    int numeroDePoliza = info.Cell("B1").GetValue<int>();
    int numeroDeCuentaBBVA = info.Cell("B2").GetValue<int>();
    int numeroDeCuentaSantander = info.Cell("B3").GetValue<int>();
    int numeroDeCuentaSumatoria = info.Cell("B4").GetValue<int>();
    int numeroDeCuentaIVA1 = info.Cell("B5").GetValue<int>();
    int numeroDeCuentaIVA2 = info.Cell("B6").GetValue<int>();
    int nPoliza = 1;

    // La fila en la cual se comienzan a agregar los movimientos deacuerdo al fomato es en la fila 8
    int posicionPrimerMovimiento = 23;
    int posicionUltimoMovimiento = 23;

    // Asignar el no de poliza para el template
    poliza.Cell($"D{posicionUltimoMovimiento - 1}").SetValue(numeroDePoliza);

    // Ciclo para elaboracion de polizas
    int numeroDeColumnas = bbva.Row(1).LastCellUsed().Address.ColumnNumber;

    for (int i = 0; i < numeroDeColumnas; i++)
    {
        posicionUltimoMovimiento = capturarMovimientosDelDia(bbva, poliza, numeroDeCuentaBBVA, posicionUltimoMovimiento, i + 1);
        posicionUltimoMovimiento = capturarMovimientosDelDia(santander, poliza, numeroDeCuentaSantander, posicionUltimoMovimiento, i + 1);

        sumatoriaMovimientos(poliza, posicionPrimerMovimiento, posicionUltimoMovimiento, numeroDeCuentaSumatoria);
        posicionUltimoMovimiento++;

        capturarIvaDelDia(iva, poliza, numeroDeCuentaIVA1, numeroDeCuentaIVA2, i + 1, posicionUltimoMovimiento);
        posicionUltimoMovimiento += 2;

        if (i != numeroDeColumnas - 1)
        {
            numeroDePoliza = nuevaPoliza(poliza, bbva, info, posicionUltimoMovimiento, numeroDePoliza, nPoliza);
            nPoliza++;
            posicionUltimoMovimiento++;
            posicionPrimerMovimiento = posicionUltimoMovimiento;
        }

    }

    newfilepath += $"poliza{numeroDePoliza}.xlsx";

    // Guardar los cambios en el archivo
    workbook.SaveAs(newfilepath);
}

Console.WriteLine("Archivo actualizado correctamente.");

//Converir XLSX a XLS
using (FileStream archivoEntrada = new FileStream(newfilepath, FileMode.Open, FileAccess.Read))
{
    XSSFWorkbook archivoXLSX = new XSSFWorkbook(archivoEntrada);
    HSSFWorkbook archivoXLS = new HSSFWorkbook();

    // Copiar unicamente la hoja de poliza
    ISheet polizaXLSX = archivoXLSX.GetSheetAt(4);
    ISheet polizaXLS = archivoXLS.CreateSheet(polizaXLSX.SheetName);

    // Copiar las filas de la hoja XLSX a la hoja XLS
    for (int rowIndex = 0; rowIndex <= polizaXLSX.LastRowNum; rowIndex++)
    {
        IRow xlsxRow = polizaXLSX.GetRow(rowIndex);
        IRow xlsRow = polizaXLS.CreateRow(rowIndex);

        if (xlsxRow != null)
        {
            for (int colIndex = 0; colIndex < xlsxRow.LastCellNum; colIndex++)
            {
                ICell xlsxCell = xlsxRow.GetCell(colIndex);
                ICell xlsCell = xlsRow.CreateCell(colIndex);

                if (xlsxCell != null)
                {
                    switch (xlsxCell.CellType)
                    {
                        case NPOI.SS.UserModel.CellType.String:
                            xlsCell.SetCellValue(xlsxCell.StringCellValue);
                            break;
                        case NPOI.SS.UserModel.CellType.Numeric:
                            xlsCell.SetCellValue(xlsxCell.NumericCellValue);
                            if (DateUtil.IsCellDateFormatted(xlsxCell))
                            {
                                ICellStyle dateCellStyle = archivoXLS.CreateCellStyle();
                                IDataFormat format = archivoXLS.CreateDataFormat();
                                dateCellStyle.DataFormat = format.GetFormat("yyyyMMdd");
                                xlsCell.CellStyle = dateCellStyle;
                            }
                            break;
                        
                    }
                }
            }
        }
    }
    // Guardr archivo
    using (FileStream archivoSalida = new FileStream(filePathxls, FileMode.Create, FileAccess.Write))
    {
        archivoXLS.Write(archivoSalida);
    }
    archivoXLS.Close();
}

Console.WriteLine("Archivo convertido correctamente.");



int capturarMovimientosDelDia(IXLWorksheet hojaOrigen, IXLWorksheet hojaDestino, int numeroDeCuenta, int filaPoliza, int ColumnaOrigen)
{
    decimal monto;
    int numeroDeFilas = hojaOrigen.Column(ColumnaOrigen).LastCellUsed().Address.RowNumber;
    for (int i = 1; i < numeroDeFilas; i++)
    {
        monto = hojaOrigen.Cell(i + 1, ColumnaOrigen).GetValue<decimal>();
        hojaDestino.Cell($"A{filaPoliza}").SetValue("M1");
        hojaDestino.Cell($"B{filaPoliza}").SetValue(numeroDeCuenta);
        hojaDestino.Cell($"D{filaPoliza}").SetValue(0);
        hojaDestino.Cell($"E{filaPoliza}").SetValue(Math.Truncate(monto*100)/100);
        hojaDestino.Cell($"F{filaPoliza}").SetValue(0);
        hojaDestino.Cell($"G{filaPoliza}").SetValue(0);
        filaPoliza++;
    }
    return filaPoliza;
}

void sumatoriaMovimientos(IXLWorksheet hojaPoliza, int posicionInicio, int posicionFinal, int numeroDeCuenta)
{
    decimal suma = 0;
    for (int i = posicionInicio; i < posicionFinal; i++)
    {
        suma += hojaPoliza.Cell($"E{i}").GetValue<decimal>();
    }
    hojaPoliza.Cell($"A{posicionFinal}").SetValue("M1");
    hojaPoliza.Cell($"B{posicionFinal}").SetValue(numeroDeCuenta);
    hojaPoliza.Cell($"D{posicionFinal}").SetValue(1);
    hojaPoliza.Cell($"E{posicionFinal}").SetValue(Math.Truncate(suma*100)/100);
    hojaPoliza.Cell($"F{posicionFinal}").SetValue(0);
    hojaPoliza.Cell($"G{posicionFinal}").SetValue(0);
}

void capturarIvaDelDia(IXLWorksheet hojaIva, IXLWorksheet hojaPoliza, int numeroDeCuenta, int numeroDeCuenta2, int columna, int posicionFinal)
{
    decimal iva = hojaIva.Cell(2, columna).GetValue<decimal>();
    hojaPoliza.Cell($"A{posicionFinal}").SetValue("M1");
    hojaPoliza.Cell($"B{posicionFinal}").SetValue(numeroDeCuenta);
    hojaPoliza.Cell($"D{posicionFinal}").SetValue(0);
    hojaPoliza.Cell($"E{posicionFinal}").SetValue(Math.Truncate(iva*100)/100);
    hojaPoliza.Cell($"F{posicionFinal}").SetValue(0);
    hojaPoliza.Cell($"G{posicionFinal}").SetValue(0);
    //
    posicionFinal++;
    hojaPoliza.Cell($"A{posicionFinal}").SetValue("M1");
    hojaPoliza.Cell($"B{posicionFinal}").SetValue(numeroDeCuenta2);
    hojaPoliza.Cell($"D{posicionFinal}").SetValue(1);
    hojaPoliza.Cell($"E{posicionFinal}").SetValue(Math.Truncate(iva*100)/100);
    hojaPoliza.Cell($"F{posicionFinal}").SetValue(0);
    hojaPoliza.Cell($"G{posicionFinal}").SetValue(0);
}

int nuevaPoliza(IXLWorksheet hojaPoliza, IXLWorksheet bbva, IXLWorksheet hojaInfo, int posicionFinal, int numeroDePoliza, int nPoliza)
{
    numeroDePoliza++;

    var fecha = bbva.Cell(1, nPoliza+1).GetDateTime();
    var concepto = hojaPoliza.Cell(22, 7).GetString();

    hojaPoliza.Cell(posicionFinal, 1).SetValue("P");
    hojaPoliza.Cell(posicionFinal, 2).Style.DateFormat.Format = "yyyyMMdd";
    hojaPoliza.Cell(posicionFinal, 2).Value = fecha;
    hojaPoliza.Cell(posicionFinal, 3).SetValue(1);
    hojaPoliza.Cell(posicionFinal, 4).SetValue(numeroDePoliza);
    hojaPoliza.Cell(posicionFinal, 5).SetValue(1);
    hojaPoliza.Cell(posicionFinal, 6).SetValue("0");
    hojaPoliza.Cell(posicionFinal, 7).Value = concepto;
    hojaPoliza.Cell(posicionFinal, 8).SetValue(11);
    hojaPoliza.Cell(posicionFinal, 9).SetValue(0);
    hojaPoliza.Cell(posicionFinal, 10).SetValue(0);

    // incrementar numero de poliza
    hojaInfo.Cell("B1").SetValue(numeroDePoliza);
    hojaPoliza.Cell($"D{posicionFinal}").SetValue(numeroDePoliza);
    return numeroDePoliza;
}
