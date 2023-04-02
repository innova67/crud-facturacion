import * as XlsxPopulate from "https://cdn.jsdelivr.net/npm/xlsx-populate/browser/xlsx-populate.min.js";

function hola() {
  console.log("hola de vuelta!!!");
}

function guardarInformacion() {

  console.log("guardando info");

	var nombre = document.getElementById("nombre").value;
	var modelo = document.getElementById("modelo").value;
	var identificacion = document.getElementById("identificacion").value;
	var fecha = document.getElementById("fecha").value;
	var codigoQR = document.getElementById("codigoQR").value;

	// Combinar los datos ingresados en un objeto
	var datos = {
		nombre: nombre,
		modelo: modelo,
		identificacion: identificacion,
		fecha: fecha,
		codigoQR: codigoQR,
	};

	// Crear un nuevo archivo Excel en blanco
	XlsxPopulate.fromBlankAsync()
		.then((workbook) => {
			// Agregar una hoja de trabajo y escribir los encabezados
			var hoja = workbook.addSheet("Datos");
			hoja.cell("A1").value("Nombre");
			hoja.cell("B1").value("Modelo");
			hoja.cell("C1").value("Identificación");
			hoja.cell("D1").value("Fecha");
			hoja.cell("E1").value("Código QR");
			// Escribir los datos ingresados en la primera fila
			hoja.cell("A2").value(nombre);
			hoja.cell("B2").value(modelo);
			hoja.cell("C2").value(identificacion);
			hoja.cell("D2").value(fecha);
			hoja.cell("E2").value(codigoQR);
			// Guardar el archivo Excel con los datos
			return workbook.toFileAsync("nuevo_archivo.xlsx");
		})
		.then(() => {
			console.log("Datos guardados correctamente en el archivo Excel!");
		})
		.catch((error) => {
			console.log(error);
		});
}

console.log("hola, sistemas iniciados")