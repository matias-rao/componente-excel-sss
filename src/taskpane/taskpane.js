/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
var FileSaver = require("file-saver");
var dayjs = require("dayjs");
var utc = require("dayjs/plugin/utc");


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("eliminar_espacios").onclick = eliminar_espacios;
    document.getElementById("procesar_lineas").onclick = procesar_lineas;
    document.getElementById("exportar_txt").onclick = exportar_txt;
  }

  //Chequea el lenguaje predeterminado del documento
  const myLanguage = Office.context.contentLanguage;

  if(myLanguage != 'es-AR'){
    Office.context.ui.displayDialogAsync('https://discapacidad-iatros.netlify.app/lenguaje.html', { height: 35, width: 30, displayInIframe: true },
      function (asyncResult) {
        dialog = asyncResult.value;
      }
    );
  }
});

export async function eliminar_espacios() {
  try {
    await Excel.run(async (context) => {
      //Sheet data
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let rango = sheet.getUsedRange(true);
      rango.load("values");

      await context.sync();

      //eliminar espacios en blanco
      rango.values = rango.values.map((e) => e.map((v) => v.toString().trim()));
    });
  } catch (error) {
    console.error(error);
  }
}

export async function procesar_lineas() {
  try {
    await Excel.run(async (context) => {
      //Nombre de la Sheet activa
      let active_sheet = context.workbook.worksheets.getActiveWorksheet();
      active_sheet.load("name");
      await context.sync();

      const active_sheet_name_original = active_sheet.name;

      active_sheet.name = active_sheet_name_original.replace(/\s+/g, '');
      await context.sync();

      const active_sheet_name = active_sheet.name;

      //Agrego sheet de exportacion y la oculto
      let export_sheet = context.workbook.worksheets.add(`Exportar-${active_sheet_name + Math.floor(Math.random() * 11)}`);
      export_sheet.load("name");

      //Rango y contador de rows en la sheet de exportacion
      let export_used_range = export_sheet.getUsedRange(true);
      export_used_range.load("rowCount");

      //Rango y contador de rows en la sheet de data
      let active_sheet_used_range = active_sheet.getUsedRange(true);
      active_sheet_used_range.load(["rowCount"]);

      //Primera row sheet export
      let primera_row_export = export_sheet.getRange("A1:S1");
      primera_row_export.load("values");

      await context.sync();

      eliminar_espacios();

      //pop-up / dialog office api
      let dialog;
      
      Office.context.ui.displayDialogAsync('https://discapacidad-iatros.netlify.app/loading.html', { height: 30, width: 30, displayInIframe: true },
        function (asyncResult) {
          dialog = asyncResult.value;
        }
      );

      primera_row_export.values = [
        [
          `=${active_sheet_name}!A1`,
          `=${active_sheet_name}!B1`,
          `=${active_sheet_name}!C1`,
          `=${active_sheet_name}!D1 & REPT(" ",40-LEN(${active_sheet_name}!D1))`,
          `=${active_sheet_name}!E1`,
          `=${active_sheet_name}!F1`,
          `=REPT(0,11-LEN(${active_sheet_name}!G1)) & ${active_sheet_name}!G1`,
          `=REPT(0,2-LEN(${active_sheet_name}!H1)) & ${active_sheet_name}!H1`,
          `=${active_sheet_name}!I1`,
          `=${active_sheet_name}!J1`,
          `=REPT(0,14-LEN(${active_sheet_name}!K1)) & ${active_sheet_name}!K1`,
          `=REPT(0,5-LEN(${active_sheet_name}!L1)) & ${active_sheet_name}!L1`,
          `=REPT(0,8-LEN(${active_sheet_name}!M1)) & ${active_sheet_name}!M1`,
          `=REPT(0,10-LEN(${active_sheet_name}!N1)) & ${active_sheet_name}!N1`,
          `=REPT(0,10-LEN(${active_sheet_name}!O1)) & ${active_sheet_name}!O1`,
          `=REPT(0,3-LEN(${active_sheet_name}!P1)) & ${active_sheet_name}!P1`,
          `=REPT(0,6-LEN(${active_sheet_name}!Q1)) & ${active_sheet_name}!Q1`,
          `=REPT(0,2-LEN(${active_sheet_name}!R1)) & ${active_sheet_name}!R1`,
          `=${active_sheet_name}!S1`,
        ],
      ];

      //Ultima fila de la sheet de exportacion
      let last_row_export = primera_row_export.getLastRow();

      await context.sync();

      for (let i = 1; i < active_sheet_used_range.rowCount; i++) {
        last_row_export.getOffsetRange(1, 0).copyFrom(
          last_row_export,
          Excel.RangeCopyType.all,
          true, // skipBlanks
          false // transpose
        );

        export_used_range = export_sheet.getUsedRange(true);
        last_row_export = export_used_range.getLastRow();

        await context.sync();
      }

      exportar_txt(active_sheet_name_original, export_sheet.name);

      dialog.close();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function exportar_txt(active_name, export_name) {
  try {
    await Excel.run(async (context) => {
      //Nombre de la Sheet activa
      let active_sheet = context.workbook.worksheets.getActiveWorksheet();
      active_sheet.load("name");

      await context.sync();

      //Sheet Export
      let export_sheet = context.workbook.worksheets.getItem(export_name);
      let export_used_range = export_sheet.getUsedRange(true);
      export_used_range.load("values");

      //Fecha de inicio de Excel
      dayjs.extend(utc);
      const start_date = dayjs.utc("1900-01-01");

      await context.sync();

      // Valores para exportar de la ultima sheet
      let values = export_used_range.values;
      values = values
        .map((r) => {
          // Cambio los valores de las Celdas E y J a formato fecha
          // Se restan 2 en la funcion add debido a que segun Microsoft el año 1900 es bisiesto
          // https://learn.microsoft.com/es-es/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
          if (typeof r[4] == "number") {
            r[4] = start_date.add(r[4] - 2, "day").format("DD/MM/YYYY");
          }

          if (typeof r[9] == "number") {
            r[9] = start_date.add(r[9] - 2, "day").format("DD/MM/YYYY");
          }

          return r.join("|");
        })
        .join("\n");

      let blob = new Blob([values], {
        type: "text/plain;charset=utf-8",
      });

      FileSaver.saveAs(blob, "export-sss.txt");

      // Eliminar sheet exportacion
      export_sheet.delete();
      active_sheet.name = active_name;
    });
  } catch (error) {
    console.error(error);
  }
}
