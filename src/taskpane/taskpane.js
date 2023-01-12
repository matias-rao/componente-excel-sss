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
});

export async function eliminar_espacios() {
  try {
    await Excel.run(async (context) => {
      //Sheet data
      let sheet = context.workbook.worksheets.getFirst();
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
      //Sheets Excel
      let export_sheet = context.workbook.worksheets.getItem("Exportar");
      let active_sheet = context.workbook.worksheets.getActiveWorksheet();

      //Nombre de la Sheet activa
      active_sheet.load('name');
      await context.sync();

      const active_sheet_name = active_sheet.name;
    
      //Rango y contador de rows en la sheet de exportacion
      let export_used_range = export_sheet.getUsedRange(true);
      export_used_range.load("rowCount");

      //Rango y contador de rows en la sheet de data
      let active_sheet_used_range = active_sheet.getUsedRange(true);
      active_sheet_used_range.load(["rowCount"]);
      
      //Primera row sheet export
      let primera_row_export = export_sheet.getRange("A1:S1");
      primera_row_export.load('values');
      
      await context.sync();

      eliminar_espacios();
      
      // Limpiar sheet de exportacion
      if(export_used_range.rowCount > 1){
        export_used_range.clear();
      }

      primera_row_export.values = [[ 
        `=${active_sheet_name}!A2`,
        `=${active_sheet_name}!B2`,
        `=${active_sheet_name}!C2`, 
        `=${active_sheet_name}!D2 & REPT(" ",40-LEN(${active_sheet_name}!D2))`,
        `=${active_sheet_name}!E2`,
        `=${active_sheet_name}!F2`,
        `=REPT(0,11-LEN(${active_sheet_name}!G2)) & ${active_sheet_name}!G2`,
        `=REPT(0,2-LEN(${active_sheet_name}!H2)) & ${active_sheet_name}!H2`,
        `=${active_sheet_name}!I2`,
        `=${active_sheet_name}!J2`,
        `=REPT(0,14-LEN(${active_sheet_name}!K2)) & ${active_sheet_name}!K2`,
        `=REPT(0,5-LEN(${active_sheet_name}!L2)) & ${active_sheet_name}!L2`,
        `=REPT(0,8-LEN(${active_sheet_name}!M2)) & ${active_sheet_name}!M2`,
        `=REPT(0,10-LEN(${active_sheet_name}!N2)) & ${active_sheet_name}!N2`,
        `=REPT(0,10-LEN(${active_sheet_name}!O2)) & ${active_sheet_name}!O2`,
        `=REPT(0,3-LEN(${active_sheet_name}!P2)) & ${active_sheet_name}!P2`,
        `=REPT(0,6-LEN(${active_sheet_name}!Q2)) & ${active_sheet_name}!Q2`,
        `=${active_sheet_name}!R2`,
        `=${active_sheet_name}!S2`,
      ]];

      //Ultima fila de la sheet de exportacion
      let last_row_export = export_used_range.getLastRow();
      last_row_export.load("address");
      
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

      exportar_txt();

    });
  } catch (error) {
    console.error(error);
  }
}

export async function exportar_txt() {
  try {
    await Excel.run(async (context) => {
      //Sheet Export
      let sheet = context.workbook.worksheets.getItem("Exportar");
      let rango = sheet.getUsedRange(true);
      rango.load(["address", "values", "rowCount"]);

      //Fecha de inicio de Excel
      dayjs.extend(utc);
      const start_date = dayjs.utc("1900-01-01");

      await context.sync();

      // Valores para exportar de la ultima sheet
      let values = rango.values;
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
    });
  } catch (error) {
    console.error(error);
  }
}