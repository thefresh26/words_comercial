import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';
import fs from 'fs';

const data = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const outPath = process.argv[3];

const bold = (text) => new TextRun({ text, bold: true, font: "Arial", size: 22 });
const normal = (text) => new TextRun({ text, font: "Arial", size: 22 });
const space = () => new Paragraph({ children: [new TextRun("")] });

const linea = (label, valor) => new Paragraph({
  indent: { left: 720 },
  children: [
    new TextRun({ text: "- ", font: "Arial", size: 22 }),
    bold(`${label} `),
    normal(valor),
  ]
});

const parrafo = (texto) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  children: [new TextRun({ text: texto, font: "Arial", size: 22 })],
  spacing: { after: 120 },
});

const firmante = (nombre, cargo) => [
  new Paragraph({
    children: [new TextRun({ text: nombre, font: "Arial", size: 22, bold: true })],
  }),
  new Paragraph({
    children: [new TextRun({ text: cargo, font: "Arial", size: 22 })],
    spacing: { after: 160 },
  }),
];

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      }
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [new TextRun({ text: "CERTIFICACIÓN DE DEBIDA DILIGENCIA", font: "Arial", size: 24, bold: true })],
      }),
      space(),
      new Paragraph({
        children: [new TextRun({ text: "OFERENTE", font: "Arial", size: 22, bold: true })],
        spacing: { after: 120 },
      }),
      linea("Nombre:", data.nombre),
      linea("Cédula:", data.cedula),
      linea("Número de la Puja:", data.id_puja),
      space(),
      new Paragraph({
        children: [new TextRun({ text: "DEBIDA DILIGENCIA LA/FT/PADM/ST/C", font: "Arial", size: 22, bold: true })],
        spacing: { after: 120 },
      }),
      parrafo("Luego de hecha la consulta en listas restrictivas y vinculantes, en conjunto con la información aportada por el oferente y resultados de fuentes de bases abiertas, el suscrito oficial de cumplimiento, no identifica riesgos de LA/FT/PADM/ST/C asociados con el oferente relacionado."),
      parrafo("El oficial de cumplimiento, en concordancia con la matriz de riesgos y los manuales pertinentes califica el nivel de riesgo de la contraparte en: BAJO."),
      space(),
      new Paragraph({
        children: [new TextRun({ text: "DEBIDA DILIGENCIA FINANCIERA", font: "Arial", size: 22, bold: true })],
        spacing: { after: 120 },
      }),
      parrafo("Según información financiera reportada por la plataforma de DATA CREDITO, desde el criterio de riesgo financiero y operativo se cataloga como viable debido a que se observa correlación en los ingresos, no se observan posibles transacciones ni crecimientos exponenciales inusuales en los ingresos."),
      space(),
      space(),
      new Paragraph({
        children: [new TextRun({ text: "Firman,", font: "Arial", size: 22, bold: true })],
        spacing: { after: 240 },
      }),
      ...firmante("JAIRO ALFONSO JARAMILLO MARRIAGA", "Director Financiero"),
      ...firmante("JULIO RESTREPO", "Oficial de Cumplimiento"),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Carrera 53 No.64-63 Prado Barranquilla, Colombia   info@activosporcolombia.com", font: "Arial", size: 18, color: "605e5c" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "WhatsApp: +57 311 676 7447   Teléfono 60(5) 3468973", font: "Arial", size: 18, color: "605e5c" })],
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log('OK');
});
