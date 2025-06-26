const express = require("express")
const cors = require("cors")
const fs = require("fs")
const path = require("path")
const PizZip = require("pizzip")
const Docxtemplater = require("docxtemplater")
const { google } = require("googleapis")
require('dotenv').config();

const app = express()
app.use(cors())
app.use(express.json())

const SHEET_ID = "1qWXMepGrgxjZK9QLPxcLcCDlHtsLBIH9Fo0GJDgr_Go" 

function numeroATexto(numero) {
  const unidades = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"]
  const decenas = ["", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"]
  const especiales = [
    "diez",
    "once",
    "doce",
    "trece",
    "catorce",
    "quince",
    "dieciséis",
    "diecisiete",
    "dieciocho",
    "diecinueve",
  ]
  const centenas = [
    "",
    "ciento",
    "doscientos",
    "trescientos",
    "cuatrocientos",
    "quinientos",
    "seiscientos",
    "setecientos",
    "ochocientos",
    "novecientos",
  ]

  if (numero === 0) return "cero"
  if (numero === 100) return "cien"
  if (numero === 1000000) return "un millón"

  let texto = ""

  // Millones
  if (numero >= 1000000) {
    const millones = Math.floor(numero / 1000000)
    if (millones === 1) {
      texto += "un millón "
    } else {
      texto += numeroATexto(millones) + " millones "
    }
    numero %= 1000000
  }

  // Miles
  if (numero >= 1000) {
    const miles = Math.floor(numero / 1000)
    if (miles === 1) {
      texto += "mil "
    } else {
      texto += numeroATexto(miles) + " mil "
    }
    numero %= 1000
  }

  // Centenas
  if (numero >= 100) {
    const cent = Math.floor(numero / 100)
    texto += centenas[cent] + " "
    numero %= 100
  }

  // Decenas y unidades
  if (numero >= 20) {
    const dec = Math.floor(numero / 10)
    const uni = numero % 10
    texto += decenas[dec]
    if (uni > 0) {
      texto += " y " + unidades[uni]
    }
  } else if (numero >= 10) {
    texto += especiales[numero - 10]
  } else if (numero > 0) {
    texto += unidades[numero]
  }

  return texto.trim()
}

function montoATexto(monto) {
  const numero = Math.floor(monto)
  const textoNumero = numeroATexto(numero)

  // Capitalizar primera letra
  const textoCapitalizado = textoNumero.charAt(0).toUpperCase() + textoNumero.slice(1)

  return `${textoCapitalizado} pesos colombianos`
}


async function guardarEnGoogleSheets(data) {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth: await auth.getClient() });
  const fila = [
    data.nombre || "",
    data.correo || "",
    data.Subtotal_1 || "",
    data.Subtotal_2 || "",
    data.Total || "",
    data.texto || "",
  ];
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "Sheet1!A:F",
    valueInputOption: "USER_ENTERED",
    resource: { values: [fila] },
  });
}

app.post("/generar-word", async (req, res) => {
  try {
    const data = req.body

    // Asignar los mismos valores fijos que en Python
    data["Acompañamie"] = "$ 1.516.141"
    data["Diseño_Calculo"] = "$ 23.918.292"
    data["Diseño_Sanitario"] = "$ 20.501.393"

    // Función auxiliar para extraer números de strings con formato de moneda
    const extraerNumero = (texto) => {
      if (!texto) return 0
      return parseInt((texto.toString().match(/\d+/g) || []).join("") || "0")
    }

    // Función auxiliar para formatear moneda
    const formatearMoneda = (numero) => {
      return "$" + numero.toLocaleString("es-CO")
    }

    // Calcular Subtotal 1 si no está presente o es 0
    let subtotal1 = extraerNumero(data["Subtotal_1"])
    if (!subtotal1) {
      subtotal1 =
        extraerNumero(data["Diseño_Ar"]) +
        extraerNumero(data["Diseño_Calculo"]) +
        extraerNumero(data["Acompañamie"])
    }
    // Calcular Subtotal 2 si no está presente o es 0
    let subtotal2 = extraerNumero(data["Subtotal_2"])
    if (!subtotal2) {
      subtotal2 =
        extraerNumero(data["Diseño_Calculo"]) +
        extraerNumero(data["Diseño_Sanitario"]) +
        extraerNumero(data["Presupuesta"])
    }
    // Calcular total
    const total = subtotal1 + subtotal2
    data["Subtotal_1"] = formatearMoneda(subtotal1)
    data["Subtotal_2"] = formatearMoneda(subtotal2)
    data["Total"] = formatearMoneda(total)
    data["texto"] = montoATexto(total)

    // Guardar en Google Sheets
    await guardarEnGoogleSheets(data);

    // Usar plantilla.docx
    const templatePath = path.join(__dirname, "plantilla.docx")
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({
        error: "No se encontró la plantilla Word en: " + templatePath,
      })
    }

    const content = fs.readFileSync(templatePath, "binary")
    const zip = new PizZip(content)
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    })

    await doc.resolveData(data)
    doc.render()

    const buf = doc.getZip().generate({ type: "nodebuffer" })
    const filename = `cotizacion_${Date.now()}.docx`
    const filepath = path.join(__dirname, filename)

    fs.writeFileSync(filepath, buf)

    res.download(filepath, filename, (err) => {
      if (err) {
        console.error("Error enviando archivo:", err)
      }
      // Limpiar archivo temporal
      try {
        fs.unlinkSync(filepath)
      } catch (cleanupErr) {
        console.error("Error limpiando archivo temporal:", cleanupErr)
      }
    })
  } catch (error) {
    console.error("Error procesando solicitud:", error)
    res.status(500).json({
      error: "Error interno del servidor",
      details: error.message,
    })
  }
})

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log('Servidor Node.js escuchando en puerto ' + PORT);
});