const express = require("express")
const cors = require("cors")
const fs = require("fs")
const path = require("path")
const PizZip = require("pizzip")
const Docxtemplater = require("docxtemplater")

const app = express()
app.use(cors())
app.use(express.json())

// Función para convertir números a texto en español
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

// Función para convertir monto en pesos a texto
function montoATexto(monto) {
  const numero = Math.floor(monto)
  const textoNumero = numeroATexto(numero)

  // Capitalizar primera letra
  const textoCapitalizado = textoNumero.charAt(0).toUpperCase() + textoNumero.slice(1)

  return `${textoCapitalizado} pesos colombianos`
}

app.post("/generar-word", async (req, res) => {
  try {
    const data = req.body

    // Valores fijos según especificación
    data["Acompañamie"] = "$ 1.516.141"
    data["Diseño_Calculo"] = "$ 23.918.292"
    data["Diseño_Sanitario"] = "$ 20.501.393"

    // Función auxiliar para extraer números de strings con formato de moneda
    const extraerNumero = (texto) => {
      if (!texto) return 0
      return Number.parseFloat(texto.toString().replace(/[^0-9]/g, "")) || 0
    }

    // Función auxiliar para formatear moneda
    const formatearMoneda = (numero) => {
      return new Intl.NumberFormat("es-CO", {
        style: "currency",
        currency: "COP",
        minimumFractionDigits: 0,
        maximumFractionDigits: 0,
      }).format(numero)
    }

    // Calcular Subtotal 1 si no está presente
    let subtotal1Numero = 0
    if (!data["Subtotal_1"]) {
      // Subtotal 1 = Diseño Arquitectónico + Diseño Estructural + Acompañamiento
      const diseno = extraerNumero(data["Diseño_Ar"])
      const estructural = extraerNumero(data["Diseño_Calculo"])
      const acompanamiento = extraerNumero(data["Acompañamie"])

      subtotal1Numero = diseno + estructural + acompanamiento
      data["Subtotal_1"] = formatearMoneda(subtotal1Numero)
    } else {
      subtotal1Numero = extraerNumero(data["Subtotal_1"])
    }

    // Calcular Subtotal 2 si no está presente
    let subtotal2Numero = 0
    if (!data["Subtotal_2"]) {
      // Subtotal 2 = Diseño Eléctrico + Diseño Sanitario + Presupuesto
      const electrico = extraerNumero(data["Diseño_Calculo"])
      const sanitario = extraerNumero(data["Diseño_Sanitario"])
      const presupuesto = extraerNumero(data["Presupuesta"])

      subtotal2Numero = electrico + sanitario + presupuesto
      data["Subtotal_2"] = formatearMoneda(subtotal2Numero)
    } else {
      subtotal2Numero = extraerNumero(data["Subtotal_2"])
    }

    // CALCULAR EL TOTAL FINAL (Subtotal_1 + Subtotal_2)
    const totalNumero = subtotal1Numero + subtotal2Numero
    data["Total"] = formatearMoneda(totalNumero)

    // CONVERTIR EL TOTAL A TEXTO
    data["texto"] = montoATexto(totalNumero)

    // Log para debugging
    console.log("📊 Cálculos realizados:")
    console.log(`   Subtotal 1: ${data["Subtotal_1"]} (${subtotal1Numero})`)
    console.log(`   Subtotal 2: ${data["Subtotal_2"]} (${subtotal2Numero})`)
    console.log(`   TOTAL: ${data["Total"]} (${totalNumero})`)
    console.log(`   TEXTO: ${data["texto"]}`)

    const templatePath = path.join(__dirname, "cotizacion_final.docx")

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
    const filename = `cotizacion_saave_${Date.now()}.docx`
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