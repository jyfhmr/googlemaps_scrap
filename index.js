//   divs De cada item:  clase: Z8fK3b

//   TITULOS DE LAS AGENCIAS: NrDZNb

//   numero: usdlKText
var cuantity = 0

//  sitio donde aparece como llegar y si tiene siito web .Rwjeuc
var robot = require("@todesktop/robotjs-prebuild");
const puppeteer = require("puppeteer");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

var sites = [
  { c: "Bogotá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Medellín", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Cali", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Barranquilla", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Cartagena", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Cúcuta", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Bucaramanga", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Pereira", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Santa Marta", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Ibagué", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Manizales", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Villavicencio", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Neiva", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Pasto", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Montería", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Armenia", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Popayán", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Sincelejo", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Valledupar", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Tunja", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Riohacha", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Yopal", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Florencia", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Quibdó", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "San Andrés", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Leticia", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Mocoa", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Mitú", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Inírida", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Arauca", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Buenaventura", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Palmira", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Itagüí", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Soledad", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Soacha", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Envigado", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Tuluá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Dosquebradas", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Apartadó", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Girón", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Floridablanca", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Bello", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Barrancabermeja", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Sabanalarga", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Girardot", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Fusagasugá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Zipaquirá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Chía", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Rionegro", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Malambo", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Magangué", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Maicao", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Pitalito", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Ipiales", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Buga", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "San José del Guaviare", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Duitama", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Sogamoso", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Aguachica", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Espinal", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Facatativá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Turbo", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Chigorodó", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Carepa", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Caucasia", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Carmen de Bolívar", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Sabaneta", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Funza", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Madrid", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Ciénaga", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Tumaco", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Yumbo", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Jamundí", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Cajicá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Candelaria", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "La Dorada", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Chiquinquirá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Lorica", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Guacarí", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Barbosa", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Puerto Boyacá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Puerto Asís", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Sahagún", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Planeta Rica", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Uribia", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "El Bagre", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Carepa", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Arjona", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "El Carmen de Viboral", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Granada", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Pacho", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Chocontá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Cáqueza", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Gachetá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Caparrapí", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Viotá", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Anapoima", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Silvania", p: "Colombia", t: "Agencias de Fotografía" },
  { c: "Subachoque", p: "Colombia", t: "Agencias de Fotografía" }
];


async function scrapeGoogleMaps(c, p, t) {

 /*
  setTimeout(()=>{

    scrapeGoogleMaps(c,p,5)

    
    robot.keyTap("f5");
    
    setTimeout(()=>{
      robot.moveMouse(200, 450);
    },2000)
   

  },5000)
 */

  const browser = await puppeteer.launch({
    headless: false, // Puedes cambiar esto a "true" si no deseas que se abra el navegador de forma visible
  });
  const page = await browser.newPage();

  try {
    const url = `https://www.google.com/maps/search/${c}+${p}+${t}`;
    await page.goto(url, { waitUntil: "load" }); // Espera hasta que la página esté completamente cargada

   setTimeout(()=>{
    robot.moveMouse(200, 600);
   },2000)

    robot.moveMouse(200, 500);

   const scrollInterval = setInterval(() => {
      robot.scrollMouse(0, -30); // La dirección y cantidad del scroll (-10 es hacia abajo)
    }, 100); // Intervalo en milisegundos

    // Espera a que hagas el scroll manualmente y aparezca el elemento .PbZDve
    await page.waitForSelector(".PbZDve", { timeout: 300000 });

    const content = await page.content();
    const $ = cheerio.load(content);

    const combinedArray = [];

    $(".Z8fK3b").each((index, parentElement) => {
        const textObject = {};
  
        // Busca el texto del elemento .NrDZNb dentro del elemento padre .Z8fK3b
        const nrDZNbText = $(parentElement).find(".NrDZNb").text().trim();
        textObject.Nombre = nrDZNbText || "NOCONOCIDO";
  
        const usdlKText = $(parentElement).find(".UsdlK").text().trim();
        var cleanedText = usdlKText.replace(/[+\s-]/g, "") + ",";
        if (cleanedText === ",") {
          cleanedText = "584123091835,";
        }
        textObject.Tlf = cleanedText;
  
        // Busca el primer elemento de la clase .Rwjeuc
        const rwjeucElement = $(".Rwjeuc").eq(index);
  
        // Busca el elemento 'a' dentro de rwjeucElement
        const urlSitioElement = rwjeucElement.find("a");
  
        // Obtiene el atributo 'href' del elemento 'a'
        const urlSitio = urlSitioElement.length
          ? urlSitioElement.attr("href")
          : "NOCONOCIDO";
  
        textObject.urlSitio = urlSitio;
  
        combinedArray.push(textObject);
      });

    console.log(combinedArray);
    console.log(combinedArray.length);
    cuantity+= combinedArray.length
    // Genera un archivo Excel
    await generateExcel(combinedArray, c, p, t,scrollInterval);
  } catch (error) {
    console.error("Error occurred:", error);
  } finally {
    await browser.close();
  }
}

async function generateExcel(data, c, p, t,scrollInterval) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Agencias de Marketing");

  // Agregar encabezados de columna
  sheet.addRow(["Nombre", "Teléfono", "URL Sitio"]);

  // Agregar datos a las filas
  data.forEach((item) => {
    sheet.addRow([item.Nombre, item.Tlf, item.urlSitio]);
  });

  // Guardar el archivo Excel
  await workbook.xlsx.writeFile(`${c}_${p}_${t}_${data.length}.xlsx`);
  clearInterval(scrollInterval)
  console.log("Archivo Excel generado correctamente.");
}

(async () => {
  for (let i = 0; i < sites.length; i++) {
    await scrapeGoogleMaps(sites[i].c, sites[i].p, sites[i].t);
  }
  console.log("NÚMEROS ",cuantity)
})();
