const PRECIO_EDH = 0.48;
const PRECIO_MAS = 0.23;
const DESIGNADO = "Stephanie Alejandra Estrada Echeverría";
const STORAGE_KEY = "cierre_diario_dh";

const clientes = [
  { vendedor: "OSCAR SALVADOR ANAYA", ref: "APOPA" },
  { vendedor: "KELVIN SALVADOR ALVAREZ", ref: "APOPA" },
  { vendedor: "ADRIANA", ref: "APOPA" },
  { vendedor: "CLAUDIA FLORES", ref: "APOPA" },
  { vendedor: "MATIAS PALMA", ref: "APOPA" },
  { vendedor: "MARCOS MEJIA", ref: "APOPA" },
  { vendedor: "ALEJANDRA", ref: "APOPA" },
  { vendedor: "ERICK", ref: "APOPA" },
  { vendedor: "FELIPE ORELLANA", ref: "APOPA" },
  { vendedor: "HNA", ref: "APOPA" },
  { vendedor: "CARLOS ALBERTO", ref: "APOPA" },
  { vendedor: "CARMEN BRISUELA", ref: "APOPA" },
  { vendedor: "CARMEN DISTRITO", ref: "DISTRITO" },
  { vendedor: "YOLANDA", ref: "CONDOM." },
  { vendedor: "JORGE", ref: "FLORES" },
  { vendedor: "NEJAPA DAVID", ref: "NEJAPA" }
];

const tbody = document.querySelector("tbody");
clientes.forEach(c => {
  tbody.innerHTML += `
    <tr>
      <td>${c.vendedor}</td>
      <td>${c.ref}</td>
      <td><input type="number" min="0"></td>
      <td><input type="number" min="0"></td>
      <td><input type="number" min="0"></td>
      <td><input type="number" min="0"></td>
      <td class="edhTotal">0.00</td>
      <td class="masTotal">0.00</td>
      <td class="filaTotal">0.00</td>
    </tr>
  `;
});
document.addEventListener("input", () => {
  calcular();
  guardarLocal();
});
function calcular() {
  let totalGeneral = 0;

  document.querySelectorAll("tbody tr").forEach(fila => {
    let entEDH = Number(fila.children[2].children[0].value || 0);
    let entMAS = Number(fila.children[3].children[0].value || 0);
    let devEDH = Number(fila.children[4].children[0].value || 0);
    let devMAS = Number(fila.children[5].children[0].value || 0);

    let edh = (entEDH - devEDH) * PRECIO_EDH;
    let mas = (entMAS - devMAS) * PRECIO_MAS;
    let total = edh + mas;

    fila.querySelector(".edhTotal").innerText = edh.toFixed(2);
    fila.querySelector(".masTotal").innerText = mas.toFixed(2);
    fila.querySelector(".filaTotal").innerText = total.toFixed(2);

    totalGeneral += total;
  });

  document.getElementById("totalGeneral").innerText = totalGeneral.toFixed(2);
  return totalGeneral;
}
function guardarLocal() {
  let data = {
    fecha: document.getElementById("fecha").value,
    filas: []
  };

  document.querySelectorAll("tbody tr").forEach(fila => {
    data.filas.push({
      entEDH: fila.children[2].children[0].value,
      entMAS: fila.children[3].children[0].value,
      devEDH: fila.children[4].children[0].value,
      devMAS: fila.children[5].children[0].value
    });
  });

  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

function cargarLocal() {
  let data = JSON.parse(localStorage.getItem(STORAGE_KEY));
  if (!data) return;

  document.getElementById("fecha").value = data.fecha;

  document.querySelectorAll("tbody tr").forEach((fila, i) => {
    fila.children[2].children[0].value = data.filas[i]?.entEDH || "";
    fila.children[3].children[0].value = data.filas[i]?.entMAS || "";
    fila.children[4].children[0].value = data.filas[i]?.devEDH || "";
    fila.children[5].children[0].value = data.filas[i]?.devMAS || "";
  });

  calcular();
}
function exportarExcel() {
  const total = calcular();

  let wsData = [
    ["DESIGNADO", DESIGNADO],
    ["FECHA", document.getElementById("fecha").value],
    [],
    ["VENDEDOR","REFERENCIA","ENT EDH","ENT MAS","DEV EDH","DEV MAS","$ EDH","$ MAS","TOTAL"]
  ];

  document.querySelectorAll("tbody tr").forEach((fila, i) => {
    let entEDH = Number(fila.children[2].children[0].value || 0);
    let entMAS = Number(fila.children[3].children[0].value || 0);
    let devEDH = Number(fila.children[4].children[0].value || 0);
    let devMAS = Number(fila.children[5].children[0].value || 0);

    wsData.push([
      clientes[i].vendedor,
      clientes[i].ref,
      entEDH,
      entMAS,
      devEDH,
      devMAS,
      (entEDH - devEDH) * PRECIO_EDH,
      (entMAS - devMAS) * PRECIO_MAS,
      ((entEDH - devEDH) * PRECIO_EDH) +
      ((entMAS - devMAS) * PRECIO_MAS)
    ]);
  });

  wsData.push([]);
  wsData.push(["","","","","","","","TOTAL", total]);

  let wb = XLSX.utils.book_new();
  let ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "CIERRE DIARIO");

  XLSX.writeFile(
    wb,
    `CIERRE_DH_${document.getElementById("fecha").value || "SIN_FECHA"}.xlsx`
  );
}
function limpiarDatos() {
  if (!confirm("¿Seguro que deseas iniciar un nuevo día?")) return;

  document.querySelectorAll("tbody input").forEach(inp => {
    inp.value = "";
  });

  document.getElementById("fecha").value = "";
  guardarLocal();
  calcular();
}
cargarLocal();
