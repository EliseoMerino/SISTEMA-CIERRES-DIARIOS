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
  let totalEDH = 0;
  let totalMAS = 0;

  document.querySelectorAll("tbody tr").forEach(fila => {
    const entEDH = Number(fila.children[2].children[0].value || 0);
    const entMAS = Number(fila.children[3].children[0].value || 0);
    const devEDH = Number(fila.children[4].children[0].value || 0);
    const devMAS = Number(fila.children[5].children[0].value || 0);
    totalEDH += entEDH;
    totalMAS += entMAS;

    const dineroEDH = (entEDH - devEDH) * PRECIO_EDH;
    const dineroMAS = (entMAS - devMAS) * PRECIO_MAS;
    const totalFila = dineroEDH + dineroMAS;

    fila.querySelector(".edhTotal").innerText = dineroEDH.toFixed(2);
    fila.querySelector(".masTotal").innerText = dineroMAS.toFixed(2);
    fila.querySelector(".filaTotal").innerText = totalFila.toFixed(2);

    totalGeneral += totalFila;
  });

  document.getElementById("totalGeneral").innerText = totalGeneral.toFixed(2);

  if (document.getElementById("totalEDHtabla"))
    document.getElementById("totalEDHtabla").innerText = totalEDH;

  if (document.getElementById("totalMAStabla"))
    document.getElementById("totalMAStabla").innerText = totalMAS;

  return { totalGeneral, totalEDH, totalMAS };
}

function guardarLocal() {
  const data = {
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
  const data = JSON.parse(localStorage.getItem(STORAGE_KEY));
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
  let totalGeneral = 0;
  let totalEDH = 0;
  let totalMAS = 0;

  const wsData = [
    ["DESIGNADO", DESIGNADO],
    ["FECHA", document.getElementById("fecha").value],
    [],
    ["VENDEDOR","REFERENCIA","ENT EDH","ENT MAS","DEV EDH","DEV MAS","$ EDH","$ MAS","TOTAL"]
  ];

  document.querySelectorAll("tbody tr").forEach((fila, i) => {
    const entEDH = Number(fila.children[2].children[0].value || 0);
    const entMAS = Number(fila.children[3].children[0].value || 0);
    const devEDH = Number(fila.children[4].children[0].value || 0);
    const devMAS = Number(fila.children[5].children[0].value || 0);

    totalEDH += entEDH;
    totalMAS += entMAS;

    const dineroEDH = (entEDH - devEDH) * PRECIO_EDH;
    const dineroMAS = (entMAS - devMAS) * PRECIO_MAS;
    const totalFila = dineroEDH + dineroMAS;

    totalGeneral += totalFila;

    wsData.push([
      clientes[i].vendedor,
      clientes[i].ref,
      entEDH,
      entMAS,
      devEDH,
      devMAS,
      dineroEDH,
      dineroMAS,
      totalFila
    ]);
  });

  wsData.push([]);
  wsData.push(["TOTAL EDH ENTREGADO", totalEDH]);
  wsData.push(["TOTAL MAS ENTREGADO", totalMAS]);
  wsData.push(["TOTAL $ (CON DEVOLUCIONES)", totalGeneral]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "CIERRE DIARIO");

  XLSX.writeFile(
    wb,
    `CIERRE_DH_${document.getElementById("fecha").value || "SIN_FECHA"}.xlsx`
  );
}

function limpiarDatos() {
  if (!confirm("¿Seguro que deseas iniciar un nuevo día?")) return;

  document.querySelectorAll("tbody input").forEach(inp => inp.value = "");
  document.getElementById("fecha").value = "";

  guardarLocal();
  calcular();
}

cargarLocal();
