// ========================================
// DASHBOARD JAS PREMIUM
// APP.JS
// ========================================

Chart.defaults.responsive = true;
Chart.defaults.maintainAspectRatio = false;
Chart.defaults.animation = false;

let datosOriginales = [];
let datosFiltrados = [];

let charts = {};

// ========================================
// INICIO
// ========================================

document.addEventListener("DOMContentLoaded", () => {

    document
        .getElementById("btnActualizar")
        .addEventListener(
            "click",
            cargarExcel
        );

});

// ========================================
// CARGAR EXCEL
// ========================================

function cargarExcel() {

    const archivo =
        document.getElementById(
            "excelFile"
        ).files[0];

    if (!archivo) {

        alert(
            "Seleccione un Excel"
        );

        return;
    }

    const reader =
        new FileReader();

    reader.onload =
        function(e){

        const data =
            new Uint8Array(
                e.target.result
            );

        const workbook =
            XLSX.read(data,{
                type:'array'
            });

        const hoja =
            workbook.SheetNames[0];

        const worksheet =
            workbook.Sheets[hoja];

        datosOriginales =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval:null
                }
            );

        datosFiltrados =
            [...datosOriginales];

        actualizarFecha();

        llenarFiltros();

	document
.getElementById("filtroClase")
.addEventListener(
    "change",
    aplicarFiltros
);

document
.getElementById("filtroMovimiento")
.addEventListener(
    "change",
    aplicarFiltros
);

document
.getElementById("filtroCET")
.addEventListener(
    "change",
    aplicarFiltros
);

document
.getElementById("filtroSKU")
.addEventListener(
    "change",
    aplicarFiltros
);

document
.getElementById("filtroCliente")
.addEventListener(
    "change",
    aplicarFiltros
);

        actualizarKPIs();

        construirGraficos();

        console.log(
            "Registros:",
            datosOriginales.length
        );
    };

    reader.readAsArrayBuffer(
        archivo
    );
}

// ========================================
// FECHA
// ========================================

function actualizarFecha(){

    document
        .getElementById(
            "ultimaActualizacion"
        )
        .innerText =

        "Actualizado: " +
        new Date()
            .toLocaleString();
}

// ========================================
// FILTROS
// ========================================

function llenarFiltros(){

    llenarSelect(
        "filtroClase",
        "Clase"
    );

    llenarSelect(
        "filtroMovimiento",
        "Movimiento"
    );

    llenarSelect(
        "filtroCET",
        "CET"
    );

    llenarSelect(
        "filtroSKU",
        "SKU"
    );

    llenarSelect(
        "filtroCliente",
        "Cliente"
    );
}

function llenarSelect(
    id,
    campo
){

    const select =
        document.getElementById(
            id
        );

    const valores =
        [...new Set(
            datosOriginales.map(
                r => r[campo]
            )
        )]
        .filter(Boolean)
        .sort();

    select.innerHTML =
        '<option value="">Todos</option>';

    valores.forEach(v=>{

        const option =
            document.createElement(
                "option"
            );

        option.value = v;
        option.textContent = v;

        select.appendChild(
            option
        );
    });
}

// ========================================
// APLICAR FILTROS
// ========================================

function aplicarFiltros(){

    const clase =
        document
            .getElementById(
                "filtroClase"
            ).value;

    const movimiento =
        document
            .getElementById(
                "filtroMovimiento"
            ).value;

    const cet =
        document
            .getElementById(
                "filtroCET"
            ).value;

    const sku =
        document
            .getElementById(
                "filtroSKU"
            ).value;

    const cliente =
        document
            .getElementById(
                "filtroCliente"
            ).value;

    datosFiltrados =
        datosOriginales.filter(r =>

            (!clase ||
                r["Clase"] == clase)

            &&

            (!movimiento ||
                r["Movimiento"] == movimiento)

            &&

            (!cet ||
                r["CET"] == cet)

            &&

            (!sku ||
                r["SKU"] == sku)

            &&

            (!cliente ||
                r["Cliente"] == cliente)
        );

    actualizarKPIs();

    construirGraficos();
}

// ========================================
// EVENTOS FILTROS
// ========================================

[
    "filtroClase",
    "filtroMovimiento",
    "filtroCET",
    "filtroSKU",
    "filtroCliente"
]
.forEach(id=>{

    document.addEventListener(
        "change",
        e => {

            if(
                e.target.id === id
            ){

                aplicarFiltros();
            }
        }
    );

});

// ========================================
// KPI
// ========================================

function actualizarKPIs(){

    document
        .getElementById(
            "kpiMovimientos"
        )
        .innerText =
        datosFiltrados.length
            .toLocaleString();

    document
        .getElementById(
            "kpiClientes"
        )
        .innerText =

        new Set(
            datosFiltrados.map(
                r=>r["Cliente"]
            )
        ).size;

    document
        .getElementById(
            "kpiSKU"
        )
        .innerText =

        new Set(
            datosFiltrados.map(
                r=>r["SKU"]
            )
        ).size;

    const cajas =
        datosFiltrados.reduce(
            (a,b)=>

                a +
                Number(
                    b[
                        "Cajas fisica"
                    ] || 0
                ),

            0
        );

    document
        .getElementById(
            "kpiCajas"
        )
        .innerText =
        cajas.toLocaleString();
}

// ========================================
// EXPORTAR PNG
// ========================================

function exportChart(id){

    const canvas =
        document.getElementById(
            id
        );

    const link =
        document.createElement(
            "a"
        );

    link.download =
        id + ".png";

    link.href =
        canvas.toDataURL();

    link.click();
}

// ========================================
// DESTRUIR GRAFICOS
// ========================================

function destruirGraficos(){

    Object.values(charts)
        .forEach(chart => {

            if(chart)
                chart.destroy();
        });

    charts = {};
}

// ========================================
// CONSTRUIR TODOS
// ========================================

function construirGraficos(){

    destruirGraficos();

    construirTrend();

    construirClase();

    construirDona();

    construirAcumulado();

    construirTopClientes();

    construirTopSKU();

    construirRechazosCliente();

    construirRechazosCET();

    construirHeatmap();
}

// ========================================
// GRAFICO 1
// TENDENCIA TEMPORAL
// ========================================

function construirTrend(){

    const fechas = {};

    datosFiltrados.forEach(r=>{

        let fecha =
            r["Fe.contabilización"] ||
            r["Fecha de documento"] ||
            r["Fecha de entrada"];

        if(!fecha) return;

        fecha =
            String(fecha)
            .split(" ")[0];

        fechas[fecha] =
            (fechas[fecha] || 0)
            +
            Number(
                r["Cajas fisica"] || 0
            );
    });

    const labels =
        Object.keys(fechas)
            .sort();

    const valores =
        labels.map(
            x => fechas[x]
        );

    charts.trend =
        new Chart(
            document.getElementById(
                "chartTrend"
            ),
            {

                type:"line",

                data:{
                    labels,

                    datasets:[{

                        label:
                            "Cajas",

                        data:valores,

                        borderColor:
                            "#38bdf8",

                        tension:.3
                    }]
                }
            }
        );
}

// ========================================
// GRAFICO 2
// CLASE VS CAJAS
// ========================================

function construirClase(){

    const grupos = {};

    datosFiltrados.forEach(r=>{

        const clase =
            r["Clase"] ||
            "Sin Clase";

        grupos[clase] =
            (grupos[clase] || 0)
            +
            Number(
                r["Cajas fisica"] || 0
            );
    });

    charts.clase =
        new Chart(
            document.getElementById(
                "chartClaseMov"
            ),
            {

                type:"bar",

                data:{

                    labels:
                        Object.keys(
                            grupos
                        ),

                    datasets:[{

                        label:
                            "Cajas",

                        data:
                            Object.values(
                                grupos
                            ),

                        backgroundColor:
                            "#0078d4"
                    }]
                }
            }
        );
}

// ========================================
// GRAFICO 3
// DONA
// ========================================

function construirDona(){

    const grupos = {};

    datosFiltrados.forEach(r=>{

        const clase =
            r["Clase"] ||
            "Otros";

        grupos[clase] =
            (grupos[clase] || 0)
            + 1;
    });

    charts.dona =
        new Chart(
            document.getElementById(
                "chartDona"
            ),
            {

                type:"doughnut",

                data:{

                    labels:
                        Object.keys(
                            grupos
                        ),

                    datasets:[{

                        data:
                            Object.values(
                                grupos
                            )
                    }]
                }
            }
        );
}

// ========================================
// GRAFICO 4
// ACUMULADO
// ========================================

function construirAcumulado(){

    const fechas = {};

    datosFiltrados.forEach(r=>{

        let fecha =
            r["Fe.contabilización"];

        if(!fecha) return;

        fecha =
            String(fecha)
            .split(" ")[0];

        fechas[fecha] =
            (fechas[fecha] || 0)
            +
            Number(
                r["Cajas fisica"] || 0
            );
    });

    const labels =
        Object.keys(
            fechas
        ).sort();

    let acumulado = 0;

    const valores =
        labels.map(f=>{

            acumulado +=
                fechas[f];

            return acumulado;
        });

    charts.area =
        new Chart(
            document.getElementById(
                "chartArea"
            ),
            {

                type:"line",

                data:{

                    labels,

                    datasets:[{

                        label:
                            "Acumulado",

                        data:valores,

                        fill:true,

                        borderColor:
                            "#10b981",

                        backgroundColor:
                            "rgba(16,185,129,.2)"
                    }]
                }
            }
        );
}

// ========================================
// GRAFICO 5
// TOP CLIENTES
// ========================================

function construirTopClientes(){

    const clientes = {};

    datosFiltrados.forEach(r=>{

        const cliente =
            r["Nombre de Cliente"] ||
            r["Cliente"];

        clientes[cliente] =
            (clientes[cliente] || 0)
            +
            Number(
                r["Cajas fisica"] || 0
            );
    });

    const top =
        Object.entries(clientes)
            .sort(
                (a,b)=>
                    b[1]-a[1]
            )
            .slice(0,10);

    charts.ranking =
        new Chart(
            document.getElementById(
                "chartRanking"
            ),
            {

                type:"bar",

                data:{

                    labels:
                        top.map(
                            x=>x[0]
                        ),

                    datasets:[{

                        label:
                            "Top Clientes",

                        data:
                            top.map(
                                x=>x[1]
                            )
                    }]
                },

                options:{
                    indexAxis:'y'
                }
            }
        );
}

// ========================================
// GRAFICO 6
// TOP SKU
// ========================================

function construirTopSKU(){

    const sku = {};

    datosFiltrados.forEach(r=>{

        const codigo =
            r["SKU"];

        sku[codigo] =
            (sku[codigo] || 0)
            +
            Number(
                r["Cajas fisica"] || 0
            );
    });

    const top =
        Object.entries(sku)
            .sort(
                (a,b)=>
                    b[1]-a[1]
            )
            .slice(0,10);

    charts.topsku =
        new Chart(
            document.getElementById(
                "chartSKU"
            ),
            {

                type:"bar",

                data:{

                    labels:
                        top.map(
                            x=>x[0]
                        ),

                    datasets:[{

                        label:
                            "Top SKU",

                        data:
                            top.map(
                                x=>x[1]
                            ),

                        backgroundColor:
                            "#f59e0b"
                    }]
                }
            }
        );
}

// ========================================
// RECHAZOS CLIENTE
// ========================================

function construirRechazosCliente(){

    const datos = {};

    datosFiltrados
        .filter(r => {

            return String(
                r["Movimiento"] || ""
            )
            .toLowerCase()
            .includes("rechazo");

        })
        .forEach(r => {

            const cliente =
                r["Nombre de Cliente"] ||
                r["Cliente"] ||
                "Sin Cliente";

            datos[cliente] =
                (datos[cliente] || 0)
                +
                Number(
                    r["Cajas fisica"] || 0
                );
        });

    const top =
        Object.entries(datos)
            .sort(
                (a,b)=>b[1]-a[1]
            )
            .slice(0,15);

    const canvas =
        document.getElementById(
            "chartRechazosCliente"
        );

    if(!canvas) return;

    charts.rechazosCliente =
        new Chart(canvas, {

            type:"bar",

            data:{

                labels:
                    top.map(
                        x=>x[0]
                    ),

                datasets:[{

                    label:
                        "Rechazos",

                    data:
                        top.map(
                            x=>x[1]
                        ),

                    backgroundColor:
                        "#ef4444"
                }]
            },

            options:{
                indexAxis:'y'
            }
        });
}


// ========================================
// RECHAZOS CET
// ========================================

function construirRechazosCET(){

    const datos = {};

    datosFiltrados
        .filter(r => {

            return String(
                r["Movimiento"] || ""
            )
            .toLowerCase()
            .includes("rechazo");

        })
        .forEach(r => {

            const cet =
                r["CET"] ||
                "Sin CET";

            datos[cet] =
                (datos[cet] || 0)
                +
                Number(
                    r["Cajas fisica"] || 0
                );
        });

    const top =
        Object.entries(datos)
            .sort(
                (a,b)=>b[1]-a[1]
            )
            .slice(0,15);

    const canvas =
        document.getElementById(
            "chartRechazosCET"
        );

    if(!canvas) return;

    charts.rechazosCET =
        new Chart(canvas, {

            type:"bar",

            data:{

                labels:
                    top.map(
                        x=>x[0]
                    ),

                datasets:[{

                    label:
                        "Rechazos CET",

                    data:
                        top.map(
                            x=>x[1]
                        ),

                    backgroundColor:
                        "#f97316"
                }]
            }
        });
}

// ========================================
// HEATMAP
// ========================================

function construirHeatmap(){

    const contenedor =
        document.getElementById(
            "heatmapContainer"
        );

    if(!contenedor) return;

    contenedor.innerHTML = "";

    const clases =
        [...new Set(
            datosFiltrados.map(
                x=>x["Clase"]
            )
        )]
        .filter(Boolean);

    const movimientos =
        [...new Set(
            datosFiltrados.map(
                x=>x["Movimiento"]
            )
        )]
        .filter(Boolean);

    let html =
        '<table class="heatmap-table">';

    html +=
        '<tr><th>Clase</th>';

    movimientos.forEach(m=>{

        html +=
            `<th>${m}</th>`;
    });

    html +=
        '</tr>';

    clases.forEach(clase=>{

        html +=
            `<tr><th>${clase}</th>`;

        movimientos.forEach(m=>{

            const total =
                datosFiltrados.filter(
                    x =>

                        x["Clase"] === clase

                        &&

                        x["Movimiento"] === m
                ).length;

            const intensidad =
                Math.min(
                    total / 500,
                    1
                );

            html += `
            <td
            style="
            background:
            rgba(
                0,
                120,
                212,
                ${intensidad}
            );
            ">
            ${total}
            </td>
            `;
        });

        html += '</tr>';
    });

    html += '</table>';

    contenedor.innerHTML =
        html;
}