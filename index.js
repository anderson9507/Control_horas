//Funcion para mostrar la fecha actual en el input de fecha
    window.addEventListener('DOMContentLoaded', function () {
        const inputFecha = document.getElementById('Fecha');
        if (inputFecha) {
            const hoy = new Date();
            const yyyy = hoy.getFullYear();
            const mm = String(hoy.getMonth() + 1).padStart(2, '0');
            const dd = String(hoy.getDate()).padStart(2, '0');
            inputFecha.value = `${yyyy}-${mm}-${dd}`;
        }
    });

document.querySelector('form').addEventListener('submit', function (event) {
    event.preventDefault(); // Evitar el envío del formulario

    // Obtener los valores de los campos y asignar 0 si están vacíos
    const nombre = document.getElementById('Nombre').value || "DatosTrabajador";
    if(nombre){
        document.getElementById('Nombre').value = nombre.replace(/\s+/g, '');
    }
    const fecha = document.getElementById('Fecha').value;
    const horas = parseFloat(document.getElementById('Horas').value) || 0;
    const dias = parseFloat(document.getElementById('Dias').value) || 0;
    const horasExtra = parseFloat(document.getElementById('HorasExtra').value) || 0;
    const horasNocturnas = parseFloat(document.getElementById('HorasNocturnas').value) || 0;
    const horasExtraNocturnas = parseFloat(document.getElementById('HorasExtraNocturnas').value) || 0;
    const horasFaltantes = parseFloat(document.getElementById('HorasFaltantes').value) || 0;
    const sextoDia = parseFloat(document.getElementById('SextoDia').value) || 0;

    // Calcular el total de horas trabajadas
    const totalHoras = (horas * dias) + horasExtra - horasFaltantes;
    // const totalSextoDia = (8 * sextoDia); // Calcular las horas del sexto día
    // const totalNocturnas = (horasNocturnas + horasExtraNocturnas); // Calcular las horas nocturnas
    const totalFinal = parseInt(totalHoras);

    // Crear un objeto para almacenar los datos
    const datosTrabajador = {
        "Fecha": fecha,
        "Nombre": nombre,
        "Horas por Día": horas,
        "Días Trabajados": dias,
        "Horas Extra": horasExtra,
        "Horas Nocturnas": horasNocturnas,
        "Horas Extra Nocturnas": horasExtraNocturnas,
        "Horas Faltantes": horasFaltantes,
        "Sexto Día": sextoDia,
        "Total Horas": totalFinal

    };

    // Mostrar los datos en la tabla
    const tablaResultados = document.getElementById('tablaResultados');
    const nombreArchivo = document.getElementById('nombreArchivo');
    if(nombreArchivo) {
        nombreArchivo.textContent = nombre;
    }

    tablaResultados.innerHTML = `
        <tr>
            <td>${datosTrabajador["Fecha"]}</td>
            <td>${datosTrabajador["Horas por Día"]}</td>
            <td>${datosTrabajador["Días Trabajados"]}</td>
            <td>${datosTrabajador["Horas Extra"]}</td>
            <td>${datosTrabajador["Horas Nocturnas"]}</td>
            <td>${datosTrabajador["Horas Faltantes"]}</td>
            <td>${datosTrabajador["Sexto Día"]}</td>
            <td>${datosTrabajador["Total Horas"]}</td>
        </tr>
    `;

    // ocultar el formulario
    document.getElementById('formulario').classList.add('hidden');
    // Mostrar la tabla quitando la clase 'hidden'
    const tablaContainer = document.getElementById('tablaContainer');
    tablaContainer.classList.remove('hidden');

    // Guardar los datos en una variable global para usarlos en otras funciones
    window.datosTrabajador = datosTrabajador;

    // Restablecer los valores de los inputs
    event.target.reset();

    const inputFecha = document.getElementById('Fecha');
    if (inputFecha) {
        const hoy = new Date();
        const yyyy = hoy.getFullYear();
        const mm = String(hoy.getMonth() + 1).padStart(2, '0');
        const dd = String(hoy.getDate()).padStart(2, '0');
        inputFecha.value = `${yyyy}-${mm}-${dd}`;
    }
});

// Botón para subir a Google Drive
document.getElementById('btnSubirDrive').addEventListener('click', async function () {
    const datosTrabajador = window.datosTrabajador;
    if (!datosTrabajador) {
        Swal.fire({
            title: 'Ups!',
            text: "Primero llena y envía el formulario.",
            icon: 'info',
            confirmButtonText: 'Aceptar',
            width: '40%'
        });
        return;
    }

    const CLIENT_ID = '824044692752-g7aoj9jpao8lte2tbknp0eefvbhldqc2.apps.googleusercontent.com';

    const tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: 'https://www.googleapis.com/auth/drive.file',
        callback: async (tokenResponse) => {
            if (!tokenResponse || !tokenResponse.access_token) {
                Swal.fire({
                    title: 'Error',
                    text: "No se pudo obtener el token de acceso. Por favor, inténtalo de nuevo.",
                    icon: 'error',
                    confirmButtonText: 'Aceptar',
                    width: '40%'
                });
                return;
            }

            const accessToken = tokenResponse.access_token;

            const nombreArchivoLimpio = (datosTrabajador["Nombre"] || "DatosTrabajador")
    .replace(/[\s\\\/\?\*\[\]\:]+/g, '');

            // Generar archivo Excel
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.json_to_sheet([datosTrabajador]);
            XLSX.utils.book_append_sheet(workbook, worksheet, "hoja1"); // Nombre de la hoja limpio
            const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
            const excelFile = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            

            const metadata = {
                name: nombreArchivoLimpio + ".xlsx", // Nombre de archivo limpio y personalizado
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            };

            const formData = new FormData();
            formData.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
            formData.append("file", excelFile);

            try {
                const response = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${accessToken}`
                    },
                    body: formData
                });

                if (!response.ok) {
                    throw new Error(`Error HTTP ${response.status}`);
                }

                const data = await response.json();
                Swal.fire({
                    title: 'Éxito',
                    text: `Archivo subido exitosamente: ${data.name}`,
                    icon: 'success',
                    confirmButtonText: 'Aceptar',
                    width: '40%'
                });
            } catch (error) {
                console.error("Error al subir el archivo:", error);
                Swal.fire({
                    title: 'Error',
                    text: "Error al subir el archivo a Google Drive.",
                    icon: 'error',
                    confirmButtonText: 'Aceptar',
                    width: '40%'
                });
            }
        }
    });

    // Solicitar el token (desencadena login si es necesario)
    tokenClient.requestAccessToken();
});

// Botón para descargar PDF
document.getElementById('btnDescargarPDF').addEventListener('click', function () {
    const datosTrabajador = window.datosTrabajador;
    if (!datosTrabajador) {
        Swal.fire({
            title: 'Error',
            text: "Primero llena y envía el formulario.",
            icon: 'info',
            confirmButtonText: 'Aceptar',
            width: '40%'
        });
        return;
    }

    const doc = new window.jspdf.jsPDF();
    doc.text("Datos del Trabajador", 10, 10);

    const headers = [["Campo", "Valor"]];
    const rows = Object.entries(datosTrabajador);

    doc.autoTable({
        head: headers,
        body: rows,
        startY: 20
    });

    // abrir el PDF en una nueva pestaña
    const pdfData = doc.output('bloburl');
    window.open(pdfData, "_blank");
});

    // Limpiar la tabla
document.getElementById('btnnuevo').addEventListener('click', function () {
    const tablaResultados = document.getElementById('tablaResultados');
    tablaResultados.innerHTML = '';

    // Ocultar la tabla
    const tablaContainer = document.getElementById('tablaContainer');
    tablaContainer.classList.add('hidden');

    // Mostrar el formulario
    document.getElementById('formulario').classList.remove('hidden');

    // Restablecer los valores de los inputs
    document.querySelector('form').reset();

    const inputFecha = document.getElementById('Fecha');
    if (inputFecha) {
        const hoy = new Date();
        const yyyy = hoy.getFullYear();
        const mm = String(hoy.getMonth() + 1).padStart(2, '0');
        const dd = String(hoy.getDate()).padStart(2, '0');
        inputFecha.value = `${yyyy}-${mm}-${dd}`;
    }
});



