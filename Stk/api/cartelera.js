async function cargarCartelera() {
            const container = document.getElementById('events-container');
            const tituloSeccion = document.getElementById('main-title');
            
            try {
                // 🔗 Conexión real con tu backend ligado a SharePoint
                const response = await fetch(`https://ugm-2.vercel.app/api/cartelera?nocache=${Date.now()}`);
                if (!response.ok) throw new Error('Error API');
                
                const eventosRaw = await response.json();
                container.innerHTML = '';
                
                if (!eventosRaw || eventosRaw.length === 0 || eventosRaw.error) {
                    container.innerHTML = '<div class="no-events">No hay eventos registrados en SharePoint</div>';
                    container.classList.add('fade-in');
                    rotationTimeout = setTimeout(siguientePantalla, 10000);
                    return;
                }

                const ahoraChile = new Date(new Date().toLocaleString("en-US", {timeZone: "America/Santiago"}));
                const hoyChileString = ahoraChile.toLocaleDateString('sv-SE'); 

                const listaProcesada = [];

                eventosRaw.forEach(ev => {
                    // Mapeamos los campos según los nombres tradicionales que entrega tu API estable
                    const titulo = ev.title || ev.Title || "Evento sin título";
                    const sala = ev.sala || ev.Sala || "Por definir";
                    const tiempo = ev.casillaTiempo || ev.Fecha || ev.EventDate || null;
                    
                    if (tiempo && typeof tiempo === 'string') {
                        const dateObjeto = new Date(tiempo);
                        
                        if (!isNaN(dateObjeto.getTime())) {
                            const esTodoElDia = tiempo.includes("T00:00:00") || tiempo.includes("T02:00:00");
                            
                            let fechaTexto = dateObjeto.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });
                            let horaTexto = dateObjeto.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'America/Santiago' });
                            
                            if (esTodoElDia) {
                                fechaTexto = tiempo.split('T')[0];
                                horaTexto = "Todo el día";
                            }

                            const esHoy = (fechaTexto === hoyChileString);
                            let etiquetaEstado = "Hoy";

                            if (esHoy && !esTodoElDia) {
                                const horaEventoObj = new Date(dateObjeto.toLocaleString("en-US", {timeZone: "America/Santiago"}));
                                const diferenciaMinutos = (ahoraChile - horaEventoObj) / 60000;

                                // Filtro dinámico de 30 minutos
                                if (diferenciaMinutos > 30) {
                                    return; 
                                }
                                
                                if (diferenciaMinutos >= 0 && diferenciaMinutos <= 30) {
                                    etiquetaEstado = "En curso";
                                }
                            }

                            listaProcesada.push({
                                title: titulo,
                                sala: sala,
                                rawDestino: ev.Destinatario || ev.destinatario || null, // Captura si la API lo manda
                                hora: horaTexto,
                                fechaStr: fechaTexto,
                                esHoy: esHoy,
                                etiquetaEstado: etiquetaEstado
                            });
                        }
                    }
                });

                let muestraFinal = [];

                if (modoVistaSemanas) {
                    tituloSeccion.innerText = "PRÓXIMOS EVENTOS";
                    const fechaActual = new Date();
                    const diaSemana = fechaActual.getDay(); 
                    let diasHastaFinDeSemana = 7 - diaSemana;
                    if (diaSemana === 0) diasHastaFinDeSemana = 0; 
                    if (diaSemana === 5 || diaSemana === 6 || diaSemana === 0) diasHastaFinDeSemana += 7;

                    const fechaLimiteObj = new Date();
                    fechaLimiteObj.setDate(fechaActual.getDate() + diasHastaFinDeSemana);
                    const limiteString = fechaLimiteObj.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

                    muestraFinal = listaProcesada.filter(ev => !ev.esHoy && ev.fechaStr > hoyChileString && ev.fechaStr <= limiteString);
                } else {
                    tituloSeccion.innerText = "EVENTOS DE HOY";
                    muestraFinal = listaProcesada.filter(ev => ev.esHoy);
                }

                if (muestraFinal.length === 0 && !modoVistaSemanas) {
                    tituloSeccion.innerText = "PRÓXIMOS EVENTOS";
                    const fechaActual = new Date();
                    const diaSemana = fechaActual.getDay();
                    let diasHastaFinDeSemana = 7 - diaSemana;
                    if (diaSemana === 0) diasHastaFinDeSemana = 0;
                    if (diaSemana === 5 || diaSemana === 6 || diaSemana === 0) diasHastaFinDeSemana += 7;

                    const fechaLimiteObj = new Date();
                    fechaLimiteObj.setDate(fechaActual.getDate() + diasHastaFinDeSemana);
                    const limiteString = fechaLimiteObj.toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' });

                    muestraFinal = listaProcesada.filter(ev => !ev.esHoy && ev.fechaStr > hoyChileString && ev.fechaStr <= limiteString);
                }

                if (muestraFinal.length === 0) {
                    container.innerHTML = '<div class="no-events">No hay más eventos programados</div>';
                    container.classList.add('fade-in');
                    modoVistaSemanas = !modoVistaSemanas;
                    rotationTimeout = setTimeout(siguientePantalla, 8000);
                    return;
                }

                muestraFinal.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

                container.innerHTML = '';
                
                muestraFinal.forEach(ev => {
                    let etiquetaDia = ev.etiquetaEstado;
                    let claseExtra = "";

                    if (etiquetaDia === "En curso") {
                        claseExtra = "badge-en-curso";
                    }
                    
                    if (!ev.esHoy) {
                        try {
                            const [aaaa, mes, dd] = ev.fechaStr.split('-');
                            const d = new Date(parseInt(aaaa,10), parseInt(mes,10) - 1, parseInt(dd,10), 12, 0, 0);
                            const nombreDia = d.toLocaleDateString('es-CL', { weekday: 'short' }).replace('.', '');
                            const numDia = d.getDate();
                            etiquetaDia = `${nombreDia.charAt(0).toUpperCase() + nombreDia.slice(1)} ${numDia}`;
                        } catch (e) {
                            etiquetaDia = "Próximo";
                        }
                    }

                    // Motor híbrido: Lee la columna de SharePoint, y si viene vacía deduce por palabra clave del título
                    const metaPublico = analizarDestinatario(ev.title, ev.rawDestino);

                    const card = document.createElement('div');
                    card.className = 'event-card';
                    
                    card.innerHTML = `
                        <div class="event-info">
                            <div class="event-title">${ev.title}</div>
                            <div class="event-badges">
                                <div class="event-room">${ev.sala}</div>
                                <div class="event-destinatario" style="background-color: ${metaPublico.color};">${metaPublico.texto}</div>
                            </div>
                        </div>
                        <div class="event-time">
                            <div class="event-day-label ${claseExtra}">${etiquetaDia}</div>
                            <div>${ev.hora}</div>
                        </div>
                    `;
                    container.appendChild(card);
                });

                container.classList.add('fade-in');
                setTimeout(iniciarManejoPantalla, 400);

            } catch (error) {
                console.error(error);
                container.innerHTML = '<div class="no-events">⚠️ Error de conexión con SharePoint</div>';
                container.classList.add('fade-in');
                rotationTimeout = setTimeout(siguientePantalla, 10000);
            }
        }
