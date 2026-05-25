async function cargarCartelera() {
            const container = document.getElementById('events-container');
            const tituloSeccion = document.getElementById('main-title');
            
            try {
                const response = await fetch(`https://ugm-2.vercel.app/api/cartelera?nocache=${Date.now()}`);
                if (!response.ok) throw new Error('Error API');
                
                const eventosRaw = await response.json();
                container.innerHTML = '';
                
                if (!eventosRaw || eventosRaw.length === 0 || eventosRaw.error) {
                    container.innerHTML = '<div class="no-events">No hay eventos registrados en SharePoint</div>';
                    return;
                }

                // Fecha y Hora de control en tiempo real de Chile
                const hoyChileString = new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Santiago' }); 
                const horaChileString = new Date().toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'America/Santiago' });

                // Procesamos aplicando la guillotina sobre la columna "Inicio"
                const listaProcesada = eventosRaw.map(ev => {
                    let fechaTexto = "9999-12-31";
                    let horaTexto = "00:00";
                    let esHoy = false;

                    if (ev.casillaTiempo && typeof ev.casillaTiempo === 'string') {
                        const partes = ev.casillaTiempo.split('T');
                        if (partes[0] && partes[1]) {
                            fechaTexto = partes[0]; 
                            esHoy = (fechaTexto === hoyChileString);
                            
                            const subPartesHora = partes[1].split(':');
                            let horaOriginal = parseInt(subPartesHora[0], 10);
                            let minutosOriginales = subPartesHora[1] || "00";

                            if (!isNaN(horaOriginal)) {
                                // CORRECCIÓN AQUÍ: Quitamos 2 horas exactas al tiempo de SharePoint
                                let horaAjustada = (horaOriginal - 2 + 24) % 24;
                                horaTexto = `${String(horaAjustada).padStart(2, '0')}:${minutosOriginales.substring(0,2)}`;
                            }
                        }
                    }

                    return {
                        title: ev.title,
                        sala: ev.sala,
                        hora: horaTexto,
                        fechaStr: fechaTexto,
                        esHoy: esHoy
                    };
                });

                let muestraFinal = [];

                if (modoVistaSemanas) {
                    tituloSeccion.innerText = "PRÓXIMOS EVENTOS";
                    muestraFinal = listaProcesada.filter(ev => !ev.esHoy && ev.fechaStr > hoyChileString);
                } else {
                    tituloSeccion.innerText = "EVENTOS DE HOY";
                    // Filtra eventos de hoy que aún no terminen/pasen de la hora actual
                    muestraFinal = listaProcesada.filter(ev => ev.esHoy && ev.hora >= horaChileString);
                }

                if (muestraFinal.length === 0 && !modoVistaSemanas) {
                    tituloSeccion.innerText = "PRÓXIMOS EVENTOS";
                    muestraFinal = listaProcesada.filter(ev => !ev.esHoy && ev.fechaStr > hoyChileString);
                }

                if (muestraFinal.length === 0) {
                    container.innerHTML = '<div class="no-events">No hay más eventos programados</div>';
                    return;
                }

                muestraFinal.sort((a, b) => `${a.fechaStr}T${a.hora}`.localeCompare(`${b.fechaStr}T${b.hora}`));

                container.innerHTML = '';
                
                muestraFinal.forEach(ev => {
                    let etiquetaDia = "Hoy";
                    
                    if (!ev.esHoy && ev.fechaStr !== "9999-12-31") {
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

                    const card = document.createElement('div');
                    card.className = 'event-card';
                    card.innerHTML = `
                        <div class="event-info">
                            <div class="event-title">${ev.title}</div>
                            <div class="event-room">${ev.sala}</div>
                        </div>
                        <div class="event-time">
                            <div class="event-day-label">${etiquetaDia}</div>
                            <div>${ev.hora}</div>
                        </div>
                    `;
                    container.appendChild(card);
                });

                modoVistaSemanas = !modoVistaSemanas;
                setTimeout(iniciarScroll, 1000);

            } catch (error) {
                console.error(error);
                container.innerHTML = '<div class="no-events">⚠️ Error de conexión</div>';
            }
        }
