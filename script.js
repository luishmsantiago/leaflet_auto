        let map = L.map('map').setView([-26.99, -48.63], 13);
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '© OpenStreetMap contributors'
        }).addTo(map);

        let excelData = [];
        let mapLayers = L.layerGroup().addTo(map);

        let modoAtivo = 'wkt';
        function mudarAba(aba) {
            modoAtivo = aba;
            document.querySelectorAll('.aba').forEach(el => el.classList.remove('ativa'));
            document.querySelectorAll('.painel').forEach(el => el.classList.remove('ativo'));
            event.target.classList.add('ativa');
            document.getElementById(`painel-${aba}`).classList.add('ativo');
        }

        const delay = ms => new Promise(res => setTimeout(res, ms));

        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            document.getElementById('status').innerText = "Lendo arquivo...";
            
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                excelData = XLSX.utils.sheet_to_json(firstSheet, {defval: ""});
                
                if (excelData.length > 0) {
                    const colunas = Object.keys(excelData[0]);
                    preencherSelects(colunas);
                    document.getElementById('status').innerText = `${excelData.length} linhas carregadas!`;
                }
            };
            reader.readAsArrayBuffer(file);
        });

        function preencherSelects(colunas) {
            const selects = ['colDesc', 'colWkt', 'colEndereco'];
            selects.forEach(id => {
                const el = document.getElementById(id);
                el.innerHTML = '<option value="">-- Selecione --</option>';
                colunas.forEach(col => {
                    el.innerHTML += `<option value="${col}">${col}</option>`;
                });
            });
        }
        
        // ==========================================
        // EXPORTAR PARA HTML INDEPENDENTE (SIMPLIFICADO)
        // ==========================================
        function downloadMapa() {
            if (mapLayers.getLayers().length === 0) return alert("Gere o mapa antes de baixar!");

            // 1. Pega os dados que estão desenhados no mapa
            let feicoes = [];
            mapLayers.eachLayer(function(layer) {
                if (layer.toGeoJSON) {
                    let geojson = layer.toGeoJSON();
                    if (!geojson.properties) geojson.properties = {};
                    
                    // Salva o texto do balãozinho (popup)
                    if (layer.getPopup && layer.getPopup()) {
                        geojson.properties.popup = layer.getPopup().getContent();
                    }
                    
                    // Salva as cores
                    if (layer.options) {
                        geojson.properties.color = layer.options.color || '#3388ff';
                        geojson.properties.weight = layer.options.weight || 3;
                    }
                    feicoes.push(geojson);
                }
            });

            const geojsonData = { type: "FeatureCollection", features: feicoes };

            // 2. Cria as tags de script de forma separada para não quebrar o navegador
            const scriptStart = '<scr' + 'ipt>';
            const scriptEnd = '</scr' + 'ipt>';
            const leafletScript = '<scr' + 'ipt src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></scr' + 'ipt>';

            // 3. Monta a página HTML final
            const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <title>Mapa Exportado - Table2Geo</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
        body { padding: 0; margin: 0; }
        html, body, #map { height: 100%; width: 100%; }
    </style>
</head>
<body>
    <div id="map"></div>
    ${leafletScript}
    ${scriptStart}
        var map = L.map('map').setView([-26.99, -48.63], 13);
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '© OpenStreetMap contributors'
        }).addTo(map);

        var data = ${JSON.stringify(geojsonData)};

        var layer = L.geoJSON(data, {
            style: function(feature) {
                return { color: feature.properties.color, weight: feature.properties.weight };
            },
            onEachFeature: function(feature, layer) {
                if (feature.properties.popup) {
                    layer.bindPopup(feature.properties.popup);
                }
            }
        }).addTo(map);

        if (Object.keys(layer._layers).length > 0) {
            map.fitBounds(layer.getBounds());
        }
    ${scriptEnd}
</body>
</html>`;

            // 4. Baixa o arquivo
            const blob = new Blob([htmlContent], {type: "text/html;charset=utf-8"});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'meu_mapa_interativo.html';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        async function gerarMapa() {
            if (excelData.length === 0) return alert("Carregue a planilha primeiro!");
            
            mapLayers.clearLayers();
            const statusEl = document.getElementById('status');
            
            const colDesc = document.getElementById('colDesc').value;
            const tipoGeo = document.getElementById('tipoGeo').value;
            let sucessos = 0;

            for (let i = 0; i < excelData.length; i++) {
                const row = excelData[i];
                const desc = row[colDesc] || `Item ${i+1}`;
                statusEl.innerText = `Processando: ${i+1} de ${excelData.length}...`;

                try {
                    // ==========================================
                    // MODO WKT (Usando Terraformer)
                    // ==========================================
                    if (modoAtivo === 'wkt') {
                        const colWkt = document.getElementById('colWkt').value;
                        const valWkt = String(row[colWkt] || "").trim();
                        
                        if (valWkt && valWkt.toLowerCase() !== "nan" && valWkt !== "undefined") {
                            try {
                                const geojson = Terraformer.WKT.parse(valWkt);
                                L.geoJSON(geojson, {
                                    style: { 
                                        color: tipoGeo === 'Point' ? 'blue' : (tipoGeo === 'Polygon' ? '#ff7800' : '#3388ff'), 
                                        weight: 4 
                                    },
                                    pointToLayer: function (feature, latlng) {
                                        return L.marker(latlng);
                                    }
                                }).bindPopup(`<b>${desc}</b>`).addTo(mapLayers);
                                sucessos++;
                            } catch(e) {
                                console.warn(`WKT inválido na linha ${i}:`, valWkt);
                            }
                        }
                    } 
                    // ==========================================
                    // MODO ENDEREÇO
                    // ==========================================
                    else {
                        const colEnd = document.getElementById('colEndereco').value;
                        let busca = String(row[colEnd] || "").trim();
                        
                        if (!busca || busca === "" || busca.toLowerCase() === "nan") continue;

                        if (!busca.toUpperCase().includes("SC") && !busca.toLowerCase().includes("catarina")) {
                            busca += ", SC";
                        }

                        if (tipoGeo === "Point") {
                            const urlArcgis = `https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates?f=json&singleLine=${encodeURIComponent(busca + ", Brasil")}&maxLocations=1`;
                            const req = await fetch(urlArcgis);
                            if (req.ok) {
                                const res = await req.json();
                                if (res.candidates && res.candidates.length > 0) {
                                    const pt = res.candidates[0].location;
                                    L.marker([pt.y, pt.x]).bindPopup(`<b>${desc}</b>`).addTo(mapLayers);
                                    sucessos++;
                                }
                            }
                        }
                        
                        else if (tipoGeo === "LineString") {
                            // Isola apenas o nome da rua e tira os números
                            let nomeRua = busca.split(',')[0].replace(/\d+/g, '').trim(); 
                            if (!nomeRua) continue;
                            
                            // Tradutor de ruas numéricas
                            const traducaoRuas = { "3300": "Três Mil e Trezentos", "1500": "Mil e Quinhentos", "3100": "Três Mil e Cem" };
                            for (let num in traducaoRuas) {
                                nomeRua = nomeRua.replace(num, traducaoRuas[num]);
                            }

                            let cidade = "Balneário Camboriú";
                            if (busca.includes("Itajaí")) cidade = "Itajaí";
                            else if (busca.includes("Camboriú") && !busca.includes("Balneário")) cidade = "Camboriú";

                            // ==========================================
                            // NOVA ABORDAGEM: Usando NOMINATIM ao invés do Overpass
                            // ==========================================
                            const queryRua = `${nomeRua}, ${cidade}, SC`;
                            const urlNom = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(queryRua)}&format=json&polygon_geojson=1&limit=20`;
                            
                            await delay(1200); // Respeita o limite do servidor público
                            const req = await fetch(urlNom);
                            
                            if (req.ok) {
                                const res = await req.json();
                                let achouLinha = false;
                                
                                if (res.length > 0) {
                                    for (let resultado of res) {
                                        // Busca nos resultados o que for do tipo linha (LineString ou MultiLineString)
                                        if (resultado.geojson && resultado.geojson.type.includes("LineString")) {
                                            L.geoJSON(resultado.geojson, {
                                                style: { color: '#3388ff', weight: 6, opacity: 0.8 }
                                            }).bindPopup(`<b>${desc}</b>`).addTo(mapLayers);
                                            achouLinha = true;
                                        }
                                    }
                                    
                                    if (achouLinha) {
                                        sucessos++;
                                    } else {
                                        // Fallback: Se achou a rua, mas não tem o desenho da linha, bota um pino
                                        L.circleMarker([res[0].lat, res[0].lon], {radius: 8, color: 'blue'}).bindPopup(`${desc} (Apenas marcador central encontrado)`).addTo(mapLayers);
                                        sucessos++;
                                    }
                                }
                            } else {
                                console.warn(`Falha na API para a rua ${nomeRua}. Status: ${req.status}`);
                            }
                        }                      
                        else if (tipoGeo === "Polygon") {
                            let buscaLimpa = busca.replace(/Bairro/ig, "").trim();
                            const urlNom = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(buscaLimpa)}&format=json&polygon_geojson=1&limit=50`;
                            
                            await delay(1200); 
                            const req = await fetch(urlNom);
                            
                            if (req.ok) {
                                const res = await req.json();
                                let achou = false;
                                if (res.length > 0) {
                                    for (let resultado of res) {
                                        if (resultado.geojson && resultado.geojson.type.includes("Polygon")) {
                                            L.geoJSON(resultado.geojson, {
                                                style: { fillColor: '#ff7800', color: '#000', weight: 2 }
                                            }).bindPopup(`<b>${desc}</b>`).addTo(mapLayers);
                                            achou = true;
                                            break;
                                        }
                                    }
                                    if (!achou) {
                                        L.circle([res[0].lat, res[0].lon], {radius: 50, color: 'red'}).bindPopup(`${desc} (Polígono não encontrado)`).addTo(mapLayers);
                                    }
                                    sucessos++;
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.error(`Erro na linha ${i}:`, error);
                }
            }

            statusEl.innerText = `Concluído! ${sucessos} feições no mapa.`;
            statusEl.style.color = "green";
        }