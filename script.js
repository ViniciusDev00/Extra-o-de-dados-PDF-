const pdfjsLib = window['pdfjs-dist/build/pdf'];
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

document.getElementById('processBtn').addEventListener('click', async () => {
    const file = document.getElementById('pdfInput').files[0];
    if (!file) return alert("Selecione um PDF");

    const reader = new FileReader();
    reader.onload = async function() {
        const typedarray = new Uint8Array(this.result);
        const pdf = await pdfjsLib.getDocument(typedarray).promise;
        let fullData = [];
        
        let orfaNumber = file.name.match(/\d+[\/-]\d+/) ? file.name.match(/\d+[\/-]\d+/)[0] : "270/2025";

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            // Junta o texto mantendo espaçamento básico
            const lines = textContent.items.map(item => item.str.trim()).filter(t => t.length > 0);
            const pageText = lines.join(" ");

            // 1. Mapeia Onde Estão Todos os Expositores na Página
            const expositorRegex = /([A-Z0-9]{2,}\s?P?V?T?\s?-\s?C\d+.*?\( \d+\.\d+.*?\))/g;
            let expositores = [];
            let match;
            while ((match = expositorRegex.exec(pageText)) !== null) {
                expositores.push({ 
                    name: match[0].trim(), 
                    start: match.index, 
                    end: match.index + match[0].length 
                });
            }

            // 2. Processa cada Expositor dentro da sua "Zona Segura"
            for (let k = 0; k < expositores.length; k++) {
                const exp = expositores[k];
                const prevExp = expositores[k - 1];
                const nextExp = expositores[k + 1];

                // DEFINIÇÃO DA FATIA (SLICE)
                // Começa no meio do caminho do anterior (ou 0 se for o primeiro)
                // Termina no meio do caminho do próximo (ou fim da página se for o último)
                let sliceStart = prevExp ? Math.floor((prevExp.end + exp.start) / 2) : 0;
                let sliceEnd = nextExp ? Math.floor((exp.end + nextExp.start) / 2) : pageText.length;

                // Garante que pegamos texto suficiente para trás se for o primeiro da página (casos de cabeçalho)
                if (k === 0) sliceStart = Math.max(0, exp.start - 600); 

                // Extrai apenas o texto que pertence a este expositor
                let rawSlice = pageText.substring(sliceStart, sliceEnd);

                let dirFinal = "Vazio";
                let esqFinal = "Vazio";
                let resultadoUnificado = "";

                // --- DETECÇÃO DE ESTRUTURA DENTRO DA FATIA ---
                const idxDir = rawSlice.indexOf("Lateral Direita");
                const idxEsq = rawSlice.indexOf("Lateral Esquerda");
                const idxDiv = rawSlice.indexOf("Div. Interna");

                // Verifica se é o layout de tabela (Geralmente Pág 1 ou tabelas compactas)
                // Critério: "Lateral Direita" e "Esquerda" estão próximas (< 150 chars)
                const isTableLayout = (idxDir !== -1 && idxEsq !== -1 && Math.abs(idxEsq - idxDir) < 150);

                if (isTableLayout && idxDiv !== -1) {
                    // === PARSER MODO TABELA ===
                    // A linha de dados geralmente começa após "Div. Interna"
                    let contentStart = idxDiv + "Div. Interna".length;
                    
                    // Procura o fim da linha (Rodapé, Acabamento, ou fim da fatia)
                    let stops = [
                        rawSlice.indexOf("Rodapé", contentStart),
                        rawSlice.indexOf("Acabamento", contentStart),
                        rawSlice.indexOf("Acb.", contentStart)
                    ].filter(x => x !== -1);
                    
                    let contentEnd = stops.length > 0 ? Math.min(...stops) : rawSlice.length;
                    
                    let rowContent = rawSlice.substring(contentStart, contentEnd).trim();
                    rowContent = rowContent.replace(/["]/g, ""); // Remove aspas

                    // Limpa palavras proibidas (lixo da próxima coluna)
                    rowContent = cleanGarbage(rowContent);

                    // Divide usando Palavras-Chave
                    let parsed = parseByKeywords(rowContent);
                    dirFinal = parsed.dir;
                    esqFinal = parsed.esq;

                } else {
                    // === PARSER MODO LISTA (Seguro) ===
                    // Procura Lateral Direita
                    if (idxDir !== -1) {
                        let start = idxDir + "Lateral Direita".length;
                        // Para antes de: Rodapé, Div, Acabamento, Lateral Esquerda (se vier depois), ou Informações
                        let stops = [
                            rawSlice.indexOf("Rodapé", start),
                            rawSlice.indexOf("Div.", start),
                            rawSlice.indexOf("Acabamento", start),
                            rawSlice.indexOf("Informações", start),
                            (idxEsq > start) ? idxEsq : -1 // Não invadir a Esquerda
                        ].filter(x => x > start);
                        
                        let end = stops.length > 0 ? Math.min(...stops) : rawSlice.length;
                        dirFinal = rawSlice.substring(start, end).trim();
                    }

                    // Procura Lateral Esquerda
                    if (idxEsq !== -1) {
                        let start = idxEsq + "Lateral Esquerda".length;
                        let stops = [
                            rawSlice.indexOf("Rodapé", start),
                            rawSlice.indexOf("Div.", start),
                            rawSlice.indexOf("Acabamento", start),
                            rawSlice.indexOf("Informações", start),
                            (idxDir > start) ? idxDir : -1 // Não invadir a Direita (raro, mas possível)
                        ].filter(x => x > start);
                        
                        let end = stops.length > 0 ? Math.min(...stops) : rawSlice.length;
                        esqFinal = rawSlice.substring(start, end).trim();
                    }
                }

                // --- LIMPEZA FINAL ---
                dirFinal = cleanText(dirFinal);
                esqFinal = cleanText(esqFinal);

                resultadoUnificado = `${dirFinal} (Direita) | ${esqFinal} (Esquerda)`;

                // Extrai Materiais (Busca na fatia inteira)
                const matExtMatch = rawSlice.match(/Acb\.\s?Externo:\s?([A-Z0-9\s]+\s?\([A-Z0-9]+\))/i);
                const matIntMatch = rawSlice.match(/Acb\.\s?TT\/PNL:\s?([A-Z]+)/i);

                fullData.push({
                    "ORFA": orfaNumber,
                    "Expositor": exp.name,
                    "Laterais": resultadoUnificado,
                    "Material Externo": matExtMatch ? matExtMatch[1].trim() : "PRETO (P1902)",
                    "Material Interno": matIntMatch ? matIntMatch[1].trim() : "INOX"
                });
            }
        }

        const uniqueData = fullData.filter((v, i, a) => 
            a.findIndex(t => t.Expositor === v.Expositor && t["Laterais"] === v["Laterais"]) === i
        );

        if (uniqueData.length === 0) {
            alert("Nenhum expositor encontrado.");
        } else {
            exportToExcel(uniqueData);
        }
    };
    reader.readAsArrayBuffer(file);
});

// --- FUNÇÕES AUXILIARES ---

function cleanGarbage(text) {
    // Remove lixo comum que vaza da coluna "Div. Interna"
    const stopWords = ["DIVISORIAS", "DIVISORIA", "PRATELEIRAS", "PRATELEIRA", "GANCHEIRA", "GANCHEIRAS", "SUPORTE", "ILUMINADAS", "VIDRO"];
    
    // Se encontrar a palavra proibida, corta tudo dali pra frente
    let cutIndex = text.length;
    stopWords.forEach(word => {
        const regex = new RegExp(`(^|\\s)${word}`, 'i');
        const match = text.match(regex);
        if (match && match.index < cutIndex) cutIndex = match.index;
    });
    return text.substring(0, cutIndex).trim();
}

function cleanText(text) {
    if (!text || text.length < 3 || text.includes("Lateral")) return "Vazio";
    
    // Remove números soltos "01", "03" no início ou fim
    let limpo = text.replace(/\s0\d(\s|$)/g, " ").replace(/\s\d{2}(\s|$)/g, " ");
    
    // Remove pontuação de sobra
    limpo = limpo.replace(/^[,.\s]+/, "").replace(/[,.\s]+$/, "");
    
    // Aplica a limpeza de lixo (Divisorias etc) também aqui
    limpo = cleanGarbage(limpo);
    
    return limpo.length > 2 ? limpo : "Vazio";
}

function parseByKeywords(rawContent) {
    // Tenta separar usando palavras-chave conhecidas
    const keywords = ["PANORAMICA", "CEGA", "INTERMEDIARIA", "CHANFRADA", "FECHADA", "VAZIO", "Vazio"];
    let matches = [];
    
    keywords.forEach(key => {
        const regex = new RegExp(`(^|\\s)(${key})`, 'gi');
        let m;
        while ((m = regex.exec(rawContent)) !== null) {
            matches.push({ index: m.index + m[1].length, word: m[2].toUpperCase() });
        }
    });

    matches.sort((a, b) => a.index - b.index);
    matches = matches.filter((v, i, a) => i === 0 || v.index !== a[i-1].index); // Remove duplicatas de índice

    // Filtra falsos positivos de nomes compostos (ex: Chanfrada Cega)
    for (let k = 0; k < matches.length - 1; k++) {
        const cur = matches[k];
        const nxt = matches[k+1];
        
        // Se Chanfrada e Cega estão coladas (< 10 chars), é uma coisa só
        if ((cur.word === "CHANFRADA" && nxt.word === "CEGA") || 
            (cur.word === "INTERMEDIARIA" && nxt.word === "CHANFRADA")) {
            if (nxt.index - (cur.index + cur.word.length) < 10) {
                matches.splice(k+1, 1);
                k--;
            }
        }
    }

    let dir = "Vazio";
    let esq = "Vazio";

    if (matches.length >= 2) {
        let split = matches[1].index;
        dir = rawContent.substring(0, split).trim();
        esq = rawContent.substring(split).trim();
    } else if (matches.length === 1) {
        // Se só tem uma, verifica pontuação
        if (/^\s*,/.test(rawContent)) { // Começa com vírgula? Então Dir é vazia
            esq = rawContent.replace(/^\s*,/, "").trim();
        } else {
            dir = rawContent.trim();
        }
    } else {
        // Sem palavras chaves, tenta vírgula simples
        let parts = rawContent.split(",");
        dir = parts[0] ? parts[0].trim() : "Vazio";
        esq = (parts.length > 1) ? parts[1].trim() : "Vazio";
    }

    return { dir, esq };
}

function exportToExcel(data) {
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch: 10}, {wch: 40}, {wch: 90}, {wch: 20}, {wch: 15}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Producao");
    XLSX.writeFile(wb, "Relatorio_Producao.xlsx");
}